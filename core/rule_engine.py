#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TestAll Rule Engine (Core Layer)
合約：評估所有啟用規則並回傳 Alert 列表
原則：內建規則 + 使用者自訂規則、冷卻期機制、可擴充
"""

from __future__ import annotations
from dataclasses import dataclass, field
from datetime import datetime, date, timedelta
from typing import Optional, Callable
import json

from core.db.repository import (
    OCRReviewRepository, AlertHistory, AlertRule,
    ReviewedItem, DailyInventory,
    get_repo, DB_PATH,
)
from core.ocr_review_service import Alert
from core.inventory_service import InventoryService


# ============================================================
# Data structures
# ============================================================

@dataclass
class RuleEvalResult:
    """規則評估結果"""
    rule_name: str
    rule_type: str
    severity: str
    triggered: bool
    alerts: list[Alert] = field(default_factory=list)


# ============================================================
# Rule Engine
# ============================================================

class RuleEngine:
    """統一規則引擎：入庫時觸發評估"""

    def __init__(self, repo: OCRReviewRepository | None = None):
        self.repo = repo or get_repo()
        self.inventory_service = InventoryService(self.repo)

    # ----------------------------------------------------------
    # 整體評估入口
    # ----------------------------------------------------------

    def evaluate_on_save(self, new_item: ReviewedItem, recalc_inventory: bool = False) -> list[Alert]:
        """
        入庫後執行所有規則。
        若 recalc_inventory=True，先觸發庫存計算再評估。
        """
        all_alerts: list[Alert] = []

        # === 內建規則 ===
        all_alerts.extend(self._builtin_cross_library(new_item))
        all_alerts.extend(self._builtin_image_duplicate(new_item))
        all_alerts.extend(self._builtin_expiry_check(new_item))
        all_alerts.extend(self._builtin_inventory_diff_check(new_item))
        all_alerts.extend(self._builtin_consistency_check(new_item))

        # === 使用者自訂規則 ===
        user_rules = self.repo.get_enabled_alert_rules()
        for rule in user_rules:
            all_alerts.extend(self._run_user_rule(rule, new_item))

        # 冷卻去重
        all_alerts = self._apply_cooldown(all_alerts)

        # 記錄 alert_history
        now = datetime.now().isoformat()
        for alert in all_alerts:
            self.repo.insert_alert_history(AlertHistory(
                history_id=None,
                rule_id=alert.rule_id,
                alert_type=alert.alert_type or "rule",
                triggered_at=now,
                severity=alert.severity,
                message=alert.message,
                context=json.dumps(alert.context, ensure_ascii=False),
                dismissed=0,
                dismissed_at=None,
                dismissed_by=None,
                created_at="",
            ))

        return all_alerts

    # ----------------------------------------------------------
    # 內建規則
    # ----------------------------------------------------------

    def _builtin_cross_library(self, item: ReviewedItem) -> list[Alert]:
        """同品項跨庫檢測"""
        rows = self.repo.execute(
            "SELECT DISTINCT library, review_date FROM ocr_reviewed_items "
            "WHERE item_name = ? AND library != ?",
            (item.item_name, item.library),
        )
        alerts = []
        for r in rows:
            lib = dict(r)["library"]
            d = dict(r)["review_date"]
            alerts.append(Alert(
                rule_name="同品項跨庫",
                severity="warning",
                alert_type="cross_lib",
                message=f"⚠️「{item.item_name}」已於 {d} 存於「{lib}」庫，現在新增至「{item.library}」",
                context={"item": item.item_name, "existing_lib": lib,
                         "existing_date": d, "current_lib": item.library},
            ))
        return alerts

    def _builtin_image_duplicate(self, item: ReviewedItem) -> list[Alert]:
        """重複圖片檢查（已在 OCRReviewService 拒絕入庫，此為備用提醒）"""
        # OCRReviewService 的 save_reviewed_item 已拒絕重複圖片
        # 此處回傳空列表（圖片重複完全拒絕，非 alert）
        return []

    def _builtin_expiry_check(self, item: ReviewedItem) -> list[Alert]:
        """效期檢查"""
        expiry_str = self.repo.get_item_attr(item.item_name, "expiry_date")
        if not expiry_str:
            return []

        try:
            expiry = date.fromisoformat(expiry_str)
            days_left = (expiry - date.today()).days
        except (ValueError, TypeError):
            return []

        if days_left <= 2:
            sev = "critical" if days_left <= 0 else "warning"
            msg = (
                f"⏰「{item.item_name}」已到期！"
                if days_left <= 0
                else f"⏰「{item.item_name}」將於 {days_left} 天後到期（{expiry_str}）"
            )
            return [Alert(
                rule_name="效期到期提醒",
                severity=sev,
                alert_type="expiry",
                message=msg,
                context={"item": item.item_name, "expiry": expiry_str, "days_left": days_left},
            )]
        return []

    def _builtin_inventory_diff_check(self, item: ReviewedItem) -> list[Alert]:
        """庫存異常損耗檢查"""
        diffs = self.inventory_service.calculate_daily(item.review_date, item.library)
        alerts = []
        for name, diff in diffs.items():
            if diff.is_abnormal:
                alerts.append(Alert(
                    rule_name="庫存異常損耗",
                    severity="critical",
                    alert_type="inventory_diff",
                    message=f"🔴「{name}」異常損耗 {diff.loss_pct:.1f}%（容許上限 {diff.normal_loss_pct}%）",
                    context={
                        "item": name, "loss_pct": diff.loss_pct,
                        "normal_loss_pct": diff.normal_loss_pct,
                        "loss_qty": diff.loss_qty, "expected_qty": diff.expected_qty,
                    },
                ))
        return alerts

    def _builtin_consistency_check(self, item: ReviewedItem) -> list[Alert]:
        """今日/前日庫存一致性檢查"""
        mismatches = self.inventory_service.check_consistency(item.review_date, item.library)
        alerts = []
        for m in mismatches:
            if not m.is_match:
                alerts.append(Alert(
                    rule_name="庫存不符",
                    severity="critical",
                    alert_type="stock_mismatch",
                    message=m.message,
                    context={
                        "item": m.item_name, "prev_closing": m.prev_closing,
                        "today_opening": m.today_opening, "mismatch": m.mismatch,
                        "tolerance": m.tolerance,
                    },
                ))
        return alerts

    # ----------------------------------------------------------
    # 使用者自訂規則
    # ----------------------------------------------------------

    def _run_user_rule(self, rule: AlertRule, item: ReviewedItem) -> list[Alert]:
        """根據 rule_type 分發至對應處理器"""
        rule_type = rule.rule_type

        handlers: dict[str, Callable] = {
            "query": self._handle_query_rule,
            "threshold": self._handle_threshold_rule,
            "cross_reference": self._handle_cross_reference_rule,
            "date_proximity": self._handle_date_proximity_rule,
            "inventory_diff": self._handle_inventory_diff_rule,
        }

        handler = handlers.get(rule_type)
        if not handler:
            return []
        return handler(rule, item)

    def _handle_query_rule(self, rule: AlertRule, item: ReviewedItem) -> list[Alert]:
        """執行自訂 SQL 查詢規則：若回傳任何列 → 觸發"""
        check_sql = rule.check_sql or ""
        if not check_sql:
            return []

        try:
            rows = self.repo.execute(check_sql)
            if rows:
                return [Alert(
                    rule_name=rule.rule_name,
                    severity=rule.severity or "warning",
                    alert_type="user_query",
                    rule_id=rule.rule_id,
                    message=f"查詢規則觸發：{rule.rule_name}（{len(rows)} 筆）",
                    context={"rule_id": rule.rule_id, "rows": len(rows)},
                )]
        except Exception:
            # 若查詢失敗（如 SQL 語法錯誤），不中斷
            pass
        return []

    def _handle_threshold_rule(self, rule: AlertRule, item: ReviewedItem) -> list[Alert]:
        """閾值規則：檢查欄位值 vs 閾值"""
        field = rule.threshold_field or ""
        op = rule.threshold_operator or "gt"
        threshold = rule.threshold_value or "0"

        if not field:
            return []

        # 從 item 或 item_attributes 取得目前值
        actual = getattr(item, field, None)
        if actual is None:
            actual = self.repo.get_item_attr(item.item_name, field)

        if actual is None:
            return []

        # 判斷
        try:
            actual_val = float(actual)
            threshold_val = float(threshold)
        except (ValueError, TypeError):
            return []

        triggered = False
        if op == "gt" and actual_val > threshold_val:
            triggered = True
        elif op == "lt" and actual_val < threshold_val:
            triggered = True
        elif op == "ge" and actual_val >= threshold_val:
            triggered = True
        elif op == "le" and actual_val <= threshold_val:
            triggered = True
        elif op == "eq" and actual_val == threshold_val:
            triggered = True

        if triggered:
            return [Alert(
                rule_name=rule.rule_name,
                severity=rule.severity or "warning",
                alert_type="threshold",
                rule_id=rule.rule_id,
                message=f"「{item.item_name}」{field}={actual_val} {op} {threshold}",
                context={"rule_id": rule.rule_id, "field": field, "value": actual_val},
            )]
        return []

    def _handle_cross_reference_rule(self, rule: AlertRule, item: ReviewedItem) -> list[Alert]:
        """交叉檢查規則：檢查品項是否出現在其他自訂表中"""
        tables = rule.cross_tables or ""
        condition = rule.cross_condition or ""

        if not tables or not condition:
            return []

        # 驗證 tables 與 condition 僅含合法的 SQL 識別符號/運算子
        import re
        if not re.match(r'^[a-zA-Z_][a-zA-Z0-9_ ,]*$', tables):
            return []
        if not re.match(r'^[a-zA-Z_. =<>!?0-9\'%"()\s,]*$', condition):
            return []

        try:
            sql = f"SELECT COUNT(*) FROM {tables} WHERE {condition}"
            rows = self.repo.execute(sql, (item.item_name,))
            count = dict(rows[0]).get("COUNT(*)", 0) if rows else 0
            if count > 0:
                return [Alert(
                    rule_name=rule.rule_name,
                    severity=rule.severity or "warning",
                    alert_type="cross_ref",
                    rule_id=rule.rule_id,
                    message=f"交叉關聯觸發：{rule.rule_name}（{count} 筆）",
                    context={"rule_id": rule.rule_id, "count": count},
                )]
        except Exception:
            pass
        return []

    def _handle_date_proximity_rule(self, rule: AlertRule, item: ReviewedItem) -> list[Alert]:
        """日期鄰近規則：檢查品項的 expiry_date 或其他日期"""
        date_field = rule.date_field or "expiry_date"
        days_before = rule.days_before or 2
        check_stock = rule.check_if_has_stock

        expiry_str = self.repo.get_item_attr(item.item_name, date_field)
        if not expiry_str:
            return []

        try:
            expiry = date.fromisoformat(expiry_str)
            days_left = (expiry - date.today()).days
        except (ValueError, TypeError):
            return []

        if days_left <= days_before:
            if check_stock:
                stock = self.repo.get_total_stock(item.item_name)
                if stock <= 0:
                    return []

            sev = "critical" if days_left <= 0 else (rule.severity or "warning")
            return [Alert(
                rule_name=rule.rule_name,
                severity=sev,
                alert_type="date_proximity",
                rule_id=rule.rule_id,
                message=f"⏰「{item.item_name}」{days_left} 天後到期（{expiry}），庫存 {stock if check_stock else '?'}",
                context={"rule_id": rule.rule_id, "days_left": days_left, "date": expiry_str},
            )]
        return []

    def _handle_inventory_diff_rule(self, rule: AlertRule, item: ReviewedItem) -> list[Alert]:
        """庫存差異規則：可自訂 damage_loss_pct 閾值"""
        pct = rule.normal_loss_pct
        if pct is None:
            return []

        diffs = self.inventory_service.calculate_daily(item.review_date, item.library)
        alerts = []
        for name, diff in diffs.items():
            if diff.loss_pct > pct:
                alerts.append(Alert(
                    rule_name=rule.rule_name,
                    severity=rule.severity or "critical",
                    alert_type="inventory_diff",
                    rule_id=rule.rule_id,
                    message=f"🔴「{name}」自訂損耗上限 {pct}% 觸發：實際 {diff.loss_pct:.1f}%",
                    context={"rule_id": rule.rule_id, "loss_pct": diff.loss_pct, "threshold": pct},
                ))
        return alerts

    # ----------------------------------------------------------
    # 冷卻機制
    # ----------------------------------------------------------

    def _apply_cooldown(self, alerts: list[Alert], minutes: int = 60) -> list[Alert]:
        """過濾冷卻期內已觸發的相同 alert"""
        if not alerts:
            return []

        recent = self.repo.get_recent_alerts(minutes=minutes)
        # 建立 (alert_type, context) 的指紋集合
        seen = set()
        for rh in recent:
            key = (rh.alert_type, rh.message[:80])  # 用訊息前 80 字元做去重
            seen.add(key)

        filtered = []
        for a in alerts:
            key = (a.alert_type, a.message[:80])
            if key not in seen:
                filtered.append(a)
        return filtered

    # ----------------------------------------------------------
    # 效期掃描（跨庫存檢查 — 供外部定時任務調用）
    # ----------------------------------------------------------

    def scan_expiry_alerts(self, days_before: int | None = None) -> list[Alert]:
        """全局效期掃描：所有品項的到期日檢查"""
        if days_before is None:
            days_before = int(self.repo.get_config("expiry_warn_days") or "2")

        result = self.inventory_service.check_expiry(days_before)
        alerts = []
        for d in result:
            alerts.append(Alert(
                rule_name="效期到期提醒",
                severity=d["severity"],
                alert_type="expiry",
                message=d["message"],
                context={
                    "item": d["item_name"], "expiry": d["expiry_date"],
                    "days_left": d["days_left"], "stock": d["total_stock"],
                },
            ))
        return self._apply_cooldown(alerts)


# ============================================================
# Convenience
# ============================================================

def get_rule_engine(repo: OCRReviewRepository | None = None) -> RuleEngine:
    return RuleEngine(repo)


# ============================================================
# Self-test
# ============================================================

if __name__ == "__main__":
    import sys, tempfile, hashlib
    from pathlib import Path
    test_db = Path(tempfile.gettempdir()) / 'test_rule_engine.db'
    if test_db.exists(): test_db.unlink()
    sys.path.insert(0, str(Path(__file__).parent.parent))
    from core.db.repository import init_db, SQLiteReviewRepository
    from core.ocr_review_service import OCRReviewService
    init_db(test_db)
    repo = SQLiteReviewRepository(test_db)
    engine = RuleEngine(repo)
    ocr = OCRReviewService(repo)

    # 準備測試資料
    repo.upsert_canonical_item("蘋果", "水果")
    repo.upsert_item_attr("蘋果", "shelf_life_days", "14", "number")
    repo.upsert_item_attr("蘋果", "normal_loss_pct", "5.0", "number")
    tomorrow = (date.today() + timedelta(days=1)).isoformat()
    repo.upsert_item_attr("蘋果", "expiry_date", tomorrow, "date")

    # 入庫 inbound
    hash1 = hashlib.sha256(b'inbound_test').hexdigest()
    r1 = ocr.save_reviewed_item(
        review_date="2026-07-22", library="inbound", item_name="蘋果",
        ocr_raw_name="蘋果", ocr_text="# test", word_path="/tmp/t1.docx",
        quantity=100, source_image_path="/tmp/t1.jpg", source_image_hash=hash1,
    )
    assert r1.ok, f"save failed: {r1.errors}"
    item = repo.get_reviewed_item("2026-07-22", "inbound", "蘋果")
    assert item is not None

    # 評估規則
    alerts = engine.evaluate_on_save(item)
    print(f"Builtin rules triggered: {len(alerts)}")
    for a in alerts:
        print(f"  [{a.severity}] {a.rule_name}: {a.message[:60]}...")

    # 入庫到另一個庫 (跨庫)
    hash2 = hashlib.sha256(b'outbound_test').hexdigest()
    r2 = ocr.save_reviewed_item(
        review_date="2026-07-22", library="outbound", item_name="蘋果",
        ocr_raw_name="蘋果", ocr_text="|out|", word_path="/tmp/t2.docx",
        quantity=20, source_image_path="/tmp/t2.jpg", source_image_hash=hash2,
    )
    assert r2.ok
    item2 = repo.get_reviewed_item("2026-07-22", "outbound", "蘋果")
    assert item2 is not None
    alerts2 = engine.evaluate_on_save(item2)
    print(f"\nCross-library alerts: {len(alerts2)}")
    for a in alerts2:
        print(f"  [{a.severity}] {a.rule_name}: {a.message[:80]}")

    # 冷卻測試
    alerts3 = engine.evaluate_on_save(item2)
    print(f"\nCooldown: {len(alerts3)} (should be 0 after same item re-eval)")
    assert len(alerts3) == 0, "Cooldown should suppress duplicate alerts"

    print("\n✅ rule_engine OK")
    test_db.unlink(missing_ok=True)