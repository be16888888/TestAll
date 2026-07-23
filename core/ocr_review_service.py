#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TestAll OCR 稽核服務 (Core Layer)
合約：提供「回存WORD檔/入資料庫」完整流程
原則：UI/Core 分離、最小耦合、自我驗證
"""

from __future__ import annotations
from dataclasses import dataclass, field
from datetime import datetime, date, timedelta
from pathlib import Path
from typing import Optional, Protocol, runtime_checkable
import hashlib
import json

from core.db.repository import (
    SQLiteReviewRepository, OCRReviewRepository,
    ReviewedItem, CanonicalItem, DailyInventory, ItemAttribute,
    AlertHistory,
    get_repo, compute_image_hash,
    DB_PATH,
)


# ============================================================
# Data structures
# ============================================================

@dataclass
class ReviewResult:
    """入庫結果"""
    success: bool
    item_id: int = 0
    item_name: str = ""
    word_path: str = ""
    alerts: list[Alert] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
    message: str = ""

    @property
    def ok(self) -> bool:
        return self.success and not self.errors


@dataclass
class Alert:
    """規則觸發提醒"""
    rule_name: str
    severity: str          # 'info' | 'warning' | 'critical'
    message: str
    context: dict = field(default_factory=dict)
    rule_id: int = 0
    alert_type: str = ""


@dataclass
class SimilarItem:
    """模糊比對結果"""
    canonical_name: str
    similarity_score: float   # 0~1, 越高越相似


# ============================================================
# OCR Review Service (核心)
# ============================================================

class OCRReviewService:
    """OCR 稽核服務：處理「回存WORD檔/入資料庫」完整流程"""

    def __init__(self, repo: OCRReviewRepository | None = None):
        self.repo = repo or get_repo()

    # ----------------------------------------------------------
    # 1. 模糊比對 (品項名稱)
    # ----------------------------------------------------------

    def find_similar_items(
        self,
        ocr_name: str,
        min_score: float = 0.4,
        max_results: int = 5,
    ) -> list[SimilarItem]:
        """
        搜尋 canonical_items 中與 ocr_name 相似的品項。
        策略：
          1. 先查 item_name_mapping（自動匹配）
          2. 再用 LIKE + 字元重疊率篩選
          3. 按相似度降序排列
        """
        # 第一步：精確 mapping 命中
        exact = self.repo.get_canonical_by_ocr(ocr_name)
        if exact:
            return [SimilarItem(canonical_name=exact, similarity_score=1.0)]

        results: list[SimilarItem] = []
        candidates = self.repo.get_all_canonical_items(active_only=True)

        for c in candidates:
            score = self._similarity(ocr_name, c.canonical_name)
            if score >= min_score:
                results.append(SimilarItem(
                    canonical_name=c.canonical_name,
                    similarity_score=score,
                ))

        results.sort(key=lambda x: x.similarity_score, reverse=True)
        return results[:max_results]

    def _similarity(self, a: str, b: str) -> float:
        """計算兩個品名之間的相似度 (0~1)"""
        a_lower = a.strip().lower()
        b_lower = b.strip().lower()

        # 完全相等
        if a_lower == b_lower:
            return 1.0

        # 子字串匹配
        if a_lower in b_lower or b_lower in a_lower:
            return 0.85

        # 字元重疊率 (適合中文：忽略順序)
        set_a = set(a_lower)
        set_b = set(b_lower)
        if not set_a or not set_b:
            return 0.0
        overlap = len(set_a & set_b)
        max_len = max(len(set_a), len(set_b))
        return overlap / max_len

    # ----------------------------------------------------------
    # 2. 品項屬性的自動沿用
    # ----------------------------------------------------------

    def inherit_item_attributes(self, item_name: str) -> dict[str, str]:
        """
        查詢品項既有屬性。可讓 UI 判斷是否需要再次詢問使用者。
        回傳 {attr_key: attr_value, ...}
        """
        attrs = self.repo.get_item_attrs(item_name)
        return {a.attr_key: a.attr_value for a in attrs}

    def set_or_inherit_attr(
        self,
        item_name: str,
        attr_key: str,
        attr_value: str | None,
        attr_type: str = "text",
    ) -> str:
        """
        設定或沿用屬性值。
        - 若傳入 attr_value：覆蓋（使用者手動輸入）
        - 若 attr_value 為 None：沿用既有值
        - 若無既有值：回傳空字串
        回傳最終使用的值。
        """
        if attr_value is not None and attr_value.strip():
            self.repo.upsert_item_attr(item_name, attr_key, attr_value, attr_type)
            return attr_value

        existing = self.repo.get_item_attr(item_name, attr_key)
        if existing is not None:
            return existing
        return ""

    # ----------------------------------------------------------
    # 3. 圖片去重檢查
    # ----------------------------------------------------------

    def check_image_duplicate(self, image_hash: str) -> list[ReviewedItem]:
        """
        檢查圖片 hash 是否已存在。
        回傳所有重複記錄（可能 N 筆）。
        若回傳空 list → 可入庫。
        """
        return self.repo.get_by_image_hash(image_hash)

    # ----------------------------------------------------------
    # 4. 完整入庫流程 (save_reviewed_item)
    # ----------------------------------------------------------

    def save_reviewed_item(
        self,
        review_date: str,
        library: str,
        item_name: str,              # canonical_name (已確認)
        ocr_raw_name: str,
        ocr_text: str,
        word_path: str,
        quantity: float,
        source_image_path: str,
        source_image_hash: str,
        confidence: float | None = None,
        page_count: int = 1,
        reviewer: str | None = None,
        is_verified: int = 1,
        notes: str = "",
        unit: str = "公斤",
        # Phase 9: 多品項日結表欄位 (向後相容，預設 0/None)
        prev_stock: float = 0.0,
        outbound_qty: float = 0.0,
        inbound_qty: float = 0.0,
        closing_qty: float = 0.0,
        unit_price: str | None = None,
        loss_qty: float = 0.0,
        skip_dup_check: bool = False,
    ) -> ReviewResult:
        """
        核心流程：
          1. 圖片去重檢查
          2. 更新 canonical_items (last_seen_at, total_reviews)
          3. 寫入 ocr_reviewed_items (UPSERT)
          4. 查詢/沿用品項屬性 (shelf_life_days, normal_loss_pct, expiry_date)
          5. 若有 shelf_life_days 但無 expiry_date → 自動計算 expiry_date
          6. 觸發內建規則 (跨庫、效期)
          7. 記錄 alert_history
          8. 回傳 ReviewResult

        回傳 ReviewResult:
          - success=True 表示已入庫
          - alerts 包含觸發的提醒
          - errors 包含阻擋原因 (如重複圖片)
        """
        result = ReviewResult(success=False, item_name=item_name, word_path=word_path)

        # ---- Step 0: 必須欄位校驗 ----
        if not item_name.strip():
            result.errors.append("品項名稱為空")
            return result
        if not ocr_text.strip():
            result.errors.append("OCR 文字為空")
            return result
        if not source_image_hash:
            result.errors.append("圖片 hash 為空")
            return result

        # ---- Step 1: 圖片去重 ----
        if not skip_dup_check:
            duplicates = self.check_image_duplicate(source_image_hash)
            if duplicates:
                n = len(duplicates)
                dates = sorted({d.reviewed_at[:10] for d in duplicates})
                date_list = "、".join(dates)
                msg = f"此圖片已審核過 {n} 次（日期：{date_list}），拒絕重複入庫"
                result.errors.append(msg)
                return result

        # ---- Step 2: 更新/建立 canonical item ----
        self.repo.upsert_canonical_item(item_name, category=None)  # category 由使用者後續設定

        # ---- Step 3: 寫入 ocr_reviewed_items ----
        now = datetime.now().isoformat()
        item = ReviewedItem(
            id=None,
            review_date=review_date,
            library=library,
            item_name=item_name,
            ocr_raw_name=ocr_raw_name,
            ocr_text=ocr_text,
            word_path=word_path,
            quantity=quantity,
            unit=unit,
            source_image_path=source_image_path,
            source_image_hash=source_image_hash,
            confidence=confidence,
            page_count=page_count,
            reviewer=reviewer,
            reviewed_at=now,
            is_verified=is_verified,
            notes=notes,
            created_at="",
            updated_at=None,
            prev_stock=prev_stock,
            outbound_qty=outbound_qty,
            inbound_qty=inbound_qty,
            closing_qty=closing_qty,
            unit_price=unit_price,
            loss_qty=loss_qty,
        )
        item_id = self.repo.insert_reviewed_item(item)
        result.item_id = item_id

        # ---- Step 4: 品項屬性自動沿用 ----
        # shelf_life_days: 沿用第一次設定的值
        existing_attrs = self.inherit_item_attributes(item_name)
        shelf_life_days = existing_attrs.get("shelf_life_days", "")
        normal_loss_pct = existing_attrs.get("normal_loss_pct", "")

        # ---- Step 5: 若無 expiry_date 但有 shelf_life_days → 自動計算 ----
        existing_expiry = existing_attrs.get("expiry_date", "")
        if shelf_life_days and not existing_expiry:
            try:
                days = int(shelf_life_days)
                calc_expiry = (date.today() + timedelta(days=days)).isoformat()
                self.repo.upsert_item_attr(item_name, "expiry_date", calc_expiry, "date")
                existing_expiry = calc_expiry
            except (ValueError, TypeError):
                pass

        # ---- Step 6: 內建規則檢查 (輕量防禦 — 不寫入 alert_history，由 RuleEngine 統一管理) ----
        alerts: list[Alert] = []

        # 6a. 跨庫檢測（快速回報用；RuleEngine 會做正式檢查並寫入 alert_history）
        cross_lib_alerts = self._check_cross_library(item_name, library)
        alerts.extend(cross_lib_alerts)

        # 6b. 效期檢查（快速回報用；正式邏輯見 RuleEngine.scan_expiry_alerts）
        if existing_expiry:
            expiry_alerts = self._check_expiry(item_name, existing_expiry)
            alerts.extend(expiry_alerts)

        # ---- Step 7: 記錄 alert_history（由 RuleEngine 統一管理；此處保留輕量回報） ----
        # Note: 正式 alert_history 寫入已移至 RuleEngine.evaluate_on_save，
        #       此處的 alerts 僅供 UI 立即顯示，不持久化。

        result.alerts = alerts
        result.success = True
        result.message = f"✅「{item_name}」已入庫（{review_date} / {library}）"
        return result

    def save_reviewed_rows(self, rows: list[dict]) -> list[ReviewResult]:
        """Phase 9: 批次多品項入庫。rows 為 dict 清單，每列欄位：
            review_date, library, item_name, ocr_raw_name, ocr_text, word_path,
            quantity, unit, source_image_path, source_image_hash,
            prev_stock, outbound_qty, inbound_qty, closing_qty, unit_price, loss_qty
            (多品項共用同一 review_date/library/word_path/image，逐列各自 item_name + 數量欄)
        回傳每列的 ReviewResult。
        """
        results: list[ReviewResult] = []
        for idx, r in enumerate(rows):
            # 多品項日結表: 同一張圖多列，首列做圖片去重，其餘列跳過
            # (避免同一來源圖被誤判為重複而擋掉後續品項列)
            skip_dup = idx > 0
            # 補足多品項欄位預設值
            res = self.save_reviewed_item(
                review_date=r["review_date"],
                library=r["library"],
                item_name=r["item_name"],
                ocr_raw_name=r.get("ocr_raw_name", r["item_name"]),
                ocr_text=r.get("ocr_text", ""),
                word_path=r.get("word_path", ""),
                quantity=float(r.get("quantity", 0) or 0),
                source_image_path=r.get("source_image_path", ""),
                source_image_hash=r.get("source_image_hash", "no_image"),
                unit=r.get("unit", "公斤"),
                # Phase 9 多品項欄位
                prev_stock=float(r.get("prev_stock", 0) or 0),
                outbound_qty=float(r.get("outbound_qty", 0) or 0),
                inbound_qty=float(r.get("inbound_qty", 0) or 0),
                closing_qty=float(r.get("closing_qty", 0) or 0),
                unit_price=r.get("unit_price") or None,
                loss_qty=float(r.get("loss_qty", 0) or 0),
                skip_dup_check=skip_dup,
            )
            results.append(res)
        return results

    # ----------------------------------------------------------
    # 5. 內建規則
    # ----------------------------------------------------------

    def _check_cross_library(self, item_name: str, current_library: str) -> list[Alert]:
        """檢查品項是否出現在其他庫別"""
        rows = self.repo.execute(
            """SELECT DISTINCT library, review_date
               FROM ocr_reviewed_items
               WHERE item_name = ? AND library != ?""",
            (item_name, current_library),
        )
        alerts = []
        for r in rows:
            lib = dict(r)["library"]
            date_str = dict(r)["review_date"]
            alerts.append(Alert(
                rule_name="同品項跨庫",
                severity="warning",
                alert_type="cross_lib",
                message=f"⚠️「{item_name}」已於 {date_str} 存於「{lib}」庫，現在新增至「{current_library}」",
                context={"item": item_name, "existing_lib": lib, "existing_date": date_str,
                         "current_lib": current_library},
            ))
        return alerts

    def _check_expiry(self, item_name: str, expiry_date_str: str) -> list[Alert]:
        """檢查效期"""
        try:
            expiry = date.fromisoformat(expiry_date_str)
            days_left = (expiry - date.today()).days
        except (ValueError, TypeError):
            return []

        if days_left <= 2:
            sev = "critical" if days_left <= 0 else "warning"
            msg = (
                f"⏰「{item_name}」已到期！"
                if days_left <= 0
                else f"⏰「{item_name}」將於 {days_left} 天後到期（{expiry}）"
            )
            return [Alert(
                rule_name="效期到期提醒",
                severity=sev,
                alert_type="expiry",
                message=msg,
                context={"item": item_name, "expiry": expiry_date_str, "days_left": days_left},
            )]
        return []

    # ----------------------------------------------------------
    # 6. 查詢輔助
    # ----------------------------------------------------------

    def get_items_by_date_library(self, review_date: str, library: str) -> list[ReviewedItem]:
        """查詢特定日期/庫別的所有入庫記錄"""
        return self.repo.get_reviewed_items_by_date_lib(review_date, library)

    def get_item_history(self, item_name: str, limit: int = 30) -> list[ReviewedItem]:
        """查詢特定品項的入庫歷史"""
        return self.repo.get_reviewed_items_by_item(item_name, limit)

    def get_reviewed_item(self, review_date: str, library: str, item_name: str) -> Optional[ReviewedItem]:
        """查詢單筆記錄"""
        return self.repo.get_reviewed_item(review_date, library, item_name)

    def update_reviewed_item_ocr(self, item_id: int, ocr_text: str,
                                  ocr_raw_name: str | None = None,
                                  quantity: float | None = None,
                                  notes: str | None = None) -> None:
        """更新已入庫記錄的 OCR 內容（編輯後重新存檔用）"""
        fields = {"ocr_text": ocr_text}
        if ocr_raw_name is not None:
            fields["ocr_raw_name"] = ocr_raw_name
        if quantity is not None:
            fields["quantity"] = quantity
        if notes is not None:
            fields["notes"] = notes
        self.repo.update_reviewed_item(item_id, **fields)

    # ----------------------------------------------------------
    # 7. 品項屬性 CRUD (供 GUI 品項管理用)
    # ----------------------------------------------------------

    def set_item_attribute(self, item_name: str, key: str, value: str,
                           attr_type: str = "text", notes: str | None = None) -> None:
        """設定/覆蓋品項屬性"""
        self.repo.upsert_item_attr(item_name, key, value, attr_type, notes)

    def get_item_attributes(self, item_name: str) -> dict[str, str]:
        """取得品項所有屬性"""
        return self.inherit_item_attributes(item_name)

    def get_all_canonical_items(self, active_only: bool = True) -> list[CanonicalItem]:
        """取得所有品項"""
        return self.repo.get_all_canonical_items(active_only)


# ============================================================
# Convenience
# ============================================================

def get_service(repo: OCRReviewRepository | None = None) -> OCRReviewService:
    """取得 service 實例"""
    return OCRReviewService(repo or get_repo())


# ============================================================
# Self-test (run: python3 core/ocr_review_service.py)
# ============================================================

if __name__ == "__main__":
    from core.db.repository import init_db, SQLiteReviewRepository

    # 用獨立 DB 測試
    test_db = DB_PATH.parent / "test_ocr_service.db"
    if test_db.exists():
        test_db.unlink()

    import core.db.repository as repo_mod
    original_db = repo_mod.DB_PATH
    repo_mod.DB_PATH = test_db

    try:
        init_db()
        repo = SQLiteReviewRepository()
        svc = OCRReviewService(repo)

        print("=== Test 1: 模糊比對 ===")
        repo.upsert_canonical_item("蘋果", "水果")
        repo.upsert_canonical_item("西瓜", "水果")
        repo.upsert_canonical_item("火龍果", "水果")

        similars = svc.find_similar_items("頻果")  # OCR 讀錯
        print(f"'頻果' → similars: {[(s.canonical_name, s.similarity_score) for s in similars]}")
        assert similars[0].canonical_name == "蘋果"

        # mapping 自動命中
        repo.add_name_mapping("頻果", "蘋果", "user")
        similars2 = svc.find_similar_items("頻果")
        print(f"'頻果' (有mapping) → similars: {[(s.canonical_name, s.similarity_score) for s in similars2]}")
        assert similars2[0].canonical_name == "蘋果"
        assert similars2[0].similarity_score == 1.0

        print("\n=== Test 2: 屬性自動沿用 ===")
        svc.set_item_attribute("蘋果", "shelf_life_days", "14", "number")
        svc.set_item_attribute("蘋果", "normal_loss_pct", "3.5", "number")
        attrs = svc.get_item_attributes("蘋果")
        print(f"蘋果 attrs: {attrs}")
        assert attrs["shelf_life_days"] == "14"

        # 沿用測試
        val = svc.set_or_inherit_attr("蘋果", "shelf_life_days", None)  # 不傳值 → 沿用
        print(f"shelf_life_days 沿用: {val}")
        assert val == "14"

        print("\n=== Test 3: 完整入庫 (成功) ===")
        hash1 = compute_image_hash(Path("/tmp/test.jpg")) if Path("/tmp/test.jpg").exists() else "testhash007"
        if not Path("/tmp/test.jpg").exists():
            Path("/tmp/test.jpg").write_text("test image content")

        hash1 = compute_image_hash(Path("/tmp/test.jpg"))
        result = svc.save_reviewed_item(
            review_date="2026-07-22",
            library="inbound",
            item_name="蘋果",
            ocr_raw_name="頻果",
            ocr_text="# 蘋果\n|數量|單價|\n|10|20|",
            word_path="/tmp/test_apple.docx",
            quantity=10.0,
            source_image_path="/tmp/test.jpg",
            source_image_hash=hash1,
            confidence=0.95,
            reviewer="tester",
        )
        print(f"Result: ok={result.ok}, id={result.item_id}, alerts={len(result.alerts)}")
        if result.errors:
            print(f"Errors: {result.errors}")
        assert result.ok
        assert result.item_id > 0

        # 確認屬性自動沿用到新款入庫品項
        attrs2 = svc.get_item_attributes("蘋果")
        print(f"入庫後 attrs: {attrs2}")
        # 應該有自動計算的 expiry_date
        if "shelf_life_days" in attrs2:
            print(f"expiry_date: {attrs2.get('expiry_date', '(未設定)')}")

        print("\n=== Test 4: 圖片去重 ===")
        result2 = svc.save_reviewed_item(
            review_date="2026-07-22",
            library="outbound",
            item_name="西瓜",
            ocr_raw_name="西瓜",
            ocr_text="# 西瓜\n|數量|\n|5|",
            word_path="/tmp/test_watermelon.docx",
            quantity=5.0,
            source_image_path="/tmp/test.jpg",  # 同一張圖！
            source_image_hash=hash1,
        )
        print(f"去重結果: ok={result2.ok}, errors={result2.errors}")
        assert not result2.ok
        assert "拒絕重複入庫" in result2.errors[0]

        print("\n=== Test 5: 跨庫檢測 ===")
        # 使用新 hash 入庫西瓜到 inbound
        Path("/tmp/test2.jpg").write_text("test image 2 content")
        hash2 = compute_image_hash(Path("/tmp/test2.jpg"))
        result3 = svc.save_reviewed_item(
            review_date="2026-07-22",
            library="inbound",
            item_name="西瓜",
            ocr_raw_name="西瓜",
            ocr_text="# 西瓜\n|數量|\n|5|",
            word_path="/tmp/test_wm2.docx",
            quantity=5.0,
            source_image_path="/tmp/test2.jpg",
            source_image_hash=hash2,
        )
        print(f"入庫西瓜(inbound): ok={result3.ok}")
        assert result3.ok

        # 現在西瓜在 inbound 有一筆，再入庫到 outbound → 應觸發跨庫 alert
        Path("/tmp/test3.jpg").write_text("test image 3 content")
        hash3 = compute_image_hash(Path("/tmp/test3.jpg"))
        result4 = svc.save_reviewed_item(
            review_date="2026-07-22",
            library="outbound",
            item_name="西瓜",
            ocr_raw_name="西瓜",
            ocr_text="# 西瓜\n|數量|\n|3|",
            word_path="/tmp/test_wm3.docx",
            quantity=3.0,
            source_image_path="/tmp/test3.jpg",
            source_image_hash=hash3,
        )
        print(f"入庫西瓜(outbound): ok={result4.ok}, alerts={len(result4.alerts)}")
        if result4.alerts:
            for a in result4.alerts:
                print(f"  [{a.severity}] {a.message}")
        assert result4.ok
        assert len(result4.alerts) >= 1
        assert any("跨庫" in a.rule_name for a in result4.alerts)

        print("\n=== Test 6: 效期檢查 ===")
        # 設定到期日為明天
        tomorrow = (date.today() + timedelta(days=1)).isoformat()
        svc.set_item_attribute("西瓜", "expiry_date", tomorrow, "date")
        Path("/tmp/test4.jpg").write_text("test image 4 content")
        hash4 = compute_image_hash(Path("/tmp/test4.jpg"))
        result5 = svc.save_reviewed_item(
            review_date="2026-07-22",
            library="inbound",
            item_name="西瓜",
            ocr_raw_name="西瓜",
            ocr_text="# 西瓜\n|數量|\n|2|",
            word_path="/tmp/test_wm4.docx",
            quantity=2.0,
            source_image_path="/tmp/test4.jpg",
            source_image_hash=hash4,
        )
        print(f"效期檢查: ok={result5.ok}, alerts={len(result5.alerts)}")
        for a in result5.alerts:
            print(f"  [{a.severity}] {a.message}")
        assert result5.ok
        assert any("到期" in a.rule_name for a in result5.alerts)

        print("\n🎉 All tests passed!")

    finally:
        repo_mod.DB_PATH = original_db
        if test_db.exists():
            test_db.unlink()