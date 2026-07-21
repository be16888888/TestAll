#!/usr/bin/env python3
"""TestAll 庫存計算服務 (Core Layer) — 每日結算、損耗判斷、一致性驗證"""
from __future__ import annotations
from dataclasses import dataclass
from datetime import date, timedelta
from typing import Optional

from core.db.repository import (
    OCRReviewRepository, DailyInventory, InventoryAudit, get_repo, DB_PATH,
)


@dataclass
class InventoryDiff:
    snapshot_date: str; library: str; item_name: str
    opening_qty: float; inbound_qty: float; outbound_qty: float; adjustment: float
    closing_qty: float; expected_qty: float; actual_qty: float
    loss_qty: float; loss_pct: float; normal_loss_pct: float; is_abnormal: bool
    message: str

    @property
    def color(self) -> str:
        return '#ff4444' if self.is_abnormal else '#4488ff'


@dataclass
class ConsistencyCheck:
    item_name: str; library: str; date: str; prev_date: str
    prev_closing: float; today_opening: float; mismatch: float; tolerance: float
    is_match: bool; message: str


class InventoryService:

    def __init__(self, repo: OCRReviewRepository | None = None):
        self.repo = repo or get_repo()

    def calculate_daily(self, review_date: str, library: str,
                            items: list[str] | None = None) -> dict[str, "InventoryDiff"]:
        if items is None:
            canonical = self.repo.get_all_canonical_items(active_only=True)
            items = [c.canonical_name for c in canonical]
        default_loss = float(self.repo.get_config("normal_loss_default_pct") or "5.0")
        results: dict[str, InventoryDaily] = {}

        for name in items:
            prev = self.repo.get_prev_daily_inventory(review_date, library, name)
            opening = prev.closing_qty if prev else 0.0
            inflow = self._qty(review_date, library, name)
            outflow = self._qty(review_date, library, name) if library != "inbound" else 0.0
            # simplification: inflow only for inbound lib, outflow only for outbound
            if library == "inbound":
                inflow = self._sum_ocr_qty(review_date, library, name)
                outflow = 0.0
            elif library == "outbound":
                inflow = 0.0
                outflow = self._sum_ocr_qty(review_date, library, name)
            else:
                inflow = self._sum_ocr_qty(review_date, library, name)
                outflow = 0.0
            adjust = 0.0
            expected = opening + inflow - outflow - adjust
            normal_pct = float(self.repo.get_item_attr(name, "normal_loss_pct") or default_loss)
            actual = self._actual_closing(review_date, library, name) or expected
            loss_qty = actual - expected
            loss_pct = abs(loss_qty) / max(abs(expected), 0.001) * 100 if expected else 0.0
            abnormal = loss_pct > normal_pct
            status = "abnormal" if abnormal else "normal"

            import json
            src = json.dumps([r.id for r in self.repo.get_reviewed_items_by_date_lib(review_date, library)
                              if r.item_name == name and r.id])
            di = DailyInventory(
                id=None, snapshot_date=review_date, library=library, item_name=name,
                opening_qty=opening, inbound_qty=inflow, outbound_qty=outflow,
                adjustment=adjust, closing_qty=actual, expected_qty=expected,
                actual_qty=actual, loss_qty=loss_qty, loss_pct=loss_pct,
                loss_status=status, source_ids=src, notes="", created_at=""
            )
            self.repo.upsert_daily_inventory(di)

            results[name] = InventoryDiff(
                snapshot_date=review_date, library=library, item_name=name,
                opening_qty=opening, inbound_qty=inflow, outbound_qty=outflow,
                adjustment=adjust, closing_qty=actual, expected_qty=expected,
                actual_qty=actual, loss_qty=loss_qty, loss_pct=loss_pct,
                normal_loss_pct=normal_pct, is_abnormal=abnormal,
                message=(f"🔵 正常 損耗 {loss_pct:.1f}%（容許 {normal_pct}%）"
                         if not abnormal else
                         f"🔴 異常 損耗 {loss_pct:.1f}%（容許上限 {normal_pct}%）"),
            )
        return results

    def check_consistency(self, review_date: str, library: str) -> list[ConsistencyCheck]:
        prev_date = self._prev_date(review_date)
        today_map = {i.item_name: i for i in self.repo.get_daily_inventory_by_date(review_date, library)}
        prev_map = {i.item_name: i for i in self.repo.get_daily_inventory_by_date(prev_date, library)}
        mismatches = []
        for name, ti in today_map.items():
            pi = prev_map.get(name)
            if not pi:
                continue
            diff = abs(pi.closing_qty - ti.opening_qty)
            mode = self.repo.get_config("stock_mismatch_mode") or "per_item"
            if mode == "per_item":
                pct = float(self.repo.get_item_attr(name, "normal_loss_pct") or 0)
                tol = pct / 100.0 * abs(pi.closing_qty)
            else:
                tol = 0.0
            ok = diff <= tol
            ck = ConsistencyCheck(
                item_name=name, library=library, date=review_date, prev_date=prev_date,
                prev_closing=pi.closing_qty, today_opening=ti.opening_qty,
                mismatch=diff, tolerance=tol, is_match=ok,
                message=(f"✅ {name}: {pi.closing_qty} ≈ {ti.opening_qty}"
                         if ok else f"⚠️ {name}: {pi.closing_qty} ≠ {ti.opening_qty} (差 {diff:.2f})"),
            )
            if not ok:
                mismatches.append(ck)
                self.repo.insert_inventory_audit(InventoryAudit(
                    audit_id=None, audit_date=review_date, library=library, item_name=name,
                    prev_closing=pi.closing_qty, today_opening=ti.opening_qty,
                    today_inbound=ti.inbound_qty, today_outbound=ti.outbound_qty,
                    today_adjust=ti.adjustment, expected_closing=ti.closing_qty,
                    actual_closing=ti.closing_qty, diff=diff,
                    diff_pct=diff/max(abs(pi.closing_qty), 0.001)*100,
                    item_loss_pct=float(self.repo.get_item_attr(name, "normal_loss_pct") or 0),
                    check_tolerance=tol, is_abnormal=1, alert_message=ck.message,
                    inbound_source_ids=None, outbound_source_ids=None,
                    resolved=0, resolved_at=None, notes="", created_at="",
                ))
        return mismatches

    def check_expiry(self, days_before: int | None = None) -> list[dict]:
        if days_before is None:
            days_before = int(self.repo.get_config("expiry_warn_days") or "2")
        today = date.today()
        target = today + timedelta(days=days_before)
        alerts = []
        for name in self.repo.get_active_items_with_attr("expiry_date"):
            s = self.repo.get_item_attr(name, "expiry_date")
            if not s: continue
            try: exp = date.fromisoformat(s)
            except (ValueError, TypeError): continue
            dl = (exp - today).days
            if dl > days_before:
                continue
            stock = self.repo.get_total_stock(name)
            if stock <= 0:
                continue
            sev = "critical" if dl <= 0 else "warning"
            alerts.append({"item_name": name, "expiry_date": s, "days_left": dl,
                           "total_stock": stock, "severity": sev,
                           "message": (f"⏰「{name}」已到期！庫存 {stock}" if dl <= 0
                                       else f"⏰「{name}」將於 {dl} 天後到期（{s}），庫存 {stock}")})
        return alerts

    def get_abnormal_diffs(self, review_date: str, library: str) -> list[InventoryDaily]:
        return [d for d in self.calculate_daily(review_date, library).values() if d.is_abnormal]

    def _sum_ocr_qty(self, review_date: str, library: str, item_name: str) -> float:
        rs = self.repo.execute(
            "SELECT SUM(quantity) FROM ocr_reviewed_items WHERE review_date = ? AND library=? AND item_name=?",
            (review_date, library, item_name))
        return float(dict(rs[0]).get("SUM(quantity)", 0) or 0) if rs else 0.0

    def _actual_closing(self, review_date: str, library: str, item_name: str) -> float | None:
        inv = self.repo.get_daily_inventory(review_date, library, item_name)
        return inv.actual_qty if inv else None

    def _qty(self, review_date, library, item_name): ...
    def _prev_date(self, d): return (date.fromisoformat(d) - timedelta(days=1)).isoformat() if d else d


def get_inventory_service(repo=None): return InventoryService(repo)


# ============================================================
# Self-test
if __name__ == "__main__":
    import sys, tempfile, hashlib
    from pathlib import Path
    test_db = Path(tempfile.gettempdir()) / 'test_inventory.db'
    if test_db.exists(): test_db.unlink()
    sys.path.insert(0, str(Path(__file__).parent.parent))
    from core.db.repository import init_db, SQLiteReviewRepository, DB_PATH as MAIN_DB
    from core.ocr_review_service import OCRReviewService
    init_db(test_db)
    repo = SQLiteReviewRepository(test_db)
    svc = InventoryService(repo)
    ocr = OCRReviewService(repo)

    repo.upsert_canonical_item("蘋果","水果")
    repo.upsert_item_attr("蘋果","shelf_life_days","14","number")
    repo.upsert_item_attr("蘋果","normal_loss_pct","5.0","number")
    for i,(lib,qty) in enumerate([("inbound",100),("outbound",20)]):
        h = hashlib.sha256(f"img{i}".encode()).hexdigest()
        r = ocr.save_reviewed_item("2026-07-22",lib,"蘋果","蘋果",
            f"|品|量|\n|---:|---:|\n|蘋果|{qty}|\n", f"/tmp/{i}.docx", qty, f"/tmp/{i}.jpg", h)
        assert r.ok, f"save failed: {r.errors}"

    diffs = svc.calculate_daily("2026-07-22","inbound")
    for name,d in sorted(diffs.items()):
        print(f"  {name}: open={d.opening_qty} in={d.inbound_qty} out={d.outbound_qty} "
              f"cls={d.closing_qty} loss={d.loss_pct:.1f}% [{d.message[:20]}...]")

    repo.upsert_daily_inventory(DailyInventory(
        None,"2026-07-21","inbound","蘋果",100,0,0,0,100,100,100,0,0,"normal","","",""))
    ms = svc.check_consistency("2026-07-22","inbound")
    print(f"\nconsistency mismatches: {len(ms)}")
    for m in ms: print(f"  {m.message}")

    tomorrow = (date.today()+timedelta(days=1)).isoformat()
    repo.upsert_item_attr("蘋果","expiry_date",tomorrow,"date")
    repo.upsert_daily_inventory(DailyInventory(
        None, date.today().isoformat(),"inbound","蘋果",80,0,0,0,80,80,80,0,0,"normal","","",""))
    ea = svc.check_expiry(2)
    print(f"\nexpiry alerts: {len(ea)}")
    for a in ea: print(f"  [{a['severity']}] {a['message'][:80]}")
    assert len(ea) >= 1
    print("\n✅ inventory_service OK")
    test_db.unlink(missing_ok=True)