
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TestAll OCR 稽核資料庫 Repository
契約：提供所有 Core Service 需要的 CRUD 操作
設計原則：UI/Core 分離、參數化查詢、最小耦合
"""

from __future__ import annotations
import sqlite3
import json
import hashlib
from contextlib import contextmanager
from dataclasses import dataclass, field
from datetime import datetime, date
from pathlib import Path
from typing import Optional, Any, Literal
from typing import Protocol, runtime_checkable


DB_PATH = Path("/home/newuser/TestAll/data/ocr_reviewed.db")


# ============================================================
# Dataclasses (契約資料結構)
# ============================================================

@dataclass(frozen=True)
class CanonicalItem:
    item_id: int
    canonical_name: str
    category: Optional[str]
    first_seen_at: str
    last_seen_at: str
    total_reviews: int
    is_active: int
    notes: Optional[str]


@dataclass(frozen=True)
class ItemAttribute:
    item_name: str
    attr_key: str
    attr_value: str
    attr_type: str
    notes: Optional[str]
    created_at: str
    updated_at: Optional[str]


@dataclass(frozen=True)
class Library:
    lib_id: int
    lib_name: str
    lib_type: str
    is_active: int
    notes: Optional[str]


@dataclass(frozen=True)
class ReviewedItem:
    id: Optional[int]
    review_date: str
    library: str
    item_name: str
    ocr_raw_name: str
    ocr_text: str
    word_path: str
    quantity: float
    unit: str
    # Phase 9: 多品項日結表欄位
    prev_stock: float = 0.0
    outbound_qty: float = 0.0
    inbound_qty: float = 0.0
    closing_qty: float = 0.0
    unit_price: Optional[str] = None
    loss_qty: float = 0.0
    source_image_path: str
    source_image_hash: str
    confidence: Optional[float]
    page_count: int
    reviewer: Optional[str]
    reviewed_at: str
    is_verified: int
    notes: Optional[str]
    created_at: str
    updated_at: Optional[str]


@dataclass(frozen=True)
class DailyInventory:
    id: Optional[int]
    snapshot_date: str
    library: str
    item_name: str
    opening_qty: float
    inbound_qty: float
    outbound_qty: float
    adjustment: float
    closing_qty: float
    expected_qty: float
    actual_qty: Optional[float]
    loss_qty: float
    loss_pct: float
    loss_status: str
    source_ids: Optional[str]
    notes: Optional[str]
    created_at: str


@dataclass(frozen=True)
class InventoryAudit:
    audit_id: Optional[int]
    audit_date: str
    library: str
    item_name: str
    prev_closing: Optional[float]
    today_opening: float
    today_inbound: float
    today_outbound: float
    today_adjust: float
    expected_closing: float
    actual_closing: float
    diff: float
    diff_pct: float
    item_loss_pct: Optional[float]
    check_tolerance: Optional[float]
    is_abnormal: int
    alert_message: Optional[str]
    inbound_source_ids: Optional[str]
    outbound_source_ids: Optional[str]
    resolved: int
    resolved_at: Optional[str]
    notes: Optional[str]
    created_at: str


@dataclass(frozen=True)
class AlertRule:
    rule_id: Optional[int]
    rule_name: str
    rule_type: str
    description: Optional[str]
    severity: str
    check_sql: Optional[str]
    threshold_field: Optional[str]
    threshold_operator: Optional[str]
    threshold_value: Optional[str]
    cross_tables: Optional[str]
    cross_condition: Optional[str]
    date_field: Optional[str]
    days_before: Optional[int]
    check_if_has_stock: int
    normal_loss_pct: Optional[float]
    is_enabled: int
    cooldown_min: int
    target_user: Optional[str]
    created_by: Optional[str]
    created_at: str
    updated_at: Optional[str]


@dataclass(frozen=True)
class AlertHistory:
    history_id: Optional[int]
    rule_id: Optional[int]
    alert_type: str
    triggered_at: str
    severity: str
    message: str
    context: Optional[str]
    dismissed: int
    dismissed_at: Optional[str]
    dismissed_by: Optional[str]
    created_at: str


# ============================================================
# Connection Manager
# ============================================================

@contextmanager
def get_conn(db_path: Path = DB_PATH):
    """取得 SQLite 連線，自動 commit/rollback/close"""
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def init_db(db_path: Path = DB_PATH) -> None:
    """初始化資料庫：執行 schema.sql"""
    # 若 db_path.parent 下無 schema.sql，fallback 到專案 data/ 目錄
    schema_path = db_path.parent / "schema.sql"
    if not schema_path.exists():
        schema_path = Path(__file__).resolve().parent.parent.parent / "data" / "schema.sql"
    if not schema_path.exists():
        raise FileNotFoundError(f"schema.sql not found: {schema_path}")
    
    with get_conn(db_path) as conn:
        sql = schema_path.read_text(encoding='utf-8')
        conn.executescript(sql)
    
    # Phase 9 migration: 多品項欄位 (向後相容，舊 DB 自動補欄)
    _migrate_multi_item_columns(db_path)
    
    # 確保 schema 版本
    with get_conn(db_path) as conn:
        ver = conn.execute("PRAGMA user_version").fetchone()[0]
        if ver == 0:
            conn.execute("PRAGMA user_version = 1")


def _migrate_multi_item_columns(db_path: Path) -> None:
    """對 ocr_reviewed_items 補上 Phase 9 多品項欄位（若尚不存在）。"""
    new_cols = [
        ("prev_stock",   "REAL NOT NULL DEFAULT 0"),
        ("outbound_qty", "REAL NOT NULL DEFAULT 0"),
        ("inbound_qty",  "REAL NOT NULL DEFAULT 0"),
        ("closing_qty",  "REAL NOT NULL DEFAULT 0"),
        ("unit_price",   "TEXT"),
        ("loss_qty",     "REAL NOT NULL DEFAULT 0"),
    ]
    with get_conn(db_path) as conn:
        existing = {r[1] for r in conn.execute("PRAGMA table_info(ocr_reviewed_items)")}
        for col, ddl in new_cols:
            if col not in existing:
                conn.execute(f"ALTER TABLE ocr_reviewed_items ADD COLUMN {col} {ddl}")



# ============================================================
# Repository 介面 (Protocol)
# ============================================================

@runtime_checkable
class OCRReviewRepository(Protocol):
    # --- Canonical Items ---
    def upsert_canonical_item(self, name: str, category: Optional[str] = None) -> CanonicalItem: ...
    def get_canonical_item(self, name: str) -> Optional[CanonicalItem]: ...
    def get_all_canonical_items(self, active_only: bool = True) -> list[CanonicalItem]: ...
    def update_canonical_item_last_seen(self, name: str) -> None: ...
    
    # --- Item Name Mapping ---
    def add_name_mapping(self, ocr_name: str, canonical_name: str, mapped_by: str = 'user') -> None: ...
    def get_canonical_by_ocr(self, ocr_name: str) -> Optional[str]: ...
    def get_all_mappings_for_canonical(self, canonical_name: str) -> list[str]: ...
    
    # --- Item Attributes ---
    def upsert_item_attr(self, item_name: str, key: str, value: str, 
                         attr_type: str = 'text', notes: Optional[str] = None) -> None: ...
    def get_item_attr(self, item_name: str, key: str) -> Optional[str]: ...
    def get_item_attrs(self, item_name: str) -> list[ItemAttribute]: ...
    
    # --- Libraries ---
    def get_library(self, name: str) -> Optional[Library]: ...
    def get_all_libraries(self, active_only: bool = True) -> list[Library]: ...
    def ensure_library(self, name: str, lib_type: str = 'storage') -> Library: ...
    
    # --- Reviewed Items (主稽核) ---
    def insert_reviewed_item(self, item: ReviewedItem) -> int: ...
    def get_reviewed_item(self, review_date: str, library: str, item_name: str) -> Optional[ReviewedItem]: ...
    def get_reviewed_items_by_date_lib(self, review_date: str, library: str) -> list[ReviewedItem]: ...
    def get_reviewed_items_by_item(self, item_name: str, limit: int = 100) -> list[ReviewedItem]: ...
    def get_by_image_hash(self, image_hash: str) -> list[ReviewedItem]: ...
    def update_reviewed_item(self, item_id: int, **fields) -> None: ...
    
    # --- Daily Inventory ---
    def upsert_daily_inventory(self, inv: DailyInventory) -> int: ...
    def get_daily_inventory(self, date: str, library: str, item_name: str) -> Optional[DailyInventory]: ...
    def get_daily_inventory_by_date(self, date: str, library: str) -> list[DailyInventory]: ...
    def get_prev_daily_inventory(self, date: str, library: str, item_name: str) -> Optional[DailyInventory]: ...
    def get_actual_closing(self, date: str, library: str, item_name: str) -> Optional[float]: ...
    
    # --- Inventory Audit ---
    def insert_inventory_audit(self, audit: InventoryAudit) -> int: ...
    def get_inventory_audit(self, date: str, library: str, item_name: str) -> Optional[InventoryAudit]: ...
    def get_pending_audits(self, limit: int = 100) -> list[InventoryAudit]: ...
    
    # --- Alert Rules ---
    def insert_alert_rule(self, rule: AlertRule) -> int: ...
    def get_enabled_alert_rules(self) -> list[AlertRule]: ...
    def get_alert_rule(self, rule_id: int) -> Optional[AlertRule]: ...
    def update_alert_rule(self, rule_id: int, **fields) -> None: ...
    
    # --- Alert History ---
    def insert_alert_history(self, alert: AlertHistory) -> int: ...
    def get_recent_alerts(self, minutes: int = 60, alert_type: Optional[str] = None) -> list[AlertHistory]: ...
    def get_alert_history(self, limit: int = 100, dismissed: Optional[bool] = None) -> list[AlertHistory]: ...
    def dismiss_alert(self, history_id: int, dismissed_by: str) -> None: ...
    
    # --- System Config ---
    def get_config(self, key: str) -> Optional[str]: ...
    def set_config(self, key: str, value: str, config_type: str = 'text', desc: Optional[str] = None) -> None: ...
    def get_all_config(self) -> dict[str, str]: ...
    
    # --- Utility ---
    def execute(self, sql: str, params: tuple = ()) -> list[sqlite3.Row]: ...
    def execute_one(self, sql: str, params: tuple = ()) -> Optional[sqlite3.Row]: ...
    
    # --- Aggregation ---
    def get_current_stock_summary(self, item_name: str) -> list[dict]: ...
    def get_active_items_with_attr(self, attr_key: str) -> list[str]: ...
    def get_total_stock(self, item_name: str) -> float: ...


# ============================================================
# SQLite 實作
# ============================================================

class SQLiteReviewRepository:
    """SQLite 實作，符合 OCRReviewRepository 協定"""
    
    def __init__(self, db_path: Path = DB_PATH):
        self.db_path = db_path
    
    # --- Internal helpers ---
    def _row_to_canonical(self, row: sqlite3.Row) -> CanonicalItem:
        return CanonicalItem(
            item_id=row['item_id'],
            canonical_name=row['canonical_name'],
            category=row['category'],
            first_seen_at=row['first_seen_at'],
            last_seen_at=row['last_seen_at'],
            total_reviews=row['total_reviews'],
            is_active=row['is_active'],
            notes=row['notes']
        )
    
    def _row_to_library(self, row: sqlite3.Row) -> Library:
        return Library(
            lib_id=row['lib_id'],
            lib_name=row['lib_name'],
            lib_type=row['lib_type'],
            is_active=row['is_active'],
            notes=row['notes']
        )
    
    def _row_to_reviewed(self, row: sqlite3.Row) -> ReviewedItem:
        return ReviewedItem(
            id=row['id'],
            review_date=row['review_date'],
            library=row['library'],
            item_name=row['item_name'],
            ocr_raw_name=row['ocr_raw_name'],
            ocr_text=row['ocr_text'],
            word_path=row['word_path'],
            quantity=row['quantity'],
            unit=row['unit'],
            source_image_path=row['source_image_path'],
            source_image_hash=row['source_image_hash'],
            confidence=row['confidence'],
            page_count=row['page_count'],
            reviewer=row['reviewer'],
            reviewed_at=row['reviewed_at'],
            is_verified=row['is_verified'],
            notes=row['notes'],
            created_at=row['created_at'],
            updated_at=row['updated_at']
        )
    
    def _row_to_daily_inv(self, row: sqlite3.Row) -> DailyInventory:
        return DailyInventory(
            id=row['id'],
            snapshot_date=row['snapshot_date'],
            library=row['library'],
            item_name=row['item_name'],
            opening_qty=row['opening_qty'],
            inbound_qty=row['inbound_qty'],
            outbound_qty=row['outbound_qty'],
            adjustment=row['adjustment'],
            closing_qty=row['closing_qty'],
            expected_qty=row['expected_qty'],
            actual_qty=row['actual_qty'],
            loss_qty=row['loss_qty'],
            loss_pct=row['loss_pct'],
            loss_status=row['loss_status'],
            source_ids=row['source_ids'],
            notes=row['notes'],
            created_at=row['created_at']
        )
    
    def _row_to_audit(self, row: sqlite3.Row) -> InventoryAudit:
        return InventoryAudit(
            audit_id=row['audit_id'],
            audit_date=row['audit_date'],
            library=row['library'],
            item_name=row['item_name'],
            prev_closing=row['prev_closing'],
            today_opening=row['today_opening'],
            today_inbound=row['today_inbound'],
            today_outbound=row['today_outbound'],
            today_adjust=row['today_adjust'],
            expected_closing=row['expected_closing'],
            actual_closing=row['actual_closing'],
            diff=row['diff'],
            diff_pct=row['diff_pct'],
            item_loss_pct=row['item_loss_pct'],
            check_tolerance=row['check_tolerance'],
            is_abnormal=row['is_abnormal'],
            alert_message=row['alert_message'],
            inbound_source_ids=row['inbound_source_ids'],
            outbound_source_ids=row['outbound_source_ids'],
            resolved=row['resolved'],
            resolved_at=row['resolved_at'],
            notes=row['notes'],
            created_at=row['created_at']
        )
    
    def _row_to_alert_rule(self, row: sqlite3.Row) -> AlertRule:
        return AlertRule(
            rule_id=row['rule_id'],
            rule_name=row['rule_name'],
            rule_type=row['rule_type'],
            description=row['description'],
            severity=row['severity'],
            check_sql=row['check_sql'],
            threshold_field=row['threshold_field'],
            threshold_operator=row['threshold_operator'],
            threshold_value=row['threshold_value'],
            cross_tables=row['cross_tables'],
            cross_condition=row['cross_condition'],
            date_field=row['date_field'],
            days_before=row['days_before'],
            check_if_has_stock=row['check_if_has_stock'],
            normal_loss_pct=row['normal_loss_pct'],
            is_enabled=row['is_enabled'],
            cooldown_min=row['cooldown_min'],
            target_user=row['target_user'],
            created_by=row['created_by'],
            created_at=row['created_at'],
            updated_at=row['updated_at']
        )
    
    def _row_to_alert_history(self, row: sqlite3.Row) -> AlertHistory:
        return AlertHistory(
            history_id=row['history_id'],
            rule_id=row['rule_id'],
            alert_type=row['alert_type'],
            triggered_at=row['triggered_at'],
            severity=row['severity'],
            message=row['message'],
            context=row['context'],
            dismissed=row['dismissed'],
            dismissed_at=row['dismissed_at'],
            dismissed_by=row['dismissed_by'],
            created_at=row['created_at']
        )
    
    def _row_to_item_attr(self, row: sqlite3.Row) -> ItemAttribute:
        return ItemAttribute(
            item_name=row['item_name'],
            attr_key=row['attr_key'],
            attr_value=row['attr_value'],
            attr_type=row['attr_type'],
            notes=row['notes'],
            created_at=row['created_at'],
            updated_at=row['updated_at']
        )
    
    # --- Canonical Items ---
    def upsert_canonical_item(self, name: str, category: Optional[str] = None) -> CanonicalItem:
        now = datetime.now().isoformat()
        with get_conn(self.db_path) as conn:
            # Insert or update: keep existing category/notes if not null; always update last_seen_at and increment total_reviews
            conn.execute("""
                INSERT INTO canonical_items (canonical_name, category, first_seen_at, last_seen_at, total_reviews)
                VALUES (?, ?, ?, ?, 1)
                ON CONFLICT(canonical_name) DO UPDATE SET
                    category = COALESCE(canonical_items.category, excluded.category),
                    notes = COALESCE(canonical_items.notes, excluded.notes),
                    last_seen_at = excluded.last_seen_at,
                    total_reviews = canonical_items.total_reviews + 1
            """, (name, category, now, now))
            
            row = conn.execute(
                "SELECT * FROM canonical_items WHERE canonical_name = ?", (name,)
            ).fetchone()
            return self._row_to_canonical(row)
    def get_canonical_item(self, name: str) -> Optional[CanonicalItem]:
        with get_conn(self.db_path) as conn:
            row = conn.execute(
                "SELECT * FROM canonical_items WHERE canonical_name = ?", (name,)
            ).fetchone()
            return self._row_to_canonical(row) if row else None
    
    def get_all_canonical_items(self, active_only: bool = True) -> list[CanonicalItem]:
        with get_conn(self.db_path) as conn:
            sql = "SELECT * FROM canonical_items"
            if active_only:
                sql += " WHERE is_active = 1"
            sql += " ORDER BY canonical_name"
            return [self._row_to_canonical(r) for r in conn.execute(sql).fetchall()]
    
    def update_canonical_item_last_seen(self, name: str) -> None:
        now = datetime.now().isoformat()
        with get_conn(self.db_path) as conn:
            conn.execute("""
                UPDATE canonical_items SET last_seen_at = ?, total_reviews = total_reviews + 1
                WHERE canonical_name = ?
            """, (now, name))
    
    # --- Item Name Mapping ---
    def add_name_mapping(self, ocr_name: str, canonical_name: str, mapped_by: str = 'user') -> None:
        with get_conn(self.db_path) as conn:
            conn.execute("""
                INSERT INTO item_name_mapping (ocr_name, canonical_name, mapped_by)
                VALUES (?, ?, ?)
                ON CONFLICT(ocr_name, canonical_name) DO NOTHING
            """, (ocr_name, canonical_name, mapped_by))
    
    def get_canonical_by_ocr(self, ocr_name: str) -> Optional[str]:
        with get_conn(self.db_path) as conn:
            row = conn.execute("""
                SELECT canonical_name FROM item_name_mapping 
                WHERE ocr_name = ? ORDER BY created_at DESC LIMIT 1
            """, (ocr_name,)).fetchone()
            return row['canonical_name'] if row else None
    
    def get_all_mappings_for_canonical(self, canonical_name: str) -> list[str]:
        with get_conn(self.db_path) as conn:
            return [r['ocr_name'] for r in conn.execute("""
                SELECT ocr_name FROM item_name_mapping WHERE canonical_name = ?
            """, (canonical_name,)).fetchall()]
    
    # --- Item Attributes ---
    def upsert_item_attr(self, item_name: str, key: str, value: str,
                         attr_type: str = 'text', notes: Optional[str] = None) -> None:
        now = datetime.now().isoformat()
        with get_conn(self.db_path) as conn:
            conn.execute("""
                INSERT INTO item_attributes (item_name, attr_key, attr_value, attr_type, notes, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(item_name, attr_key) DO UPDATE SET
                    attr_value = excluded.attr_value,
                    attr_type = excluded.attr_type,
                    notes = COALESCE(excluded.notes, item_attributes.notes),
                    updated_at = excluded.updated_at
            """, (item_name, key, value, attr_type, notes, now, now))
    
    def get_item_attr(self, item_name: str, key: str) -> Optional[str]:
        with get_conn(self.db_path) as conn:
            row = conn.execute("""
                SELECT attr_value FROM item_attributes WHERE item_name = ? AND attr_key = ?
            """, (item_name, key)).fetchone()
            return row['attr_value'] if row else None
    
    def get_item_attrs(self, item_name: str) -> list[ItemAttribute]:
        with get_conn(self.db_path) as conn:
            return [self._row_to_item_attr(r) for r in conn.execute("""
                SELECT * FROM item_attributes WHERE item_name = ?
            """, (item_name,)).fetchall()]
    
    # --- Libraries ---
    def get_library(self, name: str) -> Optional[Library]:
        with get_conn(self.db_path) as conn:
            row = conn.execute("SELECT * FROM libraries WHERE lib_name = ?", (name,)).fetchone()
            return self._row_to_library(row) if row else None
    
    def get_all_libraries(self, active_only: bool = True) -> list[Library]:
        with get_conn(self.db_path) as conn:
            sql = "SELECT * FROM libraries"
            if active_only:
                sql += " WHERE is_active = 1"
            sql += " ORDER BY lib_name"
            return [self._row_to_library(r) for r in conn.execute(sql).fetchall()]
    
    def ensure_library(self, name: str, lib_type: str = 'storage') -> Library:
        with get_conn(self.db_path) as conn:
            row = conn.execute("SELECT * FROM libraries WHERE lib_name = ?", (name,)).fetchone()
            if row:
                return self._row_to_library(row)
            conn.execute("INSERT INTO libraries (lib_name, lib_type) VALUES (?, ?)", (name, lib_type))
            row = conn.execute("SELECT * FROM libraries WHERE lib_name = ?", (name,)).fetchone()
            return self._row_to_library(row)
    
    # --- Reviewed Items ---
    def insert_reviewed_item(self, item: ReviewedItem) -> int:
        with get_conn(self.db_path) as conn:
            cur = conn.execute("""
                INSERT INTO ocr_reviewed_items (
                    review_date, library, item_name, ocr_raw_name, ocr_text,
                    word_path, quantity, unit, source_image_path, source_image_hash,
                    confidence, page_count, reviewer, reviewed_at, is_verified, notes
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(review_date, library, item_name) DO UPDATE SET
                    ocr_raw_name = excluded.ocr_raw_name,
                    ocr_text = excluded.ocr_text,
                    word_path = excluded.word_path,
                    quantity = excluded.quantity,
                    unit = excluded.unit,
                    source_image_path = excluded.source_image_path,
                    source_image_hash = excluded.source_image_hash,
                    confidence = excluded.confidence,
                    page_count = excluded.page_count,
                    reviewer = excluded.reviewer,
                    reviewed_at = excluded.reviewed_at,
                    is_verified = excluded.is_verified,
                    notes = excluded.notes,
                    updated_at = datetime('now')
            """, (
                item.review_date, item.library, item.item_name, item.ocr_raw_name,
                item.ocr_text, item.word_path, item.quantity, item.unit,
                item.source_image_path, item.source_image_hash,
                item.confidence, item.page_count, item.reviewer,
                item.reviewed_at, item.is_verified, item.notes
            ))
            return cur.lastrowid
    
    def get_reviewed_item(self, review_date: str, library: str, item_name: str) -> Optional[ReviewedItem]:
        with get_conn(self.db_path) as conn:
            row = conn.execute("""
                SELECT * FROM ocr_reviewed_items 
                WHERE review_date = ? AND library = ? AND item_name = ?
            """, (review_date, library, item_name)).fetchone()
            return self._row_to_reviewed(row) if row else None
    
    def get_reviewed_items_by_date_lib(self, review_date: str, library: str) -> list[ReviewedItem]:
        with get_conn(self.db_path) as conn:
            return [self._row_to_reviewed(r) for r in conn.execute("""
                SELECT * FROM ocr_reviewed_items 
                WHERE review_date = ? AND library = ? ORDER BY item_name
            """, (review_date, library)).fetchall()]
    
    def get_reviewed_items_by_item(self, item_name: str, limit: int = 100) -> list[ReviewedItem]:
        with get_conn(self.db_path) as conn:
            return [self._row_to_reviewed(r) for r in conn.execute("""
                SELECT * FROM ocr_reviewed_items 
                WHERE item_name = ? ORDER BY review_date DESC LIMIT ?
            """, (item_name, limit)).fetchall()]
    
    def get_by_image_hash(self, image_hash: str) -> list[ReviewedItem]:
        """取得所有相同 hash 的審核記錄（可能多筆）"""
        with get_conn(self.db_path) as conn:
            return [self._row_to_reviewed(r) for r in conn.execute("""
                SELECT * FROM ocr_reviewed_items 
                WHERE source_image_hash = ? ORDER BY reviewed_at DESC
            """, (image_hash,)).fetchall()]
    
    def update_reviewed_item(self, item_id: int, **fields) -> None:
        if not fields:
            return
        set_clause = ", ".join(f"{k} = ?" for k in fields)
        params = list(fields.values()) + [item_id]
        with get_conn(self.db_path) as conn:
            conn.execute(f"""
                UPDATE ocr_reviewed_items SET {set_clause}, updated_at = datetime('now')
                WHERE id = ?
            """, params)
    
    # --- Daily Inventory ---
    def upsert_daily_inventory(self, inv: DailyInventory) -> int:
        with get_conn(self.db_path) as conn:
            cur = conn.execute("""
                INSERT INTO daily_inventory (
                    snapshot_date, library, item_name, opening_qty, inbound_qty,
                    outbound_qty, adjustment, closing_qty, expected_qty, actual_qty,
                    loss_qty, loss_pct, loss_status, source_ids, notes
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(snapshot_date, library, item_name) DO UPDATE SET
                    opening_qty = excluded.opening_qty,
                    inbound_qty = excluded.inbound_qty,
                    outbound_qty = excluded.outbound_qty,
                    adjustment = excluded.adjustment,
                    closing_qty = excluded.closing_qty,
                    expected_qty = excluded.expected_qty,
                    actual_qty = excluded.actual_qty,
                    loss_qty = excluded.loss_qty,
                    loss_pct = excluded.loss_pct,
                    loss_status = excluded.loss_status,
                    source_ids = excluded.source_ids,
                    notes = excluded.notes
            """, (
                inv.snapshot_date, inv.library, inv.item_name,
                inv.opening_qty, inv.inbound_qty, inv.outbound_qty,
                inv.adjustment, inv.closing_qty, inv.expected_qty,
                inv.actual_qty, inv.loss_qty, inv.loss_pct, inv.loss_status,
                inv.source_ids, inv.notes
            ))
            return cur.lastrowid
    
    def get_daily_inventory(self, date: str, library: str, item_name: str) -> Optional[DailyInventory]:
        with get_conn(self.db_path) as conn:
            row = conn.execute("""
                SELECT * FROM daily_inventory 
                WHERE snapshot_date = ? AND library = ? AND item_name = ?
            """, (date, library, item_name)).fetchone()
            return self._row_to_daily_inv(row) if row else None
    
    def get_daily_inventory_by_date(self, date: str, library: str) -> list[DailyInventory]:
        with get_conn(self.db_path) as conn:
            return [self._row_to_daily_inv(r) for r in conn.execute("""
                SELECT * FROM daily_inventory WHERE snapshot_date = ? AND library = ?
            """, (date, library)).fetchall()]
    
    def get_prev_daily_inventory(self, date: str, library: str, item_name: str) -> Optional[DailyInventory]:
        """取得指定日期之前最近一筆庫存快照"""
        with get_conn(self.db_path) as conn:
            row = conn.execute("""
                SELECT * FROM daily_inventory 
                WHERE library = ? AND item_name = ? AND snapshot_date < ?
                ORDER BY snapshot_date DESC LIMIT 1
            """, (library, item_name, date)).fetchone()
            return self._row_to_daily_inv(row) if row else None
    
    def get_actual_closing(self, date: str, library: str, item_name: str) -> Optional[float]:
        with get_conn(self.db_path) as conn:
            row = conn.execute("""
                SELECT actual_qty FROM daily_inventory 
                WHERE snapshot_date = ? AND library = ? AND item_name = ?
            """, (date, library, item_name)).fetchone()
            return row['actual_qty'] if row else None
    
    # --- Inventory Audit ---
    def insert_inventory_audit(self, audit: InventoryAudit) -> int:
        with get_conn(self.db_path) as conn:
            cur = conn.execute("""
                INSERT INTO inventory_audit (
                    audit_date, library, item_name,
                    prev_closing, today_opening, today_inbound, today_outbound, today_adjust,
                    expected_closing, actual_closing, diff, diff_pct,
                    item_loss_pct, check_tolerance, is_abnormal, alert_message,
                    inbound_source_ids, outbound_source_ids, notes
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                audit.audit_date, audit.library, audit.item_name,
                audit.prev_closing, audit.today_opening, audit.today_inbound,
                audit.today_outbound, audit.today_adjust, audit.expected_closing,
                audit.actual_closing, audit.diff, audit.diff_pct,
                audit.item_loss_pct, audit.check_tolerance, audit.is_abnormal,
                audit.alert_message, audit.inbound_source_ids, audit.outbound_source_ids,
                audit.notes
            ))
            return cur.lastrowid
    
    def get_inventory_audit(self, date: str, library: str, item_name: str) -> Optional[InventoryAudit]:
        with get_conn(self.db_path) as conn:
            row = conn.execute("""
                SELECT * FROM inventory_audit WHERE audit_date = ? AND library = ? AND item_name = ?
            """, (date, library, item_name)).fetchone()
            return self._row_to_audit(row) if row else None
    
    def get_pending_audits(self, limit: int = 100) -> list[InventoryAudit]:
        with get_conn(self.db_path) as conn:
            return [self._row_to_audit(r) for r in conn.execute("""
                SELECT * FROM inventory_audit WHERE resolved = 0 ORDER BY audit_date DESC LIMIT ?
            """, (limit,)).fetchall()]
    
    # --- Alert Rules ---
    def insert_alert_rule(self, rule: AlertRule) -> int:
        with get_conn(self.db_path) as conn:
            cur = conn.execute("""
                INSERT INTO alert_rules (
                    rule_name, rule_type, description, severity,
                    check_sql, threshold_field, threshold_operator, threshold_value,
                    cross_tables, cross_condition,
                    date_field, days_before, check_if_has_stock, normal_loss_pct,
                    is_enabled, cooldown_min, target_user, created_by
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                rule.rule_name, rule.rule_type, rule.description, rule.severity,
                rule.check_sql, rule.threshold_field, rule.threshold_operator, rule.threshold_value,
                rule.cross_tables, rule.cross_condition,
                rule.date_field, rule.days_before, rule.check_if_has_stock, rule.normal_loss_pct,
                rule.is_enabled, rule.cooldown_min, rule.target_user, rule.created_by
            ))
            return cur.lastrowid
    
    def get_enabled_alert_rules(self) -> list[AlertRule]:
        with get_conn(self.db_path) as conn:
            return [self._row_to_alert_rule(r) for r in conn.execute("""
                SELECT * FROM alert_rules WHERE is_enabled = 1 ORDER BY rule_id
            """).fetchall()]
    
    def get_alert_rule(self, rule_id: int) -> Optional[AlertRule]:
        with get_conn(self.db_path) as conn:
            row = conn.execute("SELECT * FROM alert_rules WHERE rule_id = ?", (rule_id,)).fetchone()
            return self._row_to_alert_rule(row) if row else None
    
    def update_alert_rule(self, rule_id: int, **fields) -> None:
        if not fields:
            return
        set_clause = ", ".join(f"{k} = ?" for k in fields)
        params = list(fields.values()) + [rule_id]
        with get_conn(self.db_path) as conn:
            conn.execute(f"UPDATE alert_rules SET {set_clause}, updated_at = datetime('now') WHERE rule_id = ?", params)
    
    # --- Alert History ---
    def insert_alert_history(self, alert: AlertHistory) -> int:
        with get_conn(self.db_path) as conn:
            cur = conn.execute("""
                INSERT INTO alert_history (
                    rule_id, alert_type, triggered_at, severity, message, context
                ) VALUES (?, ?, ?, ?, ?, ?)
            """, (
                alert.rule_id, alert.alert_type, alert.triggered_at,
                alert.severity, alert.message, alert.context
            ))
            return cur.lastrowid
    
    def get_recent_alerts(self, minutes: int = 60, alert_type: Optional[str] = None) -> list[AlertHistory]:
        with get_conn(self.db_path) as conn:
            sql = """
                SELECT * FROM alert_history 
                WHERE triggered_at >= datetime('now', ?)
            """
            params = [f'-{minutes} minutes']
            if alert_type:
                sql += " AND alert_type = ?"
                params.append(alert_type)
            sql += " ORDER BY triggered_at DESC"
            return [self._row_to_alert_history(r) for r in conn.execute(sql, params).fetchall()]
    
    def get_alert_history(self, limit: int = 100, dismissed: Optional[bool] = None) -> list[AlertHistory]:
        with get_conn(self.db_path) as conn:
            sql = "SELECT * FROM alert_history"
            params = []
            if dismissed is not None:
                sql += " WHERE dismissed = ?"
                params.append(1 if dismissed else 0)
            sql += " ORDER BY triggered_at DESC LIMIT ?"
            params.append(limit)
            return [self._row_to_alert_history(r) for r in conn.execute(sql, params).fetchall()]
    
    def dismiss_alert(self, history_id: int, dismissed_by: str) -> None:
        with get_conn(self.db_path) as conn:
            conn.execute("""
                UPDATE alert_history SET dismissed = 1, dismissed_at = datetime('now'), dismissed_by = ?
                WHERE history_id = ?
            """, (dismissed_by, history_id))
    
    # --- System Config ---
    def get_config(self, key: str) -> Optional[str]:
        with get_conn(self.db_path) as conn:
            row = conn.execute("SELECT config_value FROM system_config WHERE config_key = ?", (key,)).fetchone()
            return row['config_value'] if row else None
    
    def set_config(self, key: str, value: str, config_type: str = 'text', desc: Optional[str] = None) -> None:
        with get_conn(self.db_path) as conn:
            conn.execute("""
                INSERT INTO system_config (config_key, config_value, config_type, description, updated_at)
                VALUES (?, ?, ?, ?, datetime('now'))
                ON CONFLICT(config_key) DO UPDATE SET
                    config_value = excluded.config_value,
                    config_type = excluded.config_type,
                    description = COALESCE(excluded.description, system_config.description),
                    updated_at = datetime('now')
            """, (key, value, config_type, desc))
    
    def get_all_config(self) -> dict[str, str]:
        with get_conn(self.db_path) as conn:
            return {r['config_key']: r['config_value'] for r in conn.execute("SELECT config_key, config_value FROM system_config").fetchall()}
    
    # --- Utility ---
    def execute(self, sql: str, params: tuple = ()) -> list[sqlite3.Row]:
        with get_conn(self.db_path) as conn:
            return conn.execute(sql, params).fetchall()
    
    def execute_one(self, sql: str, params: tuple = ()) -> Optional[sqlite3.Row]:
        with get_conn(self.db_path) as conn:
            return conn.execute(sql, params).fetchone()
    
    # --- Aggregation ---
    def get_current_stock_summary(self, item_name: str) -> list[dict]:
        """取得品項各庫別目前庫存摘要"""
        with get_conn(self.db_path) as conn:
            rows = conn.execute("""
                SELECT library, closing_qty, snapshot_date
                FROM daily_inventory
                WHERE item_name = ?
                ORDER BY snapshot_date DESC
            """, (item_name,)).fetchall()
            
            # 取每個庫別最新的一筆
            seen = set()
            result = []
            for r in rows:
                if r['library'] not in seen:
                    seen.add(r['library'])
                    result.append({
                        'library': r['library'],
                        'closing_qty': r['closing_qty'],
                        'date': r['snapshot_date']
                    })
            return result
    
    def get_active_items_with_attr(self, attr_key: str) -> list[str]:
        """取得有指定屬性且 is_active=1 的品項名稱"""
        with get_conn(self.db_path) as conn:
            return [r['item_name'] for r in conn.execute("""
                SELECT DISTINCT ia.item_name FROM item_attributes ia
                JOIN canonical_items ci ON ci.canonical_name = ia.item_name
                WHERE ia.attr_key = ? AND ci.is_active = 1
            """, (attr_key,)).fetchall()]
    
    def get_total_stock(self, item_name: str) -> float:
        """取得品項總庫存量（所有庫別 closing_qty 相加）"""
        summary = self.get_current_stock_summary(item_name)
        return sum(s['closing_qty'] for s in summary)


# ============================================================
# 便利函數
# ============================================================

def compute_image_hash(image_path: Path) -> str:
    """計算檔案 SHA256"""
    h = hashlib.sha256()
    with open(image_path, 'rb') as f:
        for chunk in iter(lambda: f.read(8192), b''):
            h.update(chunk)
    return h.hexdigest()


# 取得 repository 實例（每次呼叫回傳新實例，開銷極低 — SQLite 連線在每次 with get_conn 內即開即關）
def get_repo(db_path: Path = DB_PATH) -> SQLiteReviewRepository:
    return SQLiteReviewRepository(db_path)


# ============================================================
# Self-test
# ============================================================

if __name__ == '__main__':
    # 快速自測
    repo = get_repo()
    
    # 先初始化資料庫（建立 schema）
    init_db()
    
    # 1. 測試連線與 schema
    repo.execute("SELECT 1")
    print("✅ DB 連線正常")
    
    # 2. 測試 canonical_items
    item = repo.upsert_canonical_item("測試蘋果", "水果")
    print(f"✅ canonical_items: {item.canonical_name} (id={item.item_id})")
    
    # 3. 測試 attributes
    repo.upsert_item_attr("測試蘋果", "shelf_life_days", "7", "number")
    repo.upsert_item_attr("測試蘋果", "normal_loss_pct", "3.5", "number")
    val = repo.get_item_attr("測試蘋果", "shelf_life_days")
    print(f"✅ item_attributes: shelf_life_days = {val}")
    
    # 4. 測試 libraries
    lib = repo.ensure_library("A倉", "storage")
    print(f"✅ libraries: {lib.lib_name} (type={lib.lib_type})")
    
    # 5. 測試 reviewed_items
    test_item = ReviewedItem(
        id=None,
        review_date="2026-07-22",
        library="A倉",
        item_name="測試蘋果",
        ocr_raw_name="測試蘋果",
        ocr_text="# 測試\n|數量|單位|\n|10|公斤|",
        word_path="/tmp/test.docx",
        quantity=10.0,
        unit="公斤",
        source_image_path="/tmp/test.jpg",
        source_image_hash="abc123",
        confidence=0.95,
        page_count=1,
        reviewer="tester",
        reviewed_at=datetime.now().isoformat(),
        is_verified=1,
        notes="",
        created_at="",
        updated_at=None
    )
    item_id = repo.insert_reviewed_item(test_item)
    print(f"✅ ocr_reviewed_items: inserted id={item_id}")
    
    # 6. 測試 daily_inventory
    inv = DailyInventory(
        id=None,
        snapshot_date="2026-07-22",
        library="A倉",
        item_name="測試蘋果",
        opening_qty=100.0,
        inbound_qty=10.0,
        outbound_qty=5.0,
        adjustment=0.0,
        closing_qty=105.0,
        expected_qty=105.0,
        actual_qty=104.0,
        loss_qty=-1.0,
        loss_pct=0.95,
        loss_status="normal",
        source_ids=json.dumps([item_id]),
        notes="",
        created_at=""
    )
    repo.upsert_daily_inventory(inv)
    print("✅ daily_inventory: upserted")
    
    # 7. 測試 alert
    repo.insert_alert_history(AlertHistory(
        history_id=None, rule_id=0, alert_type='test',
        triggered_at=datetime.now().isoformat(), severity='warning',
        message='測試提醒', context='{}',
        dismissed=0, dismissed_at=None, dismissed_by=None, created_at=''
    ))
    print("✅ alert_history: inserted")
    
    # 8. 測試 config
    repo.set_config('test_key', 'test_value', 'text', '測試用')
    cfg = repo.get_config('test_key')
    print(f"✅ system_config: {cfg}")
    
    print("\n🎉 所有自測通過！")
