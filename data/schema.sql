
-- ============================================================
-- TestAll OCR 稽核資料庫 Schema v1.0
-- 原則: UI/Core 分離、契約優先、最小耦合
-- ============================================================

PRAGMA journal_mode = WAL;
PRAGMA synchronous = NORMAL;
PRAGMA cache_size = -64000;
PRAGMA foreign_keys = ON;

-- ============================================================
-- 1. 品項主表 (canonical items)
-- ============================================================
CREATE TABLE IF NOT EXISTS canonical_items (
    item_id        INTEGER PRIMARY KEY AUTOINCREMENT,
    canonical_name TEXT    NOT NULL UNIQUE,
    category       TEXT,                        -- nullable
    first_seen_at  TEXT    NOT NULL,
    last_seen_at   TEXT    NOT NULL,
    total_reviews  INTEGER NOT NULL DEFAULT 0,
    is_active      INTEGER NOT NULL DEFAULT 1,
    notes          TEXT                         -- nullable
);

-- ============================================================
-- 2. OCR → 標準名稱對照表
-- ============================================================
CREATE TABLE IF NOT EXISTS item_name_mapping (
    ocr_name       TEXT NOT NULL,
    canonical_name TEXT NOT NULL,
    mapped_by      TEXT NOT NULL DEFAULT 'user',  -- 'user' | 'auto'
    created_at     TEXT NOT NULL DEFAULT (datetime('now')),
    PRIMARY KEY (ocr_name, canonical_name),
    FOREIGN KEY (canonical_name) REFERENCES canonical_items(canonical_name)
        ON UPDATE CASCADE ON DELETE CASCADE
);

-- ============================================================
-- 3. 品項屬性 (K-V, 供規則引擎用)
--    同品項第一次設定的值會自動記憶 (PK 保證唯一)
--    第二次之後若未手動覆蓋，沿用既有值
-- ============================================================
CREATE TABLE IF NOT EXISTS item_attributes (
    item_name  TEXT NOT NULL,
    attr_key   TEXT NOT NULL,
    attr_value TEXT NOT NULL,
    attr_type  TEXT NOT NULL DEFAULT 'text',
    notes      TEXT,
    created_at TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at TEXT,
    PRIMARY KEY (item_name, attr_key),
    FOREIGN KEY (item_name) REFERENCES canonical_items(canonical_name)
        ON UPDATE CASCADE ON DELETE CASCADE
);

-- ============================================================
-- 4. 庫別定義
-- ============================================================
CREATE TABLE IF NOT EXISTS libraries (
    lib_id    INTEGER PRIMARY KEY AUTOINCREMENT,
    lib_name  TEXT    NOT NULL UNIQUE,
    lib_type  TEXT    NOT NULL DEFAULT 'storage',
    is_active INTEGER NOT NULL DEFAULT 1,
    notes     TEXT
);

-- 預設庫別 (application-level seeds)
INSERT OR IGNORE INTO libraries (lib_name, lib_type) VALUES
    ('inbound',  'inbound'),
    ('outbound', 'outbound'),
    ('A倉',      'storage'),
    ('B倉',      'storage');

-- ============================================================
-- 5. 主稽核表 (OCR 結果 + Word 路徑)
-- ============================================================
CREATE TABLE IF NOT EXISTS ocr_reviewed_items (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,

    -- 業務主鍵 (邏輯唯一)
    review_date  TEXT    NOT NULL,           -- YYYY-MM-DD
    library      TEXT    NOT NULL,
    item_name    TEXT    NOT NULL,           -- canonical_name

    -- OCR 追溯
    ocr_raw_name TEXT    NOT NULL,           -- OCR 原始品名
    ocr_text     TEXT    NOT NULL,           -- OCR Markdown 全文

    -- 產出
    word_path    TEXT    NOT NULL,           -- .docx 路徑

    -- 數量
    quantity     REAL    NOT NULL DEFAULT 0,
    unit         TEXT    NOT NULL DEFAULT '公斤',

    -- 圖片
    source_image_path TEXT NOT NULL,
    source_image_hash TEXT NOT NULL,         -- SHA256 (去重用)

    -- OCR 品質 (nullable)
    confidence   REAL,
    page_count   INTEGER DEFAULT 1,

    -- 審核 (nullable)
    reviewer     TEXT,
    reviewed_at  TEXT    NOT NULL,
    is_verified  INTEGER NOT NULL DEFAULT 0,
    notes        TEXT,

    -- 系統欄位
    created_at   TEXT    NOT NULL DEFAULT (datetime('now')),
    updated_at   TEXT,

    FOREIGN KEY (item_name) REFERENCES canonical_items(canonical_name)
        ON UPDATE CASCADE ON DELETE RESTRICT,
    FOREIGN KEY (library)   REFERENCES libraries(lib_name)
        ON UPDATE CASCADE ON DELETE RESTRICT
);

-- 索引
CREATE UNIQUE INDEX IF NOT EXISTS idx_reviewed_uk
    ON ocr_reviewed_items(review_date, library, item_name);
CREATE INDEX IF NOT EXISTS idx_reviewed_hash
    ON ocr_reviewed_items(source_image_hash);
CREATE INDEX IF NOT EXISTS idx_reviewed_date_lib
    ON ocr_reviewed_items(review_date, library);
CREATE INDEX IF NOT EXISTS idx_reviewed_item
    ON ocr_reviewed_items(item_name, review_date);

-- ============================================================
-- 6. 每日庫存結算表
-- ============================================================
CREATE TABLE IF NOT EXISTS daily_inventory (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    snapshot_date TEXT    NOT NULL,           -- YYYY-MM-DD
    library       TEXT    NOT NULL,
    item_name     TEXT    NOT NULL,

    -- 庫存數量
    opening_qty   REAL    NOT NULL DEFAULT 0,  -- 期初 (前日結存)
    inbound_qty   REAL    NOT NULL DEFAULT 0,  -- 當日進貨
    outbound_qty  REAL    NOT NULL DEFAULT 0,  -- 當日出貨
    adjustment    REAL    NOT NULL DEFAULT 0,  -- 手動調整

    closing_qty   REAL    NOT NULL,            -- 期末 (= opening + inbound - outbound + adjust)

    -- 損耗分析
    expected_qty  REAL    NOT NULL,            -- 理論庫存
    actual_qty    REAL,                        -- 實際盤點 (nullable)
    loss_qty      REAL    NOT NULL DEFAULT 0,
    loss_pct      REAL    NOT NULL DEFAULT 0,
    loss_status   TEXT    NOT NULL DEFAULT 'normal',

    -- 溯源
    source_ids    TEXT,                        -- JSON array of ocr_reviewed_items.id
    notes         TEXT,
    created_at    TEXT    NOT NULL DEFAULT (datetime('now')),

    FOREIGN KEY (item_name) REFERENCES canonical_items(canonical_name)
        ON UPDATE CASCADE ON DELETE RESTRICT,
    FOREIGN KEY (library)   REFERENCES libraries(lib_name)
        ON UPDATE CASCADE ON DELETE RESTRICT
);

CREATE UNIQUE INDEX IF NOT EXISTS idx_daily_inv_uk
    ON daily_inventory(snapshot_date, library, item_name);
CREATE INDEX IF NOT EXISTS idx_daily_inv_loss
    ON daily_inventory(loss_status, snapshot_date);

-- ============================================================
-- 7. 庫存一致性檢查記錄
-- ============================================================
CREATE TABLE IF NOT EXISTS inventory_audit (
    audit_id     INTEGER PRIMARY KEY AUTOINCREMENT,
    audit_date   TEXT    NOT NULL,
    library      TEXT    NOT NULL,
    item_name    TEXT    NOT NULL,

    -- 公式驗證
    prev_closing     REAL,                     -- 前日 closing_qty (nullable)
    today_opening    REAL,                     -- 今日 opening_qty
    today_inbound    REAL,
    today_outbound   REAL,
    today_adjust     REAL,
    expected_closing REAL,
    actual_closing   REAL,
    diff             REAL,
    diff_pct         REAL,

    -- 損耗判斷
    item_loss_pct    REAL,                     -- 該品項的正常損耗%
    check_tolerance  REAL,                     -- prev_closing * loss_pct/100
    is_abnormal      INTEGER NOT NULL DEFAULT 0,
    alert_message    TEXT,

    -- 來源追溯
    inbound_source_ids  TEXT,
    outbound_source_ids TEXT,

    resolved      INTEGER NOT NULL DEFAULT 0,
    resolved_at   TEXT,
    notes         TEXT,
    created_at    TEXT    NOT NULL DEFAULT (datetime('now')),

    FOREIGN KEY (item_name) REFERENCES canonical_items(canonical_name)
        ON UPDATE CASCADE ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_inv_audit_date
    ON inventory_audit(audit_date, library);

-- ============================================================
-- 8. 規則定義表 (使用者自訂)
-- ============================================================
CREATE TABLE IF NOT EXISTS alert_rules (
    rule_id          INTEGER PRIMARY KEY AUTOINCREMENT,
    rule_name        TEXT NOT NULL,
    rule_type        TEXT NOT NULL,
    description      TEXT,
    severity         TEXT NOT NULL DEFAULT 'warning',

    -- query 型
    check_sql        TEXT,

    -- threshold 型
    threshold_field     TEXT,
    threshold_operator  TEXT,
    threshold_value     TEXT,

    -- cross_reference 型
    cross_tables        TEXT,
    cross_condition     TEXT,

    -- date_proximity 型
    date_field          TEXT,
    days_before         INTEGER,
    check_if_has_stock  INTEGER NOT NULL DEFAULT 1,

    -- inventory_diff 型
    normal_loss_pct     REAL,

    is_enabled      INTEGER NOT NULL DEFAULT 1,
    cooldown_min    INTEGER NOT NULL DEFAULT 60,
    target_user     TEXT,
    created_by      TEXT,
    created_at      TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at      TEXT
);

-- ============================================================
-- 9. 提醒記錄
-- ============================================================
CREATE TABLE IF NOT EXISTS alert_history (
    history_id   INTEGER PRIMARY KEY AUTOINCREMENT,
    rule_id      INTEGER,
    alert_type   TEXT    NOT NULL,
    triggered_at TEXT    NOT NULL,
    severity     TEXT    NOT NULL DEFAULT 'warning',
    message      TEXT    NOT NULL,
    context      TEXT,                        -- JSON
    dismissed    INTEGER NOT NULL DEFAULT 0,
    dismissed_at TEXT,
    dismissed_by TEXT,
    created_at   TEXT    NOT NULL DEFAULT (datetime('now'))
);

CREATE INDEX IF NOT EXISTS idx_alert_history_date
    ON alert_history(triggered_at);
CREATE INDEX IF NOT EXISTS idx_alert_history_type
    ON alert_history(alert_type, dismissed);

-- ============================================================
-- 10. 系統配置
-- ============================================================
CREATE TABLE IF NOT EXISTS system_config (
    config_key   TEXT PRIMARY KEY,
    config_value TEXT NOT NULL,
    config_type  TEXT NOT NULL DEFAULT 'text',
    description  TEXT,
    updated_at   TEXT
);

-- 預設配置
INSERT OR IGNORE INTO system_config (config_key, config_value, config_type, description)
VALUES
    ('normal_loss_default_pct',   '5.0',    'number',  '全域預設正常損耗百分比'),
    ('expiry_warn_days',          '2',      'number',  '全域到期預警天數'),
    ('stock_mismatch_mode',       'per_item','text',   '庫存一致模式: per_item(依品項損耗) | strict(完全相符)'),
    ('auto_recalc_inventory',     'true',   'boolean', '入庫後自動重算庫存');

-- ============================================================
-- Schema 版本記錄 (供 migration 使用)
-- ============================================================
PRAGMA user_version = 1;
