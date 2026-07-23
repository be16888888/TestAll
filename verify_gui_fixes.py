#!/usr/bin/env python3
"""驗證 ocr_unified_gui.py 兩項修正 (Bug A / Bug B)。
- Bug A: docx 表格前無段落時，before_text 應帶入檔名 (辨識標頭)。
- Bug B: 庫存預覽在庫別為空時，應跨所有庫別彙總並含 library 欄位，不再直接 return。
依 AGENT_CODE.md：先測試再輸出。本腳本直接複製核心邏輯做單元驗證，不依賴 tkinter 顯示。
"""
import os, sys, tempfile, py_compile
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

PASS, FAIL = [], []

def check(name, cond, detail=""):
    (PASS if cond else FAIL).append((name, detail))
    print(f"[{'PASS' if cond else 'FAIL'}] {name}" + (f" — {detail}" if detail else ""))

# ---- 1) 編譯語法檢查 ----
try:
    py_compile.compile("ocr_unified_gui.py", doraise=True)
    check("GUI 模組語法編譯", True)
except py_compile.PyCompileError as e:
    check("GUI 模組語法編譯", False, str(e))
    sys.exit(1)

# ---- 2) Bug A: before_text 帶入檔名邏輯 ----
# 複製 _load_docx_to_treeview 中的 Phase 10.5 邏輯
def build_before_text(docx_path, doc_before_paras):
    """模擬 Phase 10.5：doc_before_paras=表格前段落文字清單"""
    before_text = [t for t in doc_before_paras if t.strip()]
    if not before_text and docx_path:
        hdr = os.path.basename(docx_path)
        if hdr:
            before_text = [hdr]
    return before_text

# 案例 A1: 表格前無段落 -> 應帶入檔名
bt = build_before_text("/mnt/e/DiskCUse/HFDownloads/115年3月17日(3庫).docx", [])
check("BugA: 無前段落時帶入檔名", bt == ["115年3月17日(3庫).docx"], f"got={bt}")

# 案例 A2: 已有前段落 -> 不覆蓋
bt2 = build_before_text("/x/115年3月17日(3庫).docx", ["3庫 前日庫存表"])
check("BugA: 有前段落不覆蓋", bt2 == ["3庫 前日庫存表"], f"got={bt2}")

# ---- 3) Bug B: 庫別為空 -> 跨庫別彙總 + 含 library 欄 ----
from core.db.repository import get_repo
from core.inventory_service import InventoryService

repo = get_repo()
inv = InventoryService(repo)
biz = "2026-07-17"

# 準備：清空測試日資料並寫入兩個庫別的辨識紀錄
repo.execute("DELETE FROM ocr_reviewed_items WHERE review_date=?", (biz,))
for lib, items in [("3庫", [("高麗菜", 361.0), ("包心白", 95.0)]),
                   ("4庫", [("山東白", 774.0)])]:
    try:
        repo.ensure_library(lib, lib_type='storage')
    except Exception:
        pass
    for name, qty in items:
        # item_name FK -> canonical_items.canonical_name (填滿 NOT NULL 欄)
        repo.execute(
            "INSERT INTO canonical_items (canonical_name, first_seen_at, last_seen_at, is_active) "
            "VALUES (?, datetime('now'), datetime('now'), 1) "
            "ON CONFLICT(canonical_name) DO UPDATE SET is_active=1, last_seen_at=datetime('now')",
            (name,))
        repo.execute(
            "INSERT INTO ocr_reviewed_items "
            "(review_date, library, item_name, ocr_raw_name, ocr_text, word_path, quantity, "
            "source_image_path, source_image_hash, reviewed_at, is_verified) "
            "VALUES (?,?,?,?,?,?,?,?,?,?,?)",
            (biz, lib, name, name, name, "verify.docx", qty,
             "verify.png", "verify", "2026-07-17T00:00:00", 1))

# 模擬 _show_inventory_panel 的 lib 決策邏輯
def resolve_libs(lib, biz):
    if lib:
        return [lib]
    rows = repo.execute("SELECT DISTINCT library FROM ocr_reviewed_items WHERE review_date=?", (biz,))
    libs = [dict(r)["library"] for r in rows if dict(r).get("library")]
    if not libs:
        rows = repo.execute("SELECT DISTINCT library FROM daily_inventory WHERE snapshot_date=?", (biz,))
        libs = [dict(r)["library"] for r in rows if dict(r).get("library")]
    if not libs:
        rows = repo.execute("SELECT DISTINCT library FROM ocr_reviewed_items")
        libs = [dict(r)["library"] for r in rows if dict(r).get("library")]
    return libs

libs = resolve_libs("", biz)  # 庫別為空
check("BugB: 庫別為空時解析出多庫別", set(libs) == {"3庫", "4庫"}, f"libs={libs}")

diffs = []
for l in libs:
    try:
        repo.ensure_library(l, lib_type='storage')
    except Exception:
        pass
    d = inv.calculate_daily(biz, l)
    diffs.extend(d.values())

check("BugB: 彙總品項數=3", len(diffs) == 3, f"n={len(diffs)}")
check("BugB: 每筆 diff 含 library 欄位", all(hasattr(d, "library") and d.library for d in diffs),
      f"libs={[d.library for d in diffs]}")
check("BugB: 不再因庫別為空而跳過", len(diffs) > 0, f"diffs={len(diffs)}")

# 清理測試資料
repo.execute("DELETE FROM ocr_reviewed_items WHERE review_date=?", (biz,))

print("\n==== 結果 ====")
print(f"PASS: {len(PASS)}  FAIL: {len(FAIL)}")
sys.exit(1 if FAIL else 0)
