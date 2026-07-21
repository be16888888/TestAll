#!/usr/bin/env python3
"""
統一 OCR.Space / Nanonets / LlamaIndex GUI
- 左側原圖 / 右側按鈕啟動 LibreOffice Writer 開啟 Word 檔（可編輯+存檔+修改格式）
- 依序嘗試 LlamaIndex -> Nanonets -> OCR.Space，一個失敗自動切下一個
- 支援修改 API Key 並寫回 WebOcrAPI.json
- 黑底白字、預設字體 20
"""

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from tkinter import ttk
import os
import json
import tempfile
import webbrowser
import subprocess
import threading
import time
from pathlib import Path

from ocrspace_core import process_single as ocrspace_process, load_api_key as ocrspace_load_key
from nanonets_core import process_single as nanonets_process, load_api_key as nanonets_load_key
from llamaindex_core import process_image as llamaindex_process, load_api_key as llamaindex_load_key, get_base_url as llamaindex_get_base_url


# ---------------------------
# Style & constants
# ---------------------------
FG_COLOR = 'white'
BG_COLOR = 'black'
ENTRY_BG = '#333333'
BTN_BG = '#555555'
LOG_BG = '#1a1a1a'
LOG_FG = '#00ff00'
DEFAULT_FONT = ('Arial', 20)
TITLE_FONT = ('Arial', 20, 'bold')
NORMAL_FONT = ('Arial', 14)
SMALL_FONT = ('Arial', 12)
MONO_FONT = ('Consolas', 16)

OUTPUT_DIR = r"/mnt/e/DiskCUse/HFDownloads"


# ---------------------------
# Helpers
# ---------------------------
def time_str() -> str:
    return time.strftime("%H:%M:%S")


# ---------------------------
# Main App
# ---------------------------
class UnifiedOCRApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        root.title("統一 OCR 辨識工具")
        root.geometry("1600x920")
        root.configure(bg=BG_COLOR)
        root.minsize(1200, 720)

        # state first
        self.image_paths = []        # list of paths, multi-select
        self.current_image_idx = 0   # which image is shown
        self.latest_docx_path = None
        self.latest_docx_paths = {}  # image_path -> docx_path map
        self.running = False
        self.api_keys = {}

        # UI
        self._build_top_bar()
        self._build_middle()
        self._build_bottom_log()

        # keys
        try:
            self._load_all_keys()
        except Exception as e:
            messagebox.showerror("啟動失敗", f"WebOcrAPI.json 讀取失敗：{e}")

    # -------------------
    # Load / Save keys
    # -------------------
    def _load_all_keys(self):
        try:
            with open("WebOcrAPI.json", "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception as e:
            messagebox.showerror("啟動失敗", f"WebOcrAPI.json 讀取失敗：{e}")
            data = []

        for item in data if isinstance(data, list) else []:
            name = item.get("ModelName")
            key = item.get("api_key", "")
            self.api_keys[name] = key

        self.log(f"已載入 API Key：{list(self.api_keys.keys())}")

    def _save_key_to_json(self, model_name: str, new_key: str):
        if not new_key.strip():
            messagebox.showwarning("提示", "API Key 不可為空，不會更新。")
            return False
        try:
            with open("WebOcrAPI.json", "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception as e:
            messagebox.showerror("錯誤", f"讀取 WebOcrAPI.json 失敗：{e}")
            return False

        updated = False
        for item in data if isinstance(data, list) else []:
            if item.get("ModelName") == model_name:
                item["api_key"] = new_key.strip()
                updated = True
                break

        if not updated:
            # append new entry
            data.append({"ModelName": model_name, "api_key": new_key.strip()})

        try:
            with open("WebOcrAPI.json", "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
        except Exception as e:
            messagebox.showerror("錯誤", f"寫入 WebOcrAPI.json 失敗：{e}")
            return False

        self.api_keys[model_name] = new_key.strip()
        return True

    # -------------------
    # UI build
    # -------------------
    def _build_top_bar(self):
        bar = tk.Frame(self.root, bg=BG_COLOR)
        bar.pack(side='top', fill='x', padx=20, pady=10)

        tk.Button(
            bar, text="選擇圖片（可多選）", command=self.select_files,
            bg=BTN_BG, fg=FG_COLOR, font=SMALL_FONT, padx=20, pady=8
        ).pack(side='left', padx=10)

        # file list dropdown
        tk.Label(bar, text="圖片：", fg=FG_COLOR, bg=BG_COLOR, font=SMALL_FONT).pack(side='left', padx=(10, 2))
        self.file_var = tk.StringVar(value="尚未選擇")
        self.file_combo = ttk.Combobox(
            bar, textvariable=self.file_var, values=["尚未選擇"],
            state='readonly', width=35, font=SMALL_FONT
        )
        self.file_combo.pack(side='left', padx=2)
        self.file_combo.bind('<<ComboboxSelected>>', self._on_file_selected)

        tk.Button(
            bar, text="開始辨識", command=self.start_process,
            bg='#0066cc', fg=FG_COLOR, font=SMALL_FONT, padx=24, pady=8
        ).pack(side='left', padx=10)

        tk.Button(
            bar, text="修改 API Key", command=self.open_api_manager,
            bg='#885500', fg=FG_COLOR, font=SMALL_FONT, padx=20, pady=8
        ).pack(side='left', padx=10)

        tk.Label(bar, text="輸出：", fg=FG_COLOR, bg=BG_COLOR, font=SMALL_FONT).pack(side='left', padx=10)
        self.out_var = tk.StringVar(value=OUTPUT_DIR)
        tk.Entry(bar, textvariable=self.out_var, width=42, bg=ENTRY_BG, fg=FG_COLOR, font=SMALL_FONT).pack(side='left', padx=5)
        tk.Button(bar, text="瀏覽", command=self.select_out_dir, bg=BTN_BG, fg=FG_COLOR, font=SMALL_FONT).pack(side='left')

    def _build_middle(self):
        mid = tk.PanedWindow(self.root, orient=tk.HORIZONTAL, bg=BG_COLOR, sashwidth=8)
        mid.pack(fill='both', expand=True, padx=20, pady=8)

        # left: image
        left = tk.Frame(mid, bg=BG_COLOR)
        tk.Label(left, text="原圖", fg=FG_COLOR, bg=BG_COLOR, font=TITLE_FONT).pack(anchor='w')
        self.image_canvas = tk.Canvas(left, bg='#111111', highlightthickness=1, highlightbackground='#555555')
        self.image_canvas.pack(fill='both', expand=True)
        self.image_canvas.bind("<Configure>", lambda e: self._draw_image_preview())
        self.preview_photo = None
        mid.add(left, minsize=400)

        # right: Word 表格編輯區
        right = tk.Frame(mid, bg=BG_COLOR)
        # --- Toolbar ---
        toolbar = tk.Frame(right, bg=BG_COLOR)
        toolbar.pack(fill='x', padx=0, pady=(0, 4))
        tk.Label(toolbar, text="Word 表格編輯", fg=FG_COLOR, bg=BG_COLOR, font=TITLE_FONT).pack(side='left')
        self.word_status = tk.Label(
            toolbar, text="（尚未產出）", fg='orange', bg=BG_COLOR, font=SMALL_FONT
        )
        self.word_status.pack(side='left', padx=12)

        # file path display
        self.word_path_var = tk.StringVar(value="")
        path_bar = tk.Frame(right, bg=BG_COLOR)
        path_bar.pack(fill='x', pady=(0, 4))
        tk.Label(
            path_bar, textvariable=self.word_path_var,
            fg='#aaaaaa', bg=BG_COLOR, font=('Consolas', 11), anchor='w'
        ).pack(side='left', fill='x', expand=True)

        # --- Treeview table ---
        table_frame = tk.Frame(right, bg=BG_COLOR)
        table_frame.pack(fill='both', expand=True)

        # Vertical scrollbar
        vsb = ttk.Scrollbar(table_frame, orient='vertical')
        vsb.pack(side='right', fill='y')
        # Horizontal scrollbar
        hsb = ttk.Scrollbar(table_frame, orient='horizontal')
        hsb.pack(side='bottom', fill='x')

        self.tree = ttk.Treeview(
            table_frame,
            columns=[],
            show='headings',
            yscrollcommand=vsb.set,
            xscrollcommand=hsb.set,
            selectmode='browse'
        )
        self.tree.pack(side='left', fill='both', expand=True)
        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)

        # Style for dark theme
        style = ttk.Style()
        style.theme_use('default')
        style.configure("Treeview",
            background=LOG_BG, fieldbackground=LOG_BG, foreground=LOG_FG,
            font=MONO_FONT, rowheight=28, borderwidth=0)
        style.configure("Treeview.Heading",
            background=BTN_BG, foreground=FG_COLOR, font=SMALL_FONT, borderwidth=1)
        style.map("Treeview", background=[('selected', '#004488')])

        # Bind double-click for cell editing
        self.tree.bind('<Double-1>', self._on_cell_double_click)
        self.tree.bind('<Button-1>', self._on_tree_click)

        # --- Save button ---
        save_frame = tk.Frame(right, bg=BG_COLOR)
        save_frame.pack(fill='x', pady=(8, 0))
        self.save_word_btn = tk.Button(
            save_frame, text="💾 儲存回 Word 檔",
            command=self._save_treeview_to_docx,
            bg='#cc6600', fg=FG_COLOR, font=DEFAULT_FONT, padx=24, pady=10,
            state='disabled'
        )
        self.save_word_btn.pack(fill='x')

        # --- Open in Explorer button (keep for convenience) ---
        self.open_explorer_btn = tk.Button(
            save_frame, text="📁 用檔案總管開啟所在目錄",
            command=self._open_in_explorer,
            bg='#336633', fg=FG_COLOR, font=DEFAULT_FONT, padx=24, pady=10,
            state='disabled'
        )
        self.open_explorer_btn.pack(fill='x', pady=(6, 0))

        # Internal state for editing
        self._edit_entry = None
        self._edit_row_id = None
        self._edit_col_idx = None
        self._current_docx_path = None
        self._dirty = False

        mid.add(right, minsize=500)

    def _build_bottom_log(self):
        bottom = tk.Frame(self.root, bg=BG_COLOR)
        bottom.pack(side='bottom', fill='both', expand=False, padx=20, pady=(0, 12))
        tk.Label(bottom, text="處理日誌", fg=FG_COLOR, bg=BG_COLOR, font=SMALL_FONT).pack(anchor='w')
        self.log_area = scrolledtext.ScrolledText(
            bottom, height=8, bg=LOG_BG, fg=LOG_FG,
            font=MONO_FONT, state='disabled'
        )
        self.log_area.pack(fill='both', expand=True)

    # -------------------
    # Image preview
    # -------------------
    def select_files(self):
        paths = filedialog.askopenfilenames(
            title="選擇圖片檔案（可多選）",
            filetypes=[("Image files", "*.png *.PNG *.jpg *.JPG *.jpeg *.JPEG *.bmp *.BMP *.tiff *.TIFF *.webp *.WEBP")],
            initialdir=r"/mnt/e/DiskCUse/HFDownloads/OCRUse02"
        )
        if paths:
            self.image_paths = list(paths)
            self.current_image_idx = 0
            self.latest_docx_paths = {}
            # update dropdown
            names = [os.path.basename(p) for p in self.image_paths]
            self.file_combo['values'] = names
            self.file_var.set(names[0])
            self.log(f"已選擇 {len(paths)} 個檔案：{', '.join(names)}")
            self.latest_docx_path = None
            self._update_word_buttons()
            self._draw_image_preview()

    def _on_file_selected(self, event=None):
        idx = self.file_combo.current()
        if idx >= 0 and idx < len(self.image_paths):
            self.current_image_idx = idx
            self._draw_image_preview()
            self._update_word_buttons()

    def _current_image_path(self):
        if self.image_paths and 0 <= self.current_image_idx < len(self.image_paths):
            return self.image_paths[self.current_image_idx]
        return None

    def _draw_image_preview(self):
        path = self._current_image_path()
        if not path:
            self.image_canvas.delete("all")
            self.image_canvas.create_text(
                self.image_canvas.winfo_width() / 2,
                self.image_canvas.winfo_height() / 2,
                text="尚未選擇圖片", fill='#888888', font=SMALL_FONT
            )
            return
        try:
            from PIL import Image, ImageTk
            img = Image.open(path)
            # fit to canvas
            cw = max(self.image_canvas.winfo_width(), 400)
            ch = max(self.image_canvas.winfo_height(), 300)
            img_ratio = img.width / img.height
            canvas_ratio = cw / ch
            if img_ratio > canvas_ratio:
                new_w = cw
                new_h = int(cw / img_ratio)
            else:
                new_h = ch
                new_w = int(ch * img_ratio)
            img = img.resize((new_w, new_h), Image.Resampling.LANCZOS)
            self.preview_photo = ImageTk.PhotoImage(img)
            self.image_canvas.delete("all")
            self.image_canvas.create_image(cw // 2, ch // 2, image=self.preview_photo, anchor='center')
        except Exception as e:
            self.log(f"圖片預覽失敗：{e}")

    # -------------------
    # Output dir
    # -------------------
    def select_out_dir(self):
        path = filedialog.askdirectory(title="選擇輸出目錄", initialdir=OUTPUT_DIR)
        if path:
            self.out_var.set(path)

    # -------------------
    # API manager
    # -------------------
    def open_api_manager(self):
        win = tk.Toplevel(self.root)
        win.title("修改 API Key")
        win.geometry("780x520")
        win.configure(bg=BG_COLOR)

        tk.Label(win, text="模型", fg=FG_COLOR, bg=BG_COLOR, font=SMALL_FONT, width=12, anchor='w').grid(row=0, column=0, padx=10, pady=10)
        tk.Label(win, text="原始 API Key", fg=FG_COLOR, bg=BG_COLOR, font=SMALL_FONT, width=14, anchor='w').grid(row=0, column=1, padx=10, pady=10)

        model_names = ["llamaindex", "nanonets", "ocr.space"]
        entries = {}
        for i, name in enumerate(model_names, start=1):
            tk.Label(win, text=name, fg=FG_COLOR, bg=BG_COLOR, font=SMALL_FONT, width=12, anchor='w').grid(row=i, column=0, padx=10, pady=8, sticky='w')
            var = tk.StringVar(value=self.api_keys.get(name, ""))
            e = tk.Entry(win, textvariable=var, width=58, bg=ENTRY_BG, fg=FG_COLOR, font=SMALL_FONT)
            e.grid(row=i, column=1, padx=10, pady=8)
            entries[name] = var

        status = tk.Label(win, text="", fg='orange', bg=BG_COLOR, font=SMALL_FONT)
        status.grid(row=4, column=0, columnspan=2, sticky='w', padx=10)

        def save_all():
            changed = False
            for name, var in entries.items():
                new_key = var.get()
                if self.api_keys.get(name) != new_key:
                    if self._save_key_to_json(name, new_key):
                        status.config(text=f"已更新：{name}", fg='#00ff00')
                        changed = True
                    else:
                        status.config(text=f"更新失敗：{name}", fg='red')
                        return
            if not changed:
                status.config(text="無變更", fg='orange')
            else:
                self.log("API Key 已更新並寫回 WebOcrAPI.json")

        tk.Button(win, text="儲存全部修改", command=save_all, bg='#0066cc', fg=FG_COLOR, font=SMALL_FONT, padx=16, pady=6).grid(row=5, column=1, sticky='e', padx=10, pady=12)

    # -------------------
    # Process
    # -------------------
    def start_process(self):
        if self.running:
            messagebox.showinfo("提示", "正在處理中，請稍候。")
            return
        path = self._current_image_path()
        if not path or not os.path.exists(path):
            messagebox.showerror("錯誤", "請先選擇有效的圖片檔案")
            return

        out_dir = self.out_var.get().strip()
        if not out_dir:
            out_dir = OUTPUT_DIR
        os.makedirs(out_dir, exist_ok=True)

        self.running = True
        image_path = self._current_image_path()
        self.log(f"開始依序嘗試 [{os.path.basename(image_path)}]：LlamaIndex -> Nanonets -> OCR.Space")
        threading.Thread(target=self._process_worker, args=(out_dir, image_path,), daemon=True).start()

    def _process_worker(self, out_dir: str, image_path: str):
        errors = []
        engine_results = {}
        last_docx = None

        # 1) LlamaIndex
        try:
            self.log("[核心] LlamaIndex 開始辨識 ...")
            res = llamaindex_process(
                api_key=self.api_keys.get("llamaindex", ""),
                image_path=image_path,
                output_dir=out_dir,
                save_word=True,
                save_excel=False,
                save_html=False,
                save_md=False,
            )
            engine_results["llamaindex"] = res
            last_docx = res.get("word")
            self.log(f"[核心] LlamaIndex 完成：{res}")
        except Exception as e:
            msg = str(e)
            errors.append(f"LlamaIndex 失敗：{msg}")
            self.log(f"[核心] LlamaIndex 失敗：{msg}")

        # 2) Nanonets
        if not last_docx:
            try:
                self.log("[核心] Nanonets 開始辨識 ...")
                res = nanonets_process(
                    api_key=self.api_keys.get("nanonets", ""),
                    image_path=image_path,
                    output_dir=out_dir,
                    save_word=True,
                    save_excel=False,
                    save_md=False,
                )
                engine_results["nanonets"] = res
                last_docx = res.get("word")
                self.log(f"[核心] Nanonets 完成：{res}")
            except Exception as e:
                msg = str(e)
                errors.append(f"Nanonets 失敗：{msg}")
                self.log(f"[核心] Nanonets 失敗：{msg}")

        # 3) OCR.Space
        if not last_docx:
            try:
                self.log("[核心] OCR.Space 開始辨識 ...")
                res = ocrspace_process(
                    api_key=self.api_keys.get("ocr.space", ""),
                    image_path=image_path,
                    output_dir=out_dir,
                    save_word=True,
                    save_excel=False,
                    save_text=False,
                    save_md=False,
                )
                engine_results["ocrspace"] = res
                last_docx = res.get("word")
                self.log(f"[核心] OCR.Space 完成：{res}")
            except Exception as e:
                msg = str(e)
                errors.append(f"OCR.Space 失敗：{msg}")
                self.log(f"[核心] OCR.Space 失敗：{msg}")

        if last_docx and os.path.exists(last_docx):
            self.latest_docx_path = last_docx
            self.latest_docx_paths[image_path] = last_docx
            self.root.after(0, lambda: self._update_word_buttons())
            self.root.after(0, lambda: messagebox.showinfo("完成", f"辨識完成\nWord：{last_docx}"))
            self.log(f"完成：{last_docx}")
        elif errors:
            self.root.after(0, lambda: messagebox.showerror("全部失敗", "\n".join(errors)))
            self.log("全部引擎皆失敗。")
        else:
            self.root.after(0, lambda: messagebox.showwarning("未完成", "已處理，但未產出 Word。"))
            self.log("已處理，但未產出 Word。")

        self.running = False

    # -------------------
    # Word actions — LibreOffice Writer
    # -------------------
    def _update_word_buttons(self):
        """Enable/disable buttons and load table based on current image's docx status."""
        image = self._current_image_path()
        docx = self.latest_docx_paths.get(image) if image else None
        has_docx = docx and os.path.exists(docx)
        if docx:
            self.word_path_var.set(docx)
        elif image:
            self.word_path_var.set("（此圖片尚未產出 Word）")
        else:
            self.word_path_var.set("")
        if has_docx:
            self.save_word_btn.config(state='normal')
            self.open_explorer_btn.config(state='normal')
            self.word_status.config(text="✅ 已就緒", fg='#00ff00')
            # Load table into Treeview if not already loaded for this docx
            if self._current_docx_path != docx:
                self._load_docx_to_treeview(docx)
        else:
            self.save_word_btn.config(state='disabled')
            self.open_explorer_btn.config(state='disabled')
            # Clear tree
            self.tree.delete(*self.tree.get_children())
            self.tree['columns'] = []
            self._current_docx_path = None
            self._dirty = False
            if self.running:
                self.word_status.config(text="⏳ 辨識中...", fg='orange')
            else:
                self.word_status.config(text="（尚未產出）", fg='orange')

    def _open_in_explorer(self):
        """Open the output directory in Windows File Explorer."""
        image = self._current_image_path()
        docx = self.latest_docx_paths.get(image) if image else None
        if docx and os.path.exists(docx):
            parent = os.path.dirname(docx)
        else:
            parent = self.out_var.get().strip() or OUTPUT_DIR
        if not os.path.exists(parent):
            os.makedirs(parent, exist_ok=True)
        # Use explorer.exe via WSL interop to open in Windows
        try:
            if os.path.exists("/mnt/c/Windows/explorer.exe"):
                wsl_path = str(Path(parent).resolve())
                subprocess.Popen(
                    ["/mnt/c/Windows/explorer.exe", wsl_path],
                    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
                )
                self.log(f"已用檔案總管開啟：{parent}")
            else:
                # Fallback: use xdg-open (Linux native)
                subprocess.Popen(
                    ["xdg-open", parent],
                    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
                )
                self.log(f"已用檔案管理器開啟：{parent}")
        except Exception as e:
            self.log(f"開啟目錄失敗：{e}")
            messagebox.showerror("錯誤", f"無法開啟目錄：{e}")

    # -------------------
    # Word table editor — Treeview
    # -------------------
    def _load_docx_to_treeview(self, docx_path: str):
        """Load the first table from a .docx into the Treeview."""
        try:
            from docx import Document
            doc = Document(docx_path)
            if not doc.tables:
                self.log("Word 檔案中無表格")
                return
            table = doc.tables[0]  # use first table

            # Extract headers
            headers = [cell.text.strip() for cell in table.rows[0].cells]
            if not headers:
                self.log("表格無表頭")
                return

            # Configure tree columns
            self.tree['columns'] = headers
            for h in headers:
                self.tree.heading(h, text=h, anchor='w')
                self.tree.column(h, width=120, minwidth=80, stretch=True)

            # Clear existing rows
            self.tree.delete(*self.tree.get_children())

            # Insert data rows (skip header row 0)
            for row_idx, row in enumerate(table.rows[1:], start=1):
                values = [cell.text.strip() for cell in row.cells]
                # Pad/truncate to match header count
                if len(values) < len(headers):
                    values += [''] * (len(headers) - len(values))
                elif len(values) > len(headers):
                    values = values[:len(headers)]
                self.tree.insert('', 'end', iid=f'row_{row_idx}', values=values)

            self._current_docx_path = docx_path
            self._dirty = False
            self.log(f"已載入表格：{len(table.rows)-1} 列 × {len(headers)} 欄")
        except Exception as e:
            self.log(f"載入 Word 表格失敗：{e}")
            messagebox.showerror("錯誤", f"無法讀取 Word 表格：{e}")

    def _on_tree_click(self, event):
        """Close any open editor when clicking elsewhere."""
        if self._edit_entry:
            self._finish_edit()

    def _on_cell_double_click(self, event):
        """Start editing the double-clicked cell."""
        if self._edit_entry:
            self._finish_edit()

        region = self.tree.identify('region', event.x, event.y)
        if region != 'cell':
            return

        row_id = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)
        if not row_id or not col_id:
            return

        col_idx = int(col_id.replace('#', '')) - 1
        if col_idx < 0 or col_idx >= len(self.tree['columns']):
            return

        # Get cell bounding box
        bbox = self.tree.bbox(row_id, col_id)
        if not bbox:
            return
        x, y, w, h = bbox

        # Get current value
        current_val = self.tree.item(row_id, 'values')[col_idx]

        # Create entry widget on top of cell
        entry = tk.Entry(
            self.tree, font=MONO_FONT,
            bg=ENTRY_BG, fg=FG_COLOR, insertbackground=FG_COLOR,
            borderwidth=1, relief='solid'
        )
        entry.place(x=x, y=y, width=w, height=h)
        entry.insert(0, current_val)
        entry.select_range(0, tk.END)
        entry.focus_set()

        def on_focus_out(e):
            self._finish_edit()

        def on_return(e):
            self._finish_edit()

        entry.bind('<FocusOut>', on_focus_out)
        entry.bind('<Return>', on_return)
        entry.bind('<Escape>', lambda e: self._cancel_edit())

        self._edit_entry = entry
        self._edit_row_id = row_id
        self._edit_col_idx = col_idx

    def _finish_edit(self):
        """Commit the edited value to Treeview."""
        if not self._edit_entry:
            return
        new_val = self._edit_entry.get()
        self._edit_entry.destroy()
        self._edit_entry = None

        if self._edit_row_id and self._edit_col_idx is not None:
            values = list(self.tree.item(self._edit_row_id, 'values'))
            if 0 <= self._edit_col_idx < len(values):
                values[self._edit_col_idx] = new_val
                self.tree.item(self._edit_row_id, values=values)
                self._dirty = True

        self._edit_row_id = None
        self._edit_col_idx = None

    def _cancel_edit(self):
        """Discard edit."""
        if self._edit_entry:
            self._edit_entry.destroy()
            self._edit_entry = None
            self._edit_row_id = None
            self._edit_col_idx = None

    def _save_treeview_to_docx(self):
        """Write Treeview data back to the original .docx (first table)."""
        if not self._current_docx_path or not os.path.exists(self._current_docx_path):
            messagebox.showerror("錯誤", "找不到原始 Word 檔案")
            return
        if not self._dirty:
            messagebox.showinfo("提示", "資料未變更，無需儲存")
            return

        try:
            from docx import Document
            doc = Document(self._current_docx_path)
            if not doc.tables:
                messagebox.showerror("錯誤", "Word 檔案中無表格")
                return
            table = doc.tables[0]

            headers = list(self.tree['columns'])
            row_ids = self.tree.get_children()

            # Update each data row (skip header row 0)
            for i, row_id in enumerate(row_ids):
                values = self.tree.item(row_id, 'values')
                word_row_idx = i + 1  # row 0 is header
                if word_row_idx >= len(table.rows):
                    break
                row = table.rows[word_row_idx]
                for j, val in enumerate(values):
                    if j < len(row.cells):
                        row.cells[j].text = str(val)

            doc.save(self._current_docx_path)
            self._dirty = False
            self.log(f"已儲存回 Word：{self._current_docx_path}")
            messagebox.showinfo("完成", "已儲存回 Word 檔")
        except Exception as e:
            self.log(f"儲存 Word 失敗：{e}")
            messagebox.showerror("錯誤", f"儲存失敗：{e}")

    # -------------------
    # Logging
    # -------------------
    def log(self, msg: str):
        line = f"[{time_str()}] {msg}"
        self.log_area.config(state='normal')
        self.log_area.insert(tk.END, line + "\n")
        self.log_area.see(tk.END)
        self.log_area.config(state='disabled')


def main():
    root = tk.Tk()
    app = UnifiedOCRApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
