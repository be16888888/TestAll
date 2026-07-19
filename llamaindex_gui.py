#!/usr/bin/env python3
"""
LlamaCloud OCR 表格提取工具 - UI Layer
Separated UI that uses llamaindex_core for processing logic.
"""

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import os
from llamaindex_core import process_image, load_api_key


class LlamaCloudOCRApp:
    def __init__(self, root):
        self.root = root
        self.root.title("LlamaCloud OCR 表格提取工具 (v2)")
        self.root.geometry("650x670")
        self.root.configure(bg='black')

        fg_color = 'white'
        bg_color = 'black'
        entry_bg = '#333333'
        btn_bg = '#555555'

        title = tk.Label(root, text="LlamaCloud OCR 表格提取工具 (v2)", font=('Arial', 16, 'bold'),
                         fg=fg_color, bg=bg_color)
        title.pack(pady=10)

        # API Key 輸入
        frame_key = tk.Frame(root, bg=bg_color)
        frame_key.pack(pady=5, fill='x', padx=20)
        tk.Label(frame_key, text="API Key：", fg=fg_color, bg=bg_color,
                 font=('Arial', 10)).pack(side='left')
        self.api_entry = tk.Entry(frame_key, width=50, bg=entry_bg, fg=fg_color,
                                  insertbackground='white')
        self.api_entry.pack(side='left', padx=5)

        # 輸出目錄選擇
        frame_out = tk.Frame(root, bg=bg_color)
        frame_out.pack(pady=5, fill='x', padx=20)
        tk.Label(frame_out, text="輸出目錄：", fg=fg_color, bg=bg_color,
                 font=('Arial', 10)).pack(side='left')
        self.dir_entry = tk.Entry(frame_out, width=40, bg=entry_bg, fg=fg_color)
        self.dir_entry.pack(side='left', padx=5)
        self.dir_entry.insert(0, "/mnt/e/HFDownloads/")
        tk.Button(frame_out, text="瀏覽", command=self.select_dir,
                  bg=btn_bg, fg=fg_color).pack(side='left', padx=5)

        # 圖片選擇
        frame_file = tk.Frame(root, bg=bg_color)
        frame_file.pack(pady=10, fill='x', padx=20)
        self.file_label = tk.Label(frame_file, text="未選擇圖片", fg='gray', bg=bg_color,
                                   width=40, anchor='w')
        self.file_label.pack(side='left')
        tk.Button(frame_file, text="選擇圖片", command=self.select_file,
                  bg=btn_bg, fg=fg_color, padx=10).pack(side='right')

        # 輸出格式勾選框
        frame_options = tk.LabelFrame(root, text="儲存格式選項", bg=bg_color, fg=fg_color,
                                      font=('Arial', 10))
        frame_options.pack(pady=10, fill='x', padx=20)

        self.var_word = tk.IntVar(value=1)
        self.var_excel = tk.IntVar(value=0)
        self.var_html = tk.IntVar(value=1)
        self.var_md = tk.IntVar(value=1)

        chk_word = tk.Checkbutton(frame_options, text="儲存 WORD 檔", variable=self.var_word,
                                  bg=bg_color, fg=fg_color, selectcolor='black', font=('Arial', 10))
        chk_word.pack(side='left', padx=10)

        chk_excel = tk.Checkbutton(frame_options, text="儲存 EXCEL 檔", variable=self.var_excel,
                                   bg=bg_color, fg=fg_color, selectcolor='black', font=('Arial', 10))
        chk_excel.pack(side='left', padx=10)

        chk_html = tk.Checkbutton(frame_options, text="儲存原始 HTML 內容", variable=self.var_html,
                                  bg=bg_color, fg=fg_color, selectcolor='black', font=('Arial', 10))
        chk_html.pack(side='left', padx=10)

        chk_md = tk.Checkbutton(frame_options, text="儲存 MD 文件", variable=self.var_md,
                                bg=bg_color, fg=fg_color, selectcolor='black', font=('Arial', 10))
        chk_md.pack(side='left', padx=10)

        # 開始按鈕
        self.process_btn = tk.Button(root, text="🚀 開始辨識並轉換", command=self.process,
                                     bg='#0066cc', fg=fg_color, font=('Arial', 12, 'bold'),
                                     padx=20, pady=10)
        self.process_btn.pack(pady=15)

        # 日誌區域
        self.log_area = scrolledtext.ScrolledText(root, height=18, bg='#1a1a1a',
                                                  fg='#00ff00', font=('Consolas', 10))
        self.log_area.pack(fill='both', expand=True, padx=20, pady=10)
        self.log_area.insert(tk.END, "就緒，請填寫 API Key 並選擇圖片。\n")
        self.log_area.config(state='disabled')

        self.image_path = None

        # 自動載入（或建立）WebOcrAPI.json
        self.ensure_and_load_json()

    def ensure_and_load_json(self):
        """確保 WebOcrAPI.json 存在，並載入 API Key"""
        json_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "WebOcrAPI.json")
        
        if not os.path.exists(json_path):
            default_config = {
                "url": "https://api.cloud.llamaindex.ai",
                "base_url": "https://api.cloud.llamaindex.ai/api/v2",
                "api_key": "llx-kp6L7gAeSDVvtfS4ZXXwkFeQNG82KfcoSxpoC1Wqga1cr5Xa"
            }
            try:
                with open(json_path, 'w', encoding='utf-8') as f:
                    json.dump(default_config, f, indent=2, ensure_ascii=False)
                self.log("已自動建立 WebOcrAPI.json（內含預設 API Key）")
            except Exception as e:
                self.log(f"建立 WebOcrAPI.json 失敗：{e}")
                return

        try:
            with open(json_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
            api_key = config.get("api_key") or config.get("LLAMA_CLOUD_API_KEY") or config.get("API_KEY")
            if api_key:
                self.api_entry.delete(0, tk.END)
                self.api_entry.insert(0, api_key)
                self.log("已從 WebOcrAPI.json 自動載入 API Key")
            else:
                self.log("警告：WebOcrAPI.json 中未找到 API Key 欄位，請手動輸入")
            self.base_url = config.get("base_url") or config.get("url") or "https://api.cloud.llamaindex.ai/api/v2"
        except Exception as e:
            self.log(f"載入 WebOcrAPI.json 失敗：{e}，請手動輸入 API Key")

    def select_dir(self):
        path = filedialog.askdirectory(title="選擇輸出目錄", initialdir="/mnt/e/HFDownloads/")
        if path:
            self.dir_entry.delete(0, tk.END)
            self.dir_entry.insert(0, path)

    def select_file(self):
        path = filedialog.askopenfilename(
            title="選擇圖片檔案",
            filetypes=[("Image files", "*.png *.jpg *.jpeg *.bmp *.tiff *.pdf")],
            initialdir="/mnt/e/HFDownloads/OCRUse02/"
        )
        if path:
            self.image_path = path
            self.file_label.config(text=os.path.basename(path), fg='white')
            self.log(f"已選擇檔案：{path}")

    def log(self, msg):
        self.log_area.config(state='normal')
        self.log_area.insert(tk.END, msg + "\n")
        self.log_area.see(tk.END)
        self.log_area.config(state='disabled')

    def process(self):
        api_key = self.api_entry.get().strip()
        if not api_key:
            messagebox.showerror("錯誤", "請先輸入 API Key")
            return
        if not self.image_path or not os.path.exists(self.image_path):
            messagebox.showerror("錯誤", "請先選擇有效的圖片檔案")
            return
        if not (self.var_word.get() or self.var_excel.get() or self.var_html.get() or self.var_md.get()):
            messagebox.showerror("錯誤", "請至少勾選一種輸出格式")
            return

        out_dir = self.dir_entry.get().strip()
        if not out_dir:
            out_dir = os.getcwd()
        if not os.path.exists(out_dir):
            try:
                os.makedirs(out_dir)
            except Exception as e:
                messagebox.showerror("錯誤", f"無法建立輸出目錄：{e}")
                return

        self.process_btn.config(state='disabled')
        self.log("開始處理...")

        try:
            self.log("正在上傳圖片並呼叫 LlamaCloud API v2...")
            results = process_image(
                api_key=api_key,
                image_path=self.image_path,
                output_dir=out_dir,
                save_word=bool(self.var_word.get()),
                save_excel=bool(self.var_excel.get()),
                save_html=bool(self.var_html.get()),
                save_md=bool(self.var_md.get())
            )
            
            self.log(f"API 呼叫成功，正在處理結果...")
            
            if 'html' in results:
                self.log(f"HTML 已儲存：{results['html']}")
            if 'md' in results:
                self.log(f"MD 文件已儲存：{results['md']}")
            if 'word' in results:
                self.log(f"Word 已儲存：{results['word']}")
            if 'excel' in results:
                self.log(f"Excel 已儲存：{results['excel']}")

            if not results:
                messagebox.showinfo("完成", "處理完成，但未生成任何輸出檔案。")
            else:
                messagebox.showinfo("完成", f"轉換完成！\n檔案已儲存至：\n{out_dir}")

        except Exception as e:
            error_msg = f"錯誤：{str(e)}"
            self.log(error_msg)
            messagebox.showerror("處理失敗", error_msg)
        finally:
            self.process_btn.config(state='normal')


if __name__ == "__main__":
    root = tk.Tk()
    app = LlamaCloudOCRApp(root)
    root.mainloop()