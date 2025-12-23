import fitz  # PyMuPDF：用於處理 PDF 的核心函式庫
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinterdnd2 import DND_FILES, TkinterDnD  # 支援拖放檔案功能
import ctypes
import os
import webbrowser
import platform
import threading  # 用於非同步處理轉換，避免介面卡死
import queue      # 用於執行緒間的安全通訊
import re

# 1. 跨平台動態字體偵測
def get_system_font():
    current_os = platform.system()
    if current_os == "Windows":
        return "Microsoft JhengHei"
    elif current_os == "Darwin":
        return "PingFang TC"
    elif current_os == "Linux":
        return "Noto Sans CJK TC"
    else:
        return "Arial"

SYSTEM_FONT = get_system_font()

try:
    if platform.system() == "Windows":
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

PAGE_SIZES = {
    "原始大小": None,
    "A3 (297 x 420 mm)": (841.89, 1190.55),
    "A4 (210 x 297 mm)": (595.27, 841.89),
    "A5 (148 x 210 mm)": (419.53, 595.27),
    "A6 (105 x 148 mm)": (297.64, 419.53),
    "B4 (250 x 353 mm)": (708.66, 1000.63),
    "B5 (176 x 250 mm)": (498.90, 708.66),
    "Letter (8.5 x 11\")": (612.0, 792.0),
    "Legal (8.5 x 14\")": (612.0, 1008.0),
    "Tabloid (11 x 17\")": (792.0, 1224.0),
    "4 x 6 吋 (相片)": (288.0, 432.0),
    "5 x 7 吋 (相片)": (360.0, 504.0),
}

class PlaceholderEntry(tk.Entry):
    def __init__(self, container, placeholder, is_password=False, *args, **kwargs):
        super().__init__(container, *args, **kwargs)
        self.placeholder = placeholder
        self.placeholder_color = '#aaaaaa'
        self.default_fg_color = 'black'
        self.is_password = is_password
        self.real_show = kwargs.get('show', '*') if is_password else ''
        self.bind("<FocusIn>", self._clear_placeholder)
        self.bind("<FocusOut>", self._add_placeholder)
        self._add_placeholder()

    def _add_placeholder(self, event=None):
        if not self.get():
            self.insert(0, self.placeholder)
            self['fg'] = self.placeholder_color
            if self.is_password:
                self.config(show='')

    def _clear_placeholder(self, event=None):
        if self['fg'] == self.placeholder_color:
            self.delete(0, tk.END)
            self['fg'] = self.default_fg_color
            if self.is_password:
                self.config(show=self.real_show)

    def get_real_value(self):
        if self['fg'] == self.placeholder_color:
            return ""
        return self.get()

class FilePasswordDialog(tk.Toplevel):
    def __init__(self, parent, filename):
        super().__init__(parent)
        self.title("PDF 檔案解鎖")
        self.filename = filename
        self.password = None
        width, height = 480, 220
        self.root = parent.winfo_toplevel()
        pos_x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (width // 2)
        pos_y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (height // 2)
        self.geometry(f"{width}x{height}+{pos_x}+{pos_y}")
        self.configure(bg="white")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        content = tk.Frame(self, bg="white", padx=30, pady=20)
        content.pack(fill=tk.BOTH, expand=True)
        tk.Label(content, text="此 PDF 檔案受保護，請輸入開啟密碼：", font=(SYSTEM_FONT, 10), bg="white").pack(anchor="w")
        tk.Label(content, text=filename, font=(SYSTEM_FONT, 10, "bold"), bg="white", fg="#0056b3", wraplength=400, justify="left").pack(anchor="w", pady=(5, 15))
        self.entry = tk.Entry(content, font=(SYSTEM_FONT, 11), show="*", relief=tk.SOLID, borderwidth=1)
        self.entry.pack(fill=tk.X, pady=5)
        self.entry.focus_set()
        btn_frame = tk.Frame(content, bg="white")
        btn_frame.pack(pady=10)
        tk.Button(btn_frame, text="解鎖並加入", command=self.on_confirm, font=(SYSTEM_FONT, 10, "bold"), bg="#096dd9", fg="white", relief=tk.FLAT, padx=25, pady=5, cursor="hand2").pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="放棄此檔案", command=self.destroy, font=(SYSTEM_FONT, 10), bg="#f5f5f5", relief=tk.FLAT, padx=15, pady=5, cursor="hand2").pack(side=tk.LEFT, padx=5)
        self.bind("<Return>", lambda e: self.on_confirm())

    def on_confirm(self):
        self.password = self.entry.get()
        self.destroy()

class ImageToPdfConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("圖片轉PDF小工具 by Yu-Han Cheng")
        self.root.geometry("1200x950")
        self.root.configure(bg="#f0f2f5")
        self.root.minsize(1000, 900)
        
        # 核心資料結構
        self.file_list = []      
        self.pdf_passwords = {}  
        self.thumbnails = {}     # PhotoImage 快取
        self.doc_handles = {}    # 快取開啟的 fitz.Document 物件，避免頻繁 I/O
        self.is_converting = False 

        # 非同步縮圖載入隊列與執行緒
        self.thumb_queue = queue.Queue()
        self.thumb_thread_running = True
        self.thumb_worker = threading.Thread(target=self._thumbnail_worker, daemon=True)
        self.thumb_worker.start()

        self.setup_styles()
        self.create_widgets()

    def setup_styles(self):
        self.primary_color = "#0056b3"
        self.bg_light = "#ffffff"
        self.font_title = (SYSTEM_FONT, 16, "bold")
        self.font_header = (SYSTEM_FONT, 11, "bold")
        self.font_main = (SYSTEM_FONT, 10)
        self.font_status = (SYSTEM_FONT, 9)
        self.font_btn_big = (SYSTEM_FONT, 14, "bold")
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", font=self.font_main, rowheight=60, borderwidth=0)
        style.configure("Treeview.Heading", font=self.font_main)
        style.map("Treeview", background=[('selected', '#e1f5fe')], foreground=[('selected', 'black')])
        style.configure("TProgressbar", thickness=14)
        style.configure("TCombobox", font=self.font_main)

    def create_widgets(self):
        nav_frame = tk.Frame(self.root, bg="white", height=65)
        nav_frame.pack(fill=tk.X, side=tk.TOP)
        nav_frame.pack_propagate(False)
        tk.Frame(nav_frame, bg=self.primary_color, width=5).pack(side=tk.LEFT, fill=tk.Y, padx=(20, 0), pady=12)
        tk.Label(nav_frame, text="圖片轉PDF小工具", font=self.font_title, bg="white", fg="#333").pack(side=tk.LEFT, padx=15)
        about_link = tk.Label(nav_frame, text="關於本程式", font=self.font_main, bg="white", fg="#555", cursor="hand2")
        about_link.pack(side=tk.RIGHT, padx=30)
        about_link.bind("<Button-1>", lambda e: self.show_about())

        main_content = tk.Frame(self.root, bg="#f0f2f5", padx=25, pady=5)
        main_content.pack(fill=tk.BOTH, expand=True)

        self.src_section_frame, _ = self.create_section(main_content, "檔案來源")
        self.src_section_frame.master.pack(side=tk.TOP, fill=tk.X, pady=5)
        self.btn_select = tk.Button(self.src_section_frame, text=" ＋ 選擇檔案... ", command=self.add_files, font=self.font_main, bg="#fafafa", relief=tk.GROOVE, padx=15, pady=4)
        self.btn_select.pack(side=tk.LEFT, padx=15, pady=10)
        tk.Label(self.src_section_frame, text="請加入檔案 (亦可利用拖曳方式加入清單)", font=self.font_main, bg="white", fg="gray").pack(side=tk.LEFT)

        self.exec_section_frame, _ = self.create_section(main_content, "執行作業")
        self.exec_section_frame.master.pack(side=tk.BOTTOM, fill=tk.X, pady=5)
        exec_inner = tk.Frame(self.exec_section_frame, bg="white", padx=15, pady=10) 
        exec_inner.pack(fill=tk.X)
        left_exec_ctrl = tk.Frame(exec_inner, bg="white")
        left_exec_ctrl.pack(side=tk.LEFT, fill=tk.Y)
        self.auto_open_var = tk.BooleanVar(value=False)
        tk.Checkbutton(left_exec_ctrl, text="轉換完成後自動開啟資料夾", variable=self.auto_open_var, font=self.font_main, bg="white").pack(anchor="w")
        self.status_label = tk.Label(left_exec_ctrl, text="等待作業中...", font=self.font_status, bg="white", fg="gray")
        self.status_label.pack(anchor="w")
        self.progress = ttk.Progressbar(left_exec_ctrl, orient=tk.HORIZONTAL, length=320, mode='determinate')
        self.progress.pack(anchor="w", pady=(2, 0))
        self.btn_run = tk.Button(exec_inner, text="  🚀  開始產生 PDF  ", command=self.start_conversion_thread, bg="#096dd9", fg="white", font=self.font_btn_big, relief=tk.FLAT, padx=65, pady=10, cursor="hand2")
        self.btn_run.pack(side=tk.RIGHT, pady=5)

        self.param_section_frame, _ = self.create_section(main_content, "參數設定與文件資訊 (選填)")
        self.param_section_frame.master.pack(side=tk.BOTTOM, fill=tk.X, pady=5)
        grid_container = tk.Frame(self.param_section_frame, bg="white", padx=15, pady=10)
        grid_container.pack(fill=tk.X)
        
        row1 = tk.Frame(grid_container, bg="white")
        row1.pack(fill=tk.X, pady=8)
        tk.Label(row1, text="頁面尺寸:", font=self.font_main, bg="white").pack(side=tk.LEFT)
        self.page_size_var = tk.StringVar(value="原始大小")
        self.combo_size = ttk.Combobox(row1, textvariable=self.page_size_var, values=list(PAGE_SIZES.keys()), state="readonly", width=22); self.combo_size.pack(side=tk.LEFT, padx=5)
        tk.Label(row1, text="方向:", font=self.font_main, bg="white").pack(side=tk.LEFT, padx=(10,0))
        self.orientation_var = tk.StringVar(value="直式")
        self.combo_orient = ttk.Combobox(row1, textvariable=self.orientation_var, values=["直式", "橫式"], state="readonly", width=6); self.combo_orient.pack(side=tk.LEFT, padx=5)
        tk.Label(row1, text="圖片縮放:", font=self.font_main, bg="white").pack(side=tk.LEFT, padx=(15,0))
        self.scale_mode_var = tk.StringVar(value="自動填滿")
        self.combo_scale = ttk.Combobox(row1, textvariable=self.scale_mode_var, values=["自動填滿", "保持原尺寸"], state="readonly", width=10); self.combo_scale.pack(side=tk.LEFT, padx=5)

        row2 = tk.Frame(grid_container, bg="white")
        row2.pack(fill=tk.X, pady=2)
        self.compress_var = tk.BooleanVar(value=False)
        self.check_compress = tk.Checkbutton(row2, text="圖片壓縮", variable=self.compress_var, font=self.font_main, bg="white", command=self.toggle_compress)
        self.check_compress.pack(side=tk.LEFT)
        qual_frame = tk.Frame(row2, bg="white")
        qual_frame.pack(side=tk.LEFT, padx=(5, 5))
        tk.Label(qual_frame, text="品質:", font=self.font_main, bg="white").pack(side=tk.LEFT)
        self.quality_val_label = tk.Label(qual_frame, text="80%", font=(SYSTEM_FONT, 9, "bold"), bg="#f0f2f5", width=4)
        self.quality_scale = tk.Scale(qual_frame, from_=10, to=100, orient=tk.HORIZONTAL, length=80, bg="white", highlightthickness=0, showvalue=0, command=self.update_quality_label)
        self.quality_scale.set(80); self.quality_scale.pack(side=tk.LEFT, padx=5); self.quality_val_label.pack(side=tk.LEFT); self.quality_scale.config(state=tk.DISABLED)
        tk.Frame(row2, bg="#eee", width=1).pack(side=tk.LEFT, fill=tk.Y, padx=15)
        self.encrypt_var = tk.BooleanVar(value=False)
        self.check_encrypt = tk.Checkbutton(row2, text="PDF 加密", variable=self.encrypt_var, font=self.font_main, bg="white", command=self.toggle_encrypt)
        self.check_encrypt.pack(side=tk.LEFT)
        self.password_entry = PlaceholderEntry(row2, placeholder="設定密碼", is_password=True, font=self.font_main, width=12, relief=tk.SOLID, borderwidth=1)
        self.password_entry.pack(side=tk.LEFT, padx=5); self.password_entry.config(state=tk.DISABLED)
        tk.Frame(row2, bg="#eee", width=1).pack(side=tk.LEFT, fill=tk.Y, padx=15)
        self.auto_rotate_var = tk.BooleanVar(value=False)
        self.check_auto_rotate = tk.Checkbutton(row2, text="自動旋轉", variable=self.auto_rotate_var, font=self.font_main, bg="white")
        self.check_auto_rotate.pack(side=tk.LEFT)
        self.grayscale_var = tk.BooleanVar(value=False)
        self.check_grayscale = tk.Checkbutton(row2, text="黑白模式", variable=self.grayscale_var, font=self.font_main, bg="white")
        self.check_grayscale.pack(side=tk.LEFT, padx=5)

        row3 = tk.Frame(grid_container, bg="white")
        row3.pack(fill=tk.X, pady=8)
        meta_items = [("標題:", "meta_title", "文件標題", 10), ("作者:", "meta_author", "作者名稱", 8), ("主題:", "meta_subject", "主題內容", 10), ("關鍵字:", "meta_keywords", "逗號分隔", 10)]
        for lbl_text, attr_name, ph, w in meta_items:
            tk.Label(row3, text=lbl_text, font=self.font_main, bg="white").pack(side=tk.LEFT, padx=(5, 2))
            entry = PlaceholderEntry(row3, placeholder=ph, font=self.font_main, width=w, relief=tk.SOLID, borderwidth=1)
            entry.pack(side=tk.LEFT, padx=(0, 10), fill=tk.X, expand=True)
            setattr(self, attr_name, entry)

        self.list_section_frame, list_title_bar = self.create_section(main_content, "待處理清單")
        self.list_section_frame.master.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=5)
        self.file_count_label = tk.Label(list_title_bar, text="已選擇: 0 個項目", font=(SYSTEM_FONT, 9, "bold"), bg="#fafafa", fg=self.primary_color)
        self.file_count_label.pack(side=tk.LEFT, padx=(10, 0))
        
        list_main_container = tk.Frame(self.list_section_frame, bg="white", padx=15, pady=5)
        list_main_container.pack(fill=tk.BOTH, expand=True)
        tree_frame = tk.Frame(list_main_container, bg="white")
        tree_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.tree = ttk.Treeview(tree_frame, columns=("Index", "Type", "Name"), show='headings', selectmode='extended')
        self.tree.heading("Index", text="順序/頁碼"); self.tree.heading("Type", text="類型"); self.tree.heading("Name", text="項目名稱")
        self.tree.column("#0", width=80, anchor="center") 
        self.tree.column("Index", width=120, anchor="center"); self.tree.column("Type", width=100, anchor="center"); self.tree.column("Name", width=400)
        
        self.tree.configure(show="tree headings")
        self.tree.heading("#0", text="預覽")

        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview); self.tree.configure(yscroll=scrollbar.set); self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True); scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.bind("<Delete>", lambda e: self.remove_selected())

        self.side_btn_bar = tk.Frame(list_main_container, bg="white", padx=10)
        self.side_btn_bar.pack(side=tk.RIGHT, fill=tk.Y)        
        
        self.btn_expand = tk.Button(self.side_btn_bar, text="📂 展開 PDF", command=self.expand_selected_pdf, relief=tk.GROOVE, font=self.font_main, bg="#e6f7ff", fg="#1890ff", width=12)
        self.btn_expand.pack(pady=(0, 10))

        self.btn_up = tk.Button(self.side_btn_bar, text="▲ 上移", command=self.move_up, relief=tk.GROOVE, font=self.font_main, bg="#f8f9fa", width=12); self.btn_up.pack(pady=2)
        self.btn_down = tk.Button(self.side_btn_bar, text="▼ 下移", command=self.move_down, relief=tk.GROOVE, font=self.font_main, bg="#f8f9fa", width=12); self.btn_down.pack(pady=2)
        tk.Label(self.side_btn_bar, text="自動排序", font=self.font_status, bg="white", fg="gray").pack(pady=(5, 2))
        self.btn_sort_asc = tk.Button(self.side_btn_bar, text="A-Z 排序", command=lambda: self.sort_files(False), relief=tk.GROOVE, font=self.font_main, bg="#f8f9fa", width=12); self.btn_sort_asc.pack(pady=1)
        self.btn_sort_desc = tk.Button(self.side_btn_bar, text="Z-A 排序", command=lambda: self.sort_files(True), relief=tk.GROOVE, font=self.font_main, bg="#f8f9fa", width=12); self.btn_sort_desc.pack(pady=1)
        tk.Frame(self.side_btn_bar, bg="#eee", height=1).pack(fill=tk.X, pady=8)
        self.btn_remove = tk.Button(self.side_btn_bar, text="✕ 移除選取", command=self.remove_selected, bg="#fff1f0", fg="#cf1322", relief=tk.GROOVE, font=self.font_main, width=12); self.btn_remove.pack(pady=1)
        self.btn_clear = tk.Button(self.side_btn_bar, text="🗑 全部清空", command=self.clear_all, relief=tk.GROOVE, font=self.font_main, width=12); self.btn_clear.pack(pady=1)
        self.tree.drop_target_register(DND_FILES); self.tree.dnd_bind('<<Drop>>', self.handle_drop)

    def create_section(self, parent, title):
        container = tk.Frame(parent, bg="white", bd=1, relief="solid", highlightthickness=0)
        title_bar = tk.Frame(container, bg="#fafafa"); title_bar.pack(fill=tk.X)
        tk.Frame(title_bar, bg=self.primary_color, width=3).pack(side=tk.LEFT, fill=tk.Y, padx=(12, 6), pady=6)
        tk.Label(title_bar, text=title, font=self.font_header, bg="#fafafa", fg="#333").pack(side=tk.LEFT, pady=6)
        content = tk.Frame(container, bg="white"); content.pack(fill=tk.BOTH, expand=True)
        return content, title_bar

    # --- 效能優化相關方法 ---

    def _get_pdf_doc(self, path):
        """獲取快取的 PDF 句柄，若無則開啟。有效減少重複 I/O"""
        if path in self.doc_handles:
            return self.doc_handles[path]
        try:
            doc = fitz.open(path)
            if doc.is_encrypted:
                doc.authenticate(self.pdf_passwords.get(path, ""))
            self.doc_handles[path] = doc
            return doc
        except:
            return None

    def _thumbnail_worker(self):
        """非同步縮圖處理執行緒"""
        while self.thumb_thread_running:
            try:
                # 取得任務: (清單項目ID, 檔案路徑, 頁碼)
                item_id, path, page_idx = self.thumb_queue.get(timeout=1)
                
                cache_key = f"{path}_{page_idx}"
                if cache_key not in self.thumbnails:
                    # 在背景執行緒開啟一個暫時的 doc (避免執行緒衝突)
                    with fitz.open(path) as doc:
                        if doc.is_encrypted:
                            doc.authenticate(self.pdf_passwords.get(path, ""))
                        page = doc[page_idx]
                        pix = page.get_pixmap(matrix=fitz.Matrix(0.1, 0.1))
                        img_data = pix.tobytes("png")
                        # 由於 tk.PhotoImage 必須在主執行緒建立，所以傳回數據
                        self.root.after(0, self._update_item_thumbnail, item_id, cache_key, img_data)
                else:
                    # 若已有快取，直接更新 UI
                    self.root.after(0, lambda: self.tree.item(item_id, image=self.thumbnails[cache_key]))
                
                self.thumb_queue.task_done()
            except queue.Empty:
                continue

    def _update_item_thumbnail(self, item_id, cache_key, img_data):
        """在主執行緒更新 Treeview 中的縮圖"""
        if not self.tree.exists(item_id): return
        photo = tk.PhotoImage(data=img_data)
        self.thumbnails[cache_key] = photo
        self.tree.item(item_id, image=photo)

    def update_tree_content(self):
        """重新整理清單，優化頁碼計算並將縮圖載入排入隊列"""
        # 清空佇列，避免舊任務干擾
        while not self.thumb_queue.empty():
            try: self.thumb_queue.get_nowait()
            except queue.Empty: break

        for i in self.tree.get_children(): self.tree.delete(i)
        
        current_page_offset = 1
        for idx, item in enumerate(self.file_list):
            path = item['path']
            fname = os.path.basename(path)
            ext = fname.split('.')[-1].upper()
            
            # 使用快取資訊，不再重複打開檔案
            item_page_count = item.get('page_count', 1)
            
            if item_page_count > 1:
                index_text = f"{current_page_offset} ~ {current_page_offset + item_page_count - 1}"
            else:
                index_text = str(current_page_offset)
            current_page_offset += item_page_count

            if item['page'] is not None:
                type_text = f"{ext} (分頁)"
                display_name = f"{fname}\n(第 {item['page'] + 1} 頁)"
                target_page = item['page']
            else:
                type_text = ext
                display_name = fname
                target_page = 0
            
            # 先插入不帶圖的項目
            item_id = self.tree.insert("", tk.END, values=(index_text, type_text, display_name))
            
            # 檢查快取是否有圖，有則立刻上，無則排隊
            cache_key = f"{path}_{target_page}"
            if cache_key in self.thumbnails:
                self.tree.item(item_id, image=self.thumbnails[cache_key])
            else:
                self.thumb_queue.put((item_id, path, target_page))
            
        self.file_count_label.config(text=f"已選擇: {len(self.file_list)} 個項目")

    def expand_selected_pdf(self):
        """展開 PDF，利用快取具柄實現瞬發展開"""
        if self.is_converting: return
        sel = self.tree.selection()
        if not sel: return
        
        idxs = sorted([self.tree.index(i) for i in sel], reverse=True)
        expanded_any = False
        
        for idx in idxs:
            item = self.file_list[idx]
            path = item['path']
            if path.lower().endswith('.pdf') and item['page'] is None:
                doc = self._get_pdf_doc(path)
                if doc:
                    count = len(doc)
                    page_items = [{'path': path, 'page': p, 'page_count': 1} for p in range(count)]
                    self.file_list[idx:idx+1] = page_items
                    expanded_any = True
        
        if expanded_any:
            self.update_tree_content()

    def update_quality_label(self, val): 
        self.quality_val_label.config(text=f"{val}%")

    def show_about(self):
        about_win = tk.Toplevel(self.root); about_win.title("關於本程式")
        w, h = 650, 520
        px, py = self.root.winfo_x(), self.root.winfo_y()
        pw, ph = self.root.winfo_width(), self.root.winfo_height()
        about_win.geometry(f"{w}x{h}+{px + (pw // 2) - (w // 2)}+{py + (ph // 2) - (h // 2)}")
        about_win.configure(bg="white"); about_win.transient(self.root)
        content = tk.Frame(about_win, bg="white", padx=40, pady=25)
        content.pack(fill=tk.BOTH, expand=True)
        tk.Label(content, text="圖片轉PDF小工具", font=self.font_title, bg="white", fg=self.primary_color).pack(anchor="w")
        dev_info_frame = tk.Frame(content, bg="white")
        dev_info_frame.pack(fill=tk.X, pady=(15, 0))
        tk.Label(dev_info_frame, text="開發者：鄭郁翰 (Yu-Han Cheng)", font=self.font_main, bg="white").pack(anchor="w")
        tk.Label(dev_info_frame, text="Email：kaoshou@gmail.com", font=self.font_main, bg="white").pack(anchor="w")
        tk.Label(dev_info_frame, text="GitHub：https://github.com/kaoshou/image-pdf-converter", font=self.font_main, bg="white", fg="#0056b3", cursor="hand2").pack(anchor="w")
        tk.Frame(content, bg="#eee", height=1).pack(fill=tk.X, pady=20)
        tk.Label(content, text="專案資訊與開源聲明 (Open Source Disclosure)", font=self.font_header, bg="white").pack(anchor="w", pady=(0, 10))
        license_desc = (
            "本程式原始碼 (GitHub)：\nhttps://github.com/kaoshou/image-pdf-converter\n\n"
            "本程式核心功能基於以下的開源專案實作：\n\n"
            "• PyMuPDF (fitz)：採用 GNU AGPL v3.0 授權。\n"
            "• TkinterDnD2：採用 MIT 授權。\n"
            "• Python Standard Library：採用 PSF License 授權。\n\n"
            "免責聲明：本軟體依「現狀」提供，開發者對於因使用本程式所產生的任何損失概不負責。"
        )
        text_box = tk.Text(content, height=25, font=("Consolas", 9), bg="#f9f9f9", relief=tk.FLAT, wrap=tk.WORD, padx=12, pady=12)
        text_box.insert(tk.END, license_desc)
        text_box.config(state=tk.DISABLED); text_box.pack(fill=tk.X)

    def handle_drop(self, event):
        if not self.is_converting: self.process_incoming_files(self.root.tk.splitlist(event.data))

    def add_files(self):
        if not self.is_converting:
            files = filedialog.askopenfilenames(title="選擇檔案", filetypes=[("支援格式", "*.jpg *.jpeg *.png *.pdf *.bmp *.tiff")])
            if files: self.process_incoming_files(files)

    def process_incoming_files(self, files):
        valid = ('.jpg', '.jpeg', '.png', '.pdf', '.bmp', '.tiff'); added = False
        for f in files:
            exists = any(item['path'] == f and item['page'] is None for item in self.file_list)
            if f.lower().endswith(valid) and not exists:
                count = 1
                if f.lower().endswith('.pdf'):
                    doc = self._get_pdf_doc(f)
                    if doc:
                        if doc.is_encrypted and not self.pdf_passwords.get(f):
                            correct = False
                            while not correct:
                                dialog = FilePasswordDialog(self.root, os.path.basename(f))
                                self.root.wait_window(dialog)
                                if dialog.password is None: break
                                if doc.authenticate(dialog.password): self.pdf_passwords[f] = dialog.password; correct = True
                                else: messagebox.showerror("錯誤", "密碼不正確")
                            if not correct: continue
                        count = len(doc)
                self.file_list.append({'path': f, 'page': None, 'page_count': count})
                added = True
        if added: self.update_tree_content()

    def sort_files(self, rev):
        if not self.is_converting: 
            self.file_list.sort(key=lambda x: (os.path.basename(x['path']).lower(), x['page'] if x['page'] is not None else -1), reverse=rev)
            self.update_tree_content()

    def move_up(self):
        if self.is_converting: return
        sel = self.tree.selection(); idxs = sorted([self.tree.index(i) for i in sel])
        if not idxs or idxs[0] <= 0: return
        for idx in idxs: self.file_list[idx], self.file_list[idx-1] = self.file_list[idx-1], self.file_list[idx]
        self.update_tree_content()
        for idx in idxs: self.tree.selection_add(self.tree.get_children()[idx-1])
    
    def move_down(self):
        if self.is_converting: return
        sel = self.tree.selection(); idxs = sorted([self.tree.index(i) for i in sel], reverse=True)
        if not idxs or idxs[0] >= len(self.file_list) - 1: return
        for idx in idxs: self.file_list[idx], self.file_list[idx+1] = self.file_list[idx+1], self.file_list[idx]
        self.update_tree_content()
        for idx in idxs: self.tree.selection_add(self.tree.get_children()[idx+1])

    def remove_selected(self):
        if not self.is_converting:
            sel = self.tree.selection(); idxs = sorted([self.tree.index(i) for i in sel], reverse=True)
            for idx in idxs: 
                item = self.file_list.pop(idx)
                # 若該檔案已不存在於清單，清理快取與具柄
                if not any(it['path'] == item['path'] for it in self.file_list):
                    if item['path'] in self.doc_handles:
                        self.doc_handles[item['path']].close()
                        del self.doc_handles[item['path']]
                    self.pdf_passwords.pop(item['path'], None)
                    # 清理該檔案的所有縮圖快取
                    keys = [k for k in self.thumbnails if k.startswith(item['path'])]
                    for k in keys: del self.thumbnails[k]
            self.update_tree_content()
            
    def clear_all(self):
        if not self.is_converting and self.file_list and messagebox.askyesno("確認", "是否清空？"):
            for h in self.doc_handles.values(): h.close()
            self.file_list.clear(); self.pdf_passwords.clear(); self.thumbnails.clear(); self.doc_handles.clear(); self.update_tree_content()

    def toggle_compress(self): 
        s = tk.NORMAL if self.compress_var.get() else tk.DISABLED
        self.quality_scale.config(state=s); self.quality_val_label.config(fg="black" if self.compress_var.get() else "gray")

    def toggle_encrypt(self): 
        self.password_entry.config(state=tk.NORMAL if self.encrypt_var.get() else tk.DISABLED)

    def start_conversion_thread(self):
        if not self.file_list: return
        opw = self.password_entry.get_real_value()
        if self.encrypt_var.get() and not opw: messagebox.showwarning("警告", "請設定密碼"); return
        save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF 檔案", "*.pdf")])
        if not save_path: return
        self.is_converting = True; self.toggle_ui_state(tk.DISABLED)
        self.progress['value'] = 0; self.status_label.config(text="準備開始轉換...", fg="blue")
        threading.Thread(target=self.perform_conversion, args=(save_path,), daemon=True).start()

    def toggle_ui_state(self, state):
        btns = [self.btn_run, self.btn_select, self.btn_up, self.btn_down, self.btn_remove, self.btn_clear, self.btn_sort_asc, self.btn_sort_desc, self.btn_expand, self.check_compress, self.check_encrypt, self.combo_size, self.combo_orient, self.check_auto_rotate, self.check_grayscale, self.combo_scale]
        for b in btns: b.config(state=state)
        if state == tk.NORMAL: self.toggle_compress(); self.toggle_encrypt()
        else: self.quality_scale.config(state=tk.DISABLED); self.password_entry.config(state=tk.DISABLED)

    def perform_conversion(self, save_path):
        """核心轉換邏輯，修正方向判定問題與透明圖片報錯問題"""
        doc = fitz.open(); total = len(self.file_list)
        c, q = self.compress_var.get(), self.quality_scale.get()
        enc, opw = self.encrypt_var.get(), self.password_entry.get_real_value()
        gs, ar, sm = self.grayscale_var.get(), self.auto_rotate_var.get(), self.scale_mode_var.get()
        meta = {"title": self.meta_title.get_real_value(), "creator": "圖片轉PDF小工具", "producer": "PyMuPDF"}
        base_size = PAGE_SIZES.get(self.page_size_var.get()); target_orient = self.orientation_var.get()
        HIGH_RES_DPI = 300 / 72 

        try:
            for idx, item in enumerate(self.file_list):
                path = item['path']
                self.root.after(0, lambda i=idx+1: self.status_label.config(text=f"處理中 {i}/{total}..."))
                
                if not path.lower().endswith('.pdf'):
                    img_doc = fitz.open(path); img_page = img_doc[0]; img_rect = img_page.rect

                    if gs or c:
                        pix = img_page.get_pixmap(matrix=fitz.Matrix(HIGH_RES_DPI, HIGH_RES_DPI))
                        if gs: pix = fitz.Pixmap(fitz.csGRAY, pix)
                        if c and pix.alpha:
                            new_pix = fitz.Pixmap(fitz.csRGB, pix.width, pix.height, 0)
                            new_pix.clear_with(255); new_pix.copy(pix, pix.irect); pix = new_pix
                        img_data = pix.tobytes("jpg", jpg_quality=q) if c else pix.tobytes("png")
                        
                        if base_size:
                            tw, th = base_size if target_orient == "直式" else (base_size[1], base_size[0])
                            if ar and ((pix.width > pix.height) != (tw > th)): tw, th = th, tw
                            page = doc.new_page(width=tw, height=th)
                            rect = page.rect if sm == "自動填滿" else img_rect
                            page.insert_image(rect, stream=img_data, keep_proportion=True)
                        else:
                            page = doc.new_page(width=img_rect.width, height=img_rect.height)
                            page.insert_image(page.rect, stream=img_data)
                        pix = None
                    else:
                        if base_size:
                            tw, th = base_size if target_orient == "直式" else (base_size[1], base_size[0])
                            if ar and ((img_rect.width > img_rect.height) != (tw > th)): tw, th = th, tw
                            page = doc.new_page(width=tw, height=th)
                            rect = page.rect if sm == "自動填滿" else img_rect
                            page.insert_image(rect, filename=path, keep_proportion=True)
                        else:
                            page = doc.new_page(width=img_rect.width, height=img_rect.height)
                            page.insert_image(page.rect, filename=path)
                    img_doc.close()
                else:
                    # PDF 處理：此處使用背景執行緒開啟新的具柄，避免干擾 UI 快取
                    with fitz.open(path) as sub:
                        if sub.is_encrypted: sub.authenticate(self.pdf_passwords.get(path, ""))
                        from_p = item['page'] if item['page'] is not None else 0
                        to_p = item['page'] if item['page'] is not None else len(sub) - 1
                        
                        if base_size:
                            tw, th = base_size if target_orient == "直式" else (base_size[1], base_size[0])
                            for p_no in range(from_p, to_p + 1):
                                sp = sub[p_no]
                                lw, lh = (th, tw) if ar and ((sp.rect.width > sp.rect.height) != (tw > th)) else (tw, th)
                                page = doc.new_page(width=lw, height=lh)
                                rect = page.rect if sm == "自動填滿" else sp.rect
                                page.show_pdf_page(rect, sub, sp.number)
                        else:
                            doc.insert_pdf(sub, from_page=from_p, to_page=to_p)
                self.root.after(0, lambda v=((idx + 1) / total) * 100: self.progress.configure(value=v))

            doc.set_metadata(meta)
            if enc and opw: doc.save(save_path, garbage=4, deflate=True, encryption=fitz.PDF_ENCRYPT_AES_256, user_pw=opw, owner_pw=opw)
            else: doc.save(save_path, garbage=4, deflate=True)
            doc.close()
            self.root.after(0, lambda: self.on_conversion_success(save_path))
        except Exception as e: 
            self.root.after(0, lambda msg=str(e): self.on_conversion_error(msg))

    def on_conversion_success(self, p):
        self.is_converting = False; self.toggle_ui_state(tk.NORMAL); self.status_label.config(text="完成！", fg="green"); messagebox.showinfo("成功", "PDF 已產生")
        if self.auto_open_var.get():
            d = os.path.dirname(os.path.abspath(p))
            if platform.system() == "Windows": os.startfile(d)
            else: webbrowser.open(f"file://{d}")

    def on_conversion_error(self, m):
        self.is_converting = False; self.toggle_ui_state(tk.NORMAL); self.status_label.config(text="失敗", fg="red"); messagebox.showerror("錯誤", f"轉換出錯：\n{m}")

if __name__ == "__main__":
    root = TkinterDnD.Tk(); app = ImageToPdfConverter(root); root.mainloop()