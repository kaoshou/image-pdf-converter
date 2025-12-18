import fitz  # PyMuPDF：用於處理 PDF 的核心函式庫
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinterdnd2 import DND_FILES, TkinterDnD  # 支援拖放檔案功能
import ctypes
import os
import webbrowser
import platform
import threading  # 用於非同步處理轉換，避免介面卡死
import re

# 1. 跨平台動態字體偵測：根據作業系統選擇最適合的黑體字
def get_system_font():
    current_os = platform.system()
    if current_os == "Windows":
        return "Microsoft JhengHei"  # 微軟正黑體
    elif current_os == "Darwin":  # macOS
        return "PingFang TC"        # 蘋方體
    elif current_os == "Linux":
        return "Noto Sans CJK TC"   # Noto Sans
    else:
        return "Arial"

SYSTEM_FONT = get_system_font()

# Windows 高 DPI 支援：確保在 4K 或縮放螢幕下介面清晰不模糊
try:
    if platform.system() == "Windows":
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

# 定義標準頁面尺寸 (單位：Points, 1 point = 1/72 inch)
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

# 自定義帶有 Placeholder 功能的 Entry：實作輸入框提示文字
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
        """當輸入框為空且失去焦點時顯示提示文字"""
        if not self.get():
            self.insert(0, self.placeholder)
            self['fg'] = self.placeholder_color
            if self.is_password:
                self.config(show='')

    def _clear_placeholder(self, event=None):
        """當輸入框獲得焦點時清除提示文字"""
        if self['fg'] == self.placeholder_color:
            self.delete(0, tk.END)
            self['fg'] = self.default_fg_color
            if self.is_password:
                self.config(show=self.real_show)

    def get_real_value(self):
        """獲取真正的輸入值，排除提示文字"""
        if self['fg'] == self.placeholder_color:
            return ""
        return self.get()

# 個別檔案密碼輸入對話框：當匯入加密 PDF 時自動彈出
class FilePasswordDialog(tk.Toplevel):
    def __init__(self, parent, filename):
        super().__init__(parent)
        self.title("PDF 檔案解鎖")
        self.filename = filename
        self.password = None
        
        # 計算居中座標並設定大小
        width, height = 480, 220
        self.root = parent.winfo_toplevel()
        parent_x = self.root.winfo_x()
        parent_y = self.root.winfo_y()
        parent_width = self.root.winfo_width()
        parent_height = self.root.winfo_height()
        pos_x = parent_x + (parent_width // 2) - (width // 2)
        pos_y = parent_y + (parent_height // 2) - (height // 2)
        self.geometry(f"{width}x{height}+{pos_x}+{pos_y}")
        
        self.configure(bg="white")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set() # 鎖定與父視窗的互動，必須處理此視窗

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

        self.bind("<Return>", lambda e: self.on_confirm()) # 支援 Enter 鍵確認

    def on_confirm(self):
        self.password = self.entry.get()
        self.destroy()

# 主應用程式類別
class ImageToPdfConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("圖片轉PDF小工具")
        self.root.geometry("1150x900")
        self.root.configure(bg="#f0f2f5")
        self.root.minsize(900, 890)
        
        self.file_list = []      # 儲存待處理檔案的路徑清單
        self.pdf_passwords = {}  # 儲存匯入 PDF 時的密碼映射 {路徑: 密碼}
        self.is_converting = False # 轉換狀態旗標

        self.setup_styles()
        self.create_widgets()

    def setup_styles(self):
        """設定全域 UI 樣式與色彩"""
        self.primary_color = "#0056b3"
        self.bg_light = "#ffffff"
        self.font_title = (SYSTEM_FONT, 16, "bold")
        self.font_header = (SYSTEM_FONT, 11, "bold")
        self.font_main = (SYSTEM_FONT, 10)
        self.font_status = (SYSTEM_FONT, 9)
        self.font_btn_big = (SYSTEM_FONT, 14, "bold")
        
        style = ttk.Style()
        style.theme_use("clam")
        # 設定 Treeview (表格) 樣式
        style.configure("Treeview", font=self.font_main, rowheight=32, borderwidth=0)
        style.configure("Treeview.Heading", font=self.font_main)
        style.map("Treeview", background=[('selected', '#e1f5fe')], foreground=[('selected', 'black')])
        # 設定進度條樣式
        style.configure("TProgressbar", thickness=14)
        style.configure("TCombobox", font=self.font_main)

    def create_widgets(self):
        """建立視窗所有元件"""
        # --- 頂部導航列 ---
        nav_frame = tk.Frame(self.root, bg="white", height=65)
        nav_frame.pack(fill=tk.X, side=tk.TOP)
        nav_frame.pack_propagate(False)
        
        tk.Frame(nav_frame, bg=self.primary_color, width=5).pack(side=tk.LEFT, fill=tk.Y, padx=(20, 0), pady=12)
        tk.Label(nav_frame, text="圖片轉PDF小工具", font=self.font_title, bg="white", fg="#333").pack(side=tk.LEFT, padx=15)
        
        about_link = tk.Label(nav_frame, text="關於本程式", font=self.font_main, bg="white", fg="#555", cursor="hand2")
        about_link.pack(side=tk.RIGHT, padx=30)
        about_link.bind("<Button-1>", lambda e: self.show_about())

        # --- 主要內容容器 ---
        main_content = tk.Frame(self.root, bg="#f0f2f5", padx=25, pady=5)
        main_content.pack(fill=tk.BOTH, expand=True)

        # 1. 檔案來源區塊
        self.src_section_frame, _ = self.create_section(main_content, "檔案來源")
        self.src_section_frame.master.pack(side=tk.TOP, fill=tk.X, pady=5)
        self.btn_select = tk.Button(self.src_section_frame, text=" ＋ 選擇檔案... ", command=self.add_files, 
                  font=self.font_main, bg="#fafafa", relief=tk.GROOVE, padx=15, pady=4)
        self.btn_select.pack(side=tk.LEFT, padx=15, pady=10)
        tk.Label(self.src_section_frame, text="請加入檔案 (亦可利用拖曳方式加入清單)", 
                 font=self.font_main, bg="white", fg="gray").pack(side=tk.LEFT)

        # 4. 執行作業區塊 (放在最底部，使用 pack 順序控制)
        self.exec_section_frame, _ = self.create_section(main_content, "執行作業")
        self.exec_section_frame.master.pack(side=tk.BOTTOM, fill=tk.X, pady=5)
        exec_inner = tk.Frame(self.exec_section_frame, bg="white", padx=15, pady=10) 
        exec_inner.pack(fill=tk.X)
        
        # 左側控制：自動開啟、狀態提示、進度條
        left_exec_ctrl = tk.Frame(exec_inner, bg="white")
        left_exec_ctrl.pack(side=tk.LEFT, fill=tk.Y)
        self.auto_open_var = tk.BooleanVar(value=False)
        tk.Checkbutton(left_exec_ctrl, text="轉換完成後自動開啟資料夾", variable=self.auto_open_var, font=self.font_main, bg="white").pack(anchor="w")
        self.status_label = tk.Label(left_exec_ctrl, text="等待作業中...", font=self.font_status, bg="white", fg="gray")
        self.status_label.pack(anchor="w")
        self.progress = ttk.Progressbar(left_exec_ctrl, orient=tk.HORIZONTAL, length=320, mode='determinate')
        self.progress.pack(anchor="w", pady=(2, 0))
        
        # 右側主按鈕：開始產生 PDF
        self.btn_run = tk.Button(exec_inner, text="  🚀  開始產生 PDF  ", command=self.start_conversion_thread, 
                                 bg="#096dd9", fg="white", font=self.font_btn_big, relief=tk.FLAT, padx=65, pady=10, cursor="hand2")
        self.btn_run.pack(side=tk.RIGHT, pady=5)

        # 3. 轉換參數與文件資訊區塊 (在執行作業上方)
        self.param_section_frame, _ = self.create_section(main_content, "參數設定與文件資訊 (選填)")
        self.param_section_frame.master.pack(side=tk.BOTTOM, fill=tk.X, pady=5)
        
        grid_container = tk.Frame(self.param_section_frame, bg="white", padx=15, pady=10)
        grid_container.pack(fill=tk.X)
        
        # --- 第一列：圖片效果設定、尺寸與旋轉 ---
        row0 = tk.Frame(grid_container, bg="white")
        row0.pack(fill=tk.X, pady=2)
        
        # 壓縮與品質
        self.compress_var = tk.BooleanVar(value=False)
        self.check_compress = tk.Checkbutton(row0, text="圖片壓縮", variable=self.compress_var, font=self.font_main, bg="white", command=self.toggle_compress)
        self.check_compress.pack(side=tk.LEFT)
        qual_frame = tk.Frame(row0, bg="white")
        qual_frame.pack(side=tk.LEFT, padx=(5, 5))
        tk.Label(qual_frame, text="品質:", font=self.font_main, bg="white").pack(side=tk.LEFT)
        self.quality_val_label = tk.Label(qual_frame, text="80%", font=(SYSTEM_FONT, 9, "bold"), bg="#f0f2f5", width=4)
        self.quality_scale = tk.Scale(qual_frame, from_=10, to=100, orient=tk.HORIZONTAL, length=80, bg="white", highlightthickness=0, showvalue=0, command=self.update_quality_label)
        self.quality_scale.set(80); self.quality_scale.pack(side=tk.LEFT, padx=5); self.quality_val_label.pack(side=tk.LEFT); self.quality_scale.config(state=tk.DISABLED)

        tk.Frame(row0, bg="#eee", width=1).pack(side=tk.LEFT, fill=tk.Y, padx=15)

        # 加密設定
        self.encrypt_var = tk.BooleanVar(value=False)
        self.check_encrypt = tk.Checkbutton(row0, text="PDF 加密", variable=self.encrypt_var, font=self.font_main, bg="white", command=self.toggle_encrypt)
        self.check_encrypt.pack(side=tk.LEFT)
        self.password_entry = PlaceholderEntry(row0, placeholder="設定密碼", is_password=True, font=self.font_main, width=12, relief=tk.SOLID, borderwidth=1)
        self.password_entry.pack(side=tk.LEFT, padx=5); self.password_entry.config(state=tk.DISABLED)

        tk.Frame(row0, bg="#eee", width=1).pack(side=tk.LEFT, fill=tk.Y, padx=15)

        # 旋轉與黑白處理
        self.auto_rotate_var = tk.BooleanVar(value=True)
        self.check_auto_rotate = tk.Checkbutton(row0, text="自動旋轉", variable=self.auto_rotate_var, font=self.font_main, bg="white")
        self.check_auto_rotate.pack(side=tk.LEFT)
        self.grayscale_var = tk.BooleanVar(value=False)
        self.check_grayscale = tk.Checkbutton(row0, text="黑白模式", variable=self.grayscale_var, font=self.font_main, bg="white")
        self.check_grayscale.pack(side=tk.LEFT, padx=5)

        # --- 第二列：頁面尺寸與佈局設定 ---
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

        # --- 第三列：整合的文件中繼資料 (Metadata) ---
        row2 = tk.Frame(grid_container, bg="white")
        row2.pack(fill=tk.X, pady=2)
        
        meta_items = [("標題:", "meta_title", "文件標題", 10), ("作者:", "meta_author", "作者名稱", 8), 
                      ("主題:", "meta_subject", "主題內容", 10), ("關鍵字:", "meta_keywords", "逗號分隔", 10)]
        
        for lbl_text, attr_name, ph, w in meta_items:
            tk.Label(row2, text=lbl_text, font=self.font_main, bg="white").pack(side=tk.LEFT, padx=(5, 2))
            entry = PlaceholderEntry(row2, placeholder=ph, font=self.font_main, width=w, relief=tk.SOLID, borderwidth=1)
            entry.pack(side=tk.LEFT, padx=(0, 10), fill=tk.X, expand=True)
            setattr(self, attr_name, entry)

        # 2. 待處理清單區塊 (在中間區域，會隨視窗高度自動延展)
        self.list_section_frame, list_title_bar = self.create_section(main_content, "待處理清單")
        self.list_section_frame.master.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=5)
        self.file_count_label = tk.Label(list_title_bar, text="已選擇: 0 個檔案", font=(SYSTEM_FONT, 9, "bold"), bg="#fafafa", fg=self.primary_color)
        self.file_count_label.pack(side=tk.LEFT, padx=(10, 0))
        
        list_main_container = tk.Frame(self.list_section_frame, bg="white", padx=15, pady=5)
        list_main_container.pack(fill=tk.BOTH, expand=True)
        tree_frame = tk.Frame(list_main_container, bg="white")
        tree_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 表格欄位：頁碼範圍、格式、檔案名稱
        self.tree = ttk.Treeview(tree_frame, columns=("PageRange", "Type", "Name"), show='headings', selectmode='extended')
        self.tree.heading("PageRange", text="頁碼範圍"); self.tree.heading("Type", text="格式"); self.tree.heading("Name", text="檔案名稱")
        self.tree.column("PageRange", width=120, anchor="center"); self.tree.column("Type", width=80, anchor="center"); self.tree.column("Name", width=400)
        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview); self.tree.configure(yscroll=scrollbar.set); self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True); scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 鍵盤熱鍵：Delete 鍵移除選取
        self.tree.bind("<Delete>", lambda e: self.remove_selected())

        # 右側控制按鈕
        self.side_btn_bar = tk.Frame(list_main_container, bg="white", padx=10)
        self.side_btn_bar.pack(side=tk.RIGHT, fill=tk.Y)        
        self.btn_up = tk.Button(self.side_btn_bar, text="▲ 上移", command=self.move_up, relief=tk.GROOVE, font=self.font_main, bg="#f8f9fa", width=12); self.btn_up.pack(pady=2)
        self.btn_down = tk.Button(self.side_btn_bar, text="▼ 下移", command=self.move_down, relief=tk.GROOVE, font=self.font_main, bg="#f8f9fa", width=12); self.btn_down.pack(pady=2)
        tk.Label(self.side_btn_bar, text="自動排序", font=self.font_status, bg="white", fg="gray").pack(pady=(5, 2))
        self.btn_sort_asc = tk.Button(self.side_btn_bar, text="A-Z 排序", command=lambda: self.sort_files(False), relief=tk.GROOVE, font=self.font_main, bg="#f8f9fa", width=12); self.btn_sort_asc.pack(pady=1)
        self.btn_sort_desc = tk.Button(self.side_btn_bar, text="Z-A 排序", command=lambda: self.sort_files(True), relief=tk.GROOVE, font=self.font_main, bg="#f8f9fa", width=12); self.btn_sort_desc.pack(pady=1)
        tk.Frame(self.side_btn_bar, bg="#eee", height=1).pack(fill=tk.X, pady=8)
        self.btn_remove = tk.Button(self.side_btn_bar, text="✕ 移除選取", command=self.remove_selected, bg="#fff1f0", fg="#cf1322", relief=tk.GROOVE, font=self.font_main, width=12); self.btn_remove.pack(pady=1)
        self.btn_clear = tk.Button(self.side_btn_bar, text="🗑 全部清空", command=self.clear_all, relief=tk.GROOVE, font=self.font_main, width=12); self.btn_clear.pack(pady=1)
        
        # 拖放功能註冊
        self.tree.drop_target_register(DND_FILES); self.tree.dnd_bind('<<Drop>>', self.handle_drop)

    def create_section(self, parent, title):
        """建立具備標題列與內容區的區塊容器"""
        container = tk.Frame(parent, bg="white", bd=1, relief="solid", highlightthickness=0)
        title_bar = tk.Frame(container, bg="#fafafa"); title_bar.pack(fill=tk.X)
        tk.Frame(title_bar, bg=self.primary_color, width=3).pack(side=tk.LEFT, fill=tk.Y, padx=(12, 6), pady=6)
        tk.Label(title_bar, text=title, font=self.font_header, bg="#fafafa", fg="#333").pack(side=tk.LEFT, pady=6)
        content = tk.Frame(container, bg="white"); content.pack(fill=tk.BOTH, expand=True)
        return content, title_bar

    def update_quality_label(self, val): 
        """當壓縮品質滑桿移動時更新文字百分比"""
        self.quality_val_label.config(text=f"{val}%")

    def update_tree_content(self):
        """重新整理待處理清單，並計算累計頁碼範圍"""
        for i in self.tree.get_children(): self.tree.delete(i)
        current_page = 1
        for idx, file_path in enumerate(self.file_list):
            fname = os.path.basename(file_path); ext = fname.split('.')[-1].upper()
            pages_in_file = 1
            if ext == "PDF":
                try:
                    with fitz.open(file_path) as tmp:
                        if tmp.is_encrypted:
                            pw = self.pdf_passwords.get(file_path, "")
                            if tmp.authenticate(pw): pages_in_file = len(tmp)
                            else: pages_in_file = 0 # 密碼失效
                        else: pages_in_file = len(tmp)
                except: pages_in_file = 0
            
            if pages_in_file > 0:
                p_range = f"{current_page} ~ {current_page + pages_in_file - 1}" if pages_in_file > 1 else str(current_page)
                current_page += pages_in_file
            else: p_range = "無法讀取"
            self.tree.insert("", tk.END, values=(p_range, ext, fname))
        self.file_count_label.config(text=f"已選擇: {len(self.file_list)} 個檔案")

    def show_about(self):
        """顯示關於本程式視窗，並置中顯示"""
        about_win = tk.Toplevel(self.root); about_win.title("關於本程式")
        w, h = 650, 590; self.root.update_idletasks()
        px, py = self.root.winfo_x(), self.root.winfo_y()
        pw, ph = self.root.winfo_width(), self.root.winfo_height()
        about_win.geometry(f"{w}x{h}+{px + (pw // 2) - (w // 2)}+{py + (ph // 2) - (h // 2)}")
        about_win.configure(bg="white"); about_win.transient(self.root)
        
        content = tk.Frame(about_win, bg="white", padx=40, pady=25)
        content.pack(fill=tk.BOTH, expand=True)
        
        tk.Label(content, text="圖片轉PDF小工具", font=self.font_title, bg="white", fg=self.primary_color).pack(anchor="w")
        
        # 開發者資訊
        dev_info_frame = tk.Frame(content, bg="white")
        dev_info_frame.pack(fill=tk.X, pady=(15, 0))
        tk.Label(dev_info_frame, text="開發者：鄭郁翰 (Cheng, Yu-Han)", font=self.font_main, bg="white").pack(anchor="w")
        tk.Label(dev_info_frame, text="Email：kaoshou@gmail.com", font=self.font_main, bg="white").pack(anchor="w")
        
        tk.Frame(content, bg="#eee", height=1).pack(fill=tk.X, pady=20)
        
        # 項目資訊與開源授權聲明
        tk.Label(content, text="專案資訊與開源聲明 (Open Source Disclosure)", font=self.font_header, bg="white").pack(anchor="w", pady=(0, 10))
        
        license_desc = (
            "本程式原始碼 (GitHub)：\nhttps://github.com/kaoshou/image-pdf-converter\n\n"
            "本程式核心功能基於以下的開源專案實作：\n\n"
            "• PyMuPDF (fitz)：採用 GNU AGPL v3.0 授權，負責所有 PDF 頁面建立、合併、加密、縮放及中繼資料寫入之核心邏輯。\n"
            "  網址: https://github.com/pymupdf/PyMuPDF\n\n"
            "• TkinterDnD2：採用 MIT 授權，提供跨平台之檔案拖曳匯入介面支援。\n"
            "  網址: https://github.com/pmgagne/tkinterdnd2\n\n"
            "• Python Standard Library：採用 PSF License 授權。\n"
            "  網址: https://www.python.org/\n\n"
            "免責聲明：本軟體依「現狀」提供，不附帶任何形式的明示或暗示保證。開發者對於因使用本程式所產生的任何直接或間接損失概不負責。"
        )
        
        text_box = tk.Text(content, height=30, font=("Consolas", 9), bg="#f9f9f9", relief=tk.FLAT, wrap=tk.WORD, padx=12, pady=12)
        text_box.insert(tk.END, license_desc)
        text_box.config(state=tk.DISABLED) # 設定為唯讀
        text_box.pack(fill=tk.X)

    def handle_drop(self, event):
        """處理拖曳檔案進入視窗的事件"""
        if not self.is_converting: self.process_incoming_files(self.root.tk.splitlist(event.data))

    def add_files(self):
        """彈出檔案選擇器增加檔案"""
        if not self.is_converting:
            files = filedialog.askopenfilenames(title="選擇檔案", filetypes=[("支援格式", "*.jpg *.jpeg *.png *.pdf *.bmp *.tiff")])
            if files: self.process_incoming_files(files)

    def process_incoming_files(self, files):
        """過濾有效檔案格式，並處理加密 PDF 的密碼輸入"""
        valid = ('.jpg', '.jpeg', '.png', '.pdf', '.bmp', '.tiff'); added = False
        for f in files:
            if f.lower().endswith(valid) and f not in self.file_list:
                if f.lower().endswith('.pdf'):
                    try:
                        with fitz.open(f) as tmp:
                            if tmp.is_encrypted:
                                correct = False
                                while not correct:
                                    dialog = FilePasswordDialog(self.root, os.path.basename(f))
                                    self.root.wait_window(dialog)
                                    if dialog.password is None: break # 放棄此檔案
                                    if tmp.authenticate(dialog.password): self.pdf_passwords[f] = dialog.password; correct = True
                                    else: messagebox.showerror("錯誤", "密碼不正確")
                                if not correct: continue
                    except: pass
                self.file_list.append(f); added = True
        if added: self.update_tree_content()

    def sort_files(self, rev):
        """依檔案名稱進行排序 (A-Z 或 Z-A)"""
        if not self.is_converting: self.file_list.sort(key=lambda x: os.path.basename(x).lower(), reverse=rev); self.update_tree_content()

    def move_up(self):
        """將選取的項目在清單中上移"""
        if self.is_converting: return
        sel = self.tree.selection(); idxs = sorted([self.tree.index(i) for i in sel])
        if not idxs or idxs[0] <= 0: return
        for idx in idxs: self.file_list[idx], self.file_list[idx-1] = self.file_list[idx-1], self.file_list[idx]
        self.update_tree_content()
        for idx in idxs: self.tree.selection_add(self.tree.get_children()[idx-1])
    
    def move_down(self):
        """將選取的項目在清單中下移"""
        if self.is_converting: return
        sel = self.tree.selection(); idxs = sorted([self.tree.index(i) for i in sel], reverse=True)
        if not idxs or idxs[0] >= len(self.file_list) - 1: return
        for idx in idxs: self.file_list[idx], self.file_list[idx+1] = self.file_list[idx+1], self.file_list[idx]
        self.update_tree_content()
        for idx in idxs: self.tree.selection_add(self.tree.get_children()[idx+1])

    def remove_selected(self):
        """從清單中移除所選項目"""
        if not self.is_converting:
            sel = self.tree.selection(); idxs = sorted([self.tree.index(i) for i in sel], reverse=True)
            for idx in idxs: p = self.file_list.pop(idx); self.pdf_passwords.pop(p, None)
            self.update_tree_content()
            
    def clear_all(self):
        """清空所有待處理檔案"""
        if not self.is_converting and self.file_list and messagebox.askyesno("確認", "是否清空？"):
            self.file_list.clear(); self.pdf_passwords.clear(); self.update_tree_content()

    def toggle_compress(self): 
        """切換壓縮功能的啟用狀態"""
        s = tk.NORMAL if self.compress_var.get() else tk.DISABLED
        self.quality_scale.config(state=s); self.quality_val_label.config(fg="black" if self.compress_var.get() else "gray")

    def toggle_encrypt(self): 
        """切換 PDF 加密輸出的啟用狀態"""
        self.password_entry.config(state=tk.NORMAL if self.encrypt_var.get() else tk.DISABLED)

    def start_conversion_thread(self):
        """建立非同步執行緒開始執行 PDF 轉換程序"""
        if not self.file_list: return
        opw = self.password_entry.get_real_value()
        if self.encrypt_var.get() and not opw: messagebox.showwarning("警告", "請設定密碼"); return
        save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF 檔案", "*.pdf")])
        if not save_path: return
        self.is_converting = True; self.toggle_ui_state(tk.DISABLED)
        self.progress['value'] = 0; self.status_label.config(text="準備開始轉換...", fg="blue")
        # 啟動 Thread 以免 GUI 凍結
        threading.Thread(target=self.perform_conversion, args=(save_path,), daemon=True).start()

    def toggle_ui_state(self, state):
        """鎖定或解鎖所有功能按鈕 (轉換期間鎖定)"""
        btns = [self.btn_run, self.btn_select, self.btn_up, self.btn_down, self.btn_remove, self.btn_clear, self.btn_sort_asc, self.btn_sort_desc, self.check_compress, self.check_encrypt, self.combo_size, self.combo_orient, self.check_auto_rotate, self.check_grayscale, self.combo_scale]
        for b in btns: b.config(state=state)
        if state == tk.NORMAL: self.toggle_compress(); self.toggle_encrypt()
        else: self.quality_scale.config(state=tk.DISABLED); self.password_entry.config(state=tk.DISABLED)

    def perform_conversion(self, save_path):
        """後台轉換邏輯：優化圖片處理，避免無謂的重新編碼導致畫質下降"""
        doc = fitz.open(); total = len(self.file_list)
        c, q = self.compress_var.get(), self.quality_scale.get()
        enc, opw = self.encrypt_var.get(), self.password_entry.get_real_value()
        gs = self.grayscale_var.get(); ar = self.auto_rotate_var.get()
        sm = self.scale_mode_var.get()
        # 設定中繼資料
        meta = {"title": self.meta_title.get_real_value(), "author": self.meta_author.get_real_value(), "subject": self.meta_subject.get_real_value(), "keywords": self.meta_keywords.get_real_value(), "creator": "圖片轉PDF小工具", "producer": "PyMuPDF"}
        base_size = PAGE_SIZES.get(self.page_size_var.get()); target_orient = self.orientation_var.get()

        try:
            for idx, path in enumerate(self.file_list):
                self.root.after(0, lambda i=idx+1: self.status_label.config(text=f"處理中 {i}/{total}..."))
                
                if not path.lower().endswith('.pdf'):
                    # --- 核心畫質優化判斷 ---
                    # 當「不壓縮」且「不轉黑白」時，直接嵌入原始檔案路徑
                    if not gs and not c:
                        if base_size:
                            # 方案 A: 固定頁面尺寸，直接嵌入原始圖片數據
                            w, h = base_size if target_orient == "直式" else (base_size[1], base_size[0])
                            img_info = fitz.open(path)
                            item = img_info[0]
                            if ar: # 自動旋轉邏輯：根據原圖比例與目標頁面比例決定是否交換寬高
                                if (item.rect.width > item.rect.height and w < h) or (item.rect.width < item.rect.height and w > h):
                                    w, h = h, w
                            page = doc.new_page(width=w, height=h)
                            rect = page.rect if sm == "自動填滿" else item.rect
                            # 關鍵：使用 filename=path 直接引用原始檔案，不經過重新渲染位圖
                            page.insert_image(rect, filename=path, keep_proportion=True)
                            img_info.close()
                        else:
                            # 方案 B: 原始大小，使用無損封裝方式
                            img_temp = fitz.open(path)
                            pb = img_temp.convert_to_pdf() # 將圖片數據直接包裝成單頁 PDF
                            doc.insert_pdf(fitz.open("pdf", pb))
                            img_temp.close()
                    else:
                        # 方案 C: 需要處理（黑白或壓縮），此時才進行 Pixmap 渲染
                        img = fitz.open(path); pix = img[0].get_pixmap()
                        if gs: pix = fitz.Pixmap(fitz.csGRAY, pix) # 灰階處理
                        img_data = pix.tobytes("jpg", jpg_quality=q) if c else pix.tobytes("png")
                        
                        if base_size:
                            w, h = base_size if target_orient == "直式" else (base_size[1], base_size[0])
                            if ar: 
                                if (pix.width > pix.height and w < h) or (pix.width < pix.height and w > h): w, h = h, w
                            page = doc.new_page(width=w, height=h)
                            rect = page.rect if sm == "自動填滿" else pix.irect
                            page.insert_image(rect, stream=img_data, keep_proportion=True)
                        else:
                            pb = fitz.open("jpg" if c else "png", img_data).convert_to_pdf()
                            doc.insert_pdf(fitz.open("pdf", pb))
                        img.close()
                else:
                    # 處理 PDF 合併 (維持原邏輯)
                    with fitz.open(path) as sub:
                        if sub.is_encrypted: sub.authenticate(self.pdf_passwords.get(path, ""))
                        if base_size:
                            w, h = base_size if target_orient == "直式" else (base_size[1], base_size[0])
                            for sp in sub: 
                                if ar: 
                                    if (sp.rect.width > sp.rect.height and w < h) or (sp.rect.width < sp.rect.height and w > h): lw, lh = h, w
                                    else: lw, lh = w, h
                                else: lw, lh = w, h
                                page = doc.new_page(width=lw, height=lh)
                                rect = page.rect if sm == "自動填滿" else sp.rect
                                page.show_pdf_page(rect, sub, sp.number)
                        else: doc.insert_pdf(sub)
                
                self.root.after(0, lambda v=((idx + 1) / total) * 100: self.progress.configure(value=v))

            self.root.after(0, lambda: self.status_label.config(text="寫入資訊中...", fg="green"))
            doc.set_metadata(meta)
            # 儲存檔案
            if enc and opw: 
                doc.save(save_path, garbage=4, deflate=True, encryption=fitz.PDF_ENCRYPT_AES_256, user_pw=opw, owner_pw=opw)
            else: 
                doc.save(save_path, garbage=4, deflate=True)
            doc.close()
            self.root.after(0, lambda: self.on_conversion_success(save_path))
        except Exception as e: 
            self.root.after(0, lambda msg=str(e): self.on_conversion_error(msg))

    def on_conversion_success(self, p):
        """轉換成功後的回傳與通知"""
        self.is_converting = False; self.toggle_ui_state(tk.NORMAL); self.status_label.config(text="完成！", fg="green"); messagebox.showinfo("成功", "PDF 已產生")
        if self.auto_open_var.get():
            d = os.path.dirname(os.path.abspath(p))
            if platform.system() == "Windows": os.startfile(d)
            else: webbrowser.open(f"file://{d}")

    def on_conversion_error(self, m):
        """轉換失敗後的回傳與錯誤提示"""
        self.is_converting = False; self.toggle_ui_state(tk.NORMAL); self.status_label.config(text="失敗", fg="red"); messagebox.showerror("錯誤", f"轉換出錯：\n{m}")

if __name__ == "__main__":
    # 使用具備拖放功能的 Tk 實體
    root = TkinterDnD.Tk(); app = ImageToPdfConverter(root); root.mainloop()