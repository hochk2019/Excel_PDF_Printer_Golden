#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Golden Logistics Tools – Print & PDF (PRO4-fix)
- Fix: Scope "Tất cả sheet" sẽ in/xuất ĐÚNG các trang bạn nhập (vd: 2 hoặc 1,3,5-7) cho MỖI sheet.
- Giao diện tab, chọn trang "1,3,5-7", Excel/Word/PDF, In hoặc Xuất PDF (Gộp/Tách)
- Lọc tên, ngày sửa; tiến trình %, tạm dừng, hủy; log CSV; tự lưu config.json; tự tạo README.
"""

import os, sys, csv, time, json, threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from ui_builder import build_source_panel, build_settings_panel
from app_utils import (
    EXCEL_EXTS,
    WORD_EXTS,
    PDF_EXTS,
    parse_date,
    parse_pagespec,
    gather_files,
)
from print_handlers import (
    excel_print_first_sheet_ranges,
    excel_print_all_sheets_ranges,
    excel_export_first_sheet,
    excel_export_all_sheets_pages,
    word_print_ranges,
    word_export_pdf,
    pdf_print_ranges,
    pdf_export,
)
from app_actions import (
    open_output_folder,
    apply_file_filter,
    refresh_preset_list,
    on_preset_select,
    collect_state,
    apply_state,
    save_preset,
    delete_preset,
    apply_selected_preset,
    clear_folders,
)
APP_TITLE = "Excel & PDF Print V4.1- Golden Logistics"
CONFIG_FILE = "config.json"
README_FILE = "README_HDSD.txt"
LOGO_PATH_CANDIDATES = ["Logo moi 2- size bé.png","Logo moi 1.png","Logo cty.png","logo.png","logo_cty.png"]
ICON_PATH_CANDIDATES = ["app_icon_v41_modern.ico","app_icon_v41.ico"]
UI_COLORS = {
    "bg": "#F8FAFC",
    "panel": "#FFFFFF",
    "text": "#1E293B",
    "muted": "#475569",
    "primary": "#3B82F6",
    "primary_dark": "#1D4ED8",
    "accent": "#F97316",
    "border": "#E2E8F0",
    "success": "#22C55E",
    "danger": "#EF4444",
    "warning": "#F59E0B",
}

def ensure_readme_exists():
    if os.path.exists(README_FILE):
        return
    content = """\ufeff
==============================
HƯỚNG DẪN SỬ DỤNG (PRO4-fix)
Phiên bản: V4.1
Designer: Hoc HK
==============================
• Mới: Khi chọn "Tất cả sheet", công cụ sẽ in/xuất PDF **đúng chuỗi trang bạn nhập** trên **từng sheet**.
• VD: nhập "2" → in trang 2 của mỗi sheet; "1,3,5-7" → in tương ứng trên mỗi sheet.
"""
    with open(README_FILE, "w", encoding="utf-8-sig") as f:
        f.write(content.strip()+"\n")

def load_config():
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8-sig") as f:
            return json.load(f)
    except Exception:
        return {}

def save_config(cfg):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8-sig") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def try_imports():
    ok = True; err = ""
    try:
        import win32com.client  # noqa
        import win32print       # noqa
        import win32api         # noqa
    except Exception as e:
        ok = False; err = str(e)
    return ok, err

def get_printers():
    import win32print
    flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    names = []
    try:
        for p in win32print.EnumPrinters(flags):
            name = p[2] if isinstance(p, (list, tuple)) and len(p) > 2 else None
            if name and name not in names:
                names.append(name)
    except Exception:
        pass
    try:
        default = win32print.GetDefaultPrinter()
    except Exception:
        default = names[0] if names else ""
    return names, default

# ---------------- GUI ----------------
class App(tk.Tk):
    open_output_folder = open_output_folder
    apply_file_filter = apply_file_filter
    refresh_preset_list = refresh_preset_list
    on_preset_select = on_preset_select
    collect_state = collect_state
    apply_state = apply_state
    save_preset = save_preset
    delete_preset = delete_preset
    apply_selected_preset = apply_selected_preset
    clear_folders = clear_folders
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1060x780")
        self.minsize(980, 680)

        cfg = load_config()
        self.folders = cfg.get("folders", [])
        self.recursive = tk.BooleanVar(value=cfg.get("recursive", True))
        self.pattern = tk.StringVar(value=cfg.get("pattern", "*"))
        self.date_from = tk.StringVar(value=cfg.get("date_from", ""))
        self.date_to   = tk.StringVar(value=cfg.get("date_to", ""))

        self.use_excel = tk.BooleanVar(value=cfg.get("use_excel", True))
        self.use_word  = tk.BooleanVar(value=cfg.get("use_word", False))
        self.use_pdf   = tk.BooleanVar(value=cfg.get("use_pdf", False))

        self.page_spec = tk.StringVar(value=cfg.get("page_spec", "1"))
        self.mode_pdf_export = tk.BooleanVar(value=cfg.get("mode_pdf_export", False))
        self.pdf_merge = tk.BooleanVar(value=cfg.get("pdf_merge", True))
        self.sheet_scope = tk.StringVar(value=cfg.get("sheet_scope", "first"))  # "first" or "all"
        self.paper_size = tk.StringVar(value=cfg.get("paper_size", "A4"))       # "A4" or "A5"
        self.out_dir = tk.StringVar(value=cfg.get("out_dir", ""))

        self.enable_logging = tk.BooleanVar(value=cfg.get("enable_logging", True))
        self.log_dir = tk.StringVar(value=cfg.get("log_dir", ""))
        self.search_var = tk.StringVar(value=cfg.get("search_term", ""))
        self.sort_options = ["Tên A→Z", "Tên Z→A", "Ngày sửa mới→cũ", "Ngày sửa cũ→mới"]
        self.sort_var = tk.StringVar(value=cfg.get("sort_mode", self.sort_options[0]))
        self.presets = cfg.get("presets", {})
        self.preset_name = tk.StringVar(value="")
        self.preset_input = tk.StringVar(value="")

        self.printers, self.default_printer = get_printers()
        sel = cfg.get("selected_printer") or (self.default_printer if self.default_printer in self.printers else (self.printers[0] if self.printers else ""))
        self.selected_printer = tk.StringVar(value=sel)

        self.is_running=False; self.pause_flag=threading.Event(); self.cancel_flag=threading.Event()
        self.total_files=0; self.done_files=0
        self.current_files=[]
        self.filtered_files=[]
        self.failed_files=[]
        self.last_run_files=[]

        self.setup_styles()
        self.build_ui()
        self.refresh_preset_list()
        self.bind_updates()
        self.update_summary()
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        ensure_readme_exists()

    def resource_path(self, relative_path):
        """ Get absolute path to resource, works for dev and for PyInstaller """
        try:
            # PyInstaller creates a temp folder and stores path in _MEIPASS
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")

        return os.path.join(base_path, relative_path)

    def setup_styles(self):
        self.colors = UI_COLORS
        self.font_title = ("Segoe UI", 16, "bold")
        self.font_subtitle = ("Segoe UI", 10)
        self.font_section = ("Segoe UI", 11, "bold")
        self.font_body = ("Segoe UI", 10)
        self.option_add("*Font", "{Segoe UI} 10")
        self.configure(bg=self.colors["bg"])
        self.set_window_icon()

        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("TCombobox", fieldbackground="white", background="white")
        style.configure("TProgressbar", troughcolor=self.colors["border"], background=self.colors["primary"])

        self.btn_primary = {
            "bg": self.colors["primary"],
            "fg": "white",
            "activebackground": self.colors["primary_dark"],
            "activeforeground": "white",
            "relief": "flat",
            "bd": 0,
            "highlightthickness": 2,
            "highlightbackground": "#2563EB",
            "highlightcolor": self.colors["primary_dark"],
            "cursor": "hand2",
        }
        self.btn_secondary = {
            "bg": self.colors["border"],
            "fg": self.colors["text"],
            "activebackground": "#CBD5E1",
            "activeforeground": self.colors["text"],
            "relief": "flat",
            "bd": 0,
            "highlightthickness": 2,
            "highlightbackground": "#94A3B8",
            "highlightcolor": "#94A3B8",
            "cursor": "hand2",
        }
        self.btn_danger = {
            "bg": self.colors["danger"],
            "fg": "white",
            "activebackground": "#DC2626",
            "activeforeground": "white",
            "relief": "flat",
            "bd": 0,
            "highlightthickness": 2,
            "highlightbackground": "#DC2626",
            "highlightcolor": "#B91C1C",
            "cursor": "hand2",
        }

    def bind_updates(self):
        for v in (
            self.page_spec,
            self.mode_pdf_export,
            self.selected_printer,
            self.paper_size,
            self.out_dir,
            self.search_var,
            self.sort_var,
        ):
            v.trace_add("write", lambda *_: self.update_summary())
        if hasattr(self, "entry_pages"):
            self.entry_pages.bind("<KeyRelease>", lambda _e: self.update_summary())
        if hasattr(self, "cmb"):
            self.cmb.bind("<<ComboboxSelected>>", lambda _e: self.update_summary())
        if hasattr(self, "search_entry"):
            self.search_entry.bind("<KeyRelease>", lambda _e: self.apply_file_filter())
        if hasattr(self, "sort_combo"):
            self.sort_combo.bind("<<ComboboxSelected>>", lambda _e: self.apply_file_filter())
        if hasattr(self, "preset_combo"):
            self.preset_combo.bind("<<ComboboxSelected>>", lambda _e: self.on_preset_select())

    def _section(self, parent, title, expand=False):
        frame = tk.LabelFrame(
            parent,
            text=title,
            bg=self.colors["panel"],
            fg=self.colors["text"],
            font=self.font_section,
            padx=8,
            pady=6,
            bd=1,
            relief="solid",
        )
        frame.pack(fill="both" if expand else "x", expand=expand, padx=8, pady=6)
        return frame

    def update_summary(self):
        total = len(self.current_files)
        filtered = len(self.filtered_files)
        mode = "Xuất PDF" if self.mode_pdf_export.get() else "In"
        target = self.out_dir.get().strip() if self.mode_pdf_export.get() else self.selected_printer.get().strip()
        target = target or "Chưa chọn"
        pages = self.page_spec.get().strip() or "1"
        text = f"{filtered}/{total} file | {mode} | Giấy {self.paper_size.get()} | Trang {pages} | Đích: {target}"
        if hasattr(self, "summary_var"):
            self.summary_var.set(text)

    def toggle_logging_fields(self):
        state = "normal" if self.enable_logging.get() else "disabled"
        try:
            self.log_entry.config(state=state)
            self.btn_log_browse.config(state=state)
            self.btn_log_open.config(state=state)
        except Exception:
            pass

    def build_ui(self):
        header = tk.Frame(self, bg=self.colors["bg"])
        header.pack(fill="x", padx=12, pady=(12, 6))

        self.logo_img = None
        logo_path = None
        for cand in LOGO_PATH_CANDIDATES:
            p = self.resource_path(cand)
            if os.path.exists(p):
                logo_path = p
                break
        if not logo_path:
            for cand in LOGO_PATH_CANDIDATES:
                p = os.path.join(os.getcwd(), cand)
                if os.path.exists(p):
                    logo_path = p
                    break

        if logo_path:
            try:
                orig = tk.PhotoImage(file=logo_path)
                h = orig.height()
                factor = max(1, h // 60)
                self.logo_img = orig.subsample(factor, factor)
                tk.Label(header, image=self.logo_img, bg=self.colors["bg"]).pack(side="left", padx=8, pady=6)
            except Exception:
                pass

        title_box = tk.Frame(header, bg=self.colors["bg"])
        title_box.pack(side="left", fill="x", expand=True)
        tk.Label(
            title_box,
            text=APP_TITLE,
            font=self.font_title,
            bg=self.colors["bg"],
            fg=self.colors["text"],
        ).pack(anchor="w")
        tk.Label(
            title_box,
            text="In nhanh theo trang, phù hợp vận hành kho và xử lý loạt file.",
            font=self.font_subtitle,
            bg=self.colors["bg"],
            fg=self.colors["muted"],
        ).pack(anchor="w")
        meta_box = tk.Frame(header, bg=self.colors["bg"])
        meta_box.pack(side="right", padx=(8, 0))
        tk.Label(
            meta_box,
            text="V4.1 • Designer: Hoc HK",
            font=("Segoe UI", 9, "bold"),
            bg=self.colors["bg"],
            fg=self.colors["primary"],
        ).pack(anchor="e")

        summary = tk.Frame(self, bg=self.colors["panel"], bd=1, relief="solid")
        summary.pack(fill="x", padx=12, pady=(0, 8))
        self.summary_var = tk.StringVar()
        tk.Label(
            summary,
            textvariable=self.summary_var,
            bg=self.colors["panel"],
            fg=self.colors["text"],
            font=self.font_body,
        ).pack(anchor="w", padx=10, pady=6)

        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=12)

        paned = tk.PanedWindow(self, orient="horizontal", sashrelief="raised", sashwidth=4, bg=self.colors["bg"])
        paned.pack(fill="both", expand=True, padx=8, pady=8)

        frame_left = tk.Frame(paned, bg=self.colors["bg"])
        paned.add(frame_left, width=520)
        build_source_panel(self, frame_left)

        frame_right = tk.Frame(paned, bg=self.colors["bg"])
        paned.add(frame_right)
        build_settings_panel(self, frame_right)

        # Footer removed to prevent overlap on shorter windows.

    def set_window_icon(self):
        for cand in ICON_PATH_CANDIDATES:
            p = self.resource_path(cand)
            if os.path.exists(p):
                try:
                    self.iconbitmap(p)
                    return
                except Exception:
                    pass
        for cand in ICON_PATH_CANDIDATES:
            p = os.path.join(os.getcwd(), cand)
            if os.path.exists(p):
                try:
                    self.iconbitmap(p)
                    return
                except Exception:
                    pass

    # Helpers
    def open_log_folder(self):
        d = self.log_dir.get().strip() or os.getcwd()
        if os.path.isdir(d):
            os.startfile(d)
        else:
            messagebox.showinfo("Thông tin", "Thư mục log chưa tồn tại hoặc chưa được tạo.")

    def toggle_pdf_dir(self):
        pdf_mode = bool(self.mode_pdf_export.get())
        state_pdf = "normal" if pdf_mode else "disabled"
        state_printer = "disabled" if pdf_mode else "readonly"
        try:
            self.entry_out.config(state=state_pdf)
            self.btn_browse.config(state=state_pdf)
            self.chk_pdf_merge.config(state=state_pdf)
            self.btn_open_output.config(state=state_pdf)
        except Exception:
            pass
        try:
            self.cmb.config(state=state_printer)
            self.btn_refresh.config(state=("disabled" if pdf_mode else "normal"))
        except Exception:
            pass
        self.update_summary()

    def pick_out_dir(self):
        p=filedialog.askdirectory(title="Chọn thư mục lưu PDF")
        if p: self.out_dir.set(p)
    
    def pick_log_dir(self):
        p=filedialog.askdirectory(title="Chọn thư mục log CSV")
        if p: self.log_dir.set(p)
    def add_folder(self):
        p=filedialog.askdirectory(title="Chọn thư mục chứa file")
        if p and p not in self.folders:
            self.folders.append(p); self.listbox.insert("end", p)
            self.scan_files() # Auto scan
    def remove_selected(self):
        sel=list(self.listbox.curselection()); sel.reverse()
        for i in sel:
            p=self.listbox.get(i); 
            if p in self.folders: self.folders.remove(p)
            self.listbox.delete(i)
        self.scan_files() # Auto refresh
    def refresh_printers(self):
        self.printers, self.default_printer = get_printers()
        self.cmb['values']=self.printers
        cur=self.selected_printer.get()
        if cur not in self.printers:
            self.selected_printer.set(self.default_printer if self.default_printer in self.printers else (self.printers[0] if self.printers else ""))
        self.log("Đã làm mới danh sách máy in.")
        self.update_summary()
    
    def scan_files(self):
        dfrom=parse_date(self.date_from.get()); dto=parse_date(self.date_to.get())
        files = gather_files(self.folders, self.recursive.get(), self.pattern.get(),
                             self.use_excel.get(), self.use_word.get(), self.use_pdf.get(),
                             dfrom, dto)
        self.current_files = files
        self.apply_file_filter()
        return self.filtered_files

    def log(self, msg):
        line = f"{time.strftime('%Y-%m-%d %H:%M:%S')}  {msg}\n"
        self.txt.insert("end", line)
        if "[ERR]" in msg or "Lỗi" in msg:
            self.txt.tag_add("err", "end-2l", "end-1l")
        elif "[WARN]" in msg or "Cảnh báo" in msg:
            self.txt.tag_add("warn", "end-2l", "end-1l")
        elif "[OK]" in msg:
            self.txt.tag_add("ok", "end-2l", "end-1l")
        self.txt.see("end")
        self.txt.see("end"); self.update_idletasks()
    def update_progress(self):
        if self.total_files<=0:
            self.progress['value']=0; self.progress_label.config(text="0% (0/0)"); return
        v=int(self.done_files*100/self.total_files)
        self.progress['maximum']=100; self.progress['value']=v
        self.progress_label.config(text=f"{v}% ({self.done_files}/{self.total_files})")
        self.update_idletasks()
    def toggle_pause(self):
        if not self.is_running: return
        if not self.pause_flag.is_set():
            self.pause_flag.set(); self.pause_btn.config(text="Tiếp tục"); self.log("[WARN] Đã tạm dừng.")
        else:
            self.pause_flag.clear(); self.pause_btn.config(text="Tạm dừng"); self.log("[OK] Tiếp tục.")
    def cancel(self):
        if not self.is_running: return
        self.cancel_flag.set(); self.log("[WARN] Đang hủy...")

    def start(self, files_override=None, from_retry=False):
        if self.is_running:
            return
        ok, err = try_imports()
        if not ok:
            messagebox.showerror("Thiếu phụ thuộc", "Hãy cài: pip install pywin32 pypdf\n\n" + err)
            return
        if not from_retry and not self.folders:
            messagebox.showwarning("Thiếu thư mục", "Hãy thêm ít nhất một thư mục.")
            return
        pages, ranges = parse_pagespec(self.page_spec.get())
        dfrom = parse_date(self.date_from.get())
        dto = parse_date(self.date_to.get())
        if self.date_from.get().strip() and not dfrom:
            messagebox.showwarning("Ngày không hợp lệ", "Ngày 'từ' phải là YYYY-MM-DD")
            return
        if self.date_to.get().strip() and not dto:
            messagebox.showwarning("Ngày không hợp lệ", "Ngày 'đến' phải là YYYY-MM-DD")
            return

        files = files_override if files_override is not None else self.scan_files()
        if not files:
            messagebox.showinfo("Không có file", "Không tìm thấy file phù hợp (Hãy kiểm tra lại bộ lọc/Quét lại).")
            return

        mode_pdf = self.mode_pdf_export.get()
        out_dir = self.out_dir.get().strip()
        if mode_pdf and not out_dir:
            messagebox.showwarning("Thiếu thư mục PDF", "Hãy chọn thư mục để lưu PDF.")
            return

        target_paper = self.paper_size.get()
        if not mode_pdf and target_paper == "A5":
            if not messagebox.askyesno(
                "Xác nhận in A5",
                "Bạn đang chọn khổ giấy A5.\nMáy in sẽ được thiết lập tạm thời sang A5.\nBạn có chắc chắn không?",
            ):
                return

        log_dir = self.log_dir.get().strip() or os.getcwd()
        os.makedirs(log_dir, exist_ok=True)
        ok_csv = os.path.join(log_dir, "printed_ok.csv")
        err_csv = os.path.join(log_dir, "errors.csv")
        if self.enable_logging.get():
            if not os.path.exists(ok_csv):
                with open(ok_csv, "w", newline="", encoding="utf-8-sig") as f:
                    csv.writer(f).writerow(["time", "file", "type", "pagespec", "mode", "message"])
            if not os.path.exists(err_csv):
                with open(err_csv, "w", newline="", encoding="utf-8-sig") as f:
                    csv.writer(f).writerow(["time", "file", "type", "pagespec", "mode", "error"])

        def log_ok(path, typ, msg):
            if self.enable_logging.get():
                with open(ok_csv, "a", newline="", encoding="utf-8-sig") as f:
                    csv.writer(f).writerow(
                        [time.strftime("%Y-%m-%d %H:%M:%S"), path, typ, self.page_spec.get(), "export" if mode_pdf else "print", msg]
                    )

        def log_err(path, typ, errm):
            if self.enable_logging.get():
                with open(err_csv, "a", newline="", encoding="utf-8-sig") as f:
                    csv.writer(f).writerow(
                        [time.strftime("%Y-%m-%d %H:%M:%S"), path, typ, self.page_spec.get(), "export" if mode_pdf else "print", errm]
                    )

        self.failed_files = []
        self.last_run_files = list(files)
        self.total_files = len(files)
        self.done_files = 0
        self.update_progress()
        self.is_running = True
        self.pause_flag.clear()
        self.cancel_flag.clear()
        self.start_btn.config(state="disabled")
        self.pause_btn.config(state="normal")
        self.cancel_btn.config(state="normal")
        self.retry_btn.config(state="disabled")

        printer = self.selected_printer.get().strip() or None
        merge = self.pdf_merge.get()
        scope = self.sheet_scope.get()

        self.log(
            f"[OK] Tổng {self.total_files} file. Trang: {self.page_spec.get()} | Khổ: {target_paper} | Scope: {scope} | "
            f"{'Xuất PDF' if mode_pdf else 'In'} | {'Gộp' if merge else 'Tách'}"
        )

        def worker():
            restored_size = None
            hPrinter = None
            try:
                if not mode_pdf and printer:
                    try:
                        import win32print, win32con

                        PRINTER_ALL_ACCESS = 0xF000C
                        hPrinter = win32print.OpenPrinter(printer, {"DesiredAccess": PRINTER_ALL_ACCESS})
                        pInfo = win32print.GetPrinter(hPrinter, 2)
                        current_size = pInfo["pDevMode"].PaperSize
                        req_size = 11 if target_paper == "A5" else 9
                        if current_size != req_size:
                            self.log(f"[WARN] Đổi khổ giấy máy in từ {current_size} sang {req_size} (11=A5, 9=A4)...")
                            restored_size = current_size
                            pInfo["pDevMode"].PaperSize = req_size
                            pInfo["pDevMode"].Fields |= win32con.DM_PAPERSIZE
                            win32print.SetPrinter(hPrinter, 2, pInfo, 0)
                    except Exception as e:
                        self.log(f"[WARN] Không thể đổi khổ giấy: {e}")
                        restored_size = None

                for path in files:
                    while self.pause_flag.is_set():
                        time.sleep(0.2)
                        if self.cancel_flag.is_set():
                            break
                    if self.cancel_flag.is_set():
                        break
                    ext = os.path.splitext(path)[1].lower()
                    typ = "excel" if ext in EXCEL_EXTS else ("word" if ext in WORD_EXTS else "pdf")
                    ok2 = False
                    try:
                        if typ == "excel":
                            if not mode_pdf:
                                if scope == "first":
                                    ok2 = excel_print_first_sheet_ranges(path, ranges, printer, self.log)
                                else:
                                    ok2 = excel_print_all_sheets_ranges(path, ranges, printer, self.log)
                            else:
                                os.makedirs(out_dir, exist_ok=True)
                                if scope == "first":
                                    ok2 = excel_export_first_sheet(path, pages, ranges, printer, out_dir, merge, self.log)
                                else:
                                    ok2 = excel_export_all_sheets_pages(path, pages, ranges, printer, out_dir, merge, self.log)
                        elif typ == "word":
                            if not mode_pdf:
                                ok2 = word_print_ranges(path, ranges, printer, self.log)
                            else:
                                ok2 = word_export_pdf(path, pages, ranges, printer, out_dir, merge, self.log)
                        else:
                            if not mode_pdf:
                                ok2 = pdf_print_ranges(path, ranges, printer, self.log)
                            else:
                                ok2 = pdf_export(path, pages, ranges, out_dir, merge, self.log)
                    except Exception as e:
                        ok2 = False
                        self.log(f"[ERR] Lỗi: {path} - {e}")
                    if ok2:
                        log_ok(path, typ, "ok")
                    else:
                        log_err(path, typ, "failed")
                        self.failed_files.append(path)
                    self.done_files += 1
                    self.update_progress()

            finally:
                if hPrinter and restored_size:
                    try:
                        import win32print, win32con

                        self.log(f"[WARN] Khôi phục khổ giấy về: {restored_size}")
                        pInfo = win32print.GetPrinter(hPrinter, 2)
                        pInfo["pDevMode"].PaperSize = restored_size
                        pInfo["pDevMode"].Fields |= win32con.DM_PAPERSIZE
                        win32print.SetPrinter(hPrinter, 2, pInfo, 0)
                    except Exception as e:
                        self.log(f"[WARN] Lỗi khôi phục khổ giấy: {e}")

                if hPrinter:
                    try:
                        win32print.ClosePrinter(hPrinter)
                    except Exception:
                        pass

                if self.cancel_flag.is_set():
                    self.log("[WARN] Đã hủy theo yêu cầu.")
                else:
                    self.log("[OK] Hoàn tất.")
                if self.failed_files:
                    self.retry_btn.config(state="normal")
                self.is_running = False
                self.start_btn.config(state="normal")
                self.pause_btn.config(state="disabled")
                self.cancel_btn.config(state="disabled")
                self.pause_flag.clear()
                self.cancel_flag.clear()

        threading.Thread(target=worker, daemon=True).start()

    def retry_failed(self):
        if self.is_running:
            return
        if not self.failed_files:
            messagebox.showinfo("Không có lỗi", "Không có file lỗi để chạy lại.")
            return
        files = list(self.failed_files)
        self.failed_files = []
        self.start(files_override=files, from_retry=True)

    def check_pages(self):
        files = self.filtered_files if self.filtered_files else self.scan_files()
        if not files:
            messagebox.showinfo("Không có file", "Không có file để kiểm tra.")
            return
        try:
            from pypdf import PdfReader
        except Exception:
            messagebox.showerror("Thiếu phụ thuộc", "Hãy cài: pip install pypdf")
            return

        pages, _ = parse_pagespec(self.page_spec.get())
        pdf_checked = 0
        missing = []
        skipped = {"excel": 0, "word": 0}

        for path in files:
            ext = os.path.splitext(path)[1].lower()
            if ext in PDF_EXTS:
                try:
                    reader = PdfReader(path)
                    total = len(reader.pages)
                    invalid = [p for p in pages if p < 1 or p > total]
                    if invalid:
                        missing.append((path, total, invalid))
                    pdf_checked += 1
                except Exception as e:
                    missing.append((path, 0, [f"Lỗi đọc: {e}"]))
            elif ext in EXCEL_EXTS:
                skipped["excel"] += 1
            elif ext in WORD_EXTS:
                skipped["word"] += 1

        for path, total, invalid in missing:
            self.log(f"[WARN] PDF thiếu trang: {path} (tối đa {total}) -> {invalid}")

        summary = (
            f"Đã kiểm tra {pdf_checked} file PDF.\n"
            f"File thiếu trang: {len(missing)}.\n"
            f"Word/Excel: chưa hỗ trợ tự kiểm tra (Word: {skipped['word']}, Excel: {skipped['excel']})."
        )
        messagebox.showinfo("Kết quả kiểm tra", summary)

    def on_close(self):
        cfg={
            "folders": self.folders,
            "recursive": self.recursive.get(),
            "pattern": self.pattern.get(),
            "date_from": self.date_from.get(),
            "date_to": self.date_to.get(),
            "use_excel": self.use_excel.get(),
            "use_word": self.use_word.get(),
            "use_pdf": self.use_pdf.get(),
            "page_spec": self.page_spec.get(),
            "mode_pdf_export": self.mode_pdf_export.get(),
            "pdf_merge": self.pdf_merge.get(),
            "sheet_scope": self.sheet_scope.get(),
            "paper_size": self.paper_size.get(),
            "out_dir": self.out_dir.get(),
            "enable_logging": self.enable_logging.get(),
            "log_dir": self.log_dir.get(),
            "selected_printer": self.selected_printer.get(),
            "search_term": self.search_var.get(),
            "sort_mode": self.sort_var.get(),
            "presets": self.presets,
        }
        save_config(cfg); self.destroy()

if __name__ == "__main__":
    ensure_readme_exists()
    app=App()
    app.mainloop()
