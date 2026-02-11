import os
from tkinter import messagebox


def open_output_folder(self):
    if not self.mode_pdf_export.get():
        messagebox.showinfo("Thông tin", "Chức năng này chỉ dùng khi xuất PDF.")
        return
    out_dir = self.out_dir.get().strip()
    if not out_dir:
        messagebox.showinfo("Thông tin", "Chưa chọn thư mục PDF.")
        return
    if os.path.isdir(out_dir):
        os.startfile(out_dir)
    else:
        messagebox.showinfo("Thông tin", "Thư mục PDF chưa tồn tại.")


def apply_file_filter(self):
    files = list(self.current_files)
    term = self.search_var.get().strip().lower()
    if term:
        files = [f for f in files if term in os.path.basename(f).lower() or term in f.lower()]
    mode = self.sort_var.get()
    if mode == "Tên A→Z":
        files.sort(key=lambda p: os.path.basename(p).lower())
    elif mode == "Tên Z→A":
        files.sort(key=lambda p: os.path.basename(p).lower(), reverse=True)
    elif mode == "Ngày sửa mới→cũ":
        files.sort(key=lambda p: os.path.getmtime(p) if os.path.exists(p) else 0, reverse=True)
    elif mode == "Ngày sửa cũ→mới":
        files.sort(key=lambda p: os.path.getmtime(p) if os.path.exists(p) else 0)
    self.filtered_files = files
    if hasattr(self, "file_listbox"):
        self.file_listbox.delete(0, "end")
        for f in files:
            self.file_listbox.insert("end", f)
    if hasattr(self, "lbl_count"):
        self.lbl_count.config(text=f"Tìm thấy: {len(files)} file")
    self.update_summary()


def refresh_preset_list(self):
    if hasattr(self, "preset_combo"):
        names = sorted(self.presets.keys())
        self.preset_combo["values"] = names


def on_preset_select(self):
    name = self.preset_name.get().strip()
    if name and name in self.presets:
        self.preset_input.set(name)


def collect_state(self):
    return {
        "folders": list(self.folders),
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
    }


def apply_state(self, state):
    self.folders = list(state.get("folders", []))
    self.recursive.set(state.get("recursive", True))
    self.pattern.set(state.get("pattern", "*"))
    self.date_from.set(state.get("date_from", ""))
    self.date_to.set(state.get("date_to", ""))
    self.use_excel.set(state.get("use_excel", True))
    self.use_word.set(state.get("use_word", False))
    self.use_pdf.set(state.get("use_pdf", False))
    self.page_spec.set(state.get("page_spec", "1"))
    self.mode_pdf_export.set(state.get("mode_pdf_export", False))
    self.pdf_merge.set(state.get("pdf_merge", True))
    self.sheet_scope.set(state.get("sheet_scope", "first"))
    self.paper_size.set(state.get("paper_size", "A4"))
    self.out_dir.set(state.get("out_dir", ""))
    self.enable_logging.set(state.get("enable_logging", True))
    self.log_dir.set(state.get("log_dir", ""))
    self.selected_printer.set(state.get("selected_printer", self.selected_printer.get()))
    self.search_var.set(state.get("search_term", ""))
    self.sort_var.set(state.get("sort_mode", self.sort_options[0]))

    if hasattr(self, "listbox"):
        self.listbox.delete(0, "end")
        for p in self.folders:
            self.listbox.insert("end", p)
    self.scan_files()
    self.toggle_pdf_dir()
    self.toggle_logging_fields()


def save_preset(self):
    name = self.preset_input.get().strip()
    if not name:
        messagebox.showinfo("Thiếu tên", "Hãy nhập tên bộ cấu hình.")
        return
    if name in self.presets:
        if not messagebox.askyesno("Xác nhận", f"Đã tồn tại '{name}'. Bạn muốn cập nhật?"):
            return
    self.presets[name] = self.collect_state()
    self.preset_name.set(name)
    self.refresh_preset_list()
    self.log(f"[OK] Đã lưu cấu hình: {name}")


def delete_preset(self):
    name = self.preset_name.get().strip() or self.preset_input.get().strip()
    if not name:
        messagebox.showinfo("Thiếu tên", "Chọn cấu hình để xóa.")
        return
    if name not in self.presets:
        messagebox.showinfo("Không tồn tại", "Không tìm thấy cấu hình.")
        return
    if not messagebox.askyesno("Xác nhận", f"Xóa cấu hình '{name}'?"):
        return
    self.presets.pop(name, None)
    self.preset_name.set("")
    self.preset_input.set("")
    self.refresh_preset_list()
    self.log(f"[OK] Đã xóa cấu hình: {name}")


def apply_selected_preset(self):
    name = self.preset_name.get().strip()
    if not name or name not in self.presets:
        messagebox.showinfo("Thiếu cấu hình", "Hãy chọn một bộ cấu hình.")
        return
    self.apply_state(self.presets[name])
    self.log(f"[OK] Đã áp dụng cấu hình: {name}")


def clear_folders(self):
    if not self.folders:
        messagebox.showinfo("Thông tin", "Danh sách thư mục đang trống.")
        return
    if not messagebox.askyesno("Xác nhận", "Bạn muốn xóa tất cả thư mục nguồn?"):
        return
    self.folders = []
    if hasattr(self, "listbox"):
        self.listbox.delete(0, "end")
    self.scan_files()
    if hasattr(self, "log"):
        self.log("[OK] Đã xóa tất cả thư mục nguồn.")
