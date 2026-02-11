import tkinter as tk
from tkinter import ttk
from ui_effects import apply_button_effects, bind_tooltip, apply_focus_ring


def build_source_panel(app, parent):
    p1 = app._section(parent, "1. Nguồn dữ liệu")
    btn_row = tk.Frame(p1, bg=app.colors["panel"])
    btn_row.pack(fill="x", pady=(0, 6))
    btn_add = tk.Button(btn_row, text="Thêm thư mục", command=app.add_folder, **app.btn_primary)
    apply_button_effects(btn_add, app.colors, style="primary")
    bind_tooltip(btn_add, "Thêm thư mục vào danh sách nguồn.", app.colors)
    btn_add.pack(side="left")
    btn_remove = tk.Button(btn_row, text="Xóa đã chọn", command=app.remove_selected, **app.btn_secondary)
    apply_button_effects(btn_remove, app.colors, style="secondary")
    bind_tooltip(btn_remove, "Xóa các thư mục đang chọn khỏi danh sách.", app.colors)
    btn_remove.pack(side="left", padx=6)
    tk.Frame(btn_row, bg=app.colors["panel"]).pack(side="left", expand=True, fill="x")
    btn_clear = tk.Button(btn_row, text="Xóa hết", command=app.clear_folders, **app.btn_danger)
    apply_button_effects(btn_clear, app.colors, style="danger")
    bind_tooltip(btn_clear, "Xóa toàn bộ thư mục nguồn trong danh sách.", app.colors)
    btn_clear.pack(side="right")
    tk.Checkbutton(
        btn_row,
        text="Quét thư mục con",
        variable=app.recursive,
        command=app.scan_files,
        bg=app.colors["panel"],
        fg=app.colors["text"],
    ).pack(side="left", padx=10)

    box = tk.Frame(p1, bg=app.colors["panel"])
    box.pack(fill="both", expand=True)
    app.listbox = tk.Listbox(
        box,
        selectmode="extended",
        height=5,
        bg="white",
        fg=app.colors["text"],
        selectbackground=app.colors["primary"],
        selectforeground="white",
        relief="solid",
        bd=1,
        highlightthickness=1,
        highlightbackground=app.colors["border"],
    )
    app.listbox.pack(side="left", fill="both", expand=True)
    apply_focus_ring(app.listbox, app.colors)
    sb = tk.Scrollbar(box, orient="vertical", command=app.listbox.yview)
    sb.pack(side="left", fill="y")
    app.listbox.config(yscrollcommand=sb.set)
    for p in app.folders:
        app.listbox.insert("end", p)

    p2 = app._section(parent, "2. Lọc file")
    grid = tk.Frame(p2, bg=app.colors["panel"])
    grid.pack(fill="x")
    grid.columnconfigure(1, weight=1)

    tk.Label(grid, text="Tên file (pattern)", bg=app.colors["panel"], fg=app.colors["text"]).grid(
        row=0, column=0, sticky="w"
    )
    pattern_entry = tk.Entry(grid, textvariable=app.pattern, bg="white", relief="solid", bd=1)
    apply_focus_ring(pattern_entry, app.colors)
    pattern_entry.grid(row=0, column=1, sticky="ew", padx=6)
    btn_scan = tk.Button(grid, text="Quét file", command=app.scan_files, **app.btn_primary)
    apply_button_effects(btn_scan, app.colors, style="primary")
    bind_tooltip(btn_scan, "Quét lại danh sách file theo bộ lọc hiện tại.", app.colors)
    btn_scan.grid(row=0, column=2, sticky="e")

    tk.Label(grid, text="Ngày sửa", bg=app.colors["panel"], fg=app.colors["text"]).grid(
        row=1, column=0, sticky="w", pady=(6, 0)
    )
    date_box = tk.Frame(grid, bg=app.colors["panel"])
    date_box.grid(row=1, column=1, sticky="w", pady=(6, 0))
    date_from_entry = tk.Entry(date_box, textvariable=app.date_from, width=12, bg="white", relief="solid", bd=1)
    apply_focus_ring(date_from_entry, app.colors)
    date_from_entry.pack(side="left")
    tk.Label(date_box, text="đến", bg=app.colors["panel"], fg=app.colors["muted"]).pack(side="left", padx=6)
    date_to_entry = tk.Entry(date_box, textvariable=app.date_to, width=12, bg="white", relief="solid", bd=1)
    apply_focus_ring(date_to_entry, app.colors)
    date_to_entry.pack(side="left")

    p3 = app._section(parent, "3. Danh sách file", expand=True)
    hdr = tk.Frame(p3, bg=app.colors["panel"])
    hdr.pack(fill="x")
    app.lbl_count = tk.Label(hdr, text="0 file", fg=app.colors["primary"], bg=app.colors["panel"])
    app.lbl_count.pack(side="left")

    filter_row = tk.Frame(p3, bg=app.colors["panel"])
    filter_row.pack(fill="x", pady=(6, 0))
    tk.Label(filter_row, text="Tìm nhanh", bg=app.colors["panel"], fg=app.colors["muted"]).pack(side="left")
    app.search_entry = tk.Entry(filter_row, textvariable=app.search_var, width=18, bg="white", relief="solid", bd=1)
    apply_focus_ring(app.search_entry, app.colors)
    app.search_entry.pack(side="left", padx=6)
    tk.Label(filter_row, text="Sắp xếp", bg=app.colors["panel"], fg=app.colors["muted"]).pack(side="left", padx=(8, 0))
    app.sort_combo = ttk.Combobox(
        filter_row,
        values=app.sort_options,
        textvariable=app.sort_var,
        width=18,
        state="readonly",
    )
    apply_focus_ring(app.sort_combo, app.colors)
    app.sort_combo.pack(side="left", padx=6)

    box2 = tk.Frame(p3, bg=app.colors["panel"])
    box2.pack(fill="both", expand=True, pady=(6, 0))
    app.file_listbox = tk.Listbox(
        box2,
        selectmode="extended",
        bg="white",
        fg=app.colors["text"],
        selectbackground=app.colors["primary"],
        selectforeground="white",
        relief="solid",
        bd=1,
        highlightthickness=1,
        highlightbackground=app.colors["border"],
    )
    app.file_listbox.pack(side="left", fill="both", expand=True)
    apply_focus_ring(app.file_listbox, app.colors)
    sb2 = tk.Scrollbar(box2, orient="vertical", command=app.file_listbox.yview)
    sb2.pack(side="left", fill="y")
    app.file_listbox.config(yscrollcommand=sb2.set)


def build_settings_panel(app, parent):
    p4 = app._section(parent, "4. Thiết lập in/xuất")
    grid = tk.Frame(p4, bg=app.colors["panel"])
    grid.pack(fill="x")
    grid.columnconfigure(1, weight=1)

    tk.Label(grid, text="Loại file", bg=app.colors["panel"], fg=app.colors["text"]).grid(row=0, column=0, sticky="w")
    type_box = tk.Frame(grid, bg=app.colors["panel"])
    type_box.grid(row=0, column=1, sticky="w", padx=6)
    tk.Checkbutton(type_box, text="Excel", variable=app.use_excel, command=app.scan_files, bg=app.colors["panel"]).pack(side="left")
    tk.Checkbutton(type_box, text="Word", variable=app.use_word, command=app.scan_files, bg=app.colors["panel"]).pack(side="left", padx=6)
    tk.Checkbutton(type_box, text="PDF", variable=app.use_pdf, command=app.scan_files, bg=app.colors["panel"]).pack(side="left")

    tk.Label(grid, text="Khổ giấy", bg=app.colors["panel"], fg=app.colors["text"]).grid(row=1, column=0, sticky="w", pady=(6, 0))
    paper_box = tk.Frame(grid, bg=app.colors["panel"])
    paper_box.grid(row=1, column=1, sticky="w", padx=6, pady=(6, 0))
    tk.Radiobutton(paper_box, text="A4", variable=app.paper_size, value="A4", bg=app.colors["panel"]).pack(side="left")
    tk.Radiobutton(paper_box, text="A5", variable=app.paper_size, value="A5", bg=app.colors["panel"]).pack(side="left", padx=6)

    tk.Label(grid, text="Trang in", bg=app.colors["panel"], fg=app.colors["text"]).grid(row=2, column=0, sticky="w", pady=(6, 0))
    app.entry_pages = tk.Entry(grid, textvariable=app.page_spec, width=18, bg="white", relief="solid", bd=1)
    apply_focus_ring(app.entry_pages, app.colors)
    app.entry_pages.grid(row=2, column=1, sticky="w", padx=6, pady=(6, 0))
    tk.Label(grid, text="(vd: 1,3-5)", bg=app.colors["panel"], fg=app.colors["muted"]).grid(
        row=2, column=2, sticky="w", pady=(6, 0)
    )

    tk.Label(grid, text="Sheet Excel", bg=app.colors["panel"], fg=app.colors["text"]).grid(row=3, column=0, sticky="w", pady=(6, 0))
    sheet_box = tk.Frame(grid, bg=app.colors["panel"])
    sheet_box.grid(row=3, column=1, sticky="w", padx=6, pady=(6, 0))
    tk.Radiobutton(sheet_box, text="Sheet đầu", variable=app.sheet_scope, value="first", bg=app.colors["panel"]).pack(side="left")
    tk.Radiobutton(sheet_box, text="Tất cả", variable=app.sheet_scope, value="all", bg=app.colors["panel"]).pack(side="left", padx=6)

    p5 = app._section(parent, "5. Đầu ra")
    r1 = tk.Frame(p5, bg=app.colors["panel"])
    r1.pack(fill="x", pady=(0, 6))
    tk.Radiobutton(
        r1,
        text="In ra máy in",
        variable=app.mode_pdf_export,
        value=0,
        command=app.toggle_pdf_dir,
        bg=app.colors["panel"],
    ).pack(side="left")
    app.cmb = ttk.Combobox(r1, values=app.printers, textvariable=app.selected_printer, width=32, state="readonly")
    apply_focus_ring(app.cmb, app.colors)
    app.cmb.pack(side="left", padx=6)
    app.btn_refresh = tk.Button(r1, text="Làm mới", command=app.refresh_printers, **app.btn_secondary)
    apply_button_effects(app.btn_refresh, app.colors, style="secondary")
    bind_tooltip(app.btn_refresh, "Lấy lại danh sách máy in từ hệ thống.", app.colors)
    app.btn_refresh.pack(side="left")

    r2 = tk.Frame(p5, bg=app.colors["panel"])
    r2.pack(fill="x")
    tk.Radiobutton(
        r2,
        text="Xuất ra PDF",
        variable=app.mode_pdf_export,
        value=1,
        command=app.toggle_pdf_dir,
        bg=app.colors["panel"],
    ).pack(side="left")
    app.entry_out = tk.Entry(r2, textvariable=app.out_dir, width=30, bg="white", relief="solid", bd=1)
    apply_focus_ring(app.entry_out, app.colors)
    app.entry_out.pack(side="left", padx=6)
    app.btn_browse = tk.Button(r2, text="Chọn...", command=app.pick_out_dir, **app.btn_secondary)
    apply_button_effects(app.btn_browse, app.colors, style="secondary")
    bind_tooltip(app.btn_browse, "Chọn thư mục lưu file PDF xuất ra.", app.colors)
    app.btn_browse.pack(side="left")
    app.chk_pdf_merge = tk.Checkbutton(r2, text="Gộp 1 file", variable=app.pdf_merge, bg=app.colors["panel"])
    app.chk_pdf_merge.pack(side="left", padx=8)
    app.btn_open_output = tk.Button(r2, text="Mở thư mục PDF", command=app.open_output_folder, **app.btn_secondary)
    apply_button_effects(app.btn_open_output, app.colors, style="secondary")
    bind_tooltip(app.btn_open_output, "Mở nhanh thư mục PDF đã chọn.", app.colors)
    app.btn_open_output.pack(side="left")

    p6 = app._section(parent, "6. Cấu hình nhanh")
    preset_row = tk.Frame(p6, bg=app.colors["panel"])
    preset_row.pack(fill="x")
    tk.Label(preset_row, text="Bộ cấu hình", bg=app.colors["panel"], fg=app.colors["muted"]).pack(side="left")
    app.preset_combo = ttk.Combobox(preset_row, values=[], textvariable=app.preset_name, width=24, state="readonly")
    apply_focus_ring(app.preset_combo, app.colors)
    app.preset_combo.pack(side="left", padx=6)
    app.btn_apply_preset = tk.Button(preset_row, text="Áp dụng", command=app.apply_selected_preset, **app.btn_secondary)
    apply_button_effects(app.btn_apply_preset, app.colors, style="secondary")
    bind_tooltip(app.btn_apply_preset, "Áp dụng bộ cấu hình đang chọn.", app.colors)
    app.btn_apply_preset.pack(side="left", padx=(0, 6))

    save_row = tk.Frame(p6, bg=app.colors["panel"])
    save_row.pack(fill="x", pady=(6, 0))
    tk.Label(save_row, text="Tên mới", bg=app.colors["panel"], fg=app.colors["muted"]).pack(side="left")
    app.preset_entry = tk.Entry(save_row, textvariable=app.preset_input, width=22, bg="white", relief="solid", bd=1)
    apply_focus_ring(app.preset_entry, app.colors)
    app.preset_entry.pack(side="left", padx=6)
    app.btn_save_preset = tk.Button(save_row, text="Lưu/Cập nhật", command=app.save_preset, **app.btn_primary)
    apply_button_effects(app.btn_save_preset, app.colors, style="primary")
    bind_tooltip(app.btn_save_preset, "Lưu cấu hình hiện tại thành bộ cấu hình nhanh.", app.colors)
    app.btn_save_preset.pack(side="left", padx=(0, 6))
    app.btn_delete_preset = tk.Button(save_row, text="Xóa", command=app.delete_preset, **app.btn_danger)
    apply_button_effects(app.btn_delete_preset, app.colors, style="danger")
    bind_tooltip(app.btn_delete_preset, "Xóa bộ cấu hình đang chọn.", app.colors)
    app.btn_delete_preset.pack(side="left")

    p7 = app._section(parent, "7. Thao tác")
    app.start_btn = tk.Button(
        p7,
        text="Bắt đầu",
        command=app.start,
        font=("Segoe UI", 11, "bold"),
        **app.btn_primary,
    )
    apply_button_effects(app.start_btn, app.colors, style="primary")
    bind_tooltip(app.start_btn, "Bắt đầu in/xuất PDF theo thiết lập hiện tại.", app.colors)
    app.start_btn.pack(fill="x", ipady=6, pady=(0, 6))

    action_row = tk.Frame(p7, bg=app.colors["panel"])
    action_row.pack(fill="x")
    app.pause_btn = tk.Button(action_row, text="Tạm dừng", command=app.toggle_pause, state="disabled", **app.btn_secondary)
    apply_button_effects(app.pause_btn, app.colors, style="secondary")
    bind_tooltip(app.pause_btn, "Tạm dừng hoặc tiếp tục tiến trình.", app.colors)
    app.pause_btn.pack(side="left", expand=True, fill="x", padx=(0, 6))
    app.cancel_btn = tk.Button(action_row, text="Hủy", command=app.cancel, state="disabled", **app.btn_danger)
    apply_button_effects(app.cancel_btn, app.colors, style="danger")
    bind_tooltip(app.cancel_btn, "Dừng toàn bộ tiến trình đang chạy.", app.colors)
    app.cancel_btn.pack(side="left", expand=True, fill="x")

    util_row = tk.Frame(p7, bg=app.colors["panel"])
    util_row.pack(fill="x", pady=(6, 0))
    app.check_btn = tk.Button(util_row, text="Kiểm tra trang", command=app.check_pages, **app.btn_secondary)
    apply_button_effects(app.check_btn, app.colors, style="secondary")
    bind_tooltip(app.check_btn, "Ước lượng số trang và cảnh báo lỗi trang.", app.colors)
    app.check_btn.pack(side="left", expand=True, fill="x", padx=(0, 6))
    app.retry_btn = tk.Button(util_row, text="Chạy lại lỗi", command=app.retry_failed, state="disabled", **app.btn_secondary)
    apply_button_effects(app.retry_btn, app.colors, style="secondary")
    bind_tooltip(app.retry_btn, "Chạy lại các file bị lỗi ở lần chạy trước.", app.colors)
    app.retry_btn.pack(side="left", expand=True, fill="x")

    app.progress_label = tk.Label(p7, text="0% (0/0)", bg=app.colors["panel"], fg=app.colors["text"])
    app.progress_label.pack(anchor="w", pady=(6, 0))
    app.progress = ttk.Progressbar(p7, orient="horizontal", mode="determinate")
    app.progress.pack(fill="x", pady=(2, 0))

    p8 = app._section(parent, "8. Nhật ký", expand=True)
    tool_row = tk.Frame(p8, bg=app.colors["panel"])
    tool_row.pack(fill="x", pady=(0, 6))
    tk.Checkbutton(
        tool_row,
        text="Ghi log file",
        variable=app.enable_logging,
        command=app.toggle_logging_fields,
        bg=app.colors["panel"],
    ).pack(side="left")
    tk.Label(tool_row, text="Thư mục log", bg=app.colors["panel"], fg=app.colors["muted"]).pack(side="left", padx=(10, 4))
    app.log_entry = tk.Entry(tool_row, textvariable=app.log_dir, width=26, bg="white", relief="solid", bd=1)
    apply_focus_ring(app.log_entry, app.colors)
    app.log_entry.pack(side="left", padx=4)
    app.btn_log_browse = tk.Button(tool_row, text="Chọn...", command=app.pick_log_dir, **app.btn_secondary)
    apply_button_effects(app.btn_log_browse, app.colors, style="secondary")
    bind_tooltip(app.btn_log_browse, "Chọn thư mục lưu file log.", app.colors)
    app.btn_log_browse.pack(side="left")
    app.btn_log_open = tk.Button(tool_row, text="Mở thư mục", command=app.open_log_folder, **app.btn_secondary)
    apply_button_effects(app.btn_log_open, app.colors, style="secondary")
    bind_tooltip(app.btn_log_open, "Mở nhanh thư mục log.", app.colors)
    app.btn_log_open.pack(side="right")

    app.txt = tk.Text(p8, height=8, bg="white", fg=app.colors["text"], relief="solid", bd=1)
    app.txt.pack(fill="both", expand=True)
    app.txt.tag_config("err", foreground=app.colors["danger"])
    app.txt.tag_config("warn", foreground=app.colors["warning"])
    app.txt.tag_config("ok", foreground=app.colors["success"])

    app.toggle_pdf_dir()
    app.toggle_logging_fields()
