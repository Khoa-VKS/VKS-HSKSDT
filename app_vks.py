# -*- coding: utf-8 -*-
"""
HỆ THỐNG TỰ ĐỘNG HÓA TỔNG HỢP & GIÁM SÁT TIẾN ĐỘ ÁN HÌNH SỰ (BẢN NÂNG CẤP TOÀN DIỆN)
Viện kiểm sát nhân dân Thành phố Hồ Chí Minh

CÁC TÍNH NĂNG VÀ QUY CHUẨN ĐỒNG BỘ 100%:
1. Đúng chuẩn 23 cột Master File quy định của VKSND TP.HCM.
2. Tiêu đề nền VÀNG (#FFFF00), chữ đen in đậm, AutoFilter và Gridlines hiển thị rõ nét.
3. Cột 1 đến 8: Giữ nguyên hiện trạng từ file xuất xuống của đơn vị.
4. Cột 9 đến 14: Bóc tách chính xác từ ô Thông tin vụ án.
5. Cột 15 đến 23: Trực quan hóa tiến độ, số ngày, đơn vị và các cột ghi chú kiểm tra.
6. Sheet dữ liệu mang tên "DATA", tự động tạo kèm Sheet "BẢNG THEO DÕI" (Pivot Table theo Đơn vị: Count of MÃ VỤ ÁN, Đạt, Lỗi).
7. Giao diện Tkinter chạy đa luồng (Multi-threading), không bị đơ hoặc treo ứng dụng.
"""

import os
import sys
import threading
from datetime import datetime, date, timedelta
import pandas as pd
import openpyxl

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
if CURRENT_DIR not in sys.path:
    sys.path.insert(0, CURRENT_DIR)

from backend.core import (
    MASTER_COLUMNS,
    remove_accents,
    parse_date,
    find_header_row,
    detect_column_indices,
    extract_case_details_from_cell,
    is_case_error,
    make_compound_key,
    resolve_input_files,
    normalize_column_names,
    save_master_excel,
    add_bang_theo_doi_sheet,
    check_file_has_user_notes,
    buoc_1_tong_hop_va_loc as core_b1,
    buoc_2_cap_nhat_va_bao_cao as core_b2,
)

def buoc_1_tong_hop_va_loc(folder_or_files, output_dir=None, check_days=16, existing_master_path=None, use_timestamp=False):
    """Bọc logic Bước 1 đồng bộ với backend core"""
    return core_b1(folder_or_files, output_dir=output_dir, check_days=check_days, existing_master_path=existing_master_path, use_timestamp=use_timestamp)

def buoc_2_cap_nhat_va_bao_cao(file_tong_path, file_kt_updated_path, output_dir=None, use_timestamp=False):
    """Bọc logic Bước 2 đồng bộ với backend core"""
    return core_b2(file_tong_path, file_kt_updated_path, output_dir=output_dir, use_timestamp=use_timestamp)

# ==================== GIAO DIỆN ĐỒ HỌA & DÒNG LỆNH ====================

def run_gui():
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox, ttk
    except ImportError:
        run_cli()
        return

    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID('vksnd.tonghopan.app')
    except Exception:
        pass

    root = tk.Tk()
    root.title("Tổng hợp và giám sát việc lập hồ sơ kiểm sát điện tử - VKSND TP.HCM")
    root.geometry("730x550")
    root.resizable(False, False)

    # Style
    style = ttk.Style()
    style.theme_use('clam')

    # Banner Header
    banner = tk.Frame(root, bg="#0B2545", pady=10)
    banner.pack(fill="x")

    # Logo Viện Kiểm Sát ở góc trên bên trái màn hình
    logo_img = None
    try:
        from PIL import Image, ImageTk
        base_dir = os.path.dirname(os.path.abspath(__file__))
        logo_path = os.path.join(base_dir, "assets", "logo_vks.png")
        if not os.path.exists(logo_path):
            logo_path = os.path.abspath("assets/logo_vks.png")
        if os.path.exists(logo_path):
            im = Image.open(logo_path)
            lh = 48
            lw = int(im.width * (lh / im.height))
            im_resized = im.resize((lw, lh), Image.Resampling.LANCZOS)
            logo_img = ImageTk.PhotoImage(im_resized)
            
            ico_path = os.path.join(base_dir, "assets", "logo_vks.ico")
            if os.path.exists(ico_path):
                try:
                    root.iconbitmap(ico_path)
                except Exception:
                    root.iconphoto(False, logo_img)
            else:
                root.iconphoto(False, logo_img)
    except Exception as e:
        print("[Logo Load Error]", e)

    if logo_img:
        lbl_logo = tk.Label(banner, image=logo_img, bg="#0B2545", bd=0)
        lbl_logo.image = logo_img
        lbl_logo.place(x=16, rely=0.5, anchor="w")

    lbl_title = tk.Label(
        banner,
        text="TỔNG HỢP VÀ GIÁM SÁT VIỆC LẬP HỒ SƠ KIỂM SÁT ĐIỆN TỬ",
        font=("Segoe UI", 12, "bold"),
        fg="#FFFFFF",
        bg="#0B2545"
    )
    lbl_title.pack()

    lbl_sub = tk.Label(
        banner,
        text="VIỆN KIỂM SÁT NHÂN DÂN THÀNH PHỐ HỒ CHÍ MINH",
        font=("Segoe UI", 9, "bold"),
        fg="#FBBF24",
        bg="#0B2545"
    )
    lbl_sub.pack(pady=(2, 0))

    # Container
    main_frame = tk.Frame(root, padx=16, pady=8)
    main_frame.pack(fill="both", expand=True)

    # Cấu hình nút phụ (Browse / Chọn file) nhỏ gọn, thanh lịch
    btn_sec_cfg = {
        "font": ("Segoe UI", 8),
        "bg": "#F8FAFC",
        "fg": "#1E293B",
        "activebackground": "#E2E8F0",
        "activeforeground": "#0F172A",
        "relief": "solid",
        "bd": 1,
        "padx": 7,
        "pady": 1,
        "cursor": "hand2"
    }

    # ---------------- BƯỚC 1 ----------------
    f1 = ttk.LabelFrame(main_frame, text=" BƯỚC 1: TỔNG HỢP DỮ LIỆU ĐƠN VỊ & XUẤT DANH SÁCH KIỂM TRA ")
    f1.pack(fill="x", pady=5)
    f1.columnconfigure(1, weight=1)

    lbl_f1 = tk.Label(f1, text="Nguồn file đơn vị (*):", font=("Segoe UI", 9))
    lbl_f1.grid(row=0, column=0, padx=8, pady=4, sticky="w")

    txt_folder = ttk.Entry(f1, width=34)
    txt_folder.grid(row=0, column=1, padx=4, pady=4)

    frm_f1_btns = tk.Frame(f1)
    frm_f1_btns.grid(row=0, column=2, padx=2, pady=4)

    def chon_thu_muc():
        target = txt_folder.get().strip()
        first_item = target.split(";")[0].strip() if target else ""
        init_dir = first_item if os.path.isdir(first_item) else os.path.dirname(first_item) if os.path.isfile(first_item) else os.path.expanduser("~/Desktop")
        d = filedialog.askdirectory(title="Chọn thư mục chứa các file đơn vị", initialdir=init_dir)
        if d:
            txt_folder.delete(0, tk.END)
            txt_folder.insert(0, d)

    def chon_cac_file():
        target = txt_folder.get().strip()
        first_item = target.split(";")[0].strip() if target else ""
        init_dir = first_item if os.path.isdir(first_item) else os.path.dirname(first_item) if os.path.isfile(first_item) else os.path.expanduser("~/Desktop")
        files = filedialog.askopenfilenames(
            title="Chọn 1 hoặc nhiều file Excel của đơn vị (giữ Ctrl để chọn nhiều)",
            filetypes=[("Excel files", "*.xlsx;*.xlsm"), ("All files", "*.*")],
            initialdir=init_dir
        )
        if files:
            txt_folder.delete(0, tk.END)
            txt_folder.insert(0, "; ".join(files))

    tk.Button(frm_f1_btns, text="📁 Thư mục...", command=chon_thu_muc, **btn_sec_cfg).pack(side="left", padx=2)
    tk.Button(frm_f1_btns, text="📄 Chọn file...", command=chon_cac_file, **btn_sec_cfg).pack(side="left", padx=2)

    lbl_days = tk.Label(f1, text="Lọc từ:", font=("Segoe UI", 9))
    lbl_days.grid(row=0, column=3, padx=(10, 2), pady=4, sticky="w")
    spn_days = ttk.Spinbox(f1, from_=1, to=365, width=4)
    spn_days.set(16)
    spn_days.grid(row=0, column=4, padx=2, pady=4)
    tk.Label(f1, text="ngày", font=("Segoe UI", 9)).grid(row=0, column=5, padx=2, pady=4)

    lbl_ex = tk.Label(f1, text="File Tổng cũ (tùy chọn):", font=("Segoe UI", 9))
    lbl_ex.grid(row=1, column=0, padx=8, pady=3, sticky="w")

    txt_existing = ttk.Entry(f1, width=40)
    txt_existing.grid(row=1, column=1, padx=4, pady=3)

    def chon_file_existing():
        target = txt_existing.get().strip() or txt_folder.get().strip()
        first_item = target.split(";")[0].strip() if target else ""
        init_dir = first_item if os.path.isdir(first_item) else os.path.dirname(first_item) if os.path.isfile(first_item) else os.path.expanduser("~/Desktop")
        f = filedialog.askopenfilename(
            title="Chọn File Tổng cũ đã có (1_FILE_TONG.xlsx)",
            filetypes=[("Excel files", "*.xlsx;*.xlsm"), ("All files", "*.*")],
            initialdir=init_dir
        )
        if f:
            txt_existing.delete(0, tk.END)
            txt_existing.insert(0, f)

    tk.Button(f1, text="🔍 Chọn...", command=chon_file_existing, **btn_sec_cfg).grid(row=1, column=2, padx=4, pady=3)

    lbl_ex_hint = tk.Label(f1, text="💡 Để trống nếu tạo mới từ đầu. Nếu chọn file cũ: hệ thống chỉ nạp thêm các vụ chưa có!", font=("Segoe UI", 8, "italic"), fg="#0B2545")
    lbl_ex_hint.grid(row=2, column=0, columnspan=6, padx=8, pady=(0, 2), sticky="w")

    var_timestamp = tk.BooleanVar(value=False)
    chk_timestamp = ttk.Checkbutton(
        f1,
        text="⏱️ Đặt tên file xuất kèm ngày giờ (Timestamp) để lưu vết tránh ghi đè (VD: _20261003_1430.xlsx)",
        variable=var_timestamp
    )
    chk_timestamp.grid(row=3, column=0, columnspan=6, padx=8, pady=(1, 4), sticky="w")

    # Nút 1: Nhỏ gọn, màu Xanh Dương Hiện Đại (#0284C7) hiển thị 100% đầy đủ chữ và icon
    frm_action1 = tk.Frame(f1)
    frm_action1.grid(row=4, column=0, columnspan=6, sticky="ew", pady=(5, 8))
    btn_step1 = tk.Button(
        frm_action1,
        text="⚡ Thực hiện Bước 1: Tổng hợp & Lọc danh sách kiểm tra (>= 16 ngày)",
        bg="#0284C7",
        activebackground="#0369A1",
        activeforeground="white",
        fg="white",
        font=("Segoe UI", 9, "bold"),
        relief="flat",
        bd=0,
        padx=20,
        pady=5,
        cursor="hand2"
    )
    btn_step1.pack()

    def on_enter_b1(e):
        if str(btn_step1["state"]) != "disabled":
            btn_step1.config(bg="#0369A1")
    def on_leave_b1(e):
        if str(btn_step1["state"]) != "disabled":
            btn_step1.config(bg="#0284C7")
    btn_step1.bind("<Enter>", on_enter_b1)
    btn_step1.bind("<Leave>", on_leave_b1)

    # ---------------- BƯỚC 2 ----------------
    f2 = ttk.LabelFrame(main_frame, text=" BƯỚC 2: CẬP NHẬT KẾT QUẢ KIỂM TRA & XUẤT BÁO CÁO - DANH SÁCH LỖI ")
    f2.pack(fill="x", pady=5)
    f2.columnconfigure(1, weight=1)

    lbl_f_tong = tk.Label(f2, text="File 1_FILE_TONG.xlsx:", font=("Segoe UI", 9))
    lbl_f_tong.grid(row=0, column=0, padx=8, pady=4, sticky="w")
    txt_tong = ttk.Entry(f2, width=54)
    txt_tong.grid(row=0, column=1, padx=4, pady=4)

    def chon_file_tong():
        target = txt_tong.get().strip() or txt_folder.get().strip()
        first_item = target.split(";")[0].strip() if target else ""
        init_dir = first_item if os.path.isdir(first_item) else os.path.dirname(first_item) if os.path.isfile(first_item) else os.path.expanduser("~/Desktop")
        f = filedialog.askopenfilename(
            title="Chọn file 1_FILE_TONG.xlsx",
            filetypes=[("Excel files", "*.xlsx;*.xlsm"), ("All files", "*.*")],
            initialdir=init_dir
        )
        if f:
            txt_tong.delete(0, tk.END)
            txt_tong.insert(0, f)
    tk.Button(f2, text="🔍 Chọn...", command=chon_file_tong, **btn_sec_cfg).grid(row=0, column=2, padx=4, pady=4)

    lbl_f_kt = tk.Label(f2, text="File 2_DANH_SACH đã ghi chú:", font=("Segoe UI", 9))
    lbl_f_kt.grid(row=1, column=0, padx=8, pady=4, sticky="w")
    txt_kt = ttk.Entry(f2, width=54)
    txt_kt.grid(row=1, column=1, padx=4, pady=4)

    def chon_file_kt():
        target = txt_kt.get().strip() or txt_folder.get().strip()
        first_item = target.split(";")[0].strip() if target else ""
        init_dir = first_item if os.path.isdir(first_item) else os.path.dirname(first_item) if os.path.isfile(first_item) else os.path.expanduser("~/Desktop")
        f = filedialog.askopenfilename(
            title="Chọn file 2_DANH_SACH_CAN_KIEM_TRA.xlsx đã ghi chú",
            filetypes=[("Excel files", "*.xlsx;*.xlsm"), ("All files", "*.*")],
            initialdir=init_dir
        )
        if f:
            txt_kt.delete(0, tk.END)
            txt_kt.insert(0, f)
    tk.Button(f2, text="🔍 Chọn...", command=chon_file_kt, **btn_sec_cfg).grid(row=1, column=2, padx=4, pady=4)

    # Nút 2: Nhỏ gọn, màu Cam Ấm (#EA580C) đồng bộ kích thước và hiển thị 100% đầy đủ chữ
    frm_action2 = tk.Frame(f2)
    frm_action2.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(5, 8))
    btn_step2 = tk.Button(
        frm_action2,
        text="📊 Thực hiện Bước 2: Cập nhật kết quả & Xuất Báo cáo - Danh sách lỗi",
        bg="#EA580C",
        activebackground="#C2410C",
        activeforeground="white",
        fg="white",
        font=("Segoe UI", 9, "bold"),
        relief="flat",
        bd=0,
        padx=24,
        pady=5,
        cursor="hand2"
    )
    btn_step2.pack()

    def on_enter_b2(e):
        if str(btn_step2["state"]) != "disabled":
            btn_step2.config(bg="#C2410C")
    def on_leave_b2(e):
        if str(btn_step2["state"]) != "disabled":
            btn_step2.config(bg="#EA580C")
    btn_step2.bind("<Enter>", on_enter_b2)
    btn_step2.bind("<Leave>", on_leave_b2)

    # Khẩu hiệu 10 chữ vàng Bác Hồ dạy cán bộ ngành Kiểm sát (Hàng trên to hơn, căn giữa lấp đầy khoảng trống)
    frm_motto = tk.Frame(main_frame)
    frm_motto.pack(fill="x", expand=True, pady=(6, 2))

    lbl_motto_title = tk.Label(
        frm_motto,
        text="10 CHỮ VÀNG BÁC HỒ DẠY CÁN BỘ NGÀNH KIỂM SÁT",
        font=("Segoe UI", 11, "bold"),
        fg="#0B2545"
    )
    lbl_motto_title.pack(anchor="center")

    lbl_motto = tk.Label(
        frm_motto,
        text="CÔNG MINH - CHÍNH TRỰC - KHÁCH QUAN - THẬN TRỌNG - KHIÊM TỐN",
        font=("Segoe UI", 9, "bold"),
        fg="#B91C1C"
    )
    lbl_motto.pack(anchor="center", pady=(4, 0))

    # Thanh trạng thái & Progressbar
    status_frame = tk.Frame(root, bg="#ECEFF1", padx=10, pady=5)
    status_frame.pack(fill="x", side="bottom")

    progress = ttk.Progressbar(status_frame, mode="indeterminate")
    progress.pack(fill="x", pady=2)

    lbl_status = tk.Label(status_frame, text="Hệ thống sẵn sàng.", font=("Segoe UI", 9, "italic"), bg="#ECEFF1", fg="#37474F", anchor="w")
    lbl_status.pack(side="left", fill="x", expand=True)

    def mo_thu_muc():
        target_dir = txt_folder.get().strip()
        if not target_dir or not os.path.exists(target_dir):
            if txt_tong.get().strip() and os.path.exists(txt_tong.get().strip()):
                target_dir = os.path.dirname(txt_tong.get().strip())
        if target_dir and os.path.exists(target_dir):
            first_item = target_dir.split(";")[0].strip()
            final_dir = first_item if os.path.isdir(first_item) else os.path.dirname(first_item)
            if final_dir and os.path.exists(final_dir):
                os.startfile(final_dir)
        else:
            messagebox.showinfo("Thông báo", "Vui lòng chọn thư mục trước khi mở!")

    btn_open_folder = tk.Button(
        status_frame,
        text="📂 Mở thư mục kết quả",
        command=mo_thu_muc,
        font=("Segoe UI", 8, "bold"),
        bg="#FFFFFF",
        fg="#1E293B",
        relief="solid",
        bd=1,
        padx=8,
        pady=2,
        cursor="hand2"
    )
    btn_open_folder.pack(side="right", padx=4)

    # ---------------- ĐA LUỒNG XỬ LÝ (CHỐNG TREO GIAO DIỆN) ----------------

    def set_running_state(is_running, status_text=""):
        if is_running:
            btn_step1.config(state="disabled", bg="#94A3B8")
            btn_step2.config(state="disabled", bg="#94A3B8")
            progress.start(10)
            lbl_status.config(text=f"⏳ {status_text}", fg="#EA580C")
        else:
            btn_step1.config(state="normal", bg="#0284C7")
            btn_step2.config(state="normal", bg="#EA580C")
            progress.stop()
            lbl_status.config(text=f"✅ {status_text}", fg="#16A34A")

    def worker_buoc_1():
        d = txt_folder.get().strip()
        if not d:
            messagebox.showwarning("Cảnh báo", "Vui lòng bấm 'Thư mục' hoặc 'Chọn file' để chọn nguồn dữ liệu báo cáo của đơn vị!")
            set_running_state(False, "Chưa chọn file hoặc thư mục.")
            return

        try:
            days = int(spn_days.get().strip())
        except ValueError:
            days = 16

        existing_m = txt_existing.get().strip()
        existing_m_clean = existing_m if (existing_m and os.path.exists(existing_m)) else None

        # Xác định thư mục đầu ra
        first_item = d.split(";")[0].strip()
        out_dir = first_item if os.path.isdir(first_item) else os.path.dirname(first_item)
        if not out_dir or not os.path.exists(out_dir):
            out_dir = os.getcwd()

        use_ts = var_timestamp.get()
        standard_kt_path = os.path.join(out_dir, "2_DANH_SACH_CAN_KIEM_TRA.xlsx")

        # XÁC NHẬN NẾU FILE ĐÃ TỒN TẠI TỪ TRƯỚC
        if not use_ts and os.path.exists(standard_kt_path):
            ans = messagebox.askyesno(
                "Xác nhận cập nhật",
                f"File '2_DANH_SACH_CAN_KIEM_TRA.xlsx' đã tồn tại trong thư mục này.\n\n"
                f"Bạn đã cập nhật thông tin trước đó rồi, có cần làm lại không?\n\n"
                f"• Chọn [Yes]: Đồng ý làm lại (hệ thống sẽ tự động sao lưu bản cũ vào '📁 Backup/').\n"
                f"• Chọn [No]: Dừng lại, giữ nguyên file hiện tại.",
                icon="question"
            )
            if not ans:
                set_running_state(False, "Đã dừng thao tác, giữ nguyên file hiện tại.")
                return

        set_running_state(True, "Đang đối chiếu và tổng hợp dữ liệu...")

        def run():
            ok, res = buoc_1_tong_hop_va_loc(
                d, check_days=days,
                existing_master_path=existing_m_clean,
                use_timestamp=use_ts
            )
            root.after(0, lambda: set_running_state(False, "Đã hoàn thành Bước 1!"))
            if ok:
                file_1 = (res.get("file_tong") if isinstance(res, dict) else "") or os.path.join(out_dir, "1_FILE_TONG.xlsx")
                file_2 = (res.get("file_kt") if isinstance(res, dict) else "") or os.path.join(out_dir, "2_DANH_SACH_CAN_KIEM_TRA.xlsx")
                txt_tong.delete(0, tk.END); txt_tong.insert(0, file_1)
                txt_kt.delete(0, tk.END); txt_kt.insert(0, file_2)

                backup_note = ""
                if isinstance(res, dict) and res.get("backup_files"):
                    b_list = "\n   • ".join(os.path.basename(b) for b in res["backup_files"])
                    backup_note = f"\n\n🛡️ AN TOÀN DỮ LIỆU: Đã tự động sao lưu bản cũ vào thư mục '📁 Backup/':\n   • {b_list}"

                if isinstance(res, dict):
                    noted = res.get("noted_cases", 0)
                    unnoted = res.get("unnoted_cases", 0)
                    filtered = res.get("cases_filtered", 0)
                    under_days = res.get("under_days_cases", max(0, unnoted - filtered))

                    if res.get("mode") == "single_master":
                        msg_text = (
                            f"🎉 ĐÃ XỬ LÝ FILE DỮ LIỆU CÓ SẴN THÀNH CÔNG!\n\n"
                            f"• File nguồn: {os.path.basename(d)}\n"
                            f"• Tổng số vụ án trong file: {res.get('total_cases', 0):,} vụ\n"
                            f"• Số vụ ĐÃ GHI CHÚ làm giàu dữ liệu: {noted:,} vụ (bảo lưu nguyên vẹn)\n"
                            f"• Số vụ CHƯA GHI CHÚ làm giàu dữ liệu: {unnoted:,} vụ\n"
                            f"  ↳ Đủ từ {days} ngày trở đi (CẦN KIỂM TRA): {filtered:,} VỤ ÁN\n"
                            f"  ↳ Dưới {days} ngày (chưa đến hạn kiểm tra): {under_days:,} vụ\n\n"
                            f"📁 ĐÃ TỰ ĐỘNG XUẤT 2 FILE TẠI THƯ MỤC:\n"
                            f"1. {os.path.basename(file_1)} (Bảo toàn đầy đủ tất cả {noted:,} ghi chú)\n"
                            f"2. {os.path.basename(file_2)} (Đúng {filtered:,} vụ cần rà soát kiểm tra)"
                            f"{backup_note}"
                        )
                    elif res.get("mode") == "append" or res.get("total_existing", 0) > 0:
                        total_ex = res.get("total_existing", 0)
                        new_add = res.get("new_added", 0)
                        upd = res.get("updated_cases", 0)
                        tot = res.get("total_cases", 0)
                        file_new = res.get("file_moi_phat_sinh")

                        if file_new and os.path.exists(file_new):
                            out_files_desc = (
                                f"📁 ĐÃ TỰ ĐỘNG XUẤT 3 FILE TẠI THƯ MỤC:\n"
                                f"1. {os.path.basename(file_1)} (File Tổng gồm {tot:,} vụ đã nạp thêm)\n"
                                f"2. {os.path.basename(file_2)} (Gồm {filtered:,} vụ cần rà soát)\n"
                                f"3. {os.path.basename(file_new)} (Danh sách riêng {new_add:,} vụ án mới phát sinh)"
                            )
                        else:
                            out_files_desc = (
                                f"📁 ĐÃ TỰ ĐỘNG XUẤT 2 FILE TẠI THƯ MỤC:\n"
                                f"1. {os.path.basename(file_1)} (File Tổng gồm {tot:,} vụ đã nạp thêm)\n"
                                f"2. {os.path.basename(file_2)} (Gồm {filtered:,} vụ cần rà soát)"
                            )

                        msg_text = (
                            f"🎉 TỔNG HỢP VÀ NẠP NỐI TIẾP THÀNH CÔNG!\n\n"
                            f"📋 CHI TIẾT ĐỐI CHIẾU & NẠP THÊM:\n"
                            f"• Số file đơn vị đã xử lý: {res.get('total_files')} file\n"
                            f"• Vụ cũ sẵn có từ File Tổng: {total_ex:,} vụ\n"
                            f"• ĐÃ NẠP THÊM MỚI: +{new_add:,} VỤ ÁN MỚI PHÁT SINH\n"
                            f"• Đã cập nhật/đối chiếu thông tin: {upd:,} vụ\n"
                            f"👉 TỔNG SỐ VỤ ÁN SAU KHI GỘP: {tot:,} VỤ ÁN\n\n"
                            f"📊 KẾT QUẢ DANH SÁCH KIỂM TRA:\n"
                            f"• Số vụ đã có ghi chú trước đó: {noted:,} vụ (bảo lưu 100%)\n"
                            f"• Số vụ cần kiểm tra (>={days} ngày & chưa ghi chú): {filtered:,} VỤ ÁN\n\n"
                            f"{out_files_desc}"
                            f"{backup_note}"
                        )
                    else:
                        msg_text = (
                            f"🎉 TỔNG HỢP VÀ LỌC ÁN THÀNH CÔNG!\n\n"
                            f"• Số đơn vị / file đã xử lý: {res.get('total_files')} file\n"
                            f"• Tổng số vụ án hình sự: {res.get('total_cases', 0):,} vụ\n"
                            f"• Số vụ đã có ghi chú: {noted:,} vụ\n"
                            f"• Số vụ cần kiểm tra (>={days} ngày & chưa ghi chú): {filtered:,} vụ\n\n"
                            f"📁 ĐÃ TỰ ĐỘNG XUẤT 2 FILE TẠI THƯ MỤC:\n"
                            f"1. {os.path.basename(file_1)} (Chuẩn 22 cột nền vàng)\n"
                            f"2. {os.path.basename(file_2)} (Gồm {filtered:,} vụ cần rà soát)"
                            f"{backup_note}"
                        )
                else:
                    msg_text = str(res)
                root.after(0, lambda: messagebox.showinfo("Thành công - Bước 1", msg_text))
            else:
                err_text = res.get("message", str(res)) if isinstance(res, dict) else str(res)
                root.after(0, lambda: messagebox.showerror("Lỗi", err_text))

        threading.Thread(target=run, daemon=True).start()

    def worker_buoc_2():
        ft = txt_tong.get().strip()
        fk = txt_kt.get().strip()
        if not ft or not os.path.exists(ft) or not fk or not os.path.exists(fk):
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn đầy đủ cả 2 file:\n1. File 1_FILE_TONG.xlsx\n2. File 2_DANH_SACH_CAN_KIEM_TRA.xlsx đã ghi chú!")
            return

        use_ts = var_timestamp.get()
        set_running_state(True, "Đang cập nhật kết quả và tạo Báo cáo Infographic & File khắc phục lỗi...")

        def run():
            ok, res = buoc_2_cap_nhat_va_bao_cao(ft, fk, use_timestamp=use_ts)
            root.after(0, lambda: set_running_state(False, "Đã hoàn thành Bước 2!"))
            if ok:
                backup_note = ""
                if isinstance(res, dict) and res.get("backup_files"):
                    b_list = "\n   • ".join(os.path.basename(b) for b in res["backup_files"])
                    backup_note = f"\n\n🛡️ AN TOÀN DỮ LIỆU: Đã tự động sao lưu bản cũ vào thư mục '📁 Backup/':\n   • {b_list}"

                f_master_name = os.path.basename(res.get('file_master', '3_FILE_TONG_HOAN_THIEN.xlsx')) if isinstance(res, dict) else "3_FILE_TONG_HOAN_THIEN.xlsx"
                f_rep_name = os.path.basename(res.get('file_report', '4_BAO_CAO_VU_LOI_INFOGRAPHIC.xlsx')) if isinstance(res, dict) else "4_BAO_CAO_VU_LOI_INFOGRAPHIC.xlsx"
                f_rem_name = os.path.basename(res.get('file_remedy', '5_DANH_SACH_VU_LOI_GUI_CAC_DON_VI.xlsx')) if isinstance(res, dict) else "5_DANH_SACH_VU_LOI_GUI_CAC_DON_VI.xlsx"

                if isinstance(res, dict):
                    auto_cnt = res.get("auto_filled_dates", 0)
                    auto_note = ""
                    if auto_cnt > 0:
                        auto_note = f"\n   • 🎯 TỰ ĐỘNG BÙ NGÀY THAO TÁC ({res.get('target_date')}): {auto_cnt:,} vụ (do có ghi chú nhưng quên điền ngày)"

                    msg_text = (
                        f"🎉 CẬP NHẬT KẾT QUẢ & XUẤT BÁO CÁO THÀNH CÔNG!\n\n"
                        f"📅 Ngày kiểm tra ghi nhận: {res.get('target_date')} (Ngày thao tác trên phần mềm)\n"
                        f"📋 Thống kê đợt kiểm tra:\n"
                        f"   • Tổng số vụ đã kiểm tra: {res.get('total_checked')} vụ\n"
                        f"   • Đã đính kèm (Đạt): {res.get('total_dat')} vụ\n"
                        f"   • Phát hiện lỗi: {res.get('total_errors')} vụ ({res.get('error_rate')})"
                        f"{auto_note}\n\n"
                        f"📁 CÁC FILE KẾT QUẢ ĐÃ TẠO TẠI THƯ MỤC:\n"
                        f"1. {f_master_name}\n"
                        f"   (File Tổng hoàn thiện chuẩn 22 cột nền vàng, tự động điền đủ ngày kiểm tra + Sheet 'BẢNG THEO DÕI')\n\n"
                        f"2. {f_rep_name}\n"
                        f"   (Báo cáo Lãnh đạo kèm Thẻ KPI & Bảng chi tiết các vụ lỗi)\n\n"
                        f"3. {f_rem_name}\n"
                        f"   (File tổng 7 cột chuẩn nền xanh gửi các đơn vị yêu cầu khắc phục)\n\n"
                        f"4. Thư mục: Cac_Don_Vi_Khac_Phuc/\n"
                        f"   (Tự động chia tách từng file riêng cho từng đơn vị có lỗi để gửi ngay)"
                        f"{backup_note}"
                    )
                else:
                    msg_text = str(res)
                root.after(0, lambda: messagebox.showinfo("Hoàn thành - Bước 2", msg_text))
            else:
                err_text = res.get("message", str(res)) if isinstance(res, dict) else str(res)
                root.after(0, lambda: messagebox.showerror("Lỗi", err_text))

        threading.Thread(target=run, daemon=True).start()

    btn_step1.config(command=worker_buoc_1)
    btn_step2.config(command=worker_buoc_2)

    root.mainloop()

# ==================== CHẠY DÒNG LỆNH DỰ PHÒNG ====================

def run_cli():
    print("=" * 68)
    print("  TỔNG HỢP VÀ GIÁM SÁT VIỆC LẬP HỒ SƠ KIỂM SÁT ĐIỆN TỬ")
    print("  VIỆN KIỂM SÁT NHÂN DÂN THÀNH PHỐ HỒ CHÍ MINH")
    print("=" * 68)
    print("1. [NÚT 1] Tổng hợp 23 file đơn vị -> Xuất Danh sách kiểm tra (>= 16 ngày)")
    print("2. [NÚT 2] Đẩy file kiểm tra vào -> Cập nhật File Tổng & Xuất Báo cáo Vụ Lỗi")
    print("0. Thoát")
    print("-" * 65)

    choice = input("Vui lòng chọn chức năng (1 hoặc 2): ").strip()
    if choice == "1":
        folder = input("Nhập đường dẫn thư mục hoặc file đơn vị: ").strip()
        if not folder:
            print("❌ Chưa nhập nguồn file!")
            return
        days_str = input("Nhập số ngày kiểm tra (mặc định 16): ").strip()
        days = int(days_str) if days_str.isdigit() else 16
        print("⏳ Đang xử lý...")
        ok, msg = buoc_1_tong_hop_va_loc(folder, check_days=days)
        print(msg)
    elif choice == "2":
        f_tong = input("Nhập đường dẫn file 1_FILE_TONG.xlsx: ").strip()
        f_kt = input("Nhập đường dẫn file 2_DANH_SACH_CAN_KIEM_TRA.xlsx: ").strip()
        if not os.path.exists(f_tong) or not os.path.exists(f_kt):
            print("❌ File không tồn tại!")
            return
        print("⏳ Đang xử lý...")
        ok, msg = buoc_2_cap_nhat_va_bao_cao(f_tong, f_kt)
        print(msg)
    else:
        print("Tạm biệt!")

if __name__ == "__main__":
    try:
        run_gui()
    except Exception as e:
        run_cli()
