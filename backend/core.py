# -*- coding: utf-8 -*-
"""
HỆ THỐNG TỰ ĐỘNG HÓA TỔNG HỢP & GIÁM SÁT TIẾN ĐỘ ÁN HÌNH SỰ - MODULE BACKEND LÕI
Viện kiểm sát nhân dân Thành phố Hồ Chí Minh

QUY CHUẨN ĐỊNH DẠNG MASTER FILE (22 CỘT CHUẨN - ĐÃ BỎ HOÀN TOÀN CỘT NOTE LỖI):
1. STT
2. Thông tin vụ án          (Giữ nguyên hiện trạng từ file xuất xuống của đơn vị)
3. Loại tội                 (Giữ nguyên hiện trạng từ file xuất xuống của đơn vị)
4. Tội                      (Giữ nguyên hiện trạng từ file xuất xuống của đơn vị)
5. Thông tin bị can         (Giữ nguyên hiện trạng từ file xuất xuống của đơn vị)
6. Loại thụ lý              (Giữ nguyên hiện trạng từ file xuất xuống của đơn vị)
7. Trạng thái vụ án         (Giữ nguyên hiện trạng từ file xuất xuống của đơn vị)
8. Trạng thái biểu mẫu      (Giữ nguyên hiện trạng từ file xuất xuống của đơn vị)
9. MÃ VỤ ÁN                 (Bóc tách từ ô Thông tin vụ án)
10. TÊN VỤ ÁN               (Bóc tách từ ô Thông tin vụ án)
11. SỐ KTVA                 (Bóc tách từ ô Thông tin vụ án)
12. NGÀY KTVA               (Bóc tách từ ô Thông tin vụ án, định dạng dd/mm/yyyy)
13. NGÀY TIẾP NHẬN HỐ SƠ    (Bóc tách từ ô Thông tin vụ án, định dạng dd/mm/yyyy)
14. THỜI HẠN XỬ LÝ          (Bóc tách từ ô Thông tin vụ án, định dạng dd/mm/yyyy)
15. NOTE LÀM GIÀU DỮ LIỆU   (Chỉ điền "ĐÃ ĐÍNH KÈM", "TẠM ỔN" là không tính lỗi; có thông tin khác như LƯU TÊN FILE "DOCUMENT", Chưa đính kèm... là tính có lỗi)
16. NGÀY HIỆN TẠI SO VỚI NGÀY KTVA (Số ngày chênh lệch tính đến ngày hiện tại)
17. ĐƠN VỊ                  (Tên đơn vị gửi báo cáo, ví dụ KHU VỰC 1, KHU VỰC 2, PHÒNG 1...)
18. NOTE NGÀY KIỂM TRA      (Ngày kiểm tra dd/mm/yyyy, để trống khi chưa kiểm tra)
19. NOTE                    (Ghi chú tự do, ví dụ "ĐÃ ĐĂNG KÝ XÓA"...)
20. GHI CHÚ ĐỂ BÁO TÌNH HÌNH (Ghi chú cụ thể của cán bộ kiểm tra, lý do thay đổi hằng ngày)
21. NOTE KIỂM TRA           (Trạng thái kiểm tra: "ĐÃ KIỂM TRA", để trống khi chưa kiểm tra)
22. NGÀY BÁO TÌNH HÌNH      (Ngày báo cáo dd/mm/yyyy)

Định dạng giao diện Excel:
- Tên Sheet dữ liệu chính: DATA
- Tiêu đề nền VÀNG (#FFFF00), chữ đen in đậm (Segoe UI 10pt)
- Tự động bật mũi tên lọc AutoFilter ở dòng tiêu đề
- Bật hiển thị lưới ô (gridlines)
- Tự động tạo Sheet "BẢNG THEO DÕI" (Pivot Table theo Đơn vị: Count of MÃ VỤ ÁN, Đạt, Lỗi)
- Áp dụng thống nhất cho cả File Tổng và File Danh sách kiểm tra
"""

import os
import re
import shutil
import unicodedata
from datetime import datetime, date, timedelta
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
from openpyxl.worksheet.datavalidation import DataValidation

# DANH SÁCH 22 CỘT CHUẨN MỰC (ĐÃ LOẠI BỎ HẲN CỘT NOTE LỖI THEO YÊU CẦU CỦA NGƯỜI DÙNG)
MASTER_COLUMNS = [
    "STT",
    "Thông tin vụ án",
    "Loại tội",
    "Tội",
    "Thông tin bị can",
    "Loại thụ lý",
    "Trạng thái vụ án",
    "Trạng thái biểu mẫu",
    "MÃ VỤ ÁN",
    "TÊN VỤ ÁN",
    "SỐ KTVA",
    "NGÀY KTVA",
    "NGÀY TIẾP NHẬN HỐ SƠ",
    "THỜI HẠN XỬ LÝ",
    "NOTE LÀM GIÀU DỮ LIỆU",
    "NGÀY HIỆN TẠI SO VỚI NGÀY KTVA",
    "ĐƠN VỊ",
    "NOTE NGÀY KIỂM TRA",
    "NOTE",
    "GHI CHÚ ĐỂ BÁO TÌNH HÌNH",
    "NOTE KIỂM TRA",
    "NGÀY BÁO TÌNH HÌNH",
]

def remove_accents(input_str):
    """Loại bỏ dấu tiếng Việt và chuyển sang chữ HOA để so sánh chuỗi chính xác"""
    if not input_str:
        return ""
    nfkd = unicodedata.normalize('NFKD', str(input_str))
    cleaned = "".join([c for c in nfkd if not unicodedata.combining(c)])
    return cleaned.replace("Đ", "D").replace("đ", "d").upper().strip()

def parse_date(val):
    """Chuyển đổi đa dạng kiểu ngày tháng thành datetime.date an toàn"""
    if val is None or pd.isna(val):
        return None
    if isinstance(val, (datetime, pd.Timestamp)):
        return val.date()
    if isinstance(val, date):
        return val

    # Số serial của Excel (ví dụ 45200)
    if isinstance(val, (int, float)):
        try:
            if 20000 <= val <= 80000:
                base_date = datetime(1899, 12, 30)
                return (base_date + timedelta(days=float(val))).date()
        except Exception:
            pass

    val_str = str(val).strip()
    if not val_str or val_str.lower() in ("none", "nan", "nat", ""):
        return None

    val_date_part = re.split(r"[\sT]", val_str)[0].strip()
    formats = (
        "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y/%m/%d",
        "%d/%m/%y", "%m/%d/%Y", "%d.%m.%Y"
    )
    for fmt in formats:
        try:
            return datetime.strptime(val_date_part, fmt).date()
        except ValueError:
            pass

    match_dmy = re.search(r"(\d{1,2})[/-](\d{1,2})[/-](\d{4})", val_str)
    if match_dmy:
        d, m, y = match_dmy.groups()
        try:
            return date(int(y), int(m), int(d))
        except ValueError:
            pass

    return None

def find_header_row(ws):
    """Quét 25 dòng đầu để tìm dòng tiêu đề chuẩn xác"""
    keywords = ["MA VU AN", "THONG TIN VU AN", "SO THU LY", "SO KTVA", "TEN VU AN", "TOI DANH"]
    for row_idx, row in enumerate(ws.iter_rows(max_row=25, values_only=True), 1):
        row_str = " ".join([remove_accents(str(c)) for c in row if c is not None])
        if any(kw in row_str for kw in keywords):
            return row_idx
    return 1

def detect_column_indices(headers):
    """Khớp tên cột thông minh dựa trên từ khóa không dấu và tên chuẩn 22 cột"""
    col_map = {}
    normalized_headers = [remove_accents(str(h)).upper().strip() for h in headers]
    raw_upper_headers = [str(h).upper().strip() for h in headers]

    # 1. Khớp chính xác tên cột chuẩn trong MASTER_COLUMNS trước
    master_keys_map = {
        "THÔNG TIN VỤ ÁN": "INFO_VA",
        "LOẠI TỘI": "LOAI_TOI",
        "TỘI": "TOI",
        "THÔNG TIN BỊ CAN": "BI_CAN",
        "LOẠI THỤ LÝ": "THU_LY",
        "TRẠNG THÁI VỤ ÁN": "TRANG_THAI_VA",
        "TRẠNG THÁI BIỂU MẪU": "TRANG_THAI_BM",
        "MÃ VỤ ÁN": "MA_VA",
        "TÊN VỤ ÁN": "TEN_VA",
        "SỐ KTVA": "SO_KTVA",
        "NGÀY KTVA": "NGAY_KTVA",
        "NGÀY TIẾP NHẬN HỐ SƠ": "NGAY_TN",
        "THỜI HẠN XỬ LÝ": "HAN_XL",
        "NOTE LÀM GIÀU DỮ LIỆU": "NOTE_LAM_GIAU",
        "NGÀY HIỆN TẠI SO VỚI NGÀY KTVA": "NGAY_SO_VOI_KTVA",
        "ĐƠN VỊ": "DON_VI",
        "NOTE NGÀY KIỂM TRA": "NOTE_NGAY_KT",
        "NOTE": "NOTE",
        "GHI CHÚ ĐỂ BÁO TÌNH HÌNH": "GHI_CHU_BAO_CAO",
        "NOTE KIỂM TRA": "NOTE_KT",
        "NGÀY BÁO TÌNH HÌNH": "NGAY_BAO_TH"
    }

    for idx, (h_raw, h_norm) in enumerate(zip(raw_upper_headers, normalized_headers), 1):
        for m_col, key in master_keys_map.items():
            m_norm = remove_accents(m_col).upper().strip()
            if h_raw == m_col or h_norm == m_norm:
                if key not in col_map:
                    col_map[key] = idx
                    break

    # 2. Khớp theo từ khóa nếu chưa tìm thấy cột
    keyword_rules = {
        "INFO_VA": ["THONG TIN VU AN", "THONG TIN VU", "HO SO VU AN", "THONG TIN CHUNG"],
        "LOAI_TOI": ["LOAI TOI", "PHAN LOAI TOI"],
        "TOI": ["TOI DANH", "TOI PHAM", "TOI", "DIEU LUAT"],
        "BI_CAN": ["BI CAN", "THONG TIN BI CAN", "HO TEN BI CAN", "NGUOI BI BUOC TOI"],
        "THU_LY": ["LOAI THU LY", "THU LY", "GIAI DOAN"],
        "TRANG_THAI_VA": ["TRANG THAI VU AN", "TINH TRANG VU AN", "TRANG THAI", "TIEN DO"],
        "TRANG_THAI_BM": ["TRANG THAI BIEU MAU", "BIEU MAU"],
        "MA_VA": ["MA VU AN", "MA VU", "MA AN", "SO THU LY", "MA HO SO"],
        "TEN_VA": ["TEN VU AN", "TEN VU", "TEN AN", "NOI DUNG VU AN", "TEN VU VIEC"],
        "SO_KTVA": ["SO KTVA", "SO QD KTVA", "SO QUYET DINH KHOI TO", "SO KHOI TO", "SO QD"],
        "NGAY_KTVA": ["NGAY KTVA", "NGAY QD KTVA", "NGAY KHOI TO", "NGAY QUYET DINH", "NGAY RA QD"],
        "NGAY_TN": ["NGAY TIEP NHAN", "TIEP NHAN", "NGAY THU LY", "NGAY NHAN HO SO"],
        "HAN_XL": ["THOI HAN XU LY", "HAN XU LY", "THOI HAN", "HAN GIAI QUYET"],
        "DON_VI": ["DON VI", "CO QUAN", "VIEN KIEM SAT", "DON VI THU LY", "DON VI XU LY"],
        "NOTE_LAM_GIAU": ["NOTE LAM GIAU DU LIEU", "NOTE LAM GIAU", "LAM GIAU DU LIEU", "LAM GIAU", "NOTE GIAU"],
        "NGAY_SO_VOI_KTVA": ["NGAY HIEN TAI SO VOI NGAY KTVA", "NGAY HIEN TAI SO VOI", "SO VOI NGAY KTVA", "SO NGAY KTVA"],
        "NOTE_NGAY_KT": ["NOTE NGAY KIEM TRA", "NGAY KIEM TRA"],
        "GHI_CHU_BAO_CAO": ["GHI CHU DE BAO TINH HINH", "GHI CHU BAO TINH HINH", "GHI CHU DE BAO", "BAO TINH HINH"],
        "NOTE_KT": ["NOTE KIEM TRA"],
        "NGAY_BAO_TH": ["NGAY BAO TINH HINH", "NGAY BAO"],
        "NOTE": ["NOTE"]
    }

    for key, kw_list in keyword_rules.items():
        if key in col_map:
            continue
        for idx, h_norm in enumerate(normalized_headers, 1):
            if key == "TOI" and "LOAI" in h_norm:
                continue  # Bỏ qua cột Loại tội khi tìm Tội danh
            if key == "NOTE" and any(k in h_norm for k in ["LAM GIAU", "NGAY KIEM TRA", "KIEM TRA"]):
                continue  # Bỏ qua các cột note đặc thù
            if any(kw in h_norm for kw in kw_list):
                col_map[key] = idx
                break

    return col_map

def standardize_unit_name(val):
    """
    Chuẩn hóa tên đơn vị theo đúng yêu cầu nghiệp vụ:
    - CHỈ GIỮ LẠI KHU VỰC VÀ SỐ (ví dụ: KHU VỰC 1, KHU VỰC 2, ..., KHU VỰC 19)
    - TƯƠNG TỰ VỚI PHÒNG VÀ SỐ (ví dụ: PHÒNG 1, PHÒNG 2, PHÒNG 3)
    - Tự động cắt bỏ các hậu tố nghiệp vụ như .DT, .ST, .TT, _DT, ...
    - Chuẩn hóa các cách viết tắt: KV 1, KV.1, KV1, P 1, P.1, PHONG 1...
    """
    if val is None or pd.isna(val):
        return ""
    s = str(val).strip()
    if not s or s.lower() in ("nan", "none", ""):
        return ""
    s_upper = s.upper()

    # 1. Mẫu KHU VỰC kèm số (ví dụ: KHU VỰC 1.DT, KHU VỰC 1.ST, KHU VỰC 1_TT, KV 1, KV1.DT...)
    m_kv = re.search(r'(?:KHU\s*V[ỰU]C|KV)\s*[\.\_\-\s]*(\d+)', s_upper)
    if m_kv:
        return f"KHU VỰC {int(m_kv.group(1))}"

    # 2. Mẫu PHÒNG kèm số (ví dụ: PHÒNG 1.DT, PHÒNG 1.ST, PHÒNG 1, P 1, P.1...)
    m_p = re.search(r'(?:PH[ÒO]NG|P)\s*[\.\_\-\s]*(\d+)', s_upper)
    if m_p:
        return f"PHÒNG {int(m_p.group(1))}"

    return s.strip()

def get_best_case_sheet(wb):
    """Tìm sheet phù hợp nhất chứa dữ liệu các vụ án trong workbook openpyxl"""
    # 1. Ưu tiên số 1: Sheet tên 'DATA' (chuẩn mực của hệ thống)
    for s_name in wb.sheetnames:
        if s_name.strip().upper() == "DATA":
            return wb[s_name]

    # 2. Ưu tiên số 2: Sheet chứa header MÃ VỤ ÁN hoặc THÔNG TIN VỤ ÁN và có nhiều dòng nhất
    candidates = []
    for s_name in wb.sheetnames:
        ws = wb[s_name]
        h_row = find_header_row(ws)
        if h_row:
            row_vals = [str(c).strip().upper() for c in next(ws.iter_rows(min_row=h_row, max_row=h_row, values_only=True), []) if c is not None]
            if any("MÃ VỤ ÁN" in v or "THÔNG TIN VỤ ÁN" in v or "MA VU AN" in remove_accents(v) for v in row_vals):
                candidates.append((ws.max_row or 0, ws))
    if candidates:
        candidates.sort(key=lambda x: x[0], reverse=True)
        return candidates[0][1]

    return wb.active

def load_existing_master_df(file_path):
    """
    Tải File Tổng sẵn có (ví dụ 11082026.xlsx có hơn 3200 vụ án):
    - Tự động tìm đúng Sheet 'DATA' chứa dữ liệu các vụ án (tránh đọc nhầm sheet báo cáo/tổng hợp chỉ có 27 dòng)
    - Chuẩn hóa tên cột
    - Lấy đủ 22 cột quy chuẩn
    - Chuẩn hóa cột ĐƠN VỊ: Chỉ giữ lại KHU VỰC và số (KHU VỰC 1, KHU VỰC 2...), PHÒNG và số (PHÒNG 1, PHÒNG 2...)
    """
    wb = openpyxl.load_workbook(file_path, read_only=True)
    best_sheet = None
    for s in wb.sheetnames:
        if s.strip().upper() == "DATA":
            best_sheet = s
            break
    if not best_sheet:
        max_rows = -1
        best_sheet = wb.sheetnames[0]
        for s in wb.sheetnames:
            ws = wb[s]
            r_count = ws.max_row or 0
            if r_count > max_rows:
                max_rows = r_count
                best_sheet = s
    wb.close()

    df = pd.read_excel(file_path, sheet_name=best_sheet)
    df = normalize_column_names(df)

    for col in MASTER_COLUMNS:
        if col in df.columns:
            df[col] = df[col].astype("object").fillna("")
        else:
            df[col] = ""

    if "ĐƠN VỊ" in df.columns:
        df["ĐƠN VỊ"] = df["ĐƠN VỊ"].apply(standardize_unit_name)

    return df

def extract_case_details_from_cell(info_va, toi_val=""):
    """
    Bóc tách toàn diện các thông tin vụ án từ ô 'Thông tin vụ án'
    (Định dạng chuẩn kết xuất từ phần mềm Quản lý án của VKSND)
    """
    if not info_va or pd.isna(info_va):
        return {}

    lines = [l.strip() for l in str(info_va).split("\n") if l.strip()]
    if not lines:
        return {}

    result = {
        "ma_va": "",
        "ten_va": "",
        "so_kt": "",
        "ngay_kt": None,
        "ngay_tn": None,
        "han_xl": None
    }

    # 1. Mã vụ án (Tìm dòng bắt đầu bằng VA_ hoặc chuỗi mã chữ-số)
    for line in lines:
        if line.upper().startswith("VA_") or re.match(r"^[A-Z0-9_\-\.]{6,}$", line):
            result["ma_va"] = line
            break

    if not result["ma_va"]:
        m_mv = re.search(r"(?:Mã vụ án|Mã vụ|Số thụ lý|Mã án|Mã hồ sơ)[:\s]+([^\n,;]+)", str(info_va), re.IGNORECASE)
        if m_mv:
            result["ma_va"] = m_mv.group(1).strip()

    if not result["ma_va"] and lines:
        first_line = lines[0]
        if not re.search(r"^(?:Số QĐ|Số quyết định|Ngày|Thời hạn)", first_line, re.IGNORECASE):
            result["ma_va"] = first_line

    # 2. Ngày khởi tố vụ án (KTVA)
    m_kt = re.search(r"(?:Ngày ra QĐ khởi tố(?: vụ án)?|Ngày khởi tố(?: vụ án)?|Ngày KTVA|Ngày ra QĐ|Ngày QĐ|Ngày quyết định)[:\s]+(\d{1,2}[/-]\d{1,2}[/-]\d{4})", str(info_va), re.IGNORECASE)
    if not m_kt:
        m_kt = re.search(r"(?:Số QĐ|Số quyết định|khởi tố)[^\n\r]*?Ngày[:\s]+(\d{1,2}[/-]\d{1,2}[/-]\d{4})", str(info_va), re.IGNORECASE)
    if m_kt:
        result["ngay_kt"] = parse_date(m_kt.group(1))

    # 3. Số quyết định khởi tố vụ án (SỐ KTVA)
    m_so = re.search(r"(?:Mã số QĐ khởi tố vụ án|Số QĐ khởi tố vụ án|Số QĐ khởi tố|Số quyết định khởi tố|Số QĐ|Số quyết định)[:\s]+([^\n\r]+?)(?:\s+Ngày|$)", str(info_va), re.IGNORECASE)
    if m_so:
        result["so_kt"] = m_so.group(1).strip()

    # 4. Ngày tiếp nhận hồ sơ
    m_tn = re.search(r"(?:Ngày tiếp nhận hồ sơ|Ngày tiếp nhận|Ngày nhận hồ sơ|Ngày thụ lý)[:\s]+(\d{1,2}[/-]\d{1,2}[/-]\d{4})", str(info_va), re.IGNORECASE)
    if m_tn:
        result["ngay_tn"] = parse_date(m_tn.group(1))

    # 5. Thời hạn xử lý
    m_han = re.search(r"(?:Thời hạn xử lý|Thời hạn giải quyết|Thời hạn)[:\s]+(\d{1,2}[/-]\d{1,2}[/-]\d{4})", str(info_va), re.IGNORECASE)
    if m_han:
        result["han_xl"] = parse_date(m_han.group(1))

    # 6. Tên vụ án
    if len(lines) > 1:
        cand = lines[1]
        if not re.search(r"^(?:Số quyết định|Mã số QĐ|Số QĐ|Ngày ra QĐ|Ngày|Thời hạn)", cand, re.IGNORECASE):
            result["ten_va"] = cand

    if not result["ten_va"] and toi_val:
        t_clean = re.sub(r"^\d+\.\s*Điều\s+\d+[^:]*:\s*", "", str(toi_val)).strip()
        result["ten_va"] = f"Vụ án {t_clean}"

    return result

def is_case_error(val):
    """
    TIÊU CHÍ ĐÁNH GIÁ LỖI CỦA VKSND TP.HCM:
    - Nếu ô trống (chưa kiểm tra) -> False (không tính lỗi).
    - CHỈ CÓ điền "ĐÃ ĐÍNH KÈM", "TẠM ỔN" (hoặc Đạt, OK) là KHÔNG TÍNH LỖI.
    - CÒN CÓ ĐIỀN BẤT KỲ THÔNG TIN NÀO KHÁC (như LƯU TÊN FILE "DOCUMENT", "Chưa đính kèm tài liệu chứng cứ", "Chậm tiến độ", "Thiếu biên bản",...) -> TÍNH LÀ CÓ LỖI!
    """
    if val is None or pd.isna(val):
        return False

    val_norm = remove_accents(str(val)).strip()
    if not val_norm or val_norm in ("NAN", "NONE", ""):
        return False

    # 1. Các trường hợp chắc chắn là LỖI / VI PHẠM (chứa từ chỉ sự chưa hoàn thành hoặc thiếu sót)
    error_markers = [
        "CHUA", "KHONG", "K ", "THIEU", "CHAM", "SAI", "VI PHAM", "LOI",
        "QUA HAN", "HET HAN", "CAN KHAC PHUC", "YEU CAU"
    ]
    if any(em in val_norm for em in error_markers):
        return True

    # 2. Trường hợp ĐÃ ĐÍNH KÈM, TẠM ỔN -> KHÔNG TÍNH LỖI
    if any(k in val_norm for k in ["DA DINH KEM", "DINH KEM DU", "DINH KEM DAY DU", "TAM ON", "ON", "DAT", "TOT", "BINH THUONG", "OK", "HOP LE"]):
        return False

    # 3. QUY TẮC CỐT LÕI: Có điền bất kỳ thông tin nào khác -> ĐỀU TÍNH LÀ CÓ LỖI!
    return True

def make_compound_key(don_vi, ma_va):
    """Tạo khóa ghép ĐƠN VỊ + MÃ VỤ ÁN để loại trừ khả năng trùng mã án"""
    dv = remove_accents(str(don_vi or "")).strip()
    mv = remove_accents(str(ma_va or "")).strip()
    return f"{dv}___{mv}"

def normalize_column_names(df):
    """Chuẩn hóa đồng bộ tên cột từ các phiên bản cũ hoặc biến thể dấu tiếng Việt"""
    rename_map = {}
    for col in df.columns:
        c_norm = remove_accents(str(col)).replace(" ", "").upper()
        if "TIEPNHAN" in c_norm and "HOSO" in c_norm:
            rename_map[col] = "NGÀY TIẾP NHẬN HỐ SƠ"
        elif "GHICHUDEBAOTINHHINH" in c_norm or "GHICHUCUTHEHANGNGAY" in c_norm or col == "Column1":
            rename_map[col] = "GHI CHÚ ĐỂ BÁO TÌNH HÌNH"
        elif "KETLUANKIEMTRA" in c_norm:
            rename_map[col] = "NOTE KIỂM TRA"
    if rename_map:
        df = df.rename(columns=rename_map)
    return df

def add_bang_theo_doi_sheet(wb, df_tong=None, today_str=None, unit_summary_data=None):
    """
    Tự động tạo Sheet 'BẢNG THEO DÕI' theo đúng yêu cầu nghiệp vụ của cán bộ VKSND TP.HCM:
    - CHỈ LẤY các vụ án kiểm tra tự tay cán bộ note vào theo ngày mới nhất
    - Bảng tổng hiện ra để báo cáo lãnh đạo gồm 6 cột:
      STT | ĐƠN VỊ | TỔNG SỐ VỤ ĐÃ KIỂM TRA | SỐ VỤ ĐÃ ĐÍNH KÈM | SỐ VỤ LỖI | TỶ LỆ LỖI (%)
    - SẮP XẾP số lượng lỗi của các đơn vị TỪ NHIỀU TỚI ÍT
    - Dòng tổng kết TỔNG CỘNG (Tô màu nền xanh nhạt / xám nhạt chuẩn Pivot Table)
    """
    if "BẢNG THEO DÕI" in wb.sheetnames:
        del wb["BẢNG THEO DÕI"]

    ws = wb.create_sheet(title="BẢNG THEO DÕI")
    ws.views.sheetView[0].showGridLines = True

    font_filter_lbl = Font(name="Segoe UI", size=10, bold=True)
    font_filter_val = Font(name="Segoe UI", size=10)
    border_filter = Border(
        left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
    )
    border_pivot = Border(
        left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
    )

    t_str = today_str or datetime.now().strftime("%d/%m/%Y")
    ws["A1"] = "NOTE NGÀY KIỂM TRA"
    ws["B1"] = t_str
    ws["A2"] = "TIÊU CHÍ"
    ws["B2"] = "Vụ án kiểm tra ngày mới nhất"

    ws["A1"].font = font_filter_lbl
    ws["B1"].font = font_filter_val
    ws["B1"].alignment = Alignment(horizontal="center")
    ws["B1"].border = border_filter

    ws["A2"].font = font_filter_lbl
    ws["B2"].font = font_filter_val
    ws["B2"].alignment = Alignment(horizontal="center")
    ws["B2"].border = border_filter

    # Header Row 4: STT | ĐƠN VỊ | TỔNG SỐ VỤ ĐÃ KIỂM TRA | SỐ VỤ ĐÃ ĐÍNH KÈM | SỐ VỤ LỖI | TỶ LỆ LỖI (%)
    headers = ["STT", "ĐƠN VỊ", "TỔNG SỐ VỤ ĐÃ KIỂM TRA", "SỐ VỤ ĐÃ ĐÍNH KÈM", "SỐ VỤ LỖI", "TỶ LỆ LỖI (%)"]
    font_th = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    fill_th = PatternFill(start_color="0B2545", end_color="0B2545", fill_type="solid")

    for col_idx, h_text in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col_idx, value=h_text)
        cell.font = font_th
        cell.fill = fill_th
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border_pivot

    if unit_summary_data:
        sorted_units = unit_summary_data.get("sorted_units", [])
        total_all = unit_summary_data.get("total_kt", 0)
        total_dat = unit_summary_data.get("total_dat", 0)
        total_loi = unit_summary_data.get("total_loi", 0)
    else:
        unit_stats = {}
        if df_tong is not None:
            for _, r in df_tong.iterrows():
                rich = str(r.get("NOTE LÀM GIÀU DỮ LIỆU", "") if pd.notna(r.get("NOTE LÀM GIÀU DỮ LIỆU")) else "").strip()
                if not rich or rich.lower() in ("nan", "none", ""):
                    continue
                raw_u = r.get("ĐƠN VỊ", "")
                u = standardize_unit_name(raw_u) or "KHÔNG XÁC ĐỊNH"
                if u not in unit_stats:
                    unit_stats[u] = {"total": 0, "dat": 0, "loi": 0}
                if is_case_error(rich):
                    unit_stats[u]["loi"] += 1
                else:
                    unit_stats[u]["dat"] += 1
                unit_stats[u]["total"] += 1

        sorted_units = sorted(unit_stats.items(), key=lambda x: (x[1]["loi"], x[1]["total"]), reverse=True)
        total_all = sum(s["total"] for _, s in sorted_units)
        total_dat = sum(s["dat"] for _, s in sorted_units)
        total_loi = sum(s["loi"] for _, s in sorted_units)

    curr_row = 5
    for i, (u_name, s) in enumerate(sorted_units, 1):
        rate_val = (s["loi"] / s["total"] * 100) if s["total"] > 0 else 0.0
        ws.cell(row=curr_row, column=1, value=i).alignment = Alignment(horizontal="center")
        ws.cell(row=curr_row, column=2, value=f"  {u_name}").alignment = Alignment(horizontal="left")
        ws.cell(row=curr_row, column=3, value=s["total"]).alignment = Alignment(horizontal="center")
        ws.cell(row=curr_row, column=4, value=s["dat"]).alignment = Alignment(horizontal="center")
        c_loi = ws.cell(row=curr_row, column=5, value=s["loi"])
        c_loi.alignment = Alignment(horizontal="center")
        if s["loi"] > 0:
            c_loi.font = Font(name="Segoe UI", size=10, bold=True, color="C00000")
        c_rate = ws.cell(row=curr_row, column=6, value=f"{rate_val:.1f}%")
        c_rate.alignment = Alignment(horizontal="center")
        if s["loi"] > 0:
            c_rate.font = Font(name="Segoe UI", size=10, bold=True, color="C00000")

        for col_idx in range(1, 7):
            ws.cell(row=curr_row, column=col_idx).border = border_pivot
        curr_row += 1

    # Dòng TỔNG CỘNG
    fill_grand = PatternFill(start_color="E9EEF4", end_color="E9EEF4", fill_type="solid")
    font_grand = Font(name="Segoe UI", size=10, bold=True, color="000000")

    ws.cell(row=curr_row, column=1, value="")
    ws.cell(row=curr_row, column=2, value="TỔNG CỘNG").alignment = Alignment(horizontal="left")
    ws.cell(row=curr_row, column=3, value=total_all).alignment = Alignment(horizontal="center")
    ws.cell(row=curr_row, column=4, value=total_dat).alignment = Alignment(horizontal="center")
    c_tot_loi = ws.cell(row=curr_row, column=5, value=total_loi)
    c_tot_loi.alignment = Alignment(horizontal="center")
    rate_total = f"{(total_loi / max(total_all, 1) * 100):.1f}%"
    c_tot_rate = ws.cell(row=curr_row, column=6, value=rate_total)
    c_tot_rate.alignment = Alignment(horizontal="center")

    for col_idx in range(1, 7):
        c = ws.cell(row=curr_row, column=col_idx)
        c.font = font_grand
        c.fill = fill_grand
        c.border = border_pivot
    if total_loi > 0:
        c_tot_loi.font = Font(name="Segoe UI", size=10, bold=True, color="C00000")
        c_tot_rate.font = Font(name="Segoe UI", size=10, bold=True, color="C00000")

    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 26
    ws.column_dimensions["D"].width = 22
    ws.column_dimensions["E"].width = 18
    ws.column_dimensions["F"].width = 16

    ws.freeze_panes = "C5"


def save_master_excel(df, file_path, is_inspection_file=False, include_pivot_sheet=False, today_str=None, unit_summary_data=None):
    """
    Lưu File Excel theo đúng chuẩn 100% hình ảnh thực tế của VKSND TP.HCM:
    - Sheet chính: 'DATA' (chuẩn mực đồng bộ với file mẫu của VKS)
    - 22 cột chính xác theo đúng thứ tự quy chuẩn (ĐÃ BỎ CỘT NOTE LỖI)
    - Tiêu đề nền VÀNG (#FFFF00), chữ đen in đậm (Segoe UI 10pt)
    - Tự động bật mũi tên lọc AutoFilter ở dòng tiêu đề
    - Bật hiển thị lưới ô (gridlines)
    - Viền mỏng toàn bộ các ô dữ liệu
    - Đóng băng ngang dòng tiêu đề 1 và đóng băng dọc đến hết cột MÃ VỤ ÁN (Freeze Panes tại ô J2)
    - Nếu là file kiểm tra (is_inspection_file=True): Gắn kèm DataValidation danh sách chọn cho cột NOTE LÀM GIÀU DỮ LIỆU
    - Nếu include_pivot_sheet=True: Tự động kèm thêm Sheet 'BẢNG THEO DÕI' Pivot Table
    """
    df_out = df[[c for c in MASTER_COLUMNS if c in df.columns]].copy()
    for c in MASTER_COLUMNS:
        if c not in df_out.columns:
            df_out[c] = ""
    df_out = df_out[MASTER_COLUMNS]

    if "ĐƠN VỊ" in df_out.columns:
        df_out["ĐƠN VỊ"] = df_out["ĐƠN VỊ"].apply(standardize_unit_name)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "DATA"
    ws.views.sheetView[0].showGridLines = True

    font_hdr = Font(name="Segoe UI", size=10, bold=True, color="000000")
    fill_hdr = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    border_thin = Border(
        left=Side(style='thin', color='DDDDDD'), right=Side(style='thin', color='DDDDDD'),
        top=Side(style='thin', color='DDDDDD'), bottom=Side(style='thin', color='DDDDDD')
    )

    ws.append(list(df_out.columns))
    for c_i in range(1, len(df_out.columns) + 1):
        cell = ws.cell(row=1, column=c_i)
        cell.font = font_hdr
        cell.fill = fill_hdr
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for r in df_out.itertuples(index=False):
        row_vals = ["" if (v is None or pd.isna(v)) else v for v in r]
        ws.append(row_vals)

    ws.row_dimensions[1].height = 28

    center_cols = {1, 9, 11, 12, 13, 14, 16, 17, 18, 21, 22}
    date_cols = {12, 13, 14, 18, 22}  # NGÀY KTVA, NGÀY TIẾP NHẬN HỐ SƠ, THỜI HẠN XỬ LÝ, NOTE NGÀY KIỂM TRA, NGÀY BÁO TÌNH HÌNH
    int_cols = {1, 16}  # STT, NGÀY HIỆN TẠI SO VỚI NGÀY KTVA

    for r_idx in range(2, len(df_out) + 2):
        for c_idx in range(1, len(df_out.columns) + 1):
            cell = ws.cell(row=r_idx, column=c_idx)
            cell.border = border_thin
            v = cell.value

            # Định dạng cột NGÀY chuẩn Excel: Chuyển thành đối tượng Datetime thật kèm number_format='DD/MM/YYYY'
            # Nhờ vậy AutoFilter của Excel sẽ tự động gom nhóm Năm -> Tháng -> Ngày chính xác
            if c_idx in date_cols:
                parsed_d = parse_date(v)
                if parsed_d:
                    cell.value = datetime.combine(parsed_d, datetime.min.time())
                    cell.number_format = "DD/MM/YYYY"
                else:
                    cell.value = ""
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c_idx in int_cols:
                try:
                    if v is not None and str(v).strip() != "" and str(v).lower() not in ("nan", "none", ""):
                        cell.value = int(float(str(v).strip()))
                        cell.number_format = "#,##0"
                    else:
                        cell.value = ""
                except Exception:
                    pass
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c_idx in center_cols:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c_idx == 2:
                cell.alignment = Alignment(vertical="center", wrap_text=True)
            else:
                cell.alignment = Alignment(vertical="center")

    # Bật tính năng AutoFilter (mũi tên lọc)
    ws.auto_filter.ref = ws.dimensions

    # ĐÓNG BĂNG NGANG DÒNG TIÊU ĐỀ (DÒNG 1) VÀ ĐÓNG BĂNG DỌC ĐẾN HẾT CỘT MÃ VỤ ÁN (CỘT I / CỘT THỨ 9)
    # Ô bắt đầu cuộn tự do là J2 (từ cột TÊN VỤ ÁN và dòng dữ liệu thứ 2)
    ws.freeze_panes = "J2"

    # Căn chỉnh độ rộng từng cột trực quan, gọn gàng chuẩn mẫu VKSND TP.HCM
    col_widths = {
        "STT": 6,
        "Thông tin vụ án": 22,
        "Loại tội": 13,
        "Tội": 18,
        "Thông tin bị can": 18,
        "Loại thụ lý": 14,
        "Trạng thái vụ án": 16,
        "Trạng thái biểu mẫu": 18,
        "MÃ VỤ ÁN": 24,
        "TÊN VỤ ÁN": 24,
        "SỐ KTVA": 16,
        "NGÀY KTVA": 13,
        "NGÀY TIẾP NHẬN HỐ SƠ": 15,
        "THỜI HẠN XỬ LÝ": 15,
        "NOTE LÀM GIÀU DỮ LIỆU": 22,
        "NGÀY HIỆN TẠI SO VỚI NGÀY KTVA": 16,
        "ĐƠN VỊ": 16,
        "NOTE NGÀY KIỂM TRA": 15,
        "NOTE": 16,
        "GHI CHÚ ĐỂ BÁO TÌNH HÌNH": 30,
        "NOTE KIỂM TRA": 16,
        "NGÀY BÁO TÌNH HÌNH": 15,
    }
    for col_name, width in col_widths.items():
        if col_name in MASTER_COLUMNS:
            col_letter = get_column_letter(MASTER_COLUMNS.index(col_name) + 1)
            ws.column_dimensions[col_letter].width = width


    # Nếu yêu cầu, tạo kèm Sheet BẢNG THEO DÕI
    if include_pivot_sheet:
        t_str = today_str or datetime.now().strftime("%d/%m/%Y")
        add_bang_theo_doi_sheet(wb, df_out, t_str, unit_summary_data=unit_summary_data)

    wb.save(file_path)

def create_remedy_excel(error_cases, file_path, sheet_title="DANH_SACH_LOI_KHAC_PHUC", default_check_date=None):
    """
    Tạo File Excel danh sách các vụ án lỗi gửi các đơn vị khắc phục:
    - 7 cột theo đúng chuẩn yêu cầu nghiệp vụ:
      1. STT
      2. MÃ VỤ ÁN
      3. SỐ KTVA
      4. NGÀY KTVA (định dạng DD/MM/YYYY)
      5. NOTE LÀM GIÀU DỮ LIỆU (nội dung lỗi cụ thể, in đậm)
      6. ĐƠN VỊ (chuẩn hóa KHU VỰC..., PHÒNG...)
      7. NOTE NGÀY KIỂM TRA (định dạng DD/MM/YYYY)
    - Tiêu đề màu Xanh Teal sẫm (#156082), chữ Trắng in đậm
    - Bật AutoFilter cho toàn bộ 7 cột
    - Cố định dòng 1 và 2 cột đầu (Freeze Panes tại C2)
    - Tự động căn chỉnh độ rộng cột và viền ô chuẩn mực
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_title[:31]
    ws.views.sheetView[0].showGridLines = True

    headers = [
        "STT",
        "MÃ VỤ ÁN",
        "SỐ KTVA",
        "NGÀY KTVA",
        "NOTE LÀM GIÀU DỮ LIỆU",
        "ĐƠN VỊ",
        "NOTE NGÀY KIỂM TRA"
    ]
    ws.append(headers)
    ws.row_dimensions[1].height = 28

    font_hdr = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    fill_hdr = PatternFill(start_color="156082", end_color="156082", fill_type="solid")

    border_thin = Border(
        left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
    )

    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = font_hdr
        cell.fill = fill_hdr
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border_thin

    # Điền dữ liệu các vụ án lỗi
    for idx, r in enumerate(error_cases, 1):
        r_row = idx + 1
        if isinstance(r, dict):
            get_c = lambda k: r.get(k, "")
        elif hasattr(r, "_asdict"):
            d = r._asdict()
            get_c = lambda k: d.get(k, "")
        elif hasattr(r, "get"):
            get_c = lambda k: r.get(k, "")
        else:
            get_c = lambda k: getattr(r, k, "")

        ma_va = str(get_c("MÃ VỤ ÁN") or "").strip()
        so_kt = str(get_c("SỐ KTVA") or "").strip()
        ngay_kt_val = get_c("NGÀY KTVA")
        rich = str(get_c("NOTE LÀM GIÀU DỮ LIỆU") or "").strip()
        don_vi = standardize_unit_name(get_c("ĐƠN VỊ"))
        note_ngay_kt_val = get_c("NOTE NGÀY KIỂM TRA")
        if not note_ngay_kt_val and default_check_date:
            note_ngay_kt_val = default_check_date

        ws.cell(row=r_row, column=1, value=idx).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=r_row, column=2, value=ma_va).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=r_row, column=3, value=so_kt).alignment = Alignment(horizontal="center", vertical="center")

        # NGÀY KTVA
        c_d1 = ws.cell(row=r_row, column=4)
        p_d1 = parse_date(ngay_kt_val)
        if p_d1:
            c_d1.value = datetime.combine(p_d1, datetime.min.time())
            c_d1.number_format = "DD/MM/YYYY"
        else:
            c_d1.value = str(ngay_kt_val or "")
        c_d1.alignment = Alignment(horizontal="center", vertical="center")

        # NOTE LÀM GIÀU DỮ LIỆU (in đậm đỏ)
        c_rich = ws.cell(row=r_row, column=5, value=rich)
        c_rich.font = Font(name="Segoe UI", size=10, bold=True, color="C00000")
        c_rich.alignment = Alignment(vertical="center")

        # ĐƠN VỊ
        ws.cell(row=r_row, column=6, value=don_vi).alignment = Alignment(horizontal="center", vertical="center")

        # NOTE NGÀY KIỂM TRA
        c_d2 = ws.cell(row=r_row, column=7)
        p_d2 = parse_date(note_ngay_kt_val)
        if p_d2:
            c_d2.value = datetime.combine(p_d2, datetime.min.time())
            c_d2.number_format = "DD/MM/YYYY"
        else:
            c_d2.value = str(note_ngay_kt_val or "")
        c_d2.alignment = Alignment(horizontal="center", vertical="center")

        for c_i in range(1, 8):
            ws.cell(row=r_row, column=c_i).border = border_thin

    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "C2"

    col_widths = {
        "A": 8,
        "B": 28,
        "C": 18,
        "D": 16,
        "E": 45,
        "F": 18,
        "G": 18
    }
    for col_l, w in col_widths.items():
        ws.column_dimensions[col_l].width = w

    os.makedirs(os.path.dirname(os.path.abspath(file_path)), exist_ok=True)
    wb.save(file_path)
    return file_path

def backup_file_if_exists(file_path):
    """
    Tự động sao lưu file dự phòng trước khi ghi đè nếu file đã tồn tại:
    - Lưu vào thư mục 'Backup' cùng cấp với file
    - Đặt tên kèm timestamp: {tên_file}_backup_YYYYMMDD_HHMMSS.xlsx
    - Bảo vệ 100% dữ liệu ghi chú của cán bộ kiểm tra không bao giờ bị mất do bấm nhầm
    - Trả về đường dẫn file backup nếu đã sao lưu, ngược lại trả về None
    """
    if not file_path or not os.path.exists(file_path):
        return None
    try:
        if os.path.getsize(file_path) == 0:
            return None

        dirname = os.path.dirname(os.path.abspath(file_path))
        backup_dir = os.path.join(dirname, "Backup")
        os.makedirs(backup_dir, exist_ok=True)

        base_name = os.path.basename(file_path)
        name, ext = os.path.splitext(base_name)
        now_str = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_name = f"{name}_backup_{now_str}{ext}"
        backup_path = os.path.join(backup_dir, backup_name)

        shutil.copy2(file_path, backup_path)
        return backup_path
    except Exception:
        return None

def check_file_has_user_notes(file_path):
    """
    Kiểm tra nhanh xem file danh sách kiểm tra đã có dữ liệu ghi chú của cán bộ hay chưa:
    - Kiểm tra các cột 'NOTE LÀM GIÀU DỮ LIỆU', 'NOTE NGÀY KIỂM TRA', 'GHI CHÚ ĐỂ BÁO TÌNH HÌNH'
    - Trả về (has_notes: bool, count_noted: int)
    """
    if not file_path or not os.path.exists(file_path):
        return False, 0
    try:
        df = pd.read_excel(file_path)
        df = normalize_column_names(df)
        note_cols = ["NOTE LÀM GIÀU DỮ LIỆU", "NOTE NGÀY KIỂM TRA", "GHI CHÚ ĐỂ BÁO TÌNH HÌNH"]
        existing_cols = [c for c in note_cols if c in df.columns]
        if not existing_cols:
            return False, 0

        count = 0
        for _, r in df.iterrows():
            for c in existing_cols:
                v = str(r.get(c, "") if pd.notna(r.get(c)) else "").strip()
                if v and v.lower() not in ("nan", "none", "nat", "null", ""):
                    count += 1
                    break
        return count > 0, count
    except Exception:
        return False, 0

def resolve_input_files(folder_or_files, existing_master_path=None):
    """
    Chuẩn hóa và bóc tách danh sách file Excel của các đơn vị:
    - Nhận đường dẫn thư mục: Quét toàn bộ file .xlsx trong thư mục.
    - Nhận 1 file đơn lẻ: Chỉ xử lý file đó.
    - Nhận danh sách nhiều file (dạng list hoặc chuỗi nối nhau bởi ';'): Quét toàn bộ các file được chọn.
    - Tự động bỏ qua file tạm (~$), file hệ thống tạo ra (1_FILE_TONG, 2_DANH_SACH, ...), và file tổng cũ.
    - Tự động bỏ qua thư mục 'Backup' và 'Cac_Don_Vi_Khac_Phuc'.
    """
    raw_list = []
    if isinstance(folder_or_files, (list, tuple)):
        raw_list = list(folder_or_files)
    elif isinstance(folder_or_files, str):
        for p in folder_or_files.replace("\n", ";").split(";"):
            p_clean = p.strip().strip('"').strip("'")
            if p_clean:
                raw_list.append(p_clean)

    collected_files = []
    explicit_files = set()
    for item in raw_list:
        if os.path.isfile(item):
            p = os.path.abspath(item)
            collected_files.append(p)
            explicit_files.add(p.lower())
        elif os.path.isdir(item):
            for root_dir, dirs, fnames in os.walk(item):
                norm_root = os.path.normpath(root_dir).lower()
                if "backup" in norm_root or "cac_don_vi_khac_phuc" in norm_root:
                    continue
                for fname in fnames:
                    fp = os.path.join(root_dir, fname)
                    if os.path.isfile(fp):
                        collected_files.append(os.path.abspath(fp))

    master_abs = os.path.abspath(existing_master_path).lower() if existing_master_path else None
    clean_files = []
    seen = set()

    system_prefixes = [
        "1_FILE_TONG", "2_DANH_SACH", "3_FILE_TONG",
        "4_BAO_CAO", "5_DANH_SACH", "DANH_SACH_VU_AN_MOI_PHAT_SINH",
        "_backup_", "_khac_phuc_loi"
    ]

    for fp in collected_files:
        norm_fp = os.path.normpath(fp).lower()
        if "backup" in norm_fp or "cac_don_vi_khac_phuc" in norm_fp:
            continue

        fname = os.path.basename(fp)
        if fname.startswith("~$"):
            continue
        if not fname.lower().endswith((".xlsx", ".xlsm")):
            continue
        # Bỏ qua file hệ thống tạo ra NẾU quét từ thư mục, nhưng NẾU người dùng chỉ định đích danh file đó thì vẫn cho phép xử lý
        if norm_fp not in explicit_files and any(prefix.lower() in fname.lower() for prefix in system_prefixes):
            continue
        if master_abs and os.path.abspath(fp).lower() == master_abs:
            continue
        if os.path.abspath(fp).lower() not in seen:
            seen.add(os.path.abspath(fp).lower())
            clean_files.append(fp)

    return sorted(clean_files)

def is_already_master_file(file_path):
    """
    Kiểm tra xem một file Excel có phải là File Tổng/Bảng theo dõi đã có sẵn dữ liệu và ghi chú hay không:
    - Có sheet tên 'DATA' (chuẩn mực của hệ thống)
    - Hoặc có cột 'NOTE LÀM GIÀU DỮ LIỆU' / 'NOTE LAM GIAU'
    - Hoặc có trên 15 cột chuẩn của Master file
    """
    if not file_path or not os.path.isfile(file_path):
        return False
    try:
        wb = openpyxl.load_workbook(file_path, read_only=True)
        snames = [s.strip().upper() for s in wb.sheetnames]
        if "DATA" in snames:
            wb.close()
            return True
        for s in wb.sheetnames[:3]:
            ws = wb[s]
            for row in ws.iter_rows(max_row=5, values_only=True):
                r_str = " ".join([remove_accents(str(c)).upper() for c in row if c is not None])
                if "NOTE LAM GIAU" in r_str or "LAM GIAU DU LIEU" in r_str:
                    wb.close()
                    return True
        wb.close()
    except Exception:
        pass
    return False

# ==================== BƯỚC 1: TỔNG HỢP & LỌC ÁN ====================

def buoc_1_tong_hop_va_loc(folder_or_files, output_dir=None, check_days=16, existing_master_path=None, use_timestamp=False):
    """
    TỔNG HỢP & LỌC ÁN HÌNH SỰ:
    - Linh hoạt đầu vào: Cho phép chọn cả thư mục HOẶC chọn 1, 2... nhiều file Excel đơn lẻ.
    - Cột 1 đến 8: Giữ nguyên hiện trạng từ file xuất xuống của đơn vị.
    - Cột 9 đến 14: Bóc tách chính xác từ ô Thông tin vụ án.
    - Cột 15 đến 22: Trực quan hóa tiến độ, số ngày, đơn vị và các cột ghi chú kiểm tra (mặc định để trống).
    - ĐÃ LOẠI BỎ HOÀN TOÀN CỘT NOTE LỖI (File chuẩn 22 cột).
    - TÊN ĐƠN VỊ: CHỈ GIỮ LẠI KHU VỰC VÀ SỐ (KHU VỰC 1, KHU VỰC 2...), PHÒNG VÀ SỐ (PHÒNG 1, PHÒNG 2...). Cắt bỏ toàn bộ .DT, .ST, .TT.
    - ĐÓNG BĂNG NGANG VÀ DỌC: Cố định dòng 1 và cột MÃ VỤ ÁN (Freeze Panes tại ô J2).
    - CHẾ ĐỘ NẠP FILE TỔNG SẴN CÓ:
      + Tự động đọc đúng Sheet 'DATA' (bảo toàn toàn bộ 3200-3300 vụ án sẵn có).
      + Đối chiếu theo MÃ VỤ ÁN: Nếu đơn vị có cập nhật mới (trạng thái, biểu mẫu, ngày...) -> Cập nhật vào dòng đó nhưng BẢO TOÀN 100% các ghi chú kiểm tra cũ.
      + Nếu có vụ án mới chưa từng có -> Nạp thêm vào cuối danh sách.
      + Không bao giờ làm mất các vụ án ban đầu của File Tổng.
    """
    today = datetime.now().date()
    all_new_rows = []

    # 1. Đọc và thu thập danh sách mã vụ án từ File Tổng đã có sẵn (nếu có)
    df_existing = None
    case_index_map = {}
    total_existing = 0

    if existing_master_path and os.path.isfile(existing_master_path):
        try:
            df_existing = load_existing_master_df(existing_master_path)
            total_existing = len(df_existing)

            # Lập bản đồ chỉ mục theo MÃ VỤ ÁN
            for idx, r in df_existing.iterrows():
                ma_val = r.get("MÃ VỤ ÁN")
                if not ma_val or pd.isna(ma_val) or str(ma_val).strip().lower() in ("none", "nan", ""):
                    ext = extract_case_details_from_cell(r.get("Thông tin vụ án", ""))
                    ma_val = ext.get("ma_va", "")
                    if ma_val:
                        df_existing.at[idx, "MÃ VỤ ÁN"] = ma_val

                ma_clean = remove_accents(str(ma_val)).replace(" ", "").strip()
                if ma_clean:
                    case_index_map[ma_clean] = idx

                # Cập nhật số ngày so với hiện tại
                ngay_kt_val = r.get("NGÀY KTVA")
                ngay_kt_dt = parse_date(ngay_kt_val)
                if ngay_kt_dt:
                    df_existing.at[idx, "NGÀY HIỆN TẠI SO VỚI NGÀY KTVA"] = (today - ngay_kt_dt).days

        except Exception as e:
            return False, {
                "error": f"Lỗi đọc File Tổng cũ: {str(e)}",
                "message": f"Không thể đọc dữ liệu từ File Tổng cũ sẵn có: {str(e)}. Vui lòng kiểm tra lại file!"
            }

    # 2. Tìm danh sách file Excel của các đơn vị gửi về
    files = resolve_input_files(folder_or_files, existing_master_path=existing_master_path)

    if not output_dir:
        if files:
            output_dir = os.path.dirname(files[0])
        elif existing_master_path and os.path.isfile(existing_master_path):
            output_dir = os.path.dirname(existing_master_path)
        elif isinstance(folder_or_files, str) and os.path.isdir(folder_or_files):
            output_dir = folder_or_files
        else:
            output_dir = os.getcwd()

    if not files and df_existing is None:
        return False, {
            "error": "Không tìm thấy file Excel nào của các đơn vị!",
            "message": "Vui lòng bấm 'Chọn thư mục' hoặc 'Chọn file' để chọn các file .xlsx của các đơn vị gửi về."
        }

    # TRƯỜNG HỢP ĐẶC BIỆT: NGƯỜI DÙNG CHỈ TRUYỀN 1 FILE DUY NHẤT ĐÃ CÓ SẴN THÔNG TIN (FILE TỔNG / BẢNG THEO DÕI)
    # Không cần tổng hợp nhiều file: Xử lý trực tiếp file này dựa trên đúng 2 điều kiện nghiệp vụ:
    # 1. Cột 'NOTE LÀM GIÀU DỮ LIỆU' chưa có thông tin (ô trống)
    # 2. Cột 'NGÀY HIỆN TẠI SO VỚI NGÀY KTVA' từ check_days ngày trở đi (mặc định >= 16 ngày)
    if len(files) == 1 and df_existing is None and is_already_master_file(files[0]):
        master_file = files[0]
        try:
            df_tong = load_existing_master_df(master_file)
        except Exception as e:
            return False, {
                "error": f"Lỗi đọc file: {str(e)}",
                "message": f"Không thể đọc dữ liệu từ file '{os.path.basename(master_file)}': {str(e)}"
            }

        # Bóc tách MÃ VỤ ÁN nếu thiếu
        for idx, r in df_tong.iterrows():
            ma_val = r.get("MÃ VỤ ÁN")
            if not ma_val or pd.isna(ma_val) or str(ma_val).strip().lower() in ("none", "nan", ""):
                ext = extract_case_details_from_cell(r.get("Thông tin vụ án", ""))
                ma_val = ext.get("ma_va", "")
                if ma_val:
                    df_tong.at[idx, "MÃ VỤ ÁN"] = ma_val

        # Cập nhật số ngày so với hiện tại nếu chưa có hoặc không hợp lệ
        for idx, r in df_tong.iterrows():
            val_days = r.get("NGÀY HIỆN TẠI SO VỚI NGÀY KTVA")
            try:
                _ = int(val_days)
            except (ValueError, TypeError):
                d_kt = parse_date(r.get("NGÀY KTVA"))
                if d_kt:
                    df_tong.at[idx, "NGÀY HIỆN TẠI SO VỚI NGÀY KTVA"] = (today - d_kt).days

        df_tong["STT"] = range(1, len(df_tong) + 1)
        if "ĐƠN VỊ" in df_tong.columns:
            df_tong["ĐƠN VỊ"] = df_tong["ĐƠN VỊ"].apply(standardize_unit_name)

        for c in MASTER_COLUMNS:
            if c not in df_tong.columns:
                df_tong[c] = ""
        df_tong = df_tong[MASTER_COLUMNS]

        if use_timestamp:
            now_ts = datetime.now().strftime("%Y%m%d_%H%M")
            file_tong_path = os.path.join(output_dir, f"1_FILE_TONG_{now_ts}.xlsx")
            file_kt_path = os.path.join(output_dir, f"2_DANH_SACH_CAN_KIEM_TRA_{now_ts}.xlsx")
        else:
            file_tong_path = os.path.join(output_dir, "1_FILE_TONG.xlsx")
            file_kt_path = os.path.join(output_dir, "2_DANH_SACH_CAN_KIEM_TRA.xlsx")

        backup_tong = backup_file_if_exists(file_tong_path)
        try:
            save_master_excel(df_tong, file_tong_path, is_inspection_file=False)
        except PermissionError:
            return False, {
                "error": "File đang mở trong Excel!",
                "message": f"File '{os.path.basename(file_tong_path)}' đang được mở trong Excel. Vui lòng đóng file lại rồi thử lại!"
            }
        except Exception as e:
            return False, {
                "error": f"Lỗi lưu file: {str(e)}",
                "message": f"Không thể lưu file '{os.path.basename(file_tong_path)}'."
            }

        # LỌC DANH SÁCH CẦN KIỂM TRA THEO ĐÚNG 2 ĐIỀU KIỆN NGHIỆP VỤ:
        # 1. Cột 'NOTE LÀM GIÀU DỮ LIỆU' chưa có thông tin (ô trống)
        # 2. Cột 'NGÀY HIỆN TẠI SO VỚI NGÀY KTVA' từ check_days ngày trở đi (mặc định >= 16)
        def is_empty_rich_note(v):
            if v is None or pd.isna(v):
                return True
            s = str(v).strip()
            return s == "" or s.lower() in ("nan", "none", "nat", "null")

        mask_days = pd.to_numeric(df_tong["NGÀY HIỆN TẠI SO VỚI NGÀY KTVA"], errors="coerce") >= check_days
        mask_empty_rich = df_tong["NOTE LÀM GIÀU DỮ LIỆU"].apply(is_empty_rich_note)

        df_kt = df_tong[mask_days & mask_empty_rich].copy()
        df_kt["STT"] = range(1, len(df_kt) + 1)
        df_kt = df_kt[MASTER_COLUMNS]

        backup_kt = backup_file_if_exists(file_kt_path)
        try:
            save_master_excel(df_kt, file_kt_path, is_inspection_file=True)
        except PermissionError:
            return False, {
                "error": "File đang mở trong Excel!",
                "message": f"File '{os.path.basename(file_kt_path)}' đang được mở trong Excel. Vui lòng đóng file lại rồi thử lại!"
            }
        except Exception as e:
            return False, {
                "error": f"Lỗi lưu file: {str(e)}",
                "message": f"Không thể lưu file '{os.path.basename(file_kt_path)}'."
            }

        backups = [b for b in [backup_tong, backup_kt] if b]
        noted_cnt = int((~mask_empty_rich).sum())
        unnoted_cnt = int(mask_empty_rich.sum())
        filtered_cnt = len(df_kt)

        status_msg = (
            f"Đã xử lý trực tiếp file dữ liệu có sẵn '{os.path.basename(master_file)}' (Tổng {len(df_tong)} vụ án):\n"
            f"- Đã bảo lưu: {noted_cnt} vụ án đã note làm giàu dữ liệu (bỏ qua không kiểm tra lại).\n"
            f"- Số vụ chưa có thông tin tại cột Làm giàu dữ liệu: {unnoted_cnt} vụ.\n"
            f"- ĐÃ LỌC ĐÚNG {filtered_cnt} VỤ ÁN CHƯA NOTE VÀ TỪ {check_days} NGÀY TRỞ ĐI vào file '2_DANH_SACH_CAN_KIEM_TRA.xlsx'."
        )

        result_data = {
            "success": True,
            "mode": "single_master",
            "message": status_msg,
            "total_files": 1,
            "total_existing": len(df_tong),
            "new_added": 0,
            "updated_cases": 0,
            "skipped_duplicates": 0,
            "total_cases": len(df_tong),
            "noted_cases": noted_cnt,
            "unnoted_cases": unnoted_cnt,
            "cases_filtered": filtered_cnt,
            "under_days_cases": unnoted_cnt - filtered_cnt,
            "check_days": check_days,
            "file_tong": os.path.abspath(file_tong_path),
            "file_kt": os.path.abspath(file_kt_path),
            "file_moi_phat_sinh": None,
            "backup_files": backups,
            "output_dir": os.path.abspath(output_dir),
            "files_processed": [os.path.basename(master_file)]
        }
        return True, result_data

    updated_cases_count = 0
    new_added_count = 0

    for fpath in files:
        fname = os.path.basename(fpath)
        unit_name_raw = os.path.splitext(fname)[0].strip()
        default_unit_name = standardize_unit_name(unit_name_raw)

        try:
            wb = openpyxl.load_workbook(fpath, data_only=True)
            ws = get_best_case_sheet(wb)
        except Exception:
            continue

        h_row = find_header_row(ws)
        raw_headers = [str(c).strip() if c is not None else "" for c in next(ws.iter_rows(min_row=h_row, max_row=h_row, values_only=True))]
        col_map = detect_column_indices(raw_headers)

        for r in ws.iter_rows(min_row=h_row+1, values_only=True):
            if not any(r):
                continue

            def get_val(key, default_col_idx=None):
                if key in col_map and col_map[key] <= len(r):
                    v = r[col_map[key] - 1]
                    if v is not None and not pd.isna(v):
                        return str(v).strip()
                if default_col_idx and default_col_idx <= len(r):
                    v = r[default_col_idx - 1]
                    if v is not None and not pd.isna(v):
                        return str(v).strip()
                return ""

            # TỪ CỘT SỐ 1 ĐẾN CỘT SỐ 8: GIỮ NGUYÊN HIỆN TRẠNG FILE XUẤT XUỐNG
            info_va = get_val("INFO_VA", default_col_idx=2)
            loai_toi = get_val("LOAI_TOI", default_col_idx=3)
            toi_val = get_val("TOI", default_col_idx=4)
            bi_can = get_val("BI_CAN", default_col_idx=5)
            thu_ly = get_val("THU_LY", default_col_idx=6)
            trang_thai_va = get_val("TRANG_THAI_VA", default_col_idx=7)
            trang_thai_bm = get_val("TRANG_THAI_BM", default_col_idx=8)

            # CÁC CỘT CÒN LẠI: TÁCH THÔNG TIN TỪ Ô CỘT THÔNG TIN VỤ ÁN VÀO
            ext = extract_case_details_from_cell(info_va, toi_val)

            ma_va = get_val("MA_VA") or ext.get("ma_va", "")
            if not ma_va or ma_va.lower() in ("none", ""):
                ma_va = ext.get("ma_va", "")

            if not ma_va or ma_va.lower() in ("none", ""):
                continue

            # Đơn vị: Chỉ giữ lại KHU VỰC và số, PHÒNG và số
            don_vi_cell = get_val("DON_VI")
            final_unit = standardize_unit_name(don_vi_cell) or default_unit_name

            ten_va = get_val("TEN_VA") or ext.get("ten_va", "")
            so_ktva = get_val("SO_KTVA") or ext.get("so_kt", "")

            ngay_kt_val = r[col_map["NGAY_KTVA"] - 1] if ("NGAY_KTVA" in col_map and col_map["NGAY_KTVA"] <= len(r)) else None
            ngay_kt_dt = parse_date(ngay_kt_val) or ext.get("ngay_kt")

            ngay_tn_val = r[col_map["NGAY_TN"] - 1] if ("NGAY_TN" in col_map and col_map["NGAY_TN"] <= len(r)) else None
            ngay_tn_dt = parse_date(ngay_tn_val) or ext.get("ngay_tn")

            han_xl_val = r[col_map["HAN_XL"] - 1] if ("HAN_XL" in col_map and col_map["HAN_XL"] <= len(r)) else None
            han_xl_dt = parse_date(han_xl_val) or ext.get("han_xl")

            so_ngay = (today - ngay_kt_dt).days if ngay_kt_dt else ""

            # LẤY CÁC CỘT GHI CHÚ NẾU FILE ĐÃ CÓ SẴN (BẢO TOÀN DỮ LIỆU)
            note_lam_giau = get_val("NOTE_LAM_GIAU")
            note_ngay_kt = get_val("NOTE_NGAY_KT")
            note_val = get_val("NOTE")
            ghi_chu_th = get_val("GHI_CHU_BAO_CAO")
            note_kt_val = get_val("NOTE_KT")
            ngay_bao_th = get_val("NGAY_BAO_TH")
            file_so_ngay = get_val("NGAY_SO_VOI_KTVA")

            ma_va_clean = remove_accents(str(ma_va)).replace(" ", "").strip()

            # QUY TẮC ĐỐI CHIẾU:
            # 1. Nếu vụ án ĐÃ CÓ trong File Tổng cũ: Cập nhật thông tin mới nhất từ đơn vị, BẢO TOÀN GHI CHÚ KIỂM TRA
            if df_existing is not None and ma_va_clean in case_index_map:
                idx = case_index_map[ma_va_clean]
                if idx < len(df_existing):
                    if info_va: df_existing.at[idx, "Thông tin vụ án"] = info_va
                    if loai_toi: df_existing.at[idx, "Loại tội"] = loai_toi
                    if toi_val: df_existing.at[idx, "Tội"] = toi_val
                    if bi_can: df_existing.at[idx, "Thông tin bị can"] = bi_can
                    if thu_ly: df_existing.at[idx, "Loại thụ lý"] = thu_ly
                    if trang_thai_va: df_existing.at[idx, "Trạng thái vụ án"] = trang_thai_va
                    if trang_thai_bm: df_existing.at[idx, "Trạng thái biểu mẫu"] = trang_thai_bm
                    if ten_va: df_existing.at[idx, "TÊN VỤ ÁN"] = ten_va
                    if so_ktva: df_existing.at[idx, "SỐ KTVA"] = so_ktva
                    if ngay_kt_dt: df_existing.at[idx, "NGÀY KTVA"] = ngay_kt_dt.strftime("%d/%m/%Y")
                    if ngay_tn_dt: df_existing.at[idx, "NGÀY TIẾP NHẬN HỐ SƠ"] = ngay_tn_dt.strftime("%d/%m/%Y")
                    if han_xl_dt: df_existing.at[idx, "THỜI HẠN XỬ LÝ"] = han_xl_dt.strftime("%d/%m/%Y")
                    if final_unit: df_existing.at[idx, "ĐƠN VỊ"] = final_unit

                    if note_lam_giau and not df_existing.at[idx, "NOTE LÀM GIÀU DỮ LIỆU"]:
                        df_existing.at[idx, "NOTE LÀM GIÀU DỮ LIỆU"] = note_lam_giau
                    if note_ngay_kt and not df_existing.at[idx, "NOTE NGÀY KIỂM TRA"]:
                        df_existing.at[idx, "NOTE NGÀY KIỂM TRA"] = note_ngay_kt
                    if ghi_chu_th and not df_existing.at[idx, "GHI CHÚ ĐỂ BÁO TÌNH HÌNH"]:
                        df_existing.at[idx, "GHI CHÚ ĐỂ BÁO TÌNH HÌNH"] = ghi_chu_th

                    curr_kt = parse_date(df_existing.at[idx, "NGÀY KTVA"])
                    if curr_kt:
                        df_existing.at[idx, "NGÀY HIỆN TẠI SO VỚI NGÀY KTVA"] = (today - curr_kt).days
                else:
                    new_idx = idx - len(df_existing)
                    if 0 <= new_idx < len(all_new_rows):
                        r_d = all_new_rows[new_idx]
                        if info_va: r_d["Thông tin vụ án"] = info_va
                        if loai_toi: r_d["Loại tội"] = loai_toi
                        if toi_val: r_d["Tội"] = toi_val
                        if bi_can: r_d["Thông tin bị can"] = bi_can
                        if thu_ly: r_d["Loại thụ lý"] = thu_ly
                        if trang_thai_va: r_d["Trạng thái vụ án"] = trang_thai_va
                        if trang_thai_bm: r_d["Trạng thái biểu mẫu"] = trang_thai_bm
                        if ten_va: r_d["TÊN VỤ ÁN"] = ten_va
                        if so_ktva: r_d["SỐ KTVA"] = so_ktva
                        if ngay_kt_dt: r_d["NGÀY KTVA"] = ngay_kt_dt.strftime("%d/%m/%Y")
                        if ngay_tn_dt: r_d["NGÀY TIẾP NHẬN HỐ SƠ"] = ngay_tn_dt.strftime("%d/%m/%Y")
                        if han_xl_dt: r_d["THỜI HẠN XỬ LÝ"] = han_xl_dt.strftime("%d/%m/%Y")
                        if final_unit: r_d["ĐƠN VỊ"] = final_unit
                        if note_lam_giau and not r_d.get("NOTE LÀM GIÀU DỮ LIỆU"):
                            r_d["NOTE LÀM GIÀU DỮ LIỆU"] = note_lam_giau
                        if note_ngay_kt and not r_d.get("NOTE NGÀY KIỂM TRA"):
                            r_d["NOTE NGÀY KIỂM TRA"] = note_ngay_kt
                        if ghi_chu_th and not r_d.get("GHI CHÚ ĐỂ BÁO TÌNH HÌNH"):
                            r_d["GHI CHÚ ĐỂ BÁO TÌNH HÌNH"] = ghi_chu_th
                        curr_kt = parse_date(r_d.get("NGÀY KTVA"))
                        if curr_kt:
                            r_d["NGÀY HIỆN TẠI SO VỚI NGÀY KTVA"] = (today - curr_kt).days

                updated_cases_count += 1
                continue

            # 2. Nếu là vụ án mới chưa từng có:
            new_added_count += 1
            final_days = file_so_ngay if (file_so_ngay and file_so_ngay.isdigit()) else so_ngay
            row_dict = {
                "STT": len(all_new_rows) + 1,
                "Thông tin vụ án": info_va,
                "Loại tội": loai_toi,
                "Tội": toi_val,
                "Thông tin bị can": bi_can,
                "Loại thụ lý": thu_ly,
                "Trạng thái vụ án": trang_thai_va,
                "Trạng thái biểu mẫu": trang_thai_bm,
                "MÃ VỤ ÁN": ma_va,
                "TÊN VỤ ÁN": ten_va,
                "SỐ KTVA": so_ktva,
                "NGÀY KTVA": ngay_kt_dt.strftime("%d/%m/%Y") if ngay_kt_dt else "",
                "NGÀY TIẾP NHẬN HỐ SƠ": ngay_tn_dt.strftime("%d/%m/%Y") if ngay_tn_dt else "",
                "THỜI HẠN XỬ LÝ": han_xl_dt.strftime("%d/%m/%Y") if han_xl_dt else "",
                "NOTE LÀM GIÀU DỮ LIỆU": note_lam_giau,
                "NGÀY HIỆN TẠI SO VỚI NGÀY KTVA": final_days,
                "ĐƠN VỊ": final_unit,
                "NOTE NGÀY KIỂM TRA": note_ngay_kt,
                "NOTE": note_val,
                "GHI CHÚ ĐỂ BÁO TÌNH HÌNH": ghi_chu_th,
                "NOTE KIỂM TRA": note_kt_val,
                "NGÀY BÁO TÌNH HÌNH": ngay_bao_th
            }
            all_new_rows.append(row_dict)
            if df_existing is not None:
                case_index_map[ma_va_clean] = len(df_existing) + len(all_new_rows) - 1

    # Ghép toàn bộ dữ liệu vào File Tổng
    if df_existing is not None:
        if all_new_rows:
            df_new = pd.DataFrame(all_new_rows)
            df_tong = pd.concat([df_existing, df_new], ignore_index=True)
        else:
            df_tong = df_existing.copy()

        df_tong["STT"] = range(1, len(df_tong) + 1)
        df_tong["ĐƠN VỊ"] = df_tong["ĐƠN VỊ"].apply(standardize_unit_name)
        status_msg = (
            f"Đã bảo lưu toàn bộ {total_existing} vụ án từ File Tổng sẵn có.\n"
            f"- Đã cập nhật lại thông tin mới: {updated_cases_count} vụ từ {len(files)} file đơn vị.\n"
            f"- Thêm mới: {new_added_count} vụ án phát sinh.\n"
            f"- Tổng cộng File Tổng: {len(df_tong)} vụ án."
        )
    else:
        if not all_new_rows:
            return False, {
                "error": "Không trích xuất được dòng dữ liệu nào từ các file!",
                "message": "Các file Excel không chứa cột 'Mã vụ án' hoặc 'Thông tin vụ án'."
            }
        df_tong = pd.DataFrame(all_new_rows)
        df_tong["STT"] = range(1, len(df_tong) + 1)
        df_tong["ĐƠN VỊ"] = df_tong["ĐƠN VỊ"].apply(standardize_unit_name)
        status_msg = f"Đã tổng hợp thành công {len(files)} file đơn vị (Tổng {len(df_tong)} vụ án) và tách các vụ án từ {check_days} ngày trở lên."

    for c in MASTER_COLUMNS:
        if c not in df_tong.columns:
            df_tong[c] = ""
    df_tong = df_tong[MASTER_COLUMNS]

    if use_timestamp:
        now_ts = datetime.now().strftime("%Y%m%d_%H%M")
        file_tong_path = os.path.join(output_dir, f"1_FILE_TONG_{now_ts}.xlsx")
        file_kt_path = os.path.join(output_dir, f"2_DANH_SACH_CAN_KIEM_TRA_{now_ts}.xlsx")
        file_new_path = os.path.join(output_dir, f"DANH_SACH_VU_AN_MOI_PHAT_SINH_{now_ts}.xlsx")
    else:
        file_tong_path = os.path.join(output_dir, "1_FILE_TONG.xlsx")
        file_kt_path = os.path.join(output_dir, "2_DANH_SACH_CAN_KIEM_TRA.xlsx")
        file_new_path = os.path.join(output_dir, "DANH_SACH_VU_AN_MOI_PHAT_SINH.xlsx")

    # TỰ ĐỘNG SAO LƯU FILE CŨ NẾU ĐÃ TỒN TẠI (CHỐNG GHI ĐÈ MẤT DỮ LIỆU)
    backup_tong = backup_file_if_exists(file_tong_path)
    try:
        save_master_excel(df_tong, file_tong_path, is_inspection_file=False)
    except PermissionError:
        return False, {
            "error": "File đang mở trong Excel!",
            "message": f"File '{os.path.basename(file_tong_path)}' đang được mở trong Excel. Vui lòng đóng file Excel lại rồi thử lại!"
        }
    except Exception as e:
        return False, {
            "error": f"Lỗi lưu file: {str(e)}",
            "message": f"Không thể lưu file '{os.path.basename(file_tong_path)}'."
        }

    # TỰ ĐỘNG XUẤT FILE CÁC VỤ ÁN MỚI PHÁT SINH (NẾU CÓ VỤ ÁN MỚI TRONG CHẾ ĐỘ NẠP NỐI TIẾP)
    file_new_saved = None
    backup_new = None
    if df_existing is not None and all_new_rows:
        df_new_cases = pd.DataFrame(all_new_rows)
        df_new_cases["STT"] = range(1, len(df_new_cases) + 1)
        df_new_cases["ĐƠN VỊ"] = df_new_cases["ĐƠN VỊ"].apply(standardize_unit_name)
        for c in MASTER_COLUMNS:
            if c not in df_new_cases.columns:
                df_new_cases[c] = ""
        df_new_cases = df_new_cases[MASTER_COLUMNS]

        backup_new = backup_file_if_exists(file_new_path)
        try:
            save_master_excel(df_new_cases, file_new_path, is_inspection_file=False)
            file_new_saved = os.path.abspath(file_new_path)
        except PermissionError:
            return False, {
                "error": "File đang mở trong Excel!",
                "message": f"File '{os.path.basename(file_new_path)}' đang được mở trong Excel. Vui lòng đóng file Excel lại rồi thử lại!"
            }
        except Exception as e:
            return False, {
                "error": f"Lỗi lưu file: {str(e)}",
                "message": f"Không thể lưu file '{os.path.basename(file_new_path)}'."
            }

    # Lọc danh sách cần kiểm tra theo đúng 2 tiêu chí nghiệp vụ:
    # 1. Số ngày KTVA từ check_days trở lên (mặc định >= 16 ngày)
    # 2. CHƯA CÓ GHI CHÚ tại cột NOTE LÀM GIÀU DỮ LIỆU (ô trống hoàn toàn)
    def is_empty_rich_note(v):
        if v is None or pd.isna(v):
            return True
        s = str(v).strip()
        return s == "" or s.lower() in ("nan", "none", "nat", "null")

    mask_days = pd.to_numeric(df_tong["NGÀY HIỆN TẠI SO VỚI NGÀY KTVA"], errors="coerce") >= check_days
    mask_empty_rich = df_tong["NOTE LÀM GIÀU DỮ LIỆU"].apply(is_empty_rich_note)

    df_kt = df_tong[mask_days & mask_empty_rich].copy()
    df_kt["STT"] = range(1, len(df_kt) + 1)
    df_kt = df_kt[MASTER_COLUMNS]

    # TỰ ĐỘNG SAO LƯU FILE DANH SÁCH KIỂM TRA CŨ (NƠI CHỨA CÁC NOTE CỦA CÁN BỘ)
    backup_kt = backup_file_if_exists(file_kt_path)
    try:
        save_master_excel(df_kt, file_kt_path, is_inspection_file=True)
    except PermissionError:
        return False, {
            "error": "File đang mở trong Excel!",
            "message": f"File '{os.path.basename(file_kt_path)}' đang được mở trong Excel. Vui lòng đóng file Excel lại rồi thử lại!"
        }
    except Exception as e:
        return False, {
            "error": f"Lỗi lưu file: {str(e)}",
            "message": f"Không thể lưu file '{os.path.basename(file_kt_path)}'."
        }

    backups = [b for b in [backup_tong, backup_kt, backup_new] if b]

    noted_cnt = int((~mask_empty_rich).sum())
    unnoted_cnt = int(mask_empty_rich.sum())
    filtered_cnt = len(df_kt)

    result_data = {
        "success": True,
        "mode": "append" if (df_existing is not None) else "create",
        "message": status_msg,
        "total_files": len(files),
        "total_existing": total_existing,
        "new_added": new_added_count,
        "updated_cases": updated_cases_count,
        "skipped_duplicates": updated_cases_count,
        "total_cases": len(df_tong),
        "noted_cases": noted_cnt,
        "unnoted_cases": unnoted_cnt,
        "cases_filtered": filtered_cnt,
        "under_days_cases": unnoted_cnt - filtered_cnt,
        "check_days": check_days,
        "file_tong": os.path.abspath(file_tong_path),
        "file_kt": os.path.abspath(file_kt_path),
        "file_moi_phat_sinh": file_new_saved,
        "backup_files": backups,
        "output_dir": os.path.abspath(output_dir),
        "files_processed": [os.path.basename(f) for f in files]
    }
    return True, result_data

# ==================== BƯỚC 2: CẬP NHẬT KẾT QUẢ & BÁO CÁO ====================

def buoc_2_cap_nhat_va_bao_cao(file_tong_path, file_kt_updated_path, output_dir=None, use_timestamp=False):
    """
    CẬP NHẬT KẾT QUẢ KIỂM TRA & XUẤT BÁO CÁO:
    - Nhận file 1_FILE_TONG.xlsx và file 2_DANH_SACH_CAN_KIEM_TRA.xlsx đã ghi chú.
    - Tiêu chí đánh giá lỗi tại NOTE LÀM GIÀU DỮ LIỆU: Chỉ có điền 'ĐÃ ĐÍNH KÈM', 'TẠM ỔN' là không lỗi. Có điền thông tin khác là có lỗi.
    - Cán bộ ghi chú cụ thể trường hợp tại ô GHI CHÚ ĐỂ BÁO TÌNH HÌNH.
    - Vụ nào chưa kiểm tra (ô để trống): Bảo lưu hoặc để trống hoàn toàn trong file hoàn thiện.
    - ĐÃ LOẠI BỎ HOÀN TOÀN CỘT NOTE LỖI.
    - Xuất file 3_FILE_TONG_HOAN_THIEN.xlsx theo đúng mẫu 22 cột chuẩn nền vàng, KÈM SHEET 'BẢNG THEO DÕI' Pivot Table.
    - Xuất file 4_BAO_CAO_VU_LOI_INFOGRAPHIC.xlsx kèm thẻ KPI và Biểu đồ thống kê theo Đơn vị.
    - Xuất file 5_DANH_SACH_VU_LOI_GUI_CAC_DON_VI.xlsx gồm 7 cột chuẩn nền xanh sẫm (#156082), bật AutoFilter, Freeze Panes C2.
    - Tự động tạo thư mục 'Cac_Don_Vi_Khac_Phuc' chứa file lỗi riêng biệt của từng đơn vị để gửi ngay.
    """
    if not output_dir:
        output_dir = os.path.dirname(file_tong_path) or "."

    today_dt = datetime.now().date()
    today_str = today_dt.strftime("%d/%m/%Y")

    try:
        df_tong = pd.read_excel(file_tong_path)
        df_tong = normalize_column_names(df_tong)
        df_kt = pd.read_excel(file_kt_updated_path)
        df_kt = normalize_column_names(df_kt)
    except Exception as e:
        return False, {
            "error": f"Lỗi đọc file Excel: {str(e)}",
            "message": "Không thể đọc dữ liệu từ file tổng hoặc file danh sách kiểm tra."
        }

    # Đảm bảo các cột ghi chú có kiểu dữ liệu dạng chuỗi (object)
    note_cols = [
        "NOTE LÀM GIÀU DỮ LIỆU", "GHI CHÚ ĐỂ BÁO TÌNH HÌNH",
        "NOTE KIỂM TRA", "NOTE NGÀY KIỂM TRA", "NOTE", "NGÀY BÁO TÌNH HÌNH"
    ]
    for col in note_cols:
        if col in df_tong.columns:
            df_tong[col] = df_tong[col].astype("object").fillna("")
        else:
            df_tong[col] = ""
        if col in df_kt.columns:
            df_kt[col] = df_kt[col].astype("object").fillna("")
        else:
            df_kt[col] = ""

    notes_by_compound = {}
    notes_by_mava = {}
    auto_filled_dates_count = 0
    checked_cases_latest = []

    for _, row in df_kt.iterrows():
        don_vi = str(row.get("ĐƠN VỊ", "") if pd.notna(row.get("ĐƠN VỊ")) else "").strip()
        ma_va = str(row.get("MÃ VỤ ÁN", "") if pd.notna(row.get("MÃ VỤ ÁN")) else "").strip()

        raw_note_enrich = str(row.get("NOTE LÀM GIÀU DỮ LIỆU", "") if pd.notna(row.get("NOTE LÀM GIÀU DỮ LIỆU")) else "").strip()
        if raw_note_enrich.lower() in ("nan", "none"):
            raw_note_enrich = ""

        raw_note_daily = str(row.get("GHI CHÚ ĐỂ BÁO TÌNH HÌNH", "") if pd.notna(row.get("GHI CHÚ ĐỂ BÁO TÌNH HÌNH")) else "").strip()
        if raw_note_daily.lower() in ("nan", "none"):
            raw_note_daily = ""

        raw_note_free = str(row.get("NOTE", "") if pd.notna(row.get("NOTE")) else "").strip()
        if raw_note_free.lower() in ("nan", "none"):
            raw_note_free = ""

        raw_note_kt = str(row.get("NOTE KIỂM TRA", "") if pd.notna(row.get("NOTE KIỂM TRA")) else "").strip()
        if raw_note_kt.lower() in ("nan", "none"):
            raw_note_kt = ""

        # Lấy ngày kiểm tra do người dùng trực tiếp điền tay vào cột NOTE NGÀY KIỂM TRA
        raw_note_date_cell = row.get("NOTE NGÀY KIỂM TRA")
        d_parsed = parse_date(raw_note_date_cell)
        if d_parsed:
            raw_note_date = d_parsed.strftime("%d/%m/%Y")
        else:
            raw_note_date = str(raw_note_date_cell or "").strip()
            if raw_note_date.lower() in ("nan", "none", "nat", "null"):
                raw_note_date = ""

        # Lấy ngày báo tình hình do người dùng trực tiếp điền tay vào (nếu có)
        raw_ngay_bao = str(row.get("NGÀY BÁO TÌNH HÌNH", "") if pd.notna(row.get("NGÀY BÁO TÌNH HÌNH")) else "").strip()
        if raw_ngay_bao.lower() in ("nan", "none", "nat", "null"):
            raw_ngay_bao = ""

        # TIÊU CHÍ NGHIỆP VỤ BẮT BUỘC:
        # Một vụ án trong file kiểm tra CHỈ ĐƯỢC COI LÀ ĐÃ KIỂM TRA khi và chỉ khi
        # Cán bộ ĐÃ TRỰC TIẾP GHI CHÚ vào cột 'NOTE LÀM GIÀU DỮ LIỆU' (raw_note_enrich có nội dung)!
        # Các cột khác (NOTE, NOTE KIỂM TRA, GHI CHÚ TÌNH HÌNH) nếu còn sót từ file cũ TUYỆT ĐỐI không tính là vụ mới kiểm tra.
        has_checked = bool(raw_note_enrich)

        if has_checked:
            raw_note_kt = "ĐÃ KIỂM TRA"
            # Nếu cán bộ quên điền ngày kiểm tra: Tự động bù Ngày thao tác hiện tại
            if not raw_note_date:
                raw_note_date = today_str
                auto_filled_dates_count += 1
        else:
            # Chưa kiểm tra (ô Làm giàu dữ liệu để trống): Bỏ qua hoàn toàn, không tính vào thống kê đợt này
            raw_note_enrich = ""
            raw_note_daily = ""
            raw_note_date = ""
            raw_note_kt = ""

        note_data = {
            "has_checked": has_checked,
            "NOTE LÀM GIÀU DỮ LIỆU": raw_note_enrich,
            "GHI CHÚ ĐỂ BÁO TÌNH HÌNH": raw_note_daily,
            "NOTE": raw_note_free,
            "NOTE KIỂM TRA": raw_note_kt,
            "NOTE NGÀY KIỂM TRA": raw_note_date,  # Tự động điền ngày thao tác nếu cán bộ quên
            "NGÀY BÁO TÌNH HÌNH": raw_ngay_bao
        }

        if don_vi and ma_va:
            notes_by_compound[make_compound_key(don_vi, ma_va)] = note_data
        if ma_va:
            notes_by_mava[remove_accents(ma_va).replace(" ", "").strip()] = note_data

        if has_checked:
            row_dict = row.to_dict() if hasattr(row, "to_dict") else dict(row)
            row_dict["NOTE NGÀY KIỂM TRA"] = raw_note_date
            row_dict["NOTE LÀM GIÀU DỮ LIỆU"] = raw_note_enrich
            row_dict["NOTE KIỂM TRA"] = raw_note_kt
            row_dict["GHI CHÚ ĐỂ BÁO TÌNH HÌNH"] = raw_note_daily
            checked_cases_latest.append(row_dict)

    # Cập nhật ngược lại File Tổng
    for idx, row in df_tong.iterrows():
        don_vi = str(row.get("ĐƠN VỊ", "") if pd.notna(row.get("ĐƠN VỊ")) else "").strip()
        ma_va = str(row.get("MÃ VỤ ÁN", "") if pd.notna(row.get("MÃ VỤ ÁN")) else "").strip()
        ma_va_clean = remove_accents(ma_va).replace(" ", "").strip()
        c_key = make_compound_key(don_vi, ma_va)

        matched_data = None
        if c_key in notes_by_compound:
            matched_data = notes_by_compound[c_key]
        elif ma_va_clean in notes_by_mava:
            matched_data = notes_by_mava[ma_va_clean]

        if matched_data:
            if matched_data.get("has_checked"):
                # CẬP NHẬT KHI ĐÃ KIỂM TRA
                if matched_data.get("NOTE LÀM GIÀU DỮ LIỆU"):
                    df_tong.at[idx, "NOTE LÀM GIÀU DỮ LIỆU"] = matched_data.get("NOTE LÀM GIÀU DỮ LIỆU")
                if matched_data.get("GHI CHÚ ĐỂ BÁO TÌNH HÌNH"):
                    df_tong.at[idx, "GHI CHÚ ĐỂ BÁO TÌNH HÌNH"] = matched_data.get("GHI CHÚ ĐỂ BÁO TÌNH HÌNH")
                if matched_data.get("NOTE"):
                    df_tong.at[idx, "NOTE"] = matched_data.get("NOTE")
                if matched_data.get("NOTE KIỂM TRA"):
                    df_tong.at[idx, "NOTE KIỂM TRA"] = matched_data.get("NOTE KIỂM TRA")
                if matched_data.get("NOTE NGÀY KIỂM TRA"):
                    df_tong.at[idx, "NOTE NGÀY KIỂM TRA"] = matched_data.get("NOTE NGÀY KIỂM TRA")
                if matched_data.get("NGÀY BÁO TÌNH HÌNH"):
                    df_tong.at[idx, "NGÀY BÁO TÌNH HÌNH"] = matched_data.get("NGÀY BÁO TÌNH HÌNH")

    # Đảm bảo đúng 22 cột quy chuẩn
    for c in MASTER_COLUMNS:
        if c not in df_tong.columns:
            df_tong[c] = ""
    df_tong = df_tong[MASTER_COLUMNS]

    # 1. Xác định ngày kiểm tra của đợt này (Ưu tiên ngày thao tác hiện tại nếu có bù ngày, hoặc ngày lớn nhất trong file)
    dates_found = []
    for r in checked_cases_latest:
        d_p = parse_date(r.get("NOTE NGÀY KIỂM TRA"))
        if d_p:
            dates_found.append(d_p)

    if auto_filled_dates_count > 0:
        target_date_dt = today_dt
        target_date_str = today_str
    elif dates_found:
        target_date_dt = max(dates_found)
        target_date_str = target_date_dt.strftime("%d/%m/%Y")
    else:
        target_date_dt = today_dt
        target_date_str = today_str

    # Lọc các vụ án kiểm tra thuộc đợt kiểm tra mới nhất
    inspected_final = []
    for r in checked_cases_latest:
        d_p = parse_date(r.get("NOTE NGÀY KIỂM TRA"))
        if not d_p or d_p == target_date_dt:
            inspected_final.append(r)
    if not inspected_final and checked_cases_latest:
        inspected_final = checked_cases_latest

    # 2. Thống kê theo đơn vị (đã kiểm tra, đã đính kèm, lỗi)
    unit_stats = {}
    error_cases = []

    for r in inspected_final:
        raw_u = r.get("ĐƠN VỊ", "")
        u = standardize_unit_name(raw_u) or "KHÔNG XÁC ĐỊNH"
        if u not in unit_stats:
            unit_stats[u] = {"total": 0, "dat": 0, "loi": 0}

        rich = str(r.get("NOTE LÀM GIÀU DỮ LIỆU", "") if pd.notna(r.get("NOTE LÀM GIÀU DỮ LIỆU")) else "").strip()
        if is_case_error(rich):
            unit_stats[u]["loi"] += 1
            error_cases.append(r)
        else:
            unit_stats[u]["dat"] += 1
        unit_stats[u]["total"] += 1

    # SẮP XẾP SỐ LƯỢNG LỖI CỦA CÁC ĐƠN VỊ TỪ NHIỀU TỚI ÍT
    sorted_units = sorted(unit_stats.items(), key=lambda x: (x[1]["loi"], x[1]["total"]), reverse=True)
    total_kt = sum(s["total"] for _, s in sorted_units)
    total_dat = sum(s["dat"] for _, s in sorted_units)
    total_loi = sum(s["loi"] for _, s in sorted_units)
    total_rate = (total_loi / max(total_kt, 1) * 100) if total_kt > 0 else 0.0

    # Sắp xếp danh sách chi tiết các vụ lỗi theo thứ tự đơn vị có nhiều lỗi nhất
    unit_order_map = {u: i for i, (u, _) in enumerate(sorted_units)}
    error_cases.sort(key=lambda r: unit_order_map.get(standardize_unit_name(r.get("ĐƠN VỊ", "")), 999))

    unit_summary_data = {
        "sorted_units": sorted_units,
        "total_kt": total_kt,
        "total_dat": total_dat,
        "total_loi": total_loi,
        "total_rate": total_rate,
        "target_date_str": target_date_str
    }

    if use_timestamp:
        now_ts = datetime.now().strftime("%Y%m%d_%H%M")
        file_final_master = os.path.join(output_dir, f"3_FILE_TONG_HOAN_THIEN_{now_ts}.xlsx")
        file_report = os.path.join(output_dir, f"4_BAO_CAO_VU_LOI_INFOGRAPHIC_{now_ts}.xlsx")
        file_remedy = os.path.join(output_dir, f"5_DANH_SACH_VU_LOI_GUI_CAC_DON_VI_{now_ts}.xlsx")
    else:
        file_final_master = os.path.join(output_dir, "3_FILE_TONG_HOAN_THIEN.xlsx")
        file_report = os.path.join(output_dir, "4_BAO_CAO_VU_LOI_INFOGRAPHIC.xlsx")
        file_remedy = os.path.join(output_dir, "5_DANH_SACH_VU_LOI_GUI_CAC_DON_VI.xlsx")

    # 3. Lưu File 3_FILE_TONG_HOAN_THIEN.xlsx (kèm Sheet 'BẢNG THEO DÕI' tóm lược đợt kiểm tra mới nhất)
    backup_master = backup_file_if_exists(file_final_master)
    try:
        save_master_excel(
            df_tong, file_final_master,
            is_inspection_file=False,
            include_pivot_sheet=True,
            today_str=target_date_str,
            unit_summary_data=unit_summary_data
        )
    except PermissionError:
        return False, {
            "error": "File đang mở trong Excel!",
            "message": f"File '{os.path.basename(file_final_master)}' đang được mở trong Excel. Vui lòng đóng file Excel lại rồi thử lại!"
        }
    except Exception as e:
        return False, {
            "error": f"Lỗi lưu file: {str(e)}",
            "message": f"Không thể lưu file '{os.path.basename(file_final_master)}'."
        }

    # 4. Tạo File 4_BAO_CAO_VU_LOI_INFOGRAPHIC.xlsx (Báo cáo lãnh đạo chuẩn mẫu nghiệp vụ)
    backup_rep = backup_file_if_exists(file_report)
    wb_rep = openpyxl.Workbook()
    ws = wb_rep.active
    ws.title = "BaoCao_LanhDao"
    ws.views.sheetView[0].showGridLines = True

    # Tiêu đề báo cáo
    ws.merge_cells("A1:H1")
    t = ws["A1"]
    t.value = "BÁO CÁO KẾT QUẢ KIỂM TRA ÁN HÌNH SỰ - BÁO CÁO LÃNH ĐẠO"
    t.font = Font(name="Segoe UI", size=14, bold=True, color="C00000")
    t.alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("A2:H2")
    sub = ws["A2"]
    sub.value = (
        f"Viện kiểm sát nhân dân TP. Hồ Chí Minh — Ngày kiểm tra: {target_date_str} | "
        f"Xuất báo cáo lúc: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    )
    sub.font = Font(name="Segoe UI", size=10, italic=True)
    sub.alignment = Alignment(horizontal="center", vertical="center")

    # Thẻ KPI Tổng quan
    kpi_defs = [
        ("A4:B4", "A5:B5", "TỔNG SỐ VỤ ĐÃ KIỂM TRA", total_kt, "0B2545"),
        ("C4:D4", "C5:D5", "SỐ VỤ ĐÃ ĐÍNH KÈM (ĐẠT)", total_dat, "2E7D32"),
        ("E4:F4", "E5:F5", "SỐ VỤ PHÁT HIỆN LỖI", total_loi, "C00000"),
        ("G4:H4", "G5:H5", "TỶ LỆ LỖI TOÀN THÀNH PHỐ", f"{total_rate:.1f}%", "7F0000")
    ]
    fill_kpi = PatternFill(start_color="F7F9FA", end_color="F7F9FA", fill_type="solid")
    for top_range, val_range, title_text, val_text, color_code in kpi_defs:
        ws.merge_cells(top_range)
        ws.merge_cells(val_range)
        t_cell = ws[top_range.split(":")[0]]
        v_cell = ws[val_range.split(":")[0]]
        t_cell.value = title_text
        t_cell.font = Font(name="Segoe UI", size=9, bold=True, color="555555")
        t_cell.alignment = Alignment(horizontal="center", vertical="center")
        t_cell.fill = fill_kpi
        v_cell.value = val_text
        v_cell.font = Font(name="Segoe UI", size=16, bold=True, color=color_code)
        v_cell.alignment = Alignment(horizontal="center", vertical="center")
        v_cell.fill = fill_kpi

    border_thin = Border(
        left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
    )

    # I. BẢNG TỔNG HỢP THEO ĐƠN VỊ (BÁO CÁO LÃNH ĐẠO)
    start_sec1 = 7
    ws.cell(row=start_sec1, column=1, value="I. BẢNG TỔNG HỢP THEO ĐƠN VỊ (SẮP XẾP SỐ LƯỢNG LỖI TỪ NHIỀU ĐẾN ÍT)").font = Font(name="Segoe UI", size=11, bold=True, color="0B2545")

    headers_sec1 = ["STT", "ĐƠN VỊ", "TỔNG SỐ VỤ ĐÃ KIỂM TRA", "SỐ VỤ ĐÃ ĐÍNH KÈM", "SỐ VỤ LỖI", "TỶ LỆ LỖI (%)"]
    h_row_sec1 = start_sec1 + 1

    for col_idx, h_text in enumerate(headers_sec1, 1):
        cell = ws.cell(row=h_row_sec1, column=col_idx, value=h_text)
        cell.font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="0B2545", end_color="0B2545", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border_thin

    curr_r = h_row_sec1 + 1
    for i, (u_name, s) in enumerate(sorted_units, 1):
        rate_val = (s["loi"] / s["total"] * 100) if s["total"] > 0 else 0.0
        ws.cell(row=curr_r, column=1, value=i).alignment = Alignment(horizontal="center")
        ws.cell(row=curr_r, column=2, value=f"  {u_name}").alignment = Alignment(horizontal="left")
        ws.cell(row=curr_r, column=3, value=s["total"]).alignment = Alignment(horizontal="center")
        ws.cell(row=curr_r, column=4, value=s["dat"]).alignment = Alignment(horizontal="center")
        c_loi = ws.cell(row=curr_r, column=5, value=s["loi"])
        c_loi.alignment = Alignment(horizontal="center")
        if s["loi"] > 0:
            c_loi.font = Font(name="Segoe UI", size=10, bold=True, color="C00000")
        c_rate = ws.cell(row=curr_r, column=6, value=f"{rate_val:.1f}%")
        c_rate.alignment = Alignment(horizontal="center")
        if s["loi"] > 0:
            c_rate.font = Font(name="Segoe UI", size=10, bold=True, color="C00000")

        for c_i in range(1, 7):
            ws.cell(row=curr_r, column=c_i).border = border_thin
        curr_r += 1

    # Dòng TỔNG CỘNG
    fill_grand = PatternFill(start_color="E9EEF4", end_color="E9EEF4", fill_type="solid")
    font_grand = Font(name="Segoe UI", size=10, bold=True, color="000000")

    ws.cell(row=curr_r, column=1, value="")
    ws.cell(row=curr_r, column=2, value="TỔNG CỘNG").alignment = Alignment(horizontal="left")
    ws.cell(row=curr_r, column=3, value=total_kt).alignment = Alignment(horizontal="center")
    ws.cell(row=curr_r, column=4, value=total_dat).alignment = Alignment(horizontal="center")
    c_tot_loi = ws.cell(row=curr_r, column=5, value=total_loi)
    c_tot_loi.alignment = Alignment(horizontal="center")
    c_tot_rate = ws.cell(row=curr_r, column=6, value=f"{total_rate:.1f}%")
    c_tot_rate.alignment = Alignment(horizontal="center")

    for c_i in range(1, 7):
        c = ws.cell(row=curr_r, column=c_i)
        c.font = font_grand
        c.fill = fill_grand
        c.border = border_thin
    if total_loi > 0:
        c_tot_loi.font = Font(name="Segoe UI", size=10, bold=True, color="C00000")
        c_tot_rate.font = Font(name="Segoe UI", size=10, bold=True, color="C00000")

    # II. DANH SÁCH CHI TIẾT CÁC VỤ ÁN PHÁT HIỆN LỖI
    start_sec2 = curr_r + 2
    ws.cell(row=start_sec2, column=1, value=f"II. DANH SÁCH CHI TIẾT CÁC VỤ ÁN PHÁT HIỆN LỖI HOẶC CHẬM TIẾN ĐỘ (TỔNG CỘNG: {total_loi} VỤ)").font = Font(name="Segoe UI", size=11, bold=True, color="C00000")

    headers_sec2 = ["STT", "MÃ VỤ ÁN", "ĐƠN VỊ", "TÊN VỤ ÁN", "SỐ NGÀY KTVA", "NOTE LÀM GIÀU DỮ LIỆU", "GHI CHÚ ĐỂ BÁO TÌNH HÌNH", "ĐỀ XUẤT THAM MƯU"]
    h_row_sec2 = start_sec2 + 1

    for col_idx, h_text in enumerate(headers_sec2, 1):
        cell = ws.cell(row=h_row_sec2, column=col_idx, value=h_text)
        cell.font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border_thin

    curr_r2 = h_row_sec2 + 1
    if total_loi == 0:
        ws.merge_cells(f"A{curr_r2}:H{curr_r2}")
        msg_cell = ws[f"A{curr_r2}"]
        msg_cell.value = "✅ Không phát hiện vụ án nào có lỗi hoặc vi phạm trong đợt kiểm tra này."
        msg_cell.font = Font(name="Segoe UI", size=10, italic=True, color="2E7D32")
        msg_cell.alignment = Alignment(horizontal="center", vertical="center")
    else:
        for idx, r in enumerate(error_cases, 1):
            ma_va = str(r.get("MÃ VỤ ÁN", "") if pd.notna(r.get("MÃ VỤ ÁN")) else "")
            don_vi = standardize_unit_name(r.get("ĐƠN VỊ", "")) or str(r.get("ĐƠN VỊ", ""))
            ten_va = str(r.get("TÊN VỤ ÁN", "") if pd.notna(r.get("TÊN VỤ ÁN")) else "")
            so_ngay = r.get("NGÀY HIỆN TẠI SO VỚI NGÀY KTVA", "")
            rich_note = str(r.get("NOTE LÀM GIÀU DỮ LIỆU", "") if pd.notna(r.get("NOTE LÀM GIÀU DỮ LIỆU")) else "")
            daily_note = str(r.get("GHI CHÚ ĐỂ BÁO TÌNH HÌNH", "") if pd.notna(r.get("GHI CHÚ ĐỂ BÁO TÌNH HÌNH")) else "")

            ws.cell(row=curr_r2, column=1, value=idx).alignment = Alignment(horizontal="center")
            ws.cell(row=curr_r2, column=2, value=ma_va).alignment = Alignment(horizontal="center")
            ws.cell(row=curr_r2, column=3, value=don_vi).alignment = Alignment(horizontal="center")
            ws.cell(row=curr_r2, column=4, value=ten_va).alignment = Alignment(vertical="center")
            ws.cell(row=curr_r2, column=5, value=so_ngay).alignment = Alignment(horizontal="center")
            c_rich = ws.cell(row=curr_r2, column=6, value=rich_note)
            c_rich.font = Font(name="Segoe UI", size=10, bold=True, color="C00000")
            c_rich.alignment = Alignment(vertical="center")
            ws.cell(row=curr_r2, column=7, value=daily_note).alignment = Alignment(vertical="center")
            ws.cell(row=curr_r2, column=8, value="Yêu cầu đơn vị rà soát, khắc phục ngay").alignment = Alignment(vertical="center")

            for c_i in range(1, 9):
                ws.cell(row=curr_r2, column=c_i).border = border_thin
            curr_r2 += 1

    # Căn chỉnh độ rộng các cột trực quan, chuyên nghiệp
    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 24
    ws.column_dimensions["D"].width = 22
    ws.column_dimensions["E"].width = 16
    ws.column_dimensions["F"].width = 26
    ws.column_dimensions["G"].width = 30
    ws.column_dimensions["H"].width = 34

    try:
        wb_rep.save(file_report)
    except PermissionError:
        return False, {
            "error": "File đang mở trong Excel!",
            "message": f"File '{os.path.basename(file_report)}' đang được mở trong Excel. Vui lòng đóng file Excel lại rồi thử lại!"
        }
    except Exception as e:
        return False, {
            "error": f"Lỗi lưu file: {str(e)}",
            "message": f"Không thể lưu file '{os.path.basename(file_report)}'."
        }

    # 5. Tạo File 5_DANH_SACH_VU_LOI_GUI_CAC_DON_VI.xlsx (Theo mẫu 7 cột chuẩn của Viện)
    backup_rem = backup_file_if_exists(file_remedy)
    try:
        create_remedy_excel(error_cases, file_remedy, sheet_title="DANH_SACH_LOI_CAC_DON_VI", default_check_date=target_date_str)
    except PermissionError:
        return False, {
            "error": "File đang mở trong Excel!",
            "message": f"File '{os.path.basename(file_remedy)}' đang được mở trong Excel. Vui lòng đóng file Excel lại rồi thử lại!"
        }
    except Exception as e:
        return False, {
            "error": f"Lỗi lưu file khắc phục lỗi: {str(e)}",
            "message": f"Không thể lưu file '{os.path.basename(file_remedy)}'."
        }

    # 6. Tự động phân tách thành từng file lỗi riêng biệt cho từng đơn vị trong thư mục Cac_Don_Vi_Khac_Phuc
    remedy_unit_dir = os.path.join(output_dir, "Cac_Don_Vi_Khac_Phuc")
    unit_remedy_files = []
    if error_cases:
        os.makedirs(remedy_unit_dir, exist_ok=True)
        by_unit = {}
        for r in error_cases:
            u_clean = standardize_unit_name(r.get("ĐƠN VỊ", "")) or "KHAC"
            by_unit.setdefault(u_clean, []).append(r)

        for u_name, u_rows in by_unit.items():
            safe_u = remove_accents(u_name).replace(" ", "_").strip()
            safe_u = re.sub(r'[^\w\-]', '', safe_u)
            unit_file = os.path.join(remedy_unit_dir, f"{safe_u}_Khac_Phuc_Loi.xlsx")
            try:
                create_remedy_excel(u_rows, unit_file, sheet_title=u_name[:31], default_check_date=target_date_str)
                unit_remedy_files.append(os.path.abspath(unit_file))
            except Exception:
                pass

    backups_b2 = [b for b in [backup_master, backup_rep, backup_rem] if b]
    top_units_list = [{"ĐƠN VỊ": u, "SỐ VỤ LỖI": s["loi"], "TỔNG ĐÃ KT": s["total"]} for u, s in sorted_units]

    result_data = {
        "success": True,
        "message": f"Đã cập nhật kết quả và tạo thành công Báo cáo Lãnh đạo & File khắc phục lỗi gửi các đơn vị (Ngày kiểm tra {target_date_str}: {total_kt} vụ đã kiểm tra, {total_dat} vụ đính kèm, {total_loi} vụ phát hiện lỗi).",
        "total_cases": len(df_tong),
        "total_checked": total_kt,
        "total_dat": total_dat,
        "total_errors": total_loi,
        "error_rate": f"{total_rate:.1f}%",
        "target_date": target_date_str,
        "auto_filled_dates": auto_filled_dates_count,
        "file_master": os.path.abspath(file_final_master),
        "file_report": os.path.abspath(file_report),
        "file_remedy": os.path.abspath(file_remedy),
        "backup_files": backups_b2,
        "remedy_unit_dir": os.path.abspath(remedy_unit_dir) if error_cases else None,
        "unit_remedy_files": unit_remedy_files,
        "output_dir": os.path.abspath(output_dir),
        "top_units": top_units_list
    }
    return True, result_data

