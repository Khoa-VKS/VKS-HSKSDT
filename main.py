# -*- coding: utf-8 -*-
"""
HỆ THỐNG TỔNG HỢP & GIÁM SÁT TIẾN ĐỘ ÁN HÌNH SỰ
VIỆN KIỂM SÁT NHÂN DÂN THÀNH PHỐ HỒ CHÍ MINH

ĐIỂM KHỞI CHẠY CHÍNH (MAIN ENTRYPOINT)
Tự động kích hoạt Backend nhúng và mở Giao diện Web Frontend trên trình duyệt.
Hoàn toàn OFFLINE trên máy tính, bảo mật tuyệt đối.
"""

import os
import sys

# Đảm bảo đường dẫn thư mục gốc và backend nằm trong sys.path
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
BACKEND_DIR = os.path.join(BASE_DIR, "backend")

if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)
if BACKEND_DIR not in sys.path:
    sys.path.insert(0, BACKEND_DIR)

from backend.server import start_server

def main():
    print("=" * 72)
    print("       VIỆN KIỂM SÁT NHÂN DÂN THÀNH PHỐ HỒ CHÍ MINH")
    print("   TỔNG HỢP VÀ GIÁM SÁT VIỆC LẬP HỒ SƠ KIỂM SÁT ĐIỆN TỬ")
    print("                 (KIẾN TRÚC FE & BE NHÚNG)")
    print("=" * 72)
    print("[*] Đang khởi động Backend nhúng nội bộ...")
    print("[*] Giao diện người dùng Frontend sẽ tự động mở trên trình duyệt...")
    print("[*] Nhấn Ctrl + C tại cửa sổ này khi muốn dừng hệ thống.")
    print("-" * 72)
    
    start_server(auto_open_browser=True)

if __name__ == "__main__":
    main()
