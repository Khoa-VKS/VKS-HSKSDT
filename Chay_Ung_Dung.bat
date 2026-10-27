@echo off
chcp 65001 >nul
title HỆ THỐNG TỔNG HỢP & GIÁM SÁT TIẾN ĐỘ ÁN HÌNH SỰ - VKSND TP.HCM
color 0B

echo ==============================================================================
echo       VIỆN KIỂM SÁT NHÂN DÂN THÀNH PHỐ HỒ CHÍ MINH
echo   HỆ THỐNG TỰ ĐỘNG HÓA TỔNG HỢP VÀ GIÁM SÁT TIẾN ĐỘ ÁN HÌNH SỰ
echo                 (MÔ HÌNH FE VÀ BE NHÚNG HOÀN TOÀN)
echo ==============================================================================
echo.
echo [1/2] Đang kiểm tra môi trường Python trên máy tính...

set PYTHON_CMD=

where py >nul 2>nul
if %errorlevel% equ 0 (
    set PYTHON_CMD=py
    goto :PYTHON_FOUND
)

where python >nul 2>nul
if %errorlevel% equ 0 (
    set PYTHON_CMD=python
    goto :PYTHON_FOUND
)

echo [!] KHÔNG TÌM THẤY PYTHON TRÊN MÁY TÍNH!
echo Vui lòng tải và cài đặt Python từ https://www.python.org/
echo Khi cài đặt, hãy nhớ tích chọn "Add Python to PATH".
echo.
pause
exit /b

:PYTHON_FOUND
echo [*] Đã tìm thấy Python: %PYTHON_CMD%
echo [2/2] Đang kiểm tra thư viện hỗ trợ (pandas, openpyxl)...
%PYTHON_CMD% -m pip install --quiet pandas openpyxl

echo [*] Đang khởi động Giao diện phần mềm VKSND TP.HCM...
echo.
where pythonw >nul 2>nul
if %errorlevel% equ 0 (
    start "" pythonw "%~dp0app_vks.py"
    exit
) else (
    %PYTHON_CMD% "%~dp0app_vks.py"
)
