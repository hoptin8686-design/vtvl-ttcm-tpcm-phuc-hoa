@echo off
chcp 65001 >nul
title TỰ ĐỘNG XÂY DỰNG ĐỀ ÁN VTVL VÀ ĐẨY LÊN GITHUB & VERCEL - THPT PHỤC HÒA
color 0B

echo ======================================================================
echo   HỆ THỐNG TỰ ĐỘNG ĐỒNG BỘ GITHUB & VERCEL
echo   Đề án Vị trí việc làm Viên chức toàn diện: 03 Danh mục
echo   1. Quản lý | 2. Chuyên môn nghiệp vụ | 3. Hỗ trợ, phục vụ
echo   Căn cứ Nghị định số 232/2026/NĐ-CP ngày 26/6/2026 của Chính phủ
echo   Trường THPT Phục Hòa - Tỉnh Cao Bằng
echo ======================================================================
echo.

set "PATH=C:\Program Files\Python39;C:\Users\DMX HOA THUAN\AppData\Local\Programs\Python\Python39;C:\Users\DMX HOA THUAN\AppData\Local\Microsoft\WindowsApps;C:\Users\DMX HOA THUAN\AppData\Local\Programs\Git\cmd;C:\Users\DMX HOA THUAN\AppData\Local\Programs\gh\bin;%PATH%"
cd /d "d:\Du-an-web\web-vtvl-ttcm-tpcm-phuc-hoa"

echo [1/4] Đang tự động tạo/cập nhật bảng Excel Đề án 3 Danh mục VTVL (07 Sheet)...
python generate_excel.py
if %errorlevel% neq 0 (
    echo [!] Cảnh báo: Không thể chạy Python, kiểm tra môi trường Python.
) else (
    echo       ✓ Đã tạo và đồng bộ bảng Excel ra màn hình Desktop thành công!
)

echo.
echo [2/4] Kiểm tra trạng thái Git...
git status --short

echo.
echo [3/4] Đang thêm tập tin và tạo bản lưu (Commit)...
git add .
git diff --cached --quiet
if %errorlevel% equ 0 (
    echo       - Không có thay đổi mới cần commit.
) else (
    for /f "tokens=1-4 delims=/ " %%a in ("%date%") do set mydate=%%a-%%b-%%c
    for /f "tokens=1-2 delims=: " %%a in ("%time%") do set mytime=%%a:%%b
    git commit -m "feat: Cap nhat De an VTVL tron bo 03 Danh muc (Quan ly, Chuyen mon, Ho tro) theo Nghi dinh 232/2026/ND-CP - %date% %time%"
    echo       ✓ Đã tạo commit thành công!
)

echo.
echo [4/4] Đang tự động đẩy lên GitHub (Tự động kích hoạt Vercel & GitHub Pages)...
git push origin main

if %errorlevel% equ 0 (
    echo.
    echo ======================================================================
    echo   ✓ ĐẨY LÊN GITHUB THÀNH CÔNG!
    echo   ✓ GitHub Pages & Vercel đang tự động triển khai (CI/CD)...
    echo ======================================================================
    echo.
    echo [LINK TRUY CẬP TRỰC TUYẾN]:
    echo   1. GitHub Pages (Đang chạy): https://hoptin8686-design.github.io/vtvl-ttcm-tpcm-phuc-hoa/
    echo   2. Kho mã nguồn GitHub:     https://github.com/hoptin8686-design/vtvl-ttcm-tpcm-phuc-hoa
    echo   3. Vercel Project:           https://vercel.com/new/clone?repository-url=https%%3A%%2F%%2Fgithub.com%%2Fhoptin8686-design%%2Fvtvl-ttcm-tpcm-phuc-hoa
    echo.
    echo [FILE EXCEL ĐÃ LƯU TRÊN MÀN HÌNH MÁY TÍNH (DESKTOP)]:
    echo   - D:\Desktop\Bang-mo-ta-De-an-VTVL-THPT-Phuc-Hoa-ND232-2026.xlsx
    echo.
) else (
    echo.
    echo [!] Có lỗi xảy ra trong quá trình push. Vui lòng kiểm tra lại kết nối mạng.
)

echo Bấm phím bất kỳ để đóng cửa sổ này...
pause >nul
