@echo off
chcp 65001 >nul
title TỰ ĐỘNG ĐẨY LÊN GITHUB & VERCEL - THPT PHỤC HÒA
color 0B

echo ======================================================================
echo   HỆ THỐNG TỰ ĐỘNG ĐỒNG BỘ GITHUB ^& VERCEL
echo   Dự án: Vị trí việc làm TTCM ^& TPCM (Nghị định 232/2026/NĐ-CP)
echo   Trường THPT Phục Hòa - Tỉnh Cao Bằng
echo ======================================================================
echo.

set "PATH=C:\Users\DMX HOA THUAN\AppData\Local\Programs\Git\cmd;C:\Users\DMX HOA THUAN\AppData\Local\Programs\gh\bin;%PATH%"
cd /d "d:\Du-an-web\web-vtvl-ttcm-tpcm-phuc-hoa"

echo [1/3] Kiểm tra trạng thái Git...
git status --short

echo.
echo [2/3] Đang thêm tập tin và tạo bản lưu (Commit)...
git add .
git diff --cached --quiet
if %errorlevel% equ 0 (
    echo       - Không có thay đổi mới cần commit.
) else (
    for /f "tokens=1-4 delims=/ " %%a in ("%date%") do set mydate=%%a-%%b-%%c
    for /f "tokens=1-2 delims=: " %%a in ("%time%") do set mytime=%%a:%%b
    git commit -m "update: Dong bo tu dong ngay %date% luc %time%"
    echo       ✓ Đã tạo commit thành công!
)

echo.
echo [3/3] Đang tự động đẩy lên GitHub (Tự động kích hoạt Vercel ^& GitHub Pages)...
git push origin main

if %errorlevel% equ 0 (
    echo.
    echo ======================================================================
    echo   ✓ ĐẨY LÊN GITHUB THÀNH CÔNG!
    echo   ✓ GitHub Pages ^& Vercel đang tự động triển khai (CI/CD)...
    echo ======================================================================
    echo.
    echo [LINK TRUY CẬP TRỰC TUYẾN]:
    echo   1. GitHub Pages (Đang chạy): https://hoptin8686-design.github.io/vtvl-ttcm-tpcm-phuc-hoa/
    echo   2. Kho mã nguồn GitHub:     https://github.com/hoptin8686-design/vtvl-ttcm-tpcm-phuc-hoa
    echo   3. Vercel Project:           https://vercel.com/new/clone?repository-url=https%%3A%%2F%%2Fgithub.com%%2Fhoptin8686-design%%2Fvtvl-ttcm-tpcm-phuc-hoa
    echo.
) else (
    echo.
    echo [!] Có lỗi xảy ra trong quá trình push. Vui lòng kiểm tra lại kết nối mạng.
)

echo Bấm phím bất kỳ để đóng cửa sổ này...
pause >nul
