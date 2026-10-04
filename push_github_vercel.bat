@echo off
chcp 65001 >nul
set "PATH=C:\Users\DMX HOA THUAN\AppData\Local\Programs\Git\cmd;C:\Users\DMX HOA THUAN\AppData\Local\Programs\gh\bin;%PATH%"
set "PROJECT_DIR=d:\Du-an-web\web-vtvl-ttcm-tpcm-phuc-hoa"
set "GITHUB_REPO=vtvl-ttcm-tpcm-phuc-hoa"
set "GITHUB_USER=hoptin8686-design"

cd /d "%PROJECT_DIR%"

echo.
echo ╔══════════════════════════════════════════════════════════════╗
echo ║   BẢNG MÔ TẢ VỊ TRÍ VIỆC LÀM – TTCM & TPCM                ║
echo ║   Trường THPT Phục Hòa – Nghị định 232/2026/NĐ-CP          ║
echo ╠══════════════════════════════════════════════════════════════╣
echo ║   ĐANG KHỞI TẠO VÀ ĐẨY LÊN GITHUB + VERCEL...             ║
echo ╚══════════════════════════════════════════════════════════════╝
echo.

:: Kiểm tra Git đã được cài
git --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [LỖI] Git chưa được cài đặt. Vui lòng cài Git trước.
    pause & exit /b 1
)

:: Khởi tạo git repo nếu chưa có
if not exist ".git" (
    echo [1/6] Khởi tạo Git repository...
    git init
    git branch -M main
    echo      ✓ Đã khởi tạo git repo
) else (
    echo [1/6] Git repository đã tồn tại - bỏ qua khởi tạo
)

:: Thêm toàn bộ file
echo [2/6] Thêm tất cả file vào staging...
git add .
echo      ✓ Đã add tất cả file

:: Commit
echo [3/6] Tạo commit...
git commit -m "feat: Bảng VTVL TTCM và TPCM theo NĐ 232/2026 – THPT Phục Hòa"
if %ERRORLEVEL% EQU 0 (
    echo      ✓ Commit thành công
) else (
    echo      [!] Không có thay đổi mới hoặc lỗi commit
)

:: Kiểm tra remote
git remote get-url origin >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [4/6] Tạo repository GitHub và thêm remote...
    
    :: Thử tạo repo bằng gh CLI
    gh repo create %GITHUB_REPO% --public --description "Bảng mô tả VTVL Tổ trưởng & Tổ phó CM – THPT Phục Hòa – NĐ 232/2026" >nul 2>&1
    if %ERRORLEVEL% EQU 0 (
        echo      ✓ Đã tạo repo GitHub: %GITHUB_USER%/%GITHUB_REPO%
    ) else (
        echo      [!] gh CLI không khả dụng hoặc repo đã tồn tại
    )
    
    :: Thêm remote
    git remote add origin https://github.com/%GITHUB_USER%/%GITHUB_REPO%.git
    echo      ✓ Đã thêm remote origin
) else (
    echo [4/6] Remote origin đã tồn tại - bỏ qua
    git remote -v
)

:: Push lên GitHub
echo [5/6] Đang push lên GitHub...
git push -u origin main
if %ERRORLEVEL% EQU 0 (
    echo      ✓ ĐÃ ĐẨY LÊN GITHUB THÀNH CÔNG!
    echo      🔗 https://github.com/%GITHUB_USER%/%GITHUB_REPO%
) else (
    echo.
    echo [!] Lỗi push GitHub. Thử force push...
    git push -u origin main --force
    if %ERRORLEVEL% EQU 0 (
        echo      ✓ Force push thành công!
    ) else (
        echo      ✗ Lỗi push – Kiểm tra kết nối Internet và xác thực GitHub
    )
)

:: Deploy Vercel
echo [6/6] Đang deploy lên Vercel...
vercel --version >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    vercel --prod --yes --name %GITHUB_REPO%
    if %ERRORLEVEL% EQU 0 (
        echo      ✓ ĐÃ DEPLOY LÊN VERCEL THÀNH CÔNG!
    ) else (
        echo      [!] Lỗi Vercel deploy
    )
) else (
    echo      [!] Vercel CLI chưa được cài. Cài bằng: npm install -g vercel
    echo      [!] Sau đó chạy: vercel --prod --yes trong thư mục này
    echo      [💡] Hoặc kết nối GitHub với Vercel dashboard để tự động deploy
    echo          https://vercel.com/new
)

echo.
echo ╔══════════════════════════════════════════════════════════════╗
echo ║   HOÀN THÀNH! THÔNG TIN DỰ ÁN:                             ║
echo ╠══════════════════════════════════════════════════════════════╣
echo ║   📁 Thư mục : %PROJECT_DIR%
echo ║   🐙 GitHub  : https://github.com/%GITHUB_USER%/%GITHUB_REPO%
echo ║   🌐 Vercel  : https://%GITHUB_REPO%.vercel.app            ║
echo ║   📋 Tài liệu: VTVL TTCM & TPCM – NĐ 232/2026             ║
echo ╚══════════════════════════════════════════════════════════════╝
echo.
pause
