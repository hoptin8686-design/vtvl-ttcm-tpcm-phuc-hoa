@echo off
chcp 65001 >nul
echo.
echo ╔══════════════════════════════════════════════════════════════╗
echo ║   HƯỚNG DẪN DEPLOY LÊN VERCEL – THPT PHỤ HÒA              ║
echo ║   Dự án: VTVL TTCM & TPCM – NĐ 232/2026                   ║
echo ╚══════════════════════════════════════════════════════════════╝
echo.
echo Dự án đã được đẩy lên GitHub:
echo   🔗 https://github.com/hoptin8686-design/vtvl-ttcm-tpcm-phuc-hoa
echo.
echo ═══════════════════════════════════════════════════════════════
echo CÁCH 1: DEPLOY QUA VERCEL DASHBOARD (DỄ NHẤT - Khuyến nghị)
echo ═══════════════════════════════════════════════════════════════
echo.
echo   Bước 1: Mở trình duyệt, vào https://vercel.com/new
echo   Bước 2: Nhấn "Import Git Repository"
echo   Bước 3: Chọn repo: vtvl-ttcm-tpcm-phuc-hoa
echo   Bước 4: Để mặc định tất cả cài đặt (static site)
echo   Bước 5: Nhấn "Deploy"
echo   Bước 6: Chờ 1-2 phút → URL tự động tạo
echo.
echo ═══════════════════════════════════════════════════════════════
echo CÁCH 2: DEPLOY QUA VERCEL CLI
echo ═══════════════════════════════════════════════════════════════
echo.
echo   Bước 1: Cài Vercel CLI:   npm install -g vercel
echo   Bước 2: Đăng nhập:        vercel login
echo   Bước 3: Deploy:           vercel --prod --yes
echo.
echo ═══════════════════════════════════════════════════════════════
echo URL SAU KHI DEPLOY
echo ═══════════════════════════════════════════════════════════════
echo.
echo   Tên dự án: vtvl-ttcm-tpcm-phuc-hoa
echo   URL dự kiến: https://vtvl-ttcm-tpcm-phuc-hoa.vercel.app
echo.
echo ═══════════════════════════════════════════════════════════════
echo ĐẦY LÊN GITHUB LẦN NỮA (nếu có thay đổi)
echo ═══════════════════════════════════════════════════════════════
echo.
set "PATH=C:\Users\DMX HOA THUAN\AppData\Local\Programs\Git\cmd;C:\Users\DMX HOA THUAN\AppData\Local\Programs\gh\bin;%PATH%"
cd /d "d:\Du-an-web\web-vtvl-ttcm-tpcm-phuc-hoa"
git add .
git commit -m "update: Cap nhat noi dung VTVL TTCM TPCM"
git push origin main
echo.
echo   ✓ Đã push lên GitHub!
echo   🔗 https://github.com/hoptin8686-design/vtvl-ttcm-tpcm-phuc-hoa
echo.
pause
