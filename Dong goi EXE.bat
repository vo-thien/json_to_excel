@echo off
rem Dong goi tool thanh MOT file .exe duy nhat, chay duoc tren may khong co Python.
rem Chay lai script nay moi khi sua json_to_excel_gui.py hoac json_excel_core.py.

cd /d "%~dp0"

echo Dang dong goi... (mat khoang 20-40 giay)
echo.

python -m PyInstaller --onefile --windowed --clean --noconfirm --noupx ^
  --name "JSON sang Excel" ^
  --icon "icon.ico" ^
  --exclude-module numpy --exclude-module pandas --exclude-module matplotlib ^
  --exclude-module scipy --exclude-module PIL --exclude-module pytest ^
  "json_to_excel_gui.py"

if errorlevel 1 goto err

echo.
echo ============================================
echo  XONG! File exe nam trong thu muc dist\
echo ============================================
pause
exit /b 0

:err
echo.
echo ============================================
echo  DONG GOI THAT BAI - xem loi o phia tren
echo ============================================
pause
exit /b 1
