@echo off
rem Mo giao dien chuyen doi JSON sang Excel.
rem Dung pythonw.exe de khong hien cua so console den.

cd /d "%~dp0"

where pythonw.exe >nul 2>nul
if errorlevel 1 goto fallback

start "" pythonw.exe "json_to_excel_gui.py" %*
exit /b 0

:fallback
echo Khong tim thay pythonw.exe, chay bang python.exe...
python "json_to_excel_gui.py" %*
pause
