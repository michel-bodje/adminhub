@echo off
echo.
echo === Setting up Outlook Macros ===
echo.

start outlook.exe
timeout /t 5 >nul
python install_macros.py

pause
