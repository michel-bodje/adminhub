@echo off
cd /d "%~dp0"

pyinstaller ^
  --onefile ^
  --noconsole ^
  --name AdminHub ^
  --distpath dist ^
  --add-data "src;app/src" ^
  --add-data "web;app/web" ^
  --add-data "templates;templates" ^
  main.py

pause