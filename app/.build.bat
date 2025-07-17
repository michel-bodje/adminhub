@echo off
cd /d "%~dp0"

pyinstaller ^
  --onefile ^
  --noconsole ^
  --name AdminHub ^
  --distpath dist ^
  --add-data "web;app/web" ^
  --add-data "bin;app/bin" ^
  --add-data "src;app/src" ^
  --add-data "templates;templates" ^
  main.py

pause