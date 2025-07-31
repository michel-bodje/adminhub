@echo off
cd /d "%~dp0"

pyinstaller ^
  --onefile ^
  --noconsole ^
  --name AdminHub ^
  --distpath dist ^
  --add-data "src;src" ^
  --add-data "web;web" ^
  --add-data "templates;templates" ^
  app/src/main.py

pause