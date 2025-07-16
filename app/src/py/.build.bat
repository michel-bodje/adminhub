@echo off
REM Activate virtual environment
call ..\..\..\venv\Scripts\activate

REM Build new_matter.exe
pyinstaller --onefile  --distpath ..\..\bin new_matter.py

REM Build bill_matter.exe
pyinstaller --onefile  --distpath ..\..\bin --add-data "tesseract;tesseract" bill_matter.py

REM Build close_matter.exe
pyinstaller --onefile  --distpath ..\..\bin --add-data "tesseract;tesseract" close_matter.py

echo Build complete!
pause