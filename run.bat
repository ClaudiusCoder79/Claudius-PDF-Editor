@echo off
cd /d "%~dp0"
where python >nul 2>nul
if errorlevel 1 (
  echo Python is not installed. Download it from https://www.python.org/downloads/
  echo During setup, tick "Add python.exe to PATH".
  pause
  exit /b 1
)
python -c "import pymupdf, PIL, docx, openpyxl, pptx, tkinterdnd2" >nul 2>nul || python -c "import fitz, PIL, docx, openpyxl, pptx, tkinterdnd2" >nul 2>nul
if errorlevel 1 (
  echo Installing required packages...
  python -m pip install -r requirements.txt
)
python claudius.py %*
if errorlevel 1 pause
