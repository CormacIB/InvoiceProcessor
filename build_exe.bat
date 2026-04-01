@echo off
echo ============================================
echo  Coffee Lab Invoice Processor - Build EXE
echo ============================================
echo.

REM Install dependencies
echo Installing Python dependencies...
pip install pypdf pdfplumber reportlab pyinstaller
echo.

REM Build the EXE (single file, no console window)
echo Building EXE...
pyinstaller ^
  --onefile ^
  --windowed ^
  --name "CoffeeLabInvoiceProcessor" ^
  --add-data "config;config" ^
  invoice_processor.py

echo.
echo ============================================
echo  Done! EXE is in the dist\ folder.
echo  Copy dist\CoffeeLabInvoiceProcessor.exe
echo  into this folder and run it.
echo ============================================
pause
