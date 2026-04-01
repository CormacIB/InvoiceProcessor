@echo off
cd /d "%~dp0"
python invoice_processor.py
if %errorlevel% neq 0 (
    echo.
    echo ERROR: Could not start the Invoice Processor.
    echo Make sure Python is installed and run:
    echo   pip install pdfplumber pypdf reportlab
    echo.
    pause
)
