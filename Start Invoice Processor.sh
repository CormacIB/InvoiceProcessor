#!/bin/bash
cd "$(dirname "$0")"
python3 invoice_processor.py
if [ $? -ne 0 ]; then
    echo ""
    echo "ERROR: Could not start the Invoice Processor."
    echo "Make sure Python is installed and run:"
    echo "  pip install pdfplumber pypdf reportlab"
    echo ""
    read -p "Press Enter to continue..."
fi
