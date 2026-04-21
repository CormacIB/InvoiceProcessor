#!/bin/bash
echo "============================================"
echo " Coffee Lab Invoice Processor - Build App"
echo "============================================"
echo ""

echo "Installing Python dependencies..."
pip install pypdf pdfplumber reportlab pyinstaller
echo ""

echo "Building app bundle..."
pyinstaller \
  --onefile \
  --windowed \
  --name "CoffeeLabInvoiceProcessor" \
  --hidden-import=pdfminer.pdfdocument \
  --hidden-import=pdfminer.pdfpage \
  --hidden-import=pdfminer.pdfinterp \
  --hidden-import=pdfminer.converter \
  --hidden-import=pdfminer.layout \
  --hidden-import=pdfminer.high_level \
  --hidden-import=pdfminer.utils \
  --hidden-import=pdfminer.image \
  --hidden-import=charset_normalizer \
  invoice_processor.py

echo ""
echo "============================================"
echo " Done! App bundle is in the dist/ folder."
echo " Copy dist/CoffeeLabInvoiceProcessor.app"
echo " into this folder and run it."
echo "============================================"
