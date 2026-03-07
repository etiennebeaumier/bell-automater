#!/bin/bash
# Build standalone macOS app bundle
set -e

echo "Building BCECN Pricing Tool for macOS..."

python3 -m PyInstaller --onefile --windowed \
  --name "BCECN Pricing Tool" \
  --collect-data customtkinter \
  --hidden-import pdfminer.high_level \
  --add-data "parsers:parsers" \
  app.py

echo ""
echo "Build complete! App bundle is at:"
echo "  dist/BCECN Pricing Tool.app"
echo ""
echo "To share: zip the .app and send it to colleagues."
