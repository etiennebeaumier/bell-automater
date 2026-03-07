@echo off
REM Build standalone Windows executable
echo Building BCECN Pricing Tool for Windows...

python -m PyInstaller --onefile --windowed ^
  --name "BCECN Pricing Tool" ^
  --collect-data customtkinter ^
  --hidden-import pdfminer.high_level ^
  --add-data "parsers;parsers" ^
  app.py

echo.
echo Build complete! Executable is at:
echo   dist\BCECN Pricing Tool.exe
echo.
echo Share the .exe file with colleagues.
