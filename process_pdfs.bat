@echo off
echo =========================================
echo Document Issue Sheet
echo =========================================
echo.

echo Checking for required Python libraries...
python -m pip install --quiet openpyxl
echo.

echo Adding PDFs to Document Issue Sheet...
python process_pdfs.py
echo.

pause
