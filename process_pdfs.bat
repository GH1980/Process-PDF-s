@echo off
echo =========================================
echo PDF to Excel Extractor
echo =========================================
echo.

echo Checking for required Python libraries...
pip install openpyxl
echo.

echo Running PDF extraction script...
python process_pdfs.py
echo.

pause
