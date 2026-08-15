@echo off
REM Builds the Document Issue Sheet script into a standalone .exe using PyInstaller.
REM Run this on the Windows PC where you want the exe to end up.

where python >nul 2>nul
if errorlevel 1 (
    echo Python was not found on PATH. Install Python from python.org first,
    echo and make sure "Add python.exe to PATH" is checked during install.
    pause
    exit /b 1
)

echo Installing PyInstaller and required libraries (skips if already installed)...
python -m pip install --upgrade pyinstaller openpyxl

echo.
echo Building exe...
REM Ensure your Python file is named document_issue_sheet.py, or change the name below to match your file.
python -m PyInstaller --onefile --console --name "DocumentIssueSheet" process_pdfs.py

echo.
echo Done. Your exe is in the "dist" folder as DocumentIssueSheet.exe
pause
