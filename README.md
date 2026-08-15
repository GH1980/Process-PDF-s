# Process-PDF-s

Script to prepare a Document Issue Sheet workbook for a project: it copies the
"...SH-S-0001-Document Issue Sheet.xlsx" template (if one isn't already
present) and adds a new dated issue column.

## Usage

1. Put `process_pdfs.py` in the same folder as the project's Excel template
   (a file ending in `SH-S-0001-Document Issue Sheet.xlsx`).
2. Run `process_pdfs.bat` (installs `openpyxl` if needed, then runs the
   script) and answer the prompts for project number and name.

## Building a standalone .exe

Run `build_exe.bat` on Windows to package `process_pdfs.py` into a
standalone `DocumentIssueSheet.exe` with PyInstaller. The exe is written to
the `dist` folder.

A prebuilt `DocumentIssueSheet.exe` is included in this repo, so Windows
users who don't have Python installed can just download it and double-click
to run it, without needing to install Python or run `build_exe.bat`. Place
it in the same folder as the project's Excel template, same as `process_pdfs.py`.

## Requirements

- Python 3 with `openpyxl` (installed automatically by `process_pdfs.bat`)
