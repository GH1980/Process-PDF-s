# Process-PDF-s

Script to prepare a Document Issue Sheet workbook for a project: it copies the
"...SH-S-0001-Document Issue Sheet.xlsx" template (if one isn't already
present), writes the project name into `A3`, and adds today's date as a new
issue column (day / month / 2-digit year in rows 3–5, starting from column D).
Running it again on the same day reuses that day's column rather than adding a
duplicate.

## Usage

1. Put `process_pdfs.py` in the same folder as the project's Excel template
   (a file ending in `SH-S-0001-Document Issue Sheet.xlsx`).
2. Run `process_pdfs.bat` (installs `openpyxl` if needed, then runs the
   script) and answer the prompts for project number and name.

You can also skip the prompts by passing the values on the command line:

```
python process_pdfs.py --number 12345 --name "My Project" [--folder PATH]
```

If more than one template is in the folder, a placeholder one (e.g.
`xxxxx-SH-S-0001-...`) is preferred over one already numbered for a project.

## Building a standalone .exe

Run `build_exe.bat` on Windows to package `process_pdfs.py` into a
standalone `DocumentIssueSheet.exe` with PyInstaller. The exe is written to
the `dist` folder.

A prebuilt `DocumentIssueSheet.exe` is included in this repo, so Windows
users who don't have Python installed can just download it and double-click
to run it, without needing to install Python or run `build_exe.bat`. Place
it in the same folder as the project's Excel template, same as `process_pdfs.py`.
Rebuild it with `build_exe.bat` after changing `process_pdfs.py` so it stays in sync.

## Requirements

- Python 3 with `openpyxl` (installed automatically by `process_pdfs.bat`)

## Tests

```
pip install openpyxl pytest
python -m pytest tests
```
