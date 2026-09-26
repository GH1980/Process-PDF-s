# Process-PDF-s

Records the drawing PDFs in a folder as a new issue on the project's Document
Issue Sheet. It:

1. Copies the "...SH-S-0001-Document Issue Sheet.xlsx" template to
   `<project number>-SH-S-0001-Document Issue Sheet.xlsx` (if not already there).
2. Writes the project name into `A3`.
3. Adds today's date as a new issue column (day / month / 2-digit year in
   rows 3–5, starting from column D). Running it again on the same day reuses
   that day's column rather than adding a duplicate.
4. Reads every PDF's file name and lists the drawings from row 7 down:
   drawing number in column A, title in column B, and the revision in today's
   issue column.

### PDF file names

PDFs must be named `<drawing number>-<revision>-<title>.pdf`, e.g.

| File name | Number | Revision | Title |
|---|---|---|---|
| `12345-DR-S-1001-P01-General Arrangement.pdf` | 12345-DR-S-1001 | P01 | General Arrangement |
| `12345-XX-B1-DR-S-1001-P03-Basement Plan.pdf` | 12345-XX-B1-DR-S-1001 | P03 | Basement Plan |
| `12345-DR-S-1001-C01.pdf` | 12345-DR-S-1001 | C01 | *(none)* |

The revision is the first code like `P01`, `C02` or `T1` that comes straight
after the numeric part of the drawing number, so level codes such as `B1`
aren't mistaken for it. PDFs without a revision are skipped and listed on
screen.

- Drawings already in column A keep their row; new ones are added under the
  last listed drawing.
- A title from the file name is only written where column B is empty, so
  titles you've typed in are kept.
- If one drawing has several PDFs (e.g. P01 and P02), the most recently
  modified file is used.

## Usage

1. Put `process_pdfs.py` in the folder with the drawing PDFs and the project's
   Excel template (a file ending in `SH-S-0001-Document Issue Sheet.xlsx`).
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
