"""Prepare a project's Document Issue Sheet workbook and add today's issue column.

Run it from the folder holding the "...SH-S-0001-Document Issue Sheet.xlsx"
template. It creates "<project number>-SH-S-0001-Document Issue Sheet.xlsx"
from the template (if it doesn't exist yet), writes the project name, and
adds a dated issue column. Running it again on the same day reuses that
day's column instead of adding a duplicate.
"""
import argparse
import datetime
import re
import shutil
import sys
from pathlib import Path

try:
    import openpyxl
except ImportError:
    print("Error: openpyxl is not installed. Please run: pip install openpyxl")
    sys.exit(1)

TEMPLATE_SUFFIX = "SH-S-0001-Document Issue Sheet.xlsx"
FIRST_ISSUE_COL = 4  # Column D
DAY_ROW, MONTH_ROW, YEAR_ROW = 3, 4, 5
PROJECT_NAME_CELL = "A3"
SECOND_SHEET_NAME = "Sheet 2"

# Characters Windows doesn't allow in file names
INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*]')


def target_filename(proj_number):
    return f"{proj_number}-{TEMPLATE_SUFFIX}"


def find_template(folder, exclude_name):
    """Return the template to copy, preferring placeholder names like 'xxxxx-...'.

    Files already numbered for a project (e.g. '12345-...') are only used as a
    fallback, and the project's own target file is never used as its template.
    """
    candidates = sorted(
        p for p in folder.glob(f"*{TEMPLATE_SUFFIX}")
        if p.name != exclude_name and not p.name.startswith("~$")  # skip Excel lock files
    )
    if not candidates:
        return None
    placeholders = [p for p in candidates if not p.name[0].isdigit()]
    return (placeholders or candidates)[0]


def _two_digits(value):
    """Normalise a header cell to a 2-digit string so '07', 7 and '7' compare equal."""
    if value is None:
        return None
    text = str(value).strip()
    return text.zfill(2) if text.isdigit() else text


def find_issue_column(sheet, date):
    """Return (column, already_present) for the given issue date.

    Scans right from column D. If a column already holds this date it's
    reused; otherwise the first column with an empty day cell is returned.
    """
    wanted = (date.strftime("%d"), date.strftime("%m"), date.strftime("%y"))
    col = FIRST_ISSUE_COL
    while sheet.cell(row=DAY_ROW, column=col).value is not None:
        existing = tuple(
            _two_digits(sheet.cell(row=row, column=col).value)
            for row in (DAY_ROW, MONTH_ROW, YEAR_ROW)
        )
        if existing == wanted:
            return col, True
        col += 1
    return col, False


def update_workbook(path, proj_name, date):
    """Write the project name and issue date into the workbook. Returns the column used."""
    wb = openpyxl.load_workbook(path)
    sheet = wb.active

    col, already_present = find_issue_column(sheet, date)
    sheet.cell(row=DAY_ROW, column=col, value=date.strftime("%d"))
    sheet.cell(row=MONTH_ROW, column=col, value=date.strftime("%m"))
    sheet.cell(row=YEAR_ROW, column=col, value=date.strftime("%y"))
    sheet[PROJECT_NAME_CELL] = proj_name

    # Make sure the continuation sheet exists for longer drawing lists
    if SECOND_SHEET_NAME not in wb.sheetnames:
        wb.copy_worksheet(sheet).title = SECOND_SHEET_NAME

    wb.save(path)
    return col, already_present


def prompt(message, default):
    answer = input(message).strip()
    return answer or default


def parse_args(argv):
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--number", help="Revit project number, e.g. 12345")
    parser.add_argument("--name", help="Project name")
    parser.add_argument("--folder", type=Path, default=Path.cwd(),
                        help="Folder containing the template (default: current folder)")
    return parser.parse_args(argv)


def run(argv=None):
    args = parse_args(argv)
    folder = args.folder.resolve()

    proj_number = args.number or prompt("Enter Revit Project Number (e.g., 12345): ", "00000")
    proj_number = INVALID_FILENAME_CHARS.sub("", proj_number).strip() or "00000"
    proj_name = args.name or prompt("Enter Project Name: ", "Unnamed Project")

    name = target_filename(proj_number)
    target_path = folder / name

    if not target_path.exists():
        template_path = find_template(folder, exclude_name=name)
        if template_path is None:
            print(f"Error: No template matching '*{TEMPLATE_SUFFIX}' found in {folder}")
            return 1
        shutil.copy(template_path, target_path)
        print(f"Created new issue sheet: {name} from template: {template_path.name}")

    try:
        col, already_present = update_workbook(target_path, proj_name, datetime.date.today())
    except PermissionError:
        print(f"Error: Could not save '{name}'. Please make sure it is CLOSED in Excel.")
        return 1
    except Exception as e:
        print(f"Error updating '{name}': {e}")
        return 1

    letter = openpyxl.utils.get_column_letter(col)
    if already_present:
        print(f"Today's issue column already existed (column {letter}); project name updated.")
    else:
        print(f"Added today's issue date in column {letter}.")
    print(f"Success! Saved: {target_path}")
    return 0


def main():
    code = run()
    # When run as a double-clicked .exe the console closes immediately, so wait first
    if getattr(sys, "frozen", False):
        input("\nPress Enter to close...")
    sys.exit(code)


if __name__ == "__main__":
    main()
