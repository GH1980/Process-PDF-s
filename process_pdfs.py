import os
import shutil
import re
import datetime
from pathlib import Path

try:
    import openpyxl
except ImportError:
    print("Error: openpyxl is not installed. Please run: pip install openpyxl")
    exit(1)

def main():
    # Use current working directory
    target_dir = Path(os.getcwd())

    proj_number = input("Enter Revit Project Number (e.g., 12345): ").strip() or "00000"
    proj_name = input("Enter Project Name: ").strip() or "Unnamed Project"

    # Dynamically look for the template file matching the document issue sheet pattern
    template_pattern = "*SH-S-0001-Document Issue Sheet.xlsx"
    template_files = list(target_dir.glob(template_pattern))

    if template_files:
        template_path = template_files[0]
        template_filename = template_path.name
    else:
        template_filename = "xxxxx-SH-S-0001-Document Issue Sheet.xlsx"
        template_path = target_dir / template_filename

    target_filename = f"{proj_number}-SH-S-0001-Document Issue Sheet.xlsx"
    target_path = target_dir / target_filename

    if not target_path.exists():
        if template_path.exists():
            shutil.copy(template_path, target_path)
            print(f"Created new issue sheet: {target_filename} from template: {template_filename}")
        else:
            print(f"Error: Template file matching '{template_pattern}' not found in {target_dir}")
            return

    try:
        wb = openpyxl.load_workbook(target_path)
    except Exception as e:
        print(f"Error opening Excel file (make sure it is closed): {e}")
        return

    sheet1 = wb.active

    # Find next blank column starting from D (Column 4) for the new issue date
    target_col = 4
    while sheet1.cell(row=3, column=target_col).value is not None:
        target_col += 1

    # Update date header cells (Rows 3, 4, 5)
    today = datetime.date.today()
    sheet1.cell(row=3, column=target_col, value=today.strftime("%d"))
    sheet1.cell(row=4, column=target_col, value=today.strftime("%m"))
    # Changed "%Y" to "%y" to output a 2-digit year (YY)
    sheet1.cell(row=5, column=target_col, value=today.strftime("%y"))

    # Update project name
    sheet1["A3"] = proj_name

    # Handle Sheet 2: ensure blank rows and existing drawing number structures are preserved
    if "Sheet 2" in wb.sheetnames:
        sheet2 = wb["Sheet 2"]
    else:
        sheet2 = wb.copy_worksheet(sheet1)
        sheet2.title = "Sheet 2"

    print("Excel template prepared successfully at column index:", target_col)

    try:
        wb.save(target_path)
        print(f"Success! Updated sheet saved to: {target_path}")
    except PermissionError:
        print(f"CRITICAL ERROR: Could not save '{target_filename}'. Please make sure it is CLOSED in Excel.")

if __name__ == "__main__":
    main()
