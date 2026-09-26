import datetime

import openpyxl
import pytest

import process_pdfs as pp

TODAY = datetime.date.today()


def make_template(folder, name="xxxxx-SH-S-0001-Document Issue Sheet.xlsx", dates=()):
    wb = openpyxl.Workbook()
    ws = wb.active
    for offset, (d, m, y) in enumerate(dates):
        col = pp.FIRST_ISSUE_COL + offset
        ws.cell(row=3, column=col, value=d)
        ws.cell(row=4, column=col, value=m)
        ws.cell(row=5, column=col, value=y)
    wb.save(folder / name)
    return folder / name


def run(folder, number="12345", name="Test Project"):
    return pp.run(["--folder", str(folder), "--number", number, "--name", name])


def load(folder, number="12345"):
    return openpyxl.load_workbook(folder / pp.target_filename(number))


def today_cells(ws, col):
    return tuple(ws.cell(row=r, column=col).value for r in (3, 4, 5))


def test_creates_sheet_from_template(tmp_path):
    make_template(tmp_path)
    assert run(tmp_path) == 0
    wb = load(tmp_path)
    ws = wb.active
    assert ws["A3"].value == "Test Project"
    assert today_cells(ws, 4) == (TODAY.strftime("%d"), TODAY.strftime("%m"), TODAY.strftime("%y"))
    assert "Sheet 2" in wb.sheetnames


def test_appends_after_existing_dates(tmp_path):
    make_template(tmp_path, dates=[("01", "01", "20"), ("02", "02", "20")])
    run(tmp_path)
    ws = load(tmp_path).active
    assert today_cells(ws, 6)[0] == TODAY.strftime("%d")


def test_running_twice_same_day_reuses_column(tmp_path):
    make_template(tmp_path)
    run(tmp_path)
    run(tmp_path, name="Renamed")
    ws = load(tmp_path).active
    assert ws.cell(row=3, column=5).value is None
    assert ws["A3"].value == "Renamed"


def test_matches_integer_date_cells(tmp_path):
    make_template(tmp_path, dates=[(TODAY.day, TODAY.month, TODAY.year % 100)])
    run(tmp_path)
    assert load(tmp_path).active.cell(row=3, column=5).value is None


def test_prefers_placeholder_template(tmp_path):
    make_template(tmp_path, name="99999-SH-S-0001-Document Issue Sheet.xlsx", dates=[("01", "01", "20")])
    make_template(tmp_path)
    run(tmp_path)
    # Copied from the blank placeholder, so today's date lands in column D
    assert load(tmp_path).active.cell(row=3, column=4).value == TODAY.strftime("%d")


def test_missing_template_fails(tmp_path, capsys):
    assert run(tmp_path) == 1
    assert "No template" in capsys.readouterr().out


def test_strips_invalid_filename_chars(tmp_path):
    make_template(tmp_path)
    assert run(tmp_path, number="12/34?5") == 0
    assert (tmp_path / pp.target_filename("12345")).exists()


def test_prompts_when_no_args(tmp_path, monkeypatch):
    make_template(tmp_path)
    answers = iter(["", ""])
    monkeypatch.setattr("builtins.input", lambda _: next(answers))
    assert pp.run(["--folder", str(tmp_path)]) == 0
    assert load(tmp_path, "00000").active["A3"].value == "Unnamed Project"
