"""
Convert CategorySIPMaintain.xlsx (sheets: 687, D, Rest) into a long-format
upload sheet: UploadType / PartNumber / PlanDate / BatchesPerDay.

Scope: one week, starting from the date in the Summary sheet's D1 cell
(the most recent Saturday) through the following Friday.

Usage:
    python3 build_long_format.py CategorySIPMaintain.7.23.xlsx CategorySIPMaintain.7.23.long.xlsx
"""
import sys
import re
import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter


def get_start_date(wbd):
    """Week start date is read from the Summary sheet's D1 cell."""
    v = wbd['Summary']['D1'].value
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    raise ValueError(f"Summary!D1 is not a date: {v!r}")


def find_target_cols(ws, target_dates):
    """Locate the columns in row 1 whose date falls in target_dates."""
    cols = []
    for col in range(1, ws.max_column + 1):
        v = ws.cell(row=1, column=col).value
        if isinstance(v, datetime.datetime) and v.date() in target_dates:
            cols.append(col)
    cols.sort()
    return cols


def extract_qty_block_sheet(wsf, wsd, sheet_name, target_dates, results):
    """
    687 / D sheet layout: a part-number row holds the raw "batches per day"
    value; the row directly below is labeled 'QTY' and holds a formula that
    multiplies the part-number row by a batch size (e.g. '=HF2*42'). Only
    rows matching that pattern are genuine part entries -- this automatically
    skips category headers, Non-687/Total footers, and roll-up/SUM rows.
    """
    cols = find_target_cols(wsf, target_dates)
    assert len(cols) == 7, (sheet_name, cols)
    col_letter = get_column_letter(cols[0])

    for r in range(1, wsf.max_row):
        part = wsf.cell(row=r, column=1).value
        if part is None or part == 'QTY':
            continue
        if wsf.cell(row=r + 1, column=1).value != 'QTY':
            continue
        formula = wsf.cell(row=r + 1, column=cols[0]).value
        if not (isinstance(formula, str) and formula.startswith('=')):
            continue
        if not re.search(rf'(?<![A-Za-z0-9]){col_letter}{r}(?![0-9])', formula):
            continue

        for col, d in zip(cols, target_dates):
            v = wsd.cell(row=r, column=col).value
            if v is not None and v != 0:
                results.append((part, d, v))


def extract_flat_sheet(wsf, wsd, target_dates, results):
    """Rest sheet layout: no separate QTY row, values sit directly on the part-number row."""
    cols = find_target_cols(wsf, target_dates)
    assert len(cols) == 7, cols

    for r in range(2, wsf.max_row + 1):
        part = wsf.cell(row=r, column=1).value
        if part is None:
            continue
        for col, d in zip(cols, target_dates):
            v = wsd.cell(row=r, column=col).value
            if v is not None and v != 0:
                results.append((part, d, v))


def build(src_path, out_path):
    wbf = openpyxl.load_workbook(src_path, data_only=False)
    wbd = openpyxl.load_workbook(src_path, data_only=True)

    start_date = get_start_date(wbd)
    target_dates = [start_date + datetime.timedelta(days=i) for i in range(7)]
    weekend_dates = {target_dates[0], target_dates[1]}

    results = []  # (part_number, date, batches_per_day)
    for sheet_name in ('687', 'D'):
        extract_qty_block_sheet(wbf[sheet_name], wbd[sheet_name], sheet_name, target_dates, results)
    extract_flat_sheet(wbf['Rest'], wbd['Rest'], target_dates, results)

    out_wb = openpyxl.Workbook()
    ws = out_wb.active
    ws.title = 'Sheet1'

    header_fill = PatternFill(start_color='FFD9D9D9', end_color='FFD9D9D9', fill_type='solid')
    weekend_fill = PatternFill(start_color='FFFCE4D6', end_color='FFFCE4D6', fill_type='solid')
    font = Font(name='Arial', size=10)
    bold_font = Font(name='Arial', size=10, bold=True)

    for c, h in enumerate(['UploadType', 'PartNumber', 'PlanDate', 'BatchesPerDay'], start=1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = bold_font
        cell.fill = header_fill

    r = 2
    for part, d, v in results:
        ws.cell(row=r, column=1, value='ADD').font = font
        ws.cell(row=r, column=2, value=part).font = font
        dcell = ws.cell(row=r, column=3, value=datetime.datetime(d.year, d.month, d.day))
        dcell.font = font
        dcell.number_format = 'm/d/yyyy'
        ws.cell(row=r, column=4, value=v).font = font
        if d in weekend_dates:
            for c in range(1, 5):
                ws.cell(row=r, column=c).fill = weekend_fill
        r += 1

    for col, width in {'A': 12, 'B': 26, 'C': 12, 'D': 15}.items():
        ws.column_dimensions[col].width = width
    ws.freeze_panes = 'A2'

    out_wb.save(out_path)
    print(f'Wrote {r - 2} rows to {out_path}')


if __name__ == '__main__':
    src = sys.argv[1] if len(sys.argv) > 1 else 'CategorySIPMaintain.7.23.xlsx'
    out = sys.argv[2] if len(sys.argv) > 2 else 'CategorySIPMaintain.7.23.long.xlsx'
    build(src, out)
