"""
excel_style.py
================
Shared formatting helper used by battlebots_scraper.py, bot_features_scraper.py,
and analyze_correlations.py so every sheet in the final workbook looks like one
consistent, deliberately-designed report instead of a plain data dump.

Call style_workbook(path) as the LAST step after writing/appending sheets to
the .xlsx file. It's safe to call multiple times (e.g. once per script run) --
it removes and rebuilds its own formatting each time rather than stacking it.

What it does, per sheet:
  - Converts the data range into a real Excel Table (banded rows, filter
    buttons on the header, a proper named table Excel/Google Sheets recognize)
  - Bold white header text on a dark fill
  - Arial font throughout (matches the "professional font" convention used
    for spreadsheet deliverables)
  - Percentage number format on any column whose header contains "%"
  - A "x" suffix format on Lift columns (e.g. "1.29x") so it reads as a
    multiplier at a glance
  - Whole-number format on count-like columns (Wins, Losses, Ties, Robot_Count, etc.)
  - A 3-color scale (red -> yellow -> green) on Win%/Lift columns, so
    strong/weak performers are visually obvious without reading every cell
  - Frozen header row + auto-sized columns
"""

import re
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils import get_column_letter

HEADER_FILL = "1F1F1F"     # near-black, BattleBots-poster style
HEADER_FONT_COLOR = "FFFFFF"
BODY_FONT = "Arial"
TABLE_STYLE = "TableStyleMedium2"   # built-in banded-row Excel style

PCT_HEADER_HINTS = ("%", "pct")
LIFT_HEADER_HINTS = ("lift",)
INT_HEADER_HINTS = ("wins", "losses", "ties", "fights", "count", "seasons", "year")


def _sanitize_table_name(sheet_name: str) -> str:
    """Excel table names must be unique, start with a letter, and contain
    only letters/numbers/underscores."""
    name = re.sub(r"[^A-Za-z0-9_]", "_", sheet_name)
    if not name or not name[0].isalpha():
        name = "T_" + name
    return f"tbl_{name}"[:60]


def _remove_existing_tables(ws):
    # openpyxl stores tables in ws.tables (dict-like); clear before re-adding
    for name in list(ws.tables.keys()):
        del ws.tables[name]


def _header_row_values(ws):
    return [cell.value for cell in ws[1]]


def _classify_column(header) -> str:
    if header is None:
        return "text"
    h = str(header).strip().lower()
    if any(hint in h for hint in LIFT_HEADER_HINTS):
        return "lift"
    if any(hint in h for hint in PCT_HEADER_HINTS):
        return "percent"
    if any(hint in h for hint in INT_HEADER_HINTS):
        return "int"
    return "text"


def style_workbook(path: str):
    wb = load_workbook(path)

    for ws in wb.worksheets:
        if ws.max_row < 2 or ws.max_column < 1:
            continue  # nothing but a header, or empty sheet -- skip

        _remove_existing_tables(ws)

        headers = _header_row_values(ws)
        n_rows, n_cols = ws.max_row, ws.max_column
        last_col_letter = get_column_letter(n_cols)
        data_range = f"A1:{last_col_letter}{n_rows}"

        # --- Excel Table (banded rows + filter buttons + a proper name) ---
        table_name = _sanitize_table_name(ws.title)
        table = Table(displayName=table_name, ref=data_range)
        table.tableStyleInfo = TableStyleInfo(
            name=TABLE_STYLE, showRowStripes=True, showFirstColumn=False,
            showLastColumn=False, showColumnStripes=False,
        )
        ws.add_table(table)

        # --- Header styling (table style already bands rows; header needs
        #     its own explicit font/fill so it reads clearly on any theme) ---
        header_font = Font(name=BODY_FONT, bold=True, color=HEADER_FONT_COLOR, size=11)
        header_fill = PatternFill(start_color=HEADER_FILL, end_color=HEADER_FILL, fill_type="solid")
        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")

        # --- Body font + number formats, column by column ---
        for col_idx, header in enumerate(headers, start=1):
            col_letter = get_column_letter(col_idx)
            kind = _classify_column(header)

            for row_idx in range(2, n_rows + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.font = Font(name=BODY_FONT, size=10)
                if kind == "percent":
                    cell.number_format = "0.0%"
                    cell.alignment = Alignment(horizontal="center")
                elif kind == "lift":
                    cell.number_format = '0.00"x"'
                    cell.alignment = Alignment(horizontal="center")
                elif kind == "int":
                    cell.number_format = "0"
                    cell.alignment = Alignment(horizontal="center")

            # --- Color-scale heatmap on the columns that matter most ---
            if kind in ("percent", "lift") and n_rows > 2:
                rng = f"{col_letter}2:{col_letter}{n_rows}"
                rule = ColorScaleRule(
                    start_type="min", start_color="F8696B",   # red
                    mid_type="percentile", mid_value=50, mid_color="FFEB84",  # yellow
                    end_type="max", end_color="63BE7B",       # green
                )
                ws.conditional_formatting.add(rng, rule)

        # --- Freeze header, auto-size columns ---
        ws.freeze_panes = "A2"
        for col_idx in range(1, n_cols + 1):
            col_letter = get_column_letter(col_idx)
            length = max(
                (len(str(ws.cell(row=r, column=col_idx).value))
                 if ws.cell(row=r, column=col_idx).value is not None else 0)
                for r in range(1, n_rows + 1)
            )
            ws.column_dimensions[col_letter].width = min(max(length + 2, 10), 45)

        ws.sheet_view.showGridLines = False

    wb.save(path)