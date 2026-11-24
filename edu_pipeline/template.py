"""Build the multi-year Excel data-entry template.

The template mirrors how statistical survey templates work: one sheet per table,
labels in columns A-C, one column per year, input cells with built-in data
validation, Total rows and an Indicators sheet that are Excel formulas, and a
Metadata sheet for sources and footnotes.
"""

from __future__ import annotations

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from . import config as C

NAVY = "303658"
FILL_HEADER = PatternFill("solid", fgColor=NAVY)
FILL_LABEL = PatternFill("solid", fgColor="E7ECF2")
FILL_INPUT = PatternFill("solid", fgColor="FFF9E5")
FILL_FORMULA = PatternFill("solid", fgColor="EDEDED")
THIN = Side(style="thin", color="B7BFCB")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

SHEET_TITLES = {
    "Info": "Country information",
    "Population": "Population by single age and sex",
    "Enrolment": "Enrolment by level, sex and single age",
    "Teachers": "Teachers by level and sex",
    "Indicators": "Calculated indicators (do not edit)",
    "Metadata": "Metadata and data sources",
}
INPUT_NOTE = "Enter whole numbers. Codes: m = missing, a = not applicable, x = included elsewhere."
FORMULA_NOTE = "Calculated automatically from the Enrolment and Population sheets."
GER_LABEL = "Gross enrolment ratio (%)"


# ---------------------------------------------------------------- layout ----
def sheet_rows(sheet: str) -> list[tuple]:
    """Row keys (level, sex, age) for a sheet, in the order they appear."""
    if sheet == "Population":
        return [("Population", s, a) for s in C.SEXES for a in C.AGES]
    if sheet == "Enrolment":
        rows = []
        for lvl in C.LEVELS:
            for s in C.SEXES:
                rows += [(lvl, s, a) for a in C.AGES]
                rows.append((lvl, s, C.TOTAL_LABEL))
        return rows
    if sheet == "Teachers":
        return [(lvl, s, "All ages") for lvl in C.LEVELS for s in C.SEXES]
    if sheet == "Indicators":
        return [(lvl, s, GER_LABEL) for lvl in C.LEVELS for s in C.SEXES + [C.BOTH_SEXES]]
    raise ValueError(sheet)


def year_columns(ws) -> dict[int, int]:
    """Map year -> column index by reading the header row."""
    out = {}
    for col in range(C.FIRST_YEAR_COL, ws.max_column + 1):
        v = ws.cell(C.HEADER_ROW, col).value
        try:
            out[int(v)] = col
        except (TypeError, ValueError):
            continue
    return out


def row_of(sheet: str, key: tuple) -> int:
    return C.FIRST_DATA_ROW + sheet_rows(sheet).index(key)


def _total_formula(col_letter: str, sheet: str, key: tuple) -> str:
    first = row_of(sheet, (key[0], key[1], C.AGES[0]))
    last = row_of(sheet, (key[0], key[1], C.AGES[-1]))
    return f"=SUM({col_letter}{first}:{col_letter}{last})"


def _ger_formula(col_letter: str, r: int, level: str, sex: str) -> str:
    lo, hi = C.LEVELS[level]["official_ages"]
    if sex == C.BOTH_SEXES:
        num = f'SUMIFS(Enrolment!{col_letter}:{col_letter},Enrolment!$A:$A,$A{r},Enrolment!$C:$C,">=0")'
        den = (f'SUMIFS(Population!{col_letter}:{col_letter},'
               f'Population!$C:$C,">={lo}",Population!$C:$C,"<={hi}")')
    else:
        num = (f'SUMIFS(Enrolment!{col_letter}:{col_letter},Enrolment!$A:$A,$A{r},'
               f'Enrolment!$B:$B,$B{r},Enrolment!$C:$C,">=0")')
        den = (f'SUMIFS(Population!{col_letter}:{col_letter},Population!$B:$B,$B{r},'
               f'Population!$C:$C,">={lo}",Population!$C:$C,"<={hi}")')
    return f'=IFERROR(ROUND(100*{num}/{den},1),"")'


def input_validation(first_cell: str) -> DataValidation:
    """Whole number >= 0, or one of the allowed missing-value codes."""
    codes = ",".join(f'{first_cell}="{c}"' for c in C.MISSING_CODES)
    dv = DataValidation(
        type="custom",
        formula1=f"OR(AND(ISNUMBER({first_cell}),{first_cell}>=0,{first_cell}=INT({first_cell})),{codes})",
        allow_blank=True,
        showErrorMessage=True,
        errorTitle="Invalid entry",
        error="Enter a whole number of 0 or more, or a code: m, a or x.",
    )
    return dv


def style_year_column(ws, sheet: str, col: int) -> None:
    """Write header, formulas and styles for one year column of a sheet."""
    letter = get_column_letter(col)
    h = ws.cell(C.HEADER_ROW, col)
    h.font, h.fill, h.alignment, h.border = Font(bold=True, color="FFFFFF"), FILL_HEADER, Alignment(horizontal="center"), BORDER
    for key in sheet_rows(sheet):
        r = row_of(sheet, key)
        cell = ws.cell(r, col)
        cell.border = BORDER
        cell.number_format = "#,##0" if sheet != "Indicators" else "0.0"
        if sheet == "Enrolment" and key[2] == C.TOTAL_LABEL:
            cell.value = _total_formula(letter, sheet, key)
            cell.fill, cell.font = FILL_FORMULA, Font(bold=True)
        elif sheet == "Indicators":
            cell.value = _ger_formula(letter, r, key[0], key[1])
            cell.fill = FILL_FORMULA
        else:
            cell.fill = FILL_INPUT
    ws.column_dimensions[letter].width = 11


# ----------------------------------------------------------------- build ----
def build_template(years: list[int], info: dict | None = None, values: dict | None = None) -> Workbook:
    """Create a template workbook.

    values: optional {(sheet, level, sex, age, year): number_or_code} to pre-fill.
    """
    info = info or {}
    values = values or {}
    wb = Workbook()
    wb.remove(wb.active)

    for sheet in C.ALL_SHEETS:
        ws = wb.create_sheet(sheet)
        ws["A1"] = SHEET_TITLES[sheet]
        ws["A1"].font = Font(bold=True, size=13, color=NAVY)
        ws.freeze_panes = ws.cell(C.FIRST_DATA_ROW, C.FIRST_YEAR_COL)

        if sheet == "Info":
            _build_info(ws, info)
            continue
        if sheet == "Metadata":
            _build_metadata(ws)
            continue

        ws["A2"] = FORMULA_NOTE if sheet == "Indicators" else INPUT_NOTE
        ws["A2"].font = Font(italic=True, color="555555")
        third = "Indicator" if sheet == "Indicators" else "Age"
        for i, label in enumerate(["Level", "Sex", third], start=1):
            c = ws.cell(C.HEADER_ROW, i, label)
            c.font, c.fill, c.border = Font(bold=True, color="FFFFFF"), FILL_HEADER, BORDER
        ws.column_dimensions["A"].width = 30
        ws.column_dimensions["B"].width = 12
        ws.column_dimensions["C"].width = 26 if sheet == "Indicators" else 10

        for key in sheet_rows(sheet):
            r = row_of(sheet, key)
            for i, v in enumerate(key, start=1):
                c = ws.cell(r, i, v)
                c.fill, c.border = FILL_LABEL, BORDER
                if key[2] == C.TOTAL_LABEL:
                    c.font = Font(bold=True)

        for j, year in enumerate(years):
            col = C.FIRST_YEAR_COL + j
            ws.cell(C.HEADER_ROW, col, year)
            style_year_column(ws, sheet, col)
            for key in sheet_rows(sheet):
                v = values.get((sheet, *key, year))
                if v is not None:
                    ws.cell(row_of(sheet, key), col).value = v

        if sheet in C.DATA_SHEETS:
            _add_validation(ws, sheet, C.FIRST_YEAR_COL, C.FIRST_YEAR_COL + len(years) - 1)
    return wb


def _add_validation(ws, sheet: str, first_col: int, last_col: int) -> None:
    first_row = C.FIRST_DATA_ROW
    last_row = C.FIRST_DATA_ROW + len(sheet_rows(sheet)) - 1
    tl = f"{get_column_letter(first_col)}{first_row}"
    dv = input_validation(tl)
    dv.add(f"{tl}:{get_column_letter(last_col)}{last_row}")
    ws.add_data_validation(dv)


def _build_info(ws, info: dict) -> None:
    for i, label in enumerate(["Field", "Value"], start=1):
        c = ws.cell(C.HEADER_ROW - 1, i, label)
        c.font, c.fill = Font(bold=True, color="FFFFFF"), FILL_HEADER
    for k, field in enumerate(C.INFO_FIELDS):
        ws.cell(C.HEADER_ROW + k, 1, field).fill = FILL_LABEL
        v = ws.cell(C.HEADER_ROW + k, 2, info.get(field))
        v.fill = FILL_INPUT
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 34
    ws.freeze_panes = None


def _build_metadata(ws) -> None:
    for i, label in enumerate(["Sheet", "Year", "Source", "Notes"], start=1):
        c = ws.cell(C.HEADER_ROW, i, label)
        c.font, c.fill = Font(bold=True, color="FFFFFF"), FILL_HEADER
    for r in range(C.FIRST_DATA_ROW, C.FIRST_DATA_ROW + 15):
        for col in range(1, 5):
            ws.cell(r, col).fill = FILL_INPUT
            ws.cell(r, col).border = BORDER
    for letter, w in zip("ABCD", (14, 8, 40, 60)):
        ws.column_dimensions[letter].width = w
    ws.freeze_panes = None
