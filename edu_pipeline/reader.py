"""Read a filled template (in any language) into a tidy long table.

Rows are found from the labels in columns A-C and years from the header row, not
from fixed cell addresses, so a template where a country inserted a row or a year
can still be read.
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

from . import config as C
from .template import year_columns
from .translate import untranslate


class TemplateError(Exception):
    """The workbook does not have the expected structure."""


def parse_value(raw):
    """Return (value, status) for one cell."""
    if raw is None or (isinstance(raw, str) and raw.strip() == ""):
        return None, "blank"
    if isinstance(raw, bool):
        return None, "non_numeric"
    if isinstance(raw, (int, float)):
        return float(raw), "value"
    s = str(raw).strip()
    if s.startswith("="):
        return None, "formula"
    if s.lower() in C.MISSING_CODES:
        return None, C.MISSING_CODES[s.lower()]
    return None, "non_numeric"


def read_info(wb) -> dict:
    ws = wb["Info"]
    info = {}
    for r in range(C.HEADER_ROW, C.HEADER_ROW + 20):
        key = untranslate(ws.cell(r, 1).value)
        if key in C.INFO_FIELDS:
            info[key] = ws.cell(r, 2).value
    return info


def read_template(path: Path) -> tuple[dict, pd.DataFrame]:
    wb = load_workbook(path)  # keep formulas so overwritten totals can be detected
    missing = [s for s in C.ALL_SHEETS if s not in wb.sheetnames]
    if missing:
        raise TemplateError(f"{path.name}: missing sheets {missing}")
    info = read_info(wb)
    info["file"] = Path(path).name
    iso = info.get("ISO code") or Path(path).stem

    records = []
    for sheet in C.DATA_SHEETS:
        ws = wb[sheet]
        cols = year_columns(ws)
        if not cols:
            raise TemplateError(f"{path.name}: no year columns found on {sheet}")
        r = C.FIRST_DATA_ROW
        while ws.cell(r, 1).value not in (None, ""):
            level, sex, age = (untranslate(ws.cell(r, c).value) for c in (1, 2, 3))
            for year, col in cols.items():
                raw = ws.cell(r, col).value
                value, status = parse_value(raw)
                records.append({
                    "iso": iso, "sheet": sheet, "level": level, "sex": sex,
                    "age": str(age), "year": year, "value": value, "status": status,
                    "raw": None if raw is None else str(raw), "cell": ws.cell(r, col).coordinate,
                })
            r += 1
    return info, pd.DataFrame.from_records(records)


def read_folder(folder: Path) -> tuple[pd.DataFrame, pd.DataFrame, list[str]]:
    """Read every .xlsx in a folder. Returns (info table, data table, structural errors)."""
    infos, frames, errors = [], [], []
    for path in sorted(Path(folder).glob("*.xlsx")):
        if path.name.startswith("~$"):
            continue  # Excel lock file
        try:
            info, df = read_template(path)
        except TemplateError as e:
            errors.append(str(e))
            continue
        infos.append(info)
        frames.append(df)
    data = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    return pd.DataFrame(infos), data, errors
