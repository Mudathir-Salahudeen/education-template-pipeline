"""Produce French, Spanish, Portuguese and Russian versions of a template.

Only displayed text is translated. Sheet names, numbers and formulas are left as
they are, so every language version can be read by the same scripts. Labels that
formulas look up (level and sex) are translated consistently on every sheet, so
the Indicators formulas still match.
"""

from __future__ import annotations

from functools import lru_cache
from pathlib import Path

import yaml
from openpyxl import load_workbook

from . import config as C

LANGUAGES = ("fr", "es", "pt", "ru")


@lru_cache
def load_labels() -> dict[str, dict[str, str]]:
    with open(C.LABELS_FILE, encoding="utf-8") as f:
        return yaml.safe_load(f)


@lru_cache
def reverse_labels() -> dict[str, str]:
    """Any translated label -> English label."""
    out = {}
    for en, tr in load_labels().items():
        out[en] = en
        for text in tr.values():
            out.setdefault(text, en)
    return out


def translate_workbook(wb, lang: str) -> int:
    if lang not in LANGUAGES:
        raise ValueError(f"Unsupported language {lang!r}; choose from {LANGUAGES}")
    labels = load_labels()
    n = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value in labels:
                    cell.value = labels[cell.value][lang]
                    n += 1
    return n


def translate_file(src: Path, out_dir: Path, langs=LANGUAGES) -> list[Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    paths = []
    for lang in langs:
        wb = load_workbook(src)
        translate_workbook(wb, lang)
        p = out_dir / f"{src.stem}_{lang.upper()}.xlsx"
        wb.save(p)
        paths.append(p)
    return paths


def untranslate(value):
    if isinstance(value, str):
        return reverse_labels().get(value.strip(), value.strip())
    return value
