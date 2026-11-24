"""Shared configuration: template layout, education levels, ages and missing-value codes.

Everything that describes *the shape of the template* lives here, so that the
generator, the updater, the reader and the validator all agree on one layout.
"""

from pathlib import Path

# --- Education levels and their official entrance/exit ages (ISCED-style) ---
LEVELS = {
    "Primary": {"official_ages": (6, 11)},
    "Lower secondary": {"official_ages": (12, 14)},
}

# Single ages reported in the enrolment and population tables
AGES = list(range(4, 19))
SEXES = ["Male", "Female"]
BOTH_SEXES = "Both sexes"
TOTAL_LABEL = "Total"

DEFAULT_YEARS = list(range(2019, 2026))

# UIS-style special codes that countries may type instead of a number
MISSING_CODES = {
    "m": "missing",
    "a": "not applicable",
    "x": "included elsewhere",
}

# --- Workbook layout ---
DATA_SHEETS = ["Population", "Enrolment", "Teachers"]
ALL_SHEETS = ["Info", "Population", "Enrolment", "Teachers", "Indicators", "Metadata"]

TITLE_ROW = 1
NOTE_ROW = 2
HEADER_ROW = 4
FIRST_DATA_ROW = 5
LABEL_COLS = 3          # A = Level, B = Sex, C = Age
FIRST_YEAR_COL = LABEL_COLS + 1   # column D

INFO_FIELDS = ["Country", "ISO code", "Survey year", "Contact person", "Contact email"]

RULES_FILE = Path(__file__).resolve().parent.parent / "rules.yaml"
LABELS_FILE = Path(__file__).resolve().parent.parent / "labels.yaml"
