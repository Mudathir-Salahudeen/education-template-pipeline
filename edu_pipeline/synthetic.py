"""Generate realistic *synthetic* country submissions for testing and demos.

No real country data is used. Population follows birth cohorts, enrolment follows
age-specific rates, and a few submissions deliberately contain the kinds of errors
countries make, so the validation step has something to find.
"""

from __future__ import annotations

import random

from . import config as C
from .template import build_template, row_of, year_columns

COUNTRIES = [
    # name, iso, births in base year, annual birth growth, seed, injected problems
    ("Arlandia", "ARL", 520_000, 0.020, 1, []),
    ("Borvania", "BOR", 180_000, 0.028, 2, ["negative_value", "text_in_cell"]),
    ("Calmara", "CAL", 95_000, 0.012, 3, ["overwritten_total"]),
    ("Deltaria", "DEL", 340_000, 0.025, 4, ["age_misreporting"]),
    ("Estoria", "EST", 61_000, 0.008, 5, ["missing_latest_year"]),
    ("Fenwick", "FEN", 230_000, 0.022, 6, ["teacher_jump"]),
    ("Glenmora", "GLE", 145_000, 0.018, 7, []),
    ("Hollin", "HOL", 410_000, 0.030, 8, ["girls_overcounted"]),
]

# Age-specific enrolment rates by level (share of the population of that age)
PRIMARY_RATES = {4: 0.02, 5: 0.12, 6: 0.86, 7: 0.93, 8: 0.95, 9: 0.95, 10: 0.94, 11: 0.92, 12: 0.12, 13: 0.04, 14: 0.01}
LOWSEC_RATES = {11: 0.02, 12: 0.72, 13: 0.74, 14: 0.70, 15: 0.12, 16: 0.04, 17: 0.01}


def make_values(births0: int, growth: float, seed: int, years: list[int]) -> dict:
    rng = random.Random(seed)
    base_year = years[0]
    level_shift = {"Primary": rng.uniform(-0.08, 0.04), "Lower secondary": rng.uniform(-0.25, 0.05)}
    sex_gap = {"Male": 0.0, "Female": rng.uniform(-0.07, 0.01)}
    trend = rng.uniform(0.002, 0.012)  # enrolment rates improve slowly over time
    ptr = {"Primary": rng.uniform(32, 48), "Lower secondary": rng.uniform(22, 30)}
    fem_share = {"Primary": rng.uniform(0.45, 0.70), "Lower secondary": rng.uniform(0.35, 0.60)}
    values = {}
    for year in years:
        for sex in C.SEXES:
            share = 0.51 if sex == "Male" else 0.49
            for age in C.AGES:
                cohort = year - age
                births = births0 * (1 + growth) ** (cohort - (base_year - 10)) * share
                pop = round(births * 0.985 ** age * rng.uniform(0.995, 1.005))
                values[("Population", "Population", sex, age, year)] = pop
                for level, rates in (("Primary", PRIMARY_RATES), ("Lower secondary", LOWSEC_RATES)):
                    rate = rates.get(age, 0.0)
                    if rate:
                        rate = rate * (1 + level_shift[level] + sex_gap[sex] + trend * (year - base_year))
                        rate = min(rate, 0.99)
                    values[("Enrolment", level, sex, age, year)] = round(pop * rate * rng.uniform(0.98, 1.02))
        for level in C.LEVELS:
            pupils = sum(values[("Enrolment", level, s, a, year)] for s in C.SEXES for a in C.AGES)
            teachers = pupils / ptr[level] * rng.uniform(0.97, 1.03)
            share = fem_share[level] + 0.004 * (year - base_year)
            values[("Teachers", level, "Female", "All ages", year)] = round(teachers * share)
            values[("Teachers", level, "Male", "All ages", year)] = round(teachers * (1 - share))
    return values


def _cell(ws, sheet, key, year):
    return ws.cell(row_of(sheet, key), year_columns(ws)[year])


def inject(wb, problem: str, years: list[int]) -> None:
    """Introduce a realistic reporting error into a filled workbook."""
    last, prev = years[-1], years[-2]
    if problem == "negative_value":
        c = _cell(wb["Enrolment"], "Enrolment", ("Primary", "Male", 9), prev)
        c.value = -abs(c.value)
    elif problem == "text_in_cell":
        c = _cell(wb["Population"], "Population", ("Population", "Female", 10), last)
        c.value = f"{c.value:,} approx"
    elif problem == "overwritten_total":
        c = _cell(wb["Enrolment"], "Enrolment", ("Lower secondary", "Female", C.TOTAL_LABEL), last)
        c.value = 41_250  # typed over the SUM formula
    elif problem == "age_misreporting":
        # Pupils of age 8 reported as age 7 in one year
        ws = wb["Enrolment"]
        c7 = _cell(ws, "Enrolment", ("Primary", "Male", 7), prev)
        c8 = _cell(ws, "Enrolment", ("Primary", "Male", 8), prev)
        c7.value, c8.value = c7.value + int(c8.value * 0.9), int(c8.value * 0.1)
    elif problem == "missing_latest_year":
        ws = wb["Enrolment"]
        for age in C.AGES:
            _cell(ws, "Enrolment", ("Lower secondary", "Female", age), last).value = None
        _cell(ws, "Enrolment", ("Lower secondary", "Male", 15), last).value = "m"
    elif problem == "teacher_jump":
        c = _cell(wb["Teachers"], "Teachers", ("Primary", "Female", "All ages"), last)
        c.value = c.value * 3
    elif problem == "girls_overcounted":
        ws = wb["Enrolment"]
        for age in range(6, 12):
            c = _cell(ws, "Enrolment", ("Primary", "Female", age), last)
            c.value = int(c.value * 1.45)


def build_submissions(years: list[int], survey_year: int):
    """Yield (filename, workbook) for each synthetic country."""
    for name, iso, births, growth, seed, problems in COUNTRIES:
        values = make_values(births, growth, seed, years)
        info = {
            "Country": name,
            "ISO code": iso,
            "Survey year": survey_year,
            "Contact person": "Data focal point",
            "Contact email": f"stats@{name.lower()}.example",
        }
        wb = build_template(years, info=info, values=values)
        for p in problems:
            inject(wb, p, years)
        meta = wb["Metadata"]
        meta.cell(C.FIRST_DATA_ROW, 1, "Enrolment")
        meta.cell(C.FIRST_DATA_ROW, 2, years[-1])
        meta.cell(C.FIRST_DATA_ROW, 3, "Annual school census (synthetic)")
        yield f"{iso}_{name}.xlsx", wb


def population_projection(years: list[int]) -> list[dict]:
    """Synthetic population projections, used to pre-fill a new year in the templates."""
    rows = []
    for name, iso, births, growth, seed, _ in COUNTRIES:
        values = make_values(births, growth, seed + 100, years)
        for (sheet, _lvl, sex, age, year), v in values.items():
            if sheet == "Population" and year == years[-1]:
                rows.append({"iso": iso, "sheet": "Population", "level": "Population",
                             "sex": sex, "age": age, "year": year, "value": v})
    return rows

