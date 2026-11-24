"""Validation rules applied to every country submission.

Each check returns a list of issues. Thresholds and severities come from rules.yaml,
so methodology changes do not need code changes.
"""

from __future__ import annotations

from functools import lru_cache

import numpy as np
import pandas as pd
import yaml

from . import config as C

ISSUE_COLS = ["iso", "check", "severity", "sheet", "level", "sex", "age", "year", "cell", "value", "message"]


@lru_cache
def load_rules() -> dict:
    with open(C.RULES_FILE, encoding="utf-8") as f:
        return yaml.safe_load(f)["checks"]


def _issue(check, iso, message, **kw):
    rule = load_rules()[check]
    rec = {"iso": iso, "check": check, "severity": rule["severity"], "message": message}
    rec.update(kw)
    return rec


def _numeric(df):
    return df[df["status"] == "value"]


# ------------------------------------------------------------ cell checks ----
def check_cells(data: pd.DataFrame) -> list[dict]:
    out = []
    for r in data[data["status"] == "non_numeric"].itertuples():
        out.append(_issue("non_numeric", r.iso, f"'{r.raw}' is not a number or an allowed code",
                          sheet=r.sheet, level=r.level, sex=r.sex, age=r.age, year=r.year, cell=r.cell))
    num = _numeric(data)
    for r in num[num["value"] < 0].itertuples():
        out.append(_issue("negative_value", r.iso, f"Negative value {r.value:,.0f}",
                          sheet=r.sheet, level=r.level, sex=r.sex, age=r.age, year=r.year, cell=r.cell, value=r.value))
    for r in num[(num["value"] % 1) != 0].itertuples():
        out.append(_issue("non_integer", r.iso, f"Non-integer count {r.value}",
                          sheet=r.sheet, level=r.level, sex=r.sex, age=r.age, year=r.year, cell=r.cell, value=r.value))
    return out


def check_totals(data: pd.DataFrame) -> list[dict]:
    """Totals typed over the SUM formula must still equal the sum of the ages."""
    tol = load_rules()["total_mismatch"]["tolerance"]
    out = []
    totals = data[(data["age"] == C.TOTAL_LABEL) & (data["status"] == "value")]
    ages = data[data["age"].str.isdigit()]
    for r in totals.itertuples():
        parts = ages[(ages.iso == r.iso) & (ages.sheet == r.sheet) & (ages.level == r.level)
                     & (ages.sex == r.sex) & (ages.year == r.year)]
        if (parts["status"] != "value").any() and not parts.empty:
            continue  # cannot compare when ages are incomplete
        expected = parts["value"].sum()
        if abs(r.value - expected) > tol:
            out.append(_issue("total_mismatch", r.iso,
                              f"Total {r.value:,.0f} typed over the formula; sum of ages is {expected:,.0f}",
                              sheet=r.sheet, level=r.level, sex=r.sex, age=r.age, year=r.year,
                              cell=r.cell, value=r.value))
    return out


# ------------------------------------------------------- consistency checks ----
def _enrolment_by_age(data, by_sex=True):
    e = _numeric(data[(data.sheet == "Enrolment") & data.age.str.isdigit()]).copy()
    e["age_n"] = e["age"].astype(int)
    keys = ["iso", "sex", "age_n", "year"] if by_sex else ["iso", "age_n", "year"]
    return e.groupby(keys)["value"].sum()


def check_enrolment_vs_population(data: pd.DataFrame) -> list[dict]:
    max_ratio = load_rules()["enrolment_exceeds_population"]["max_ratio"]
    enr = _enrolment_by_age(data)
    p = _numeric(data[(data.sheet == "Population") & data.age.str.isdigit()]).copy()
    p["age_n"] = p["age"].astype(int)
    pop = p.groupby(["iso", "sex", "age_n", "year"])["value"].sum()
    ratio = (enr / pop.replace(0, np.nan)).dropna()
    out = []
    for (iso, sex, age, year), v in ratio[ratio > max_ratio].items():
        out.append(_issue("enrolment_exceeds_population", iso,
                          f"Enrolment at age {age} is {v:.0%} of the population of that age",
                          sheet="Enrolment", sex=sex, age=str(age), year=year, value=round(v, 3)))
    return out


def check_age_cohorts(data: pd.DataFrame) -> list[dict]:
    """Follow each cohort diagonally: age a in year t versus age a-1 in year t-1."""
    rule = load_rules()["age_cohort_break"]
    lo_age, hi_age = rule["ages"]
    enr = _enrolment_by_age(data, by_sex=False)
    out = []
    for (iso, age, year), v in enr.items():
        if not lo_age <= age <= hi_age:
            continue
        prev = enr.get((iso, age - 1, year - 1))
        if not prev:
            continue
        ratio = v / prev
        if not rule["min_ratio"] <= ratio <= rule["max_ratio"]:
            out.append(_issue("age_cohort_break", iso,
                              f"Cohort aged {age} in {year} is {ratio:.2f}x the cohort aged {age - 1} in {year - 1}",
                              sheet="Enrolment", age=str(age), year=year, value=round(ratio, 3)))
    return out


def check_time_series(data: pd.DataFrame) -> list[dict]:
    max_change = load_rules()["year_on_year_change"]["max_change"]
    d = data[(data.sheet.isin(["Enrolment", "Teachers"])) & (data.age != C.TOTAL_LABEL)]
    complete = d.groupby(["iso", "sheet", "level", "sex", "year"])["status"].apply(lambda s: (s == "value").all())
    totals = _numeric(d).groupby(["iso", "sheet", "level", "sex", "year"])["value"].sum()[complete[complete].index]
    out = []
    for (iso, sheet, level, sex, year), v in totals.items():
        prev = totals.get((iso, sheet, level, sex, year - 1))
        if prev:
            change = v / prev - 1
            if abs(change) > max_change:
                out.append(_issue("year_on_year_change", iso,
                                  f"{sheet} total changed by {change:+.0%} from {year - 1}",
                                  sheet=sheet, level=level, sex=sex, year=year, value=round(change, 3)))
    return out


def check_completeness(data: pd.DataFrame) -> list[dict]:
    latest = data["year"].max()
    d = data[(data.year == latest) & (data.age != C.TOTAL_LABEL)]
    gaps = d[d.status.isin(["blank", "missing"])]
    out = []
    for (iso, sheet, level, sex), g in gaps.groupby(["iso", "sheet", "level", "sex"]):
        out.append(_issue("missing_latest_year", iso,
                          f"{len(g)} cell(s) blank or coded 'm' for {latest}",
                          sheet=sheet, level=level, sex=sex, year=latest,
                          cell=f"{g['cell'].iloc[0]} (+{len(g) - 1} more)" if len(g) > 1 else g["cell"].iloc[0]))
    return out


# ------------------------------------------------------- indicator checks ----
def check_indicators(ind: pd.DataFrame) -> list[dict]:
    rules = load_rules()
    out = []
    ind = ind.dropna(subset=["value"])
    for r in ind.itertuples():
        kw = dict(sheet="Indicators", level=r.level, sex=r.sex, year=r.year, value=r.value)
        if r.indicator == "NER" and r.value > 100:
            out.append(_issue("ner_above_100", r.iso, f"NER is {r.value:.1f}%", **kw))
        elif r.indicator == "GER":
            rule = rules["ger_out_of_range"]
            if not rule["min"] <= r.value <= rule["max"]:
                out.append(_issue("ger_out_of_range", r.iso, f"GER is {r.value:.1f}%", **kw))
        elif r.indicator == "GPI":
            rule = rules["gpi_out_of_range"]
            if not rule["min"] <= r.value <= rule["max"]:
                out.append(_issue("gpi_out_of_range", r.iso, f"GPI is {r.value:.2f}", **kw))
        elif r.indicator == "PTR":
            rule = rules["ptr_out_of_range"]
            if not rule["min"] <= r.value <= rule["max"]:
                out.append(_issue("ptr_out_of_range", r.iso, f"Pupil-teacher ratio is {r.value:.1f}", **kw))
    return out


def run_all(data: pd.DataFrame, indicators: pd.DataFrame) -> pd.DataFrame:
    issues = []
    for check in (check_cells, check_totals, check_enrolment_vs_population,
                  check_age_cohorts, check_time_series, check_completeness):
        issues += check(data)
    issues += check_indicators(indicators)
    df = pd.DataFrame(issues, columns=ISSUE_COLS)
    df["year"] = df["year"].astype("Int64")
    order = {"error": 0, "warning": 1}
    return (df.assign(_o=df["severity"].map(order))
              .sort_values(["iso", "_o", "check", "year"]).drop(columns="_o").reset_index(drop=True))
