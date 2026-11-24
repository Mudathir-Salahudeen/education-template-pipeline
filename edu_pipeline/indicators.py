"""Calculate education indicators from the tidy data table.

Definitions (per level, sex and year):
  GER = enrolment in the level, all ages / population of the official age group x 100
  NER = enrolment of official-age pupils in the level / population of the official age group x 100
  GPI = female GER / male GER
  PTR = pupils in the level (both sexes) / teachers in the level (both sexes)

An indicator is left empty (not calculated) when any input it needs is missing,
rather than silently treating missing cells as zero.
"""

from __future__ import annotations

import numpy as np
import pandas as pd

from . import config as C


def _complete_sum(s: pd.Series) -> float:
    return s.sum() if s.notna().all() and len(s) else np.nan


def _ages(df: pd.DataFrame) -> pd.DataFrame:
    out = df[df["age"].str.isdigit()].copy()
    out["age_n"] = out["age"].astype(int)
    return out


def calculate(data: pd.DataFrame) -> pd.DataFrame:
    if data.empty:
        return pd.DataFrame(columns=["iso", "level", "sex", "indicator", "year", "value"])
    d = data.copy()
    d.loc[d["status"] != "value", "value"] = np.nan
    enrol = _ages(d[d["sheet"] == "Enrolment"])
    pop = _ages(d[d["sheet"] == "Population"])
    teach = d[d["sheet"] == "Teachers"]

    rows = []
    for level, spec in C.LEVELS.items():
        lo, hi = spec["official_ages"]
        e = enrol[enrol["level"] == level]
        p_off = pop[pop["age_n"].between(lo, hi)]
        e_off = e[e["age_n"].between(lo, hi)]
        for sex in C.SEXES + [C.BOTH_SEXES]:
            sel = (lambda x: x) if sex == C.BOTH_SEXES else (lambda x, s=sex: x[x["sex"] == s])
            num_all = sel(e).groupby(["iso", "year"])["value"].agg(_complete_sum)
            num_off = sel(e_off).groupby(["iso", "year"])["value"].agg(_complete_sum)
            den = sel(p_off).groupby(["iso", "year"])["value"].agg(_complete_sum)
            for name, num in (("GER", num_all), ("NER", num_off)):
                val = (100 * num / den.replace(0, np.nan)).round(1)
                rows += [{"iso": i, "level": level, "sex": sex, "indicator": name, "year": y, "value": v}
                         for (i, y), v in val.items()]
        pupils = e.groupby(["iso", "year"])["value"].agg(_complete_sum)
        teachers = teach[teach["level"] == level].groupby(["iso", "year"])["value"].agg(_complete_sum)
        ptr = (pupils / teachers.replace(0, np.nan)).round(1)
        rows += [{"iso": i, "level": level, "sex": C.BOTH_SEXES, "indicator": "PTR", "year": y, "value": v}
                 for (i, y), v in ptr.items()]

    out = pd.DataFrame(rows)
    ger = out[out["indicator"] == "GER"].pivot_table(index=["iso", "level", "year"], columns="sex", values="value")
    if {"Female", "Male"} <= set(ger.columns):
        gpi = (ger["Female"] / ger["Male"]).round(2).reset_index(name="value")
        gpi["sex"], gpi["indicator"] = C.BOTH_SEXES, "GPI"
        out = pd.concat([out, gpi[out.columns]], ignore_index=True)
    return out.sort_values(["iso", "level", "indicator", "sex", "year"]).reset_index(drop=True)
