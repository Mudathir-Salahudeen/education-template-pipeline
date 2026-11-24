"""Tests: run with `pytest -q` from the project folder."""

from pathlib import Path

import pandas as pd
import pytest
from openpyxl import load_workbook

from edu_pipeline import config as C
from edu_pipeline import indicators, pipeline, reader, synthetic, template, translate, update

YEARS = list(range(2019, 2026))


@pytest.fixture(scope="module")
def submissions(tmp_path_factory) -> Path:
    folder = tmp_path_factory.mktemp("subs")
    for name, wb in synthetic.build_submissions(YEARS, survey_year=2026):
        wb.save(folder / name)
    return folder


@pytest.fixture(scope="module")
def result(submissions):
    return pipeline.process(submissions)


def checks_for(result, iso):
    return set(result.issues.loc[result.issues.iso == iso, "check"])


def test_clean_countries_pass(result):
    assert checks_for(result, "ARL") == set()
    assert checks_for(result, "GLE") == set()


@pytest.mark.parametrize("iso, expected", [
    ("BOR", {"negative_value", "non_numeric"}),
    ("CAL", {"total_mismatch"}),
    ("DEL", {"age_cohort_break"}),
    ("EST", {"missing_latest_year"}),
    ("FEN", {"year_on_year_change"}),
    ("HOL", {"ner_above_100", "enrolment_exceeds_population"}),
])
def test_injected_errors_are_caught(result, iso, expected):
    assert expected <= checks_for(result, iso)


def test_indicator_formulas():
    rows = []
    for sex in C.SEXES:
        for age in C.AGES:
            rows.append(dict(iso="T", sheet="Population", level="Population", sex=sex, age=str(age), year=2025, value=100.0, status="value"))
            enrol = 90.0 if 6 <= age <= 11 else (10.0 if age == 12 else 0.0)
            rows.append(dict(iso="T", sheet="Enrolment", level="Primary", sex=sex, age=str(age), year=2025, value=enrol, status="value"))
            rows.append(dict(iso="T", sheet="Enrolment", level="Lower secondary", sex=sex, age=str(age), year=2025, value=0.0, status="value"))
        for lvl in C.LEVELS:
            rows.append(dict(iso="T", sheet="Teachers", level=lvl, sex=sex, age="All ages", year=2025, value=10.0, status="value"))
    ind = indicators.calculate(pd.DataFrame(rows))
    get = lambda i, s="Both sexes": ind[(ind.level == "Primary") & (ind.indicator == i) & (ind.sex == s)].value.iloc[0]
    # 6 official ages x 90 + 10 over-age = 550 pupils per sex; population 6 x 100 = 600
    assert get("GER", "Male") == pytest.approx(91.7, abs=0.05)
    assert get("NER", "Male") == pytest.approx(90.0)
    assert get("GPI") == pytest.approx(1.0)
    assert get("PTR") == pytest.approx(55.0)   # 1100 pupils / 20 teachers


def test_missing_input_gives_no_indicator():
    df = pd.DataFrame([
        dict(iso="T", sheet="Population", level="Population", sex="Male", age="6", year=2025, value=100.0, status="value"),
        dict(iso="T", sheet="Enrolment", level="Primary", sex="Male", age="6", year=2025, value=None, status="missing"),
    ])
    ger = indicators.calculate(df)
    assert ger[(ger.indicator == "GER") & (ger.sex == "Male")].value.isna().all()


def test_add_year_keeps_data_and_extends_formulas(submissions, tmp_path):
    src = submissions / "ARL_Arlandia.xlsx"
    wb = load_workbook(src)
    before = wb["Enrolment"]["D5"].value
    update.add_year(wb, 2026)
    ws = wb["Enrolment"]
    cols = template.year_columns(ws)
    assert 2026 in cols and ws["D5"].value == before
    total_row = template.row_of("Enrolment", ("Primary", "Male", C.TOTAL_LABEL))
    assert ws.cell(total_row, cols[2026]).value.startswith("=SUM(K")
    with pytest.raises(ValueError):
        update.add_year(wb, 2028)  # cannot skip a year


def test_prefill_never_overwrites(submissions):
    wb = load_workbook(submissions / "ARL_Arlandia.xlsx")
    rows = pd.DataFrame([dict(sheet="Population", level="Population", sex="Male", age=4, year=2025, value=1)])
    assert update.prefill(wb, rows) == 0  # 2025 already reported


def test_translated_template_reads_back_as_english(submissions, tmp_path):
    wb = load_workbook(submissions / "GLE_Glenmora.xlsx")
    translate.translate_workbook(wb, "fr")
    assert wb["Enrolment"]["A5"].value == "Primaire"
    path = tmp_path / "GLE_FR.xlsx"
    wb.save(path)
    info, df = reader.read_template(path)
    assert info["Country"] == "Glenmora"
    assert set(df.level) >= {"Primary", "Lower secondary"} and set(df.sex) >= {"Male", "Female"}


def test_missing_sheet_is_reported(tmp_path, submissions):
    wb = load_workbook(submissions / "ARL_Arlandia.xlsx")
    del wb["Teachers"]
    wb.save(tmp_path / "broken.xlsx")
    _, _, errors = reader.read_folder(tmp_path)
    assert errors and "Teachers" in errors[0]


def test_outputs_written(result, tmp_path):
    pipeline.write_all(result, tmp_path)
    for f in ("qc_log.xlsx", "database.sqlite", "consolidated_data.csv", "indicators.csv", "reports/index.html", "reports/HOL.html"):
        assert (tmp_path / f).exists(), f
