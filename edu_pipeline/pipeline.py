"""Process a folder of country submissions end to end."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import pandas as pd

from . import indicators as ind
from . import outputs, reader, validate


@dataclass
class RunResult:
    info: pd.DataFrame
    data: pd.DataFrame
    indicators: pd.DataFrame
    issues: pd.DataFrame
    unreadable: list[str]


def process(folder: Path) -> RunResult:
    info, data, unreadable = reader.read_folder(folder)
    indicators = ind.calculate(data)
    issues = validate.run_all(data, indicators) if not data.empty else pd.DataFrame(columns=validate.ISSUE_COLS)
    return RunResult(info, data, indicators, issues, unreadable)


def write_all(result: RunResult, out: Path) -> None:
    out.mkdir(parents=True, exist_ok=True)
    result.data.to_csv(out / "consolidated_data.csv", index=False)
    result.indicators.to_csv(out / "indicators.csv", index=False)
    # Stata copies for colleagues who work in Stata (R reads the CSV or SQLite directly)
    result.indicators.to_stata(out / "indicators.dta", write_index=False, version=118)
    result.data.drop(columns=["raw"]).fillna({"cell": ""}).to_stata(
        out / "consolidated_data.dta", write_index=False, version=118)
    outputs.write_database(out / "database.sqlite", result.info, result.data, result.indicators, result.issues)
    outputs.write_qc_log(out / "qc_log.xlsx", result.info, result.issues, result.unreadable)
    outputs.write_reports(out / "reports", result.info, result.indicators, result.issues)


def summary_text(result: RunResult) -> str:
    lines = [f"Read {len(result.info)} submission(s), {len(result.data):,} cells; "
             f"{len(result.unreadable)} unreadable file(s)."]
    for rec in result.info.to_dict("records"):
        iso = rec["ISO code"]
        sub = result.issues[result.issues.iso == iso]
        lines.append(f"  {iso} {rec['Country']:<10} {outputs.country_status(result.issues, iso):<22}"
                     f"{(sub.severity == 'error').sum():>3} errors {(sub.severity == 'warning').sum():>3} warnings")
    return "\n".join(lines)
