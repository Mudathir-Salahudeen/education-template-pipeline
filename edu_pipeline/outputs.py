"""Write the outputs of a processing run.

- database.sqlite      : submissions, data, indicators and QC issues in one database
- consolidated_data.csv: every reported value in long format, all countries
- indicators.csv       : calculated indicators, all countries
- qc_log.xlsx          : summary, quality-control checklist (country x check) and issue list
- reports/<ISO>.html   : one indicator review report per country, plus an index page
"""

from __future__ import annotations

import html
import sqlite3
from datetime import date
from pathlib import Path

import pandas as pd
from jinja2 import Template
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from .validate import load_rules

HEAD = PatternFill("solid", fgColor="303658")
RED = PatternFill("solid", fgColor="F8D7DA")
AMBER = PatternFill("solid", fgColor="FFF3CD")
GREEN = PatternFill("solid", fgColor="D4EDDA")


def country_status(issues: pd.DataFrame, iso: str) -> str:
    sub = issues[issues.iso == iso]
    if (sub.severity == "error").any():
        return "Needs revision"
    if len(sub):
        return "Accepted with queries"
    return "Accepted"


def checklist(info: pd.DataFrame, issues: pd.DataFrame) -> pd.DataFrame:
    checks = list(load_rules())
    rows = []
    for iso in info["ISO code"]:
        sub = issues[issues.iso == iso]
        row = {"ISO code": iso}
        for c in checks:
            n = (sub.check == c).sum()
            row[c] = "PASS" if n == 0 else f"{sub[sub.check == c].severity.iloc[0].upper()} ({n})"
        rows.append(row)
    return pd.DataFrame(rows)


def _write_sheet(ws, df: pd.DataFrame, colour_col: str | None = None) -> None:
    ws.append(list(df.columns))
    for c in ws[1]:
        c.font, c.fill = Font(bold=True, color="FFFFFF"), HEAD
    for rec in df.itertuples(index=False):
        ws.append([None if pd.isna(v) else v for v in rec])
    for i, col in enumerate(df.columns, start=1):
        width = max([len(str(col))] + [len(str(v)) for v in df[col].head(200)]) + 2
        ws.column_dimensions[get_column_letter(i)].width = min(width, 70)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            v = str(cell.value or "")
            if v.startswith("ERROR") or v in ("error", "Needs revision"):
                cell.fill = RED
            elif v.startswith("WARNING") or v in ("warning", "Accepted with queries"):
                cell.fill = AMBER
            elif v in ("PASS", "Accepted"):
                cell.fill = GREEN


def write_qc_log(path: Path, info, issues, struct_errors) -> None:
    summary = info[["ISO code", "Country", "Survey year", "file"]].copy()
    summary["Errors"] = summary["ISO code"].map(lambda i: int(((issues.iso == i) & (issues.severity == "error")).sum()))
    summary["Warnings"] = summary["ISO code"].map(lambda i: int(((issues.iso == i) & (issues.severity == "warning")).sum()))
    summary["Status"] = summary["ISO code"].map(lambda i: country_status(issues, i))
    wb = Workbook()
    _write_sheet(wb.active, summary)
    wb.active.title = "Summary"
    _write_sheet(wb.create_sheet("QC checklist"), checklist(info, issues))
    _write_sheet(wb.create_sheet("Issues"), issues)
    rules = pd.DataFrame([{"check": k, "severity": v["severity"], "description": v["description"].strip()}
                          for k, v in load_rules().items()])
    _write_sheet(wb.create_sheet("Rules"), rules)
    if struct_errors:
        _write_sheet(wb.create_sheet("Unreadable files"), pd.DataFrame({"error": struct_errors}))
    wb.save(path)


def write_database(path: Path, info, data, indicators, issues) -> None:
    if path.exists():
        path.unlink()
    with sqlite3.connect(path) as con:
        info.astype(str).to_sql("submissions", con, index=False)
        data.to_sql("data", con, index=False)
        indicators.to_sql("indicators", con, index=False)
        issues.to_sql("qc_issues", con, index=False)
        con.execute("CREATE INDEX idx_data ON data (iso, sheet, year)")
        con.execute("CREATE INDEX idx_ind ON indicators (iso, indicator, year)")


# ---------------------------------------------------------------- reports ----
def _svg_lines(series: dict[str, pd.Series], width=520, height=170) -> str:
    """Small dependency-free SVG line chart: {label: Series(index=year)}."""
    series = {k: v.dropna() for k, v in series.items() if v.notna().any()}
    if not series:
        return ""
    years = sorted({y for s in series.values() for y in s.index})
    vals = [v for s in series.values() for v in s.values]
    lo, hi = min(vals) * 0.95, max(vals) * 1.05
    pad_l, pad_b, pad_t, pad_r = 40, 24, 10, 110
    x = lambda y: pad_l + (years.index(y) / max(len(years) - 1, 1)) * (width - pad_l - pad_r)
    yy = lambda v: pad_t + (1 - (v - lo) / (hi - lo or 1)) * (height - pad_t - pad_b)
    colours = ["#303658", "#C0504D", "#4F81BD", "#9BBB59"]
    parts = [f'<svg viewBox="0 0 {width} {height}" width="{width}" height="{height}" role="img">']
    for t in (lo, (lo + hi) / 2, hi):
        parts.append(f'<line x1="{pad_l}" x2="{width - pad_r}" y1="{yy(t):.1f}" y2="{yy(t):.1f}" stroke="#e5e5e5"/>'
                     f'<text x="{pad_l - 6}" y="{yy(t) + 4:.1f}" font-size="10" text-anchor="end" fill="#666">{t:.0f}</text>')
    for y in years:
        parts.append(f'<text x="{x(y):.1f}" y="{height - 6}" font-size="10" text-anchor="middle" fill="#666">{y}</text>')
    for (label, s), col in zip(series.items(), colours):
        pts = " ".join(f"{x(y):.1f},{yy(v):.1f}" for y, v in s.items())
        parts.append(f'<polyline points="{pts}" fill="none" stroke="{col}" stroke-width="2"/>')
        ly, lv = s.index[-1], s.iloc[-1]
        parts.append(f'<text x="{x(ly) + 6:.1f}" y="{yy(lv) + 4:.1f}" font-size="11" fill="{col}">{html.escape(label)}</text>')
    parts.append("</svg>")
    return "".join(parts)


REPORT = Template("""<!DOCTYPE html><html lang="en"><head><meta charset="utf-8">
<title>{{ country }} ({{ iso }}) – indicator review</title>
<style>
body{font-family:Georgia,serif;max-width:1000px;margin:30px auto;color:#222;padding:0 16px}
h1{color:#303658;margin-bottom:4px}h2{color:#303658;border-bottom:2px solid #303658;padding-bottom:2px;margin-top:28px}
.meta{color:#555}.badge{display:inline-block;padding:3px 10px;border-radius:4px;font-weight:bold}
.Accepted{background:#D4EDDA}.Accepted.with.queries{background:#FFF3CD}.Needs.revision{background:#F8D7DA}
table{border-collapse:collapse;width:100%;font-size:13px;margin-top:8px}
th{background:#303658;color:#fff;text-align:left;padding:4px 6px}td{border-bottom:1px solid #ddd;padding:4px 6px}
td.num{text-align:right}.error{background:#F8D7DA}.warning{background:#FFF3CD}.pass{color:#2e7d32}
.charts{display:flex;gap:24px;flex-wrap:wrap}td:last-child{white-space:nowrap}
table.ind th{background:#E7ECF2;color:#222}table.ind td{text-align:right}
</style></head><body>
<h1>{{ country }} ({{ iso }})</h1>
<div class="meta">Indicator review report · survey year {{ survey_year }} · file {{ file }} · generated {{ today }}</div>
<p>Status: <span class="badge {{ status }}">{{ status }}</span> &nbsp; {{ n_err }} error(s), {{ n_warn }} warning(s)</p>

<h2>Quality-control checklist</h2>
<table><tr><th>Check</th><th>Description</th><th>Result</th></tr>
{% for c in checklist %}<tr><td>{{ c.check }}</td><td>{{ c.description }}</td>
<td class="{{ c.css }}">{{ c.result }}</td></tr>{% endfor %}</table>

<h2>Issues to resolve with the country</h2>
{% if issues %}<table><tr><th>Severity</th><th>Check</th><th>Sheet</th><th>Level</th><th>Sex</th><th>Age</th><th>Year</th><th>Cell</th><th>Detail</th></tr>
{% for i in issues %}<tr class="{{ i.severity }}"><td>{{ i.severity }}</td><td>{{ i.check }}</td><td>{{ i.sheet or '' }}</td>
<td>{{ i.level or '' }}</td><td>{{ i.sex or '' }}</td><td>{{ i.age or '' }}</td><td>{{ i.year or '' }}</td><td>{{ i.cell or '' }}</td><td>{{ i.message }}</td></tr>{% endfor %}
</table>{% else %}<p class="pass">No issues found.</p>{% endif %}

<h2>Indicators</h2>
<div class="charts">{% for c in charts %}<div><b>{{ c.title }}</b><br>{{ c.svg }}</div>{% endfor %}</div>
{{ table }}
</body></html>""")

INDEX = Template("""<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"><title>Country review reports</title>
<style>body{font-family:Georgia,serif;max-width:800px;margin:30px auto}h1{color:#303658}
table{border-collapse:collapse;width:100%}th{background:#303658;color:#fff;text-align:left;padding:5px}
td{border-bottom:1px solid #ddd;padding:5px}</style></head><body><h1>Country indicator review reports</h1>
<table><tr><th>Country</th><th>Errors</th><th>Warnings</th><th>Status</th></tr>
{% for r in rows %}<tr><td><a href="{{ r.iso }}.html">{{ r.country }} ({{ r.iso }})</a></td><td>{{ r.err }}</td><td>{{ r.warn }}</td><td>{{ r.status }}</td></tr>{% endfor %}
</table></body></html>""")


def write_reports(folder: Path, info, indicators, issues) -> None:
    folder.mkdir(parents=True, exist_ok=True)
    rules = load_rules()
    index_rows = []
    for rec in info.to_dict("records"):
        iso = rec["ISO code"]
        sub = issues[issues.iso == iso]
        status = country_status(issues, iso)
        cl = []
        for check, rule in rules.items():
            hits = sub[sub.check == check]
            cl.append({"check": check, "description": rule["description"].strip(),
                       "result": "PASS" if hits.empty else f"{len(hits)} {rule['severity']}(s)",
                       "css": "pass" if hits.empty else rule["severity"]})
        ind = indicators[indicators.iso == iso]
        charts = []
        for level in ind.level.unique():
            g = ind[(ind.level == level) & (ind.indicator == "GER")]
            svg = _svg_lines({s: g[g.sex == s].set_index("year")["value"] for s in g.sex.unique()})
            charts.append({"title": f"GER (%) – {level}", "svg": svg})
        recent = sorted(ind.year.unique())[-5:]
        table = (ind[ind.year.isin(recent)]
                 .pivot_table(index=["level", "indicator", "sex"], columns="year", values="value")
                 .to_html(classes="ind", na_rep="–", float_format=lambda v: f"{v:,.2f}".rstrip("0").rstrip(".")))
        page = REPORT.render(country=rec.get("Country"), iso=iso, survey_year=rec.get("Survey year"),
                             file=rec.get("file"), today=date.today().isoformat(), status=status,
                             n_err=int((sub.severity == "error").sum()), n_warn=int((sub.severity == "warning").sum()),
                             checklist=cl, issues=sub.astype(object).where(sub.notna(), None).to_dict("records"), charts=charts, table=table)
        (folder / f"{iso}.html").write_text(page, encoding="utf-8")
        index_rows.append({"iso": iso, "country": rec.get("Country"), "status": status,
                           "err": int((sub.severity == "error").sum()), "warn": int((sub.severity == "warning").sum())})
    (folder / "index.html").write_text(INDEX.render(rows=index_rows), encoding="utf-8")
