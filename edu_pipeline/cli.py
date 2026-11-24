"""Command-line interface.

    python -m edu_pipeline generate  --out data             # blank template + synthetic submissions
    python -m edu_pipeline translate --template data/template_blank.xlsx --out data/translated
    python -m edu_pipeline add-year  --src data/submissions --dst data/updated --year 2026 \
                                     --prefill data/population_projection_2026.csv
    python -m edu_pipeline process   --src data/submissions --out outputs
"""

from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd

from . import config as C
from . import pipeline, synthetic, template, translate, update


def cmd_generate(a):
    out = Path(a.out)
    years = list(range(a.first_year, a.last_year + 1))
    (out / "submissions").mkdir(parents=True, exist_ok=True)
    template.build_template(years).save(out / "template_blank.xlsx")
    for name, wb in synthetic.build_submissions(years, survey_year=a.last_year + 1):
        wb.save(out / "submissions" / name)
    proj = pd.DataFrame(synthetic.population_projection(years + [a.last_year + 1]))
    proj.to_csv(out / f"population_projection_{a.last_year + 1}.csv", index=False)
    print(f"Wrote blank template, {len(synthetic.COUNTRIES)} synthetic submissions and a population projection to {out}/")


def cmd_translate(a):
    paths = translate.translate_file(Path(a.template), Path(a.out), a.lang)
    print("\n".join(f"Wrote {p}" for p in paths))


def cmd_add_year(a):
    log = update.update_folder(Path(a.src), Path(a.dst), a.year, Path(a.prefill) if a.prefill else None)
    print("\n".join(log))


def cmd_process(a):
    result = pipeline.process(Path(a.src))
    pipeline.write_all(result, Path(a.out))
    print(pipeline.summary_text(result))
    print(f"\nOutputs written to {a.out}/ (qc_log.xlsx, database.sqlite, consolidated_data.csv, "
          f"indicators.csv, reports/index.html)")


def main(argv=None):
    p = argparse.ArgumentParser(prog="edu_pipeline", description="Automate multi-year education data templates.")
    sub = p.add_subparsers(dest="cmd", required=True)

    g = sub.add_parser("generate", help="create a blank template and synthetic country submissions")
    g.add_argument("--out", default="data")
    g.add_argument("--first-year", type=int, default=C.DEFAULT_YEARS[0])
    g.add_argument("--last-year", type=int, default=C.DEFAULT_YEARS[-1])
    g.set_defaults(func=cmd_generate)

    t = sub.add_parser("translate", help="create FR/ES/PT/RU versions of a template")
    t.add_argument("--template", required=True)
    t.add_argument("--out", required=True)
    t.add_argument("--lang", nargs="+", default=list(translate.LANGUAGES))
    t.set_defaults(func=cmd_translate)

    u = sub.add_parser("add-year", help="add a year to every template in a folder, optionally pre-filling it")
    u.add_argument("--src", required=True)
    u.add_argument("--dst", required=True)
    u.add_argument("--year", type=int, required=True)
    u.add_argument("--prefill", help="CSV with iso, sheet, level, sex, age, year, value")
    u.set_defaults(func=cmd_add_year)

    r = sub.add_parser("process", help="validate, calculate indicators, consolidate and report")
    r.add_argument("--src", required=True)
    r.add_argument("--out", default="outputs")
    r.set_defaults(func=cmd_process)

    a = p.parse_args(argv)
    a.func(a)


if __name__ == "__main__":
    main()
