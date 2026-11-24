"""Streamlit app: upload country templates, see validation results, download the QC log.

Run with:  streamlit run app.py
"""

import io
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

from edu_pipeline import outputs, pipeline

st.set_page_config(page_title="Education template validator", layout="wide")
st.title("Education data template validator")
st.caption("Upload one or more completed templates (any language version). "
           "Files are processed in memory and are not stored.")

files = st.file_uploader("Country templates (.xlsx)", type="xlsx", accept_multiple_files=True)
if not files:
    st.info("Tip: run `python -m edu_pipeline generate` to create sample templates to try.")
    st.stop()

with tempfile.TemporaryDirectory() as tmp:
    for f in files:
        Path(tmp, f.name).write_bytes(f.getvalue())
    result = pipeline.process(Path(tmp))
    qc_path = Path(tmp, "qc_log.xlsx")
    outputs.write_qc_log(qc_path, result.info, result.issues, result.unreadable)
    qc_bytes = qc_path.read_bytes()

for err in result.unreadable:
    st.error(err)
if result.info.empty:
    st.stop()

summary = result.info[["ISO code", "Country"]].copy()
summary["Errors"] = summary["ISO code"].map(lambda i: int(((result.issues.iso == i) & (result.issues.severity == "error")).sum()))
summary["Warnings"] = summary["ISO code"].map(lambda i: int(((result.issues.iso == i) & (result.issues.severity == "warning")).sum()))
summary["Status"] = summary["ISO code"].map(lambda i: outputs.country_status(result.issues, i))

c1, c2, c3 = st.columns(3)
c1.metric("Submissions", len(summary))
c2.metric("Need revision", int((summary.Status == "Needs revision").sum()))
c3.metric("Issues found", len(result.issues))
st.download_button("Download QC log (Excel)", qc_bytes, "qc_log.xlsx")

st.subheader("Summary")
st.dataframe(summary, hide_index=True, width="stretch")

iso = st.selectbox("Review a country", summary["ISO code"],
                   format_func=lambda i: f"{i} – {summary.set_index('ISO code').loc[i, 'Country']}")
left, right = st.columns([3, 2])
with left:
    st.subheader("Issues")
    st.dataframe(result.issues[result.issues.iso == iso].drop(columns="iso"), hide_index=True, width="stretch")
with right:
    st.subheader("Gross enrolment ratio (%)")
    ger = result.indicators[(result.indicators.iso == iso) & (result.indicators.indicator == "GER")]
    level = st.radio("Level", sorted(ger.level.unique()), horizontal=True)
    chart = ger[ger.level == level].pivot_table(index="year", columns="sex", values="value")
    st.line_chart(chart)

st.subheader("All indicators")
st.dataframe(result.indicators[result.indicators.iso == iso]
             .pivot_table(index=["level", "indicator", "sex"], columns="year", values="value"),
             width="stretch")
