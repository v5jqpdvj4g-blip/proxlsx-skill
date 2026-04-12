#!/usr/bin/env python3
"""
_build_templates.py — regenerate the 4 skeleton templates from build_xlsx.

Run from the skill root:

    cd <repo-root>
    python templates/_build_templates.py

Each skeleton is authored in the matching theme so it dogfoods the
design system. The resulting .xlsx files in templates/ are what
load_template() reads at runtime.
"""

from __future__ import annotations

import sys
from pathlib import Path

# allow `python templates/_build_templates.py` from anywhere
_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_ROOT))

from build_xlsx import (  # noqa: E402
    new_workbook,
    new_sheet,
    write_title_banner,
    write_section_header,
    write_table,
    write_kpi_band,
    write_note,
    write_source,
    set_column_widths,
    apply_freeze,
    save_and_validate,
)
from theme import default_column_widths  # noqa: E402

_TEMPLATES_DIR = _ROOT / "templates"


# ---------------------------------------------------------------------------
# Financial-model skeleton (ib)
# ---------------------------------------------------------------------------

def build_financial_model() -> Path:
    wb = new_workbook(theme="ib", first_sheet="Cover")

    # --- Cover ------------------------------------------------------------
    ws = wb["Cover"]
    set_column_widths(ws, {"A": 2, "B": 32, "C": 16, "D": 16, "E": 16, "F": 16})
    r = write_title_banner(ws, "Project Name — 3-Statement Model",
                           subtitle="Replace with project name, date, author")
    ws.cell(row=r,     column=2, value="Prepared by")
    ws.cell(row=r,     column=3, value="Your Name")
    ws.cell(row=r + 1, column=2, value="As of")
    ws.cell(row=r + 1, column=3, value="YYYY-MM-DD")
    ws.cell(row=r + 2, column=2, value="Version")
    ws.cell(row=r + 2, column=3, value="v0.1")
    for rr in range(r, r + 3):
        ws.cell(row=rr, column=2).style = "ib_label"
        ws.cell(row=rr, column=3).style = "ib_input"
        ws.cell(row=rr, column=3).number_format = "@"

    # --- Assumptions ------------------------------------------------------
    ws = new_sheet(wb, "Assumptions")
    set_column_widths(ws, {"A": 2, "B": 32, "C": 14, "D": 14, "E": 14, "F": 14, "G": 14})
    r = write_title_banner(ws, "Assumptions", subtitle="All inputs live here")
    r = write_section_header(ws, "Revenue drivers", row=r)
    info = write_table(
        ws, f"B{r}",
        headers=["Driver", "FY23A", "FY24A", "FY25E", "FY26E", "FY27E"],
        data=[
            ["Starting ARR ($mm)", 10.0, None, None, None, None],
            ["Net new ARR growth %", None, 0.40, 0.35, 0.30, 0.25],
            ["Gross margin %",       0.72, 0.73, 0.74, 0.74, 0.75],
            ["OpEx as % of revenue", 0.65, 0.60, 0.55, 0.50, 0.48],
        ],
        num_format="percent_decimal",
    )
    write_source(ws, row=info["last_data_row"] + 2,
                 text="Management model — replace with actuals before use")

    # --- IS / BS / CF / Summary placeholders ------------------------------
    for name in ("Income Statement", "Balance Sheet", "Cash Flow", "Summary"):
        ws = new_sheet(wb, name)
        set_column_widths(ws, default_column_widths("ib", "CDEFGH"))
        r = write_title_banner(ws, name, subtitle="Link cells to Assumptions; never hard-code here")
        write_note(ws, row=r, text="This sheet is a stub. Populate with formula-driven rows that reference Assumptions.")

    out = _TEMPLATES_DIR / "financial-model-skeleton.xlsx"
    save_and_validate(wb, out, recalc=False, raise_on_fail=False)
    return out


# ---------------------------------------------------------------------------
# Data-report skeleton (ib)
# ---------------------------------------------------------------------------

def build_data_report() -> Path:
    wb = new_workbook(theme="ib", first_sheet="Report")
    ws = wb["Report"]
    set_column_widths(ws, {"A": 2, "B": 30, "C": 14, "D": 14, "E": 14, "F": 14, "G": 14})

    r = write_title_banner(ws, "Data Report — Title",
                           subtitle="Subtitle / date / author")
    r = write_section_header(ws, "Summary table", row=r)
    info = write_table(
        ws, f"B{r}",
        headers=["Category", "FY23", "FY24", "FY25", "YoY Δ", "CAGR"],
        data=[
            ["Segment A", 1000, 1200, 1500, "=D{i}/C{i}-1".format(i=r + 1), None],
            ["Segment B",  500,  520,  550, "=D{i}/C{i}-1".format(i=r + 2), None],
            ["Segment C",  200,  280,  380, "=D{i}/C{i}-1".format(i=r + 3), None],
        ],
        num_format="currency",
        total=True,
    )
    write_note(ws, row=info["last_data_row"] + 2,
               text="Replace segments and periods with your own; formulas will re-calc on open.")
    write_source(ws, row=info["last_data_row"] + 3,
                 text="[System], [Query or dashboard], [Date]")

    ws2 = new_sheet(wb, "Raw Data")
    ws2["B1"].value = "Paste raw data below row 2; headers in row 2"
    ws2["B1"].style = "ib_note"

    out = _TEMPLATES_DIR / "data-report-skeleton.xlsx"
    save_and_validate(wb, out, recalc=False, raise_on_fail=False)
    return out


# ---------------------------------------------------------------------------
# Dashboard skeleton (saas)
# ---------------------------------------------------------------------------

def build_dashboard() -> Path:
    wb = new_workbook(theme="saas", first_sheet="Dashboard")
    ws = wb["Dashboard"]
    set_column_widths(ws, {"A": 2, "B": 18, "C": 18, "D": 18, "E": 18, "F": 18, "G": 18, "H": 18})

    r = write_title_banner(ws, "Product Metrics Dashboard",
                           subtitle="Replace with product / period / owner")
    r = write_kpi_band(
        ws, row=r,
        kpis=[
            ("MRR",            0, "currency"),
            ("Active users",   0, "number"),
            ("Churn",          0, "percent_decimal"),
            ("NPS",            0, "number"),
        ],
        cols_per_kpi=1,
    )
    r = write_section_header(ws, "Weekly trend", row=r)
    write_table(
        ws, f"B{r}",
        headers=["Week", "MRR", "New", "Churn", "Net"],
        data=[
            ["W01", 0, 0, 0, "=C{i}+D{i}-E{i}".format(i=r + 1)],
            ["W02", 0, 0, 0, "=C{i}+D{i}-E{i}".format(i=r + 2)],
            ["W03", 0, 0, 0, "=C{i}+D{i}-E{i}".format(i=r + 3)],
            ["W04", 0, 0, 0, "=C{i}+D{i}-E{i}".format(i=r + 4)],
        ],
        num_format="number",
    )

    ws2 = new_sheet(wb, "Data source")
    ws2["B1"].value = "Describe source: dashboard URL, query, pull date"
    ws2["B1"].style = "saas_note"

    out = _TEMPLATES_DIR / "dashboard-skeleton.xlsx"
    save_and_validate(wb, out, recalc=False, raise_on_fail=False)
    return out


# ---------------------------------------------------------------------------
# Engineering sheet skeleton (engineering)
# ---------------------------------------------------------------------------

def build_engineering_sheet() -> Path:
    wb = new_workbook(theme="engineering", first_sheet="Parameters")
    ws = wb["Parameters"]
    set_column_widths(ws, {"A": 2, "B": 32, "C": 10, "D": 14, "E": 14, "F": 14, "G": 14})

    r = write_title_banner(ws, "Equipment Parameter Sheet",
                           subtitle="Tag / service / revision")
    r = write_section_header(ws, "Process parameters", row=r)
    write_table(
        ws, f"B{r}",
        headers=["Parameter", "Unit", "Spec", "Actual", "Δ", "Notes"],
        data=[
            ["Design pressure",   "bar(g)", 35.0, None,  None, None],
            ["Design temperature", "°C",    150.0, None, None, None],
            ["Flow rate",         "kg/h",  12000.0, None, None, None],
            ["Density",           "kg/m³",  815.0, None, None, None],
        ],
        num_format="number_decimal",
    )

    out = _TEMPLATES_DIR / "engineering-sheet-skeleton.xlsx"
    save_and_validate(wb, out, recalc=False, raise_on_fail=False)
    return out


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    _TEMPLATES_DIR.mkdir(exist_ok=True)
    builders = [
        ("financial-model", build_financial_model),
        ("data-report",     build_data_report),
        ("dashboard",       build_dashboard),
        ("engineering",     build_engineering_sheet),
    ]
    for name, fn in builders:
        out = fn()
        print(f"  built {name:15s} -> {out.relative_to(_ROOT)}")


if __name__ == "__main__":
    main()
