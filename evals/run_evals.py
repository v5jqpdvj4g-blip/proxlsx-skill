#!/usr/bin/env python3
"""
run_evals.py — executable version of evals/evals.json.

evals.json describes prompts plus the expectations a model's output must
meet. This script plays each prompt through the builder API the way the
skill tells Claude to, then checks every expectation that can be verified
mechanically against the saved file. It also builds one sample workbook per
theme so a theme regression shows up as a failed check, not a pretty-but-
broken file.

Usage
-----
    python evals/run_evals.py                 # all evals + theme samples
    python evals/run_evals.py --out DIR       # keep outputs in DIR
    python evals/run_evals.py --no-recalc     # skip LibreOffice (style-only)

Exit code is 0 only when every check passes. A JSON summary is printed.
"""

from __future__ import annotations

import argparse
import json
import sys
import tempfile
import traceback
from datetime import date, timedelta
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_ROOT))

from openpyxl import load_workbook  # noqa: E402
from openpyxl.chart import LineChart, Reference  # noqa: E402

from build_xlsx import (  # noqa: E402
    load_template,
    new_sheet,
    new_workbook,
    save_and_validate,
    write_kpi_band,
    write_note,
    write_section_header,
    write_source,
    write_table,
    write_title_banner,
)
from theme import THEMES, apply_theme, get_theme, style_name  # noqa: E402

RECALC = True


class Checks:
    def __init__(self) -> None:
        self.items: list[dict] = []

    def __call__(self, name: str, ok: bool, detail: object = None) -> None:
        self.items.append({"check": name, "pass": bool(ok), **({"detail": detail} if not ok and detail is not None else {})})

    @property
    def ok(self) -> bool:
        return all(i["pass"] for i in self.items)


def _validated(c: Checks, wb, path: Path, max_slop: int = 10) -> dict:
    rep = save_and_validate(wb, path, recalc=RECALC, raise_on_fail=False)
    v = rep.get("validate") or {}
    failing = {k: x for k, x in (v.get("style_checks") or {}).items() if not x.get("pass")}
    c("save_and_validate ok", rep["ok"], {"recalc": rep.get("recalc"), "failing_checks": failing,
                                         "forbidden": v.get("forbidden_terms", {}).get("hard_hits")})
    c(f"slop_score <= {max_slop}", v.get("slop_score", 999) <= max_slop, v.get("slop_score"))
    if RECALC:
        c("recalc status success", (rep.get("recalc") or {}).get("status") == "success", rep.get("recalc"))
    return rep


def _font_of(cell) -> tuple[str, float]:
    return cell.font.name, float(cell.font.sz or 0)


def _content_cells(ws, *, numeric: bool | None = None):
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None or cell.column == 1:
                continue
            is_num = isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool)
            if numeric is None or numeric == is_num:
                yield cell


# ---------------------------------------------------------------------------
# Eval 1 — SaaS 3-year model (ib)
# ---------------------------------------------------------------------------

def eval_1(out: Path) -> Checks:
    c = Checks()
    wb = new_workbook(theme="ib", first_sheet="Model")
    ws = wb.active
    r = write_title_banner(ws, "SaaS Startup — 3-Year Model", subtitle="USD; Year 1 = first full year")
    r = write_section_header(ws, "Assumptions", row=r)
    a = write_table(ws, f"B{r}", ["Assumption", "Value"], [
        ["Year 1 revenue", 500_000],
        ["Revenue growth", 0.40],
        ["COGS % of revenue", 0.30],
        ["OpEx (flat)", 200_000],
    ], num_format="currency")
    rev0, g, cogs, opex = (f"$C${a['first_data_row'] + k}" for k in range(4))
    for k in (1, 2):
        ws.cell(row=a["first_data_row"] + k, column=3).number_format = get_theme("ib")["number_formats"]["percent"]

    r = write_section_header(ws, "Income statement", row=a["last_data_row"] + 2)
    h = r                 # header row
    i = h + 1             # first line item row
    data = [
        ["Revenue", f"={rev0}", f"=C{i}*(1+{g})", f"=D{i}*(1+{g})"],
        ["COGS", f"=-C{i}*{cogs}", f"=-D{i}*{cogs}", f"=-E{i}*{cogs}"],
        ["Gross profit", f"=C{i}+C{i+1}", f"=D{i}+D{i+1}", f"=E{i}+E{i+1}"],
        ["OpEx", f"=-{opex}", f"=-{opex}", f"=-{opex}"],
        ["EBITDA", f"=C{i+2}+C{i+3}", f"=D{i+2}+D{i+3}", f"=E{i+2}+E{i+3}"],
    ]
    t = write_table(ws, f"B{h}", ["Line item", "Year 1", "Year 2", "Year 3"], data, num_format="currency")

    chart = LineChart()
    chart.title = "Revenue"
    chart.add_data(Reference(ws, min_col=3, max_col=5, min_row=i, max_row=i), from_rows=True, titles_from_data=False)
    chart.set_categories(Reference(ws, min_col=3, max_col=5, min_row=h, max_row=h))
    ws.add_chart(chart, f"B{t['last_data_row'] + 3}")
    write_source(ws, row=t["last_data_row"] + 2, text="Prompt inputs (eval 1)")

    path = out / "saas_model.xlsx"
    _validated(c, wb, path)

    f = load_workbook(path)["Model"]
    c("growth uses absolute assumption ref", f[f"D{i}"].value == f"=C{i}*(1+{g})", f[f"D{i}"].value)
    c("no hardcoded numbers in IS block",
      all(isinstance(f.cell(row=rr, column=cc).value, str) for rr in range(i, i + 5) for cc in range(3, 6)))
    c("Arial 10 body", _font_of(f[f"B{i}"]) == ("Arial", 10.0), _font_of(f[f"B{i}"]))
    c("gridlines hidden", not f.sheet_view.showGridLines)
    c("non-default sheet name", f.title not in {"Sheet", "Sheet1"})
    c("chart present", len(f._charts) == 1, len(f._charts))
    if RECALC:
        v = load_workbook(path, data_only=True)["Model"]
        c("Year 3 revenue = 980,000", abs((v[f"E{i}"].value or 0) - 980_000) < 0.01, v[f"E{i}"].value)
        c("Year 1 EBITDA = 150,000", abs((v[f"C{i+4}"].value or 0) - 150_000) < 0.01, v[f"C{i+4}"].value)
    return c


# ---------------------------------------------------------------------------
# Eval 2 — pandas analysis of a sales file
# ---------------------------------------------------------------------------

def eval_2(out: Path) -> Checks:
    import pandas as pd

    c = Checks()
    rows = []
    regions, products = ["AMER", "EMEA", "APAC"], [f"P{n}" for n in range(1, 8)]
    d0 = date(2025, 1, 1)
    for n in range(240):
        rows.append({
            "Date": d0 + timedelta(days=(n * 3) % 365),
            "Region": regions[n % 3],
            "Product": products[(n * 5) % 7],
            "Units": 1 + (n * 7) % 13,
            "UnitPrice": 10.0 + (n % 4) * 2.5,
        })
    src = out / "sales_2025.xlsx"
    pd.DataFrame(rows).to_excel(src, index=False)

    df = pd.read_excel(src)
    df["Revenue"] = df["Units"] * df["UnitPrice"]
    by_region = df.groupby("Region")["Revenue"].sum()
    top5 = df.groupby("Product")["Units"].sum().nlargest(5)
    monthly = df.set_index(pd.to_datetime(df["Date"])).resample("MS")["Revenue"].sum()

    exp_region: dict[str, float] = {}
    exp_units: dict[str, int] = {}
    for x in rows:
        exp_region[x["Region"]] = exp_region.get(x["Region"], 0) + x["Units"] * x["UnitPrice"]
        exp_units[x["Product"]] = exp_units.get(x["Product"], 0) + x["Units"]
    c("revenue by region matches", all(abs(by_region[k] - v) < 1e-6 for k, v in exp_region.items()),
      by_region.to_dict())
    c("top-5 products by units", list(top5.index) == sorted(exp_units, key=lambda k: (-exp_units[k], k))[:5]
      or set(top5.values) == set(sorted(exp_units.values(), reverse=True)[:5]), top5.to_dict())
    c("monthly trend covers 12 months", len(monthly) == 12, len(monthly))
    c("monthly total reconciles", abs(monthly.sum() - df["Revenue"].sum()) < 1e-6)
    return c


# ---------------------------------------------------------------------------
# Eval 3 — add Assumptions to an existing workbook
# ---------------------------------------------------------------------------

def eval_3(out: Path) -> Checks:
    c = Checks()
    path = out / "existing_model.xlsx"
    wb = new_workbook(theme="ib", first_sheet="Income Statement")
    ws = wb.active
    r = write_title_banner(ws, "Existing model")
    write_table(ws, f"B{r}", ["Line item", "FY24", "FY25"], [["Revenue", 100, 120], ["Costs", 60, 70]], total=True)
    _validated(c, wb, path)
    before = [[cell.value for cell in row] for row in load_workbook(path)["Income Statement"].iter_rows()]

    wb = load_workbook(path)                     # not data_only
    apply_theme(wb, "ib")
    wb._xlsx_theme = "ib"
    ws = new_sheet(wb, "Assumptions")
    r = write_title_banner(ws, "Assumptions")
    t = write_table(ws, f"B{r}", ["Parameter", "Value"], [
        ["Revenue growth rate", 0.15],
        ["Discount rate", 0.10],
        ["Tax rate", 0.25],
    ], num_format="percent")
    _validated(c, wb, path)

    wb2 = load_workbook(path)
    c("Assumptions sheet added", "Assumptions" in wb2.sheetnames, wb2.sheetnames)
    after = [[cell.value for cell in row] for row in wb2["Income Statement"].iter_rows()]
    c("existing sheet unchanged", before == after)
    v = wb2["Assumptions"].cell(row=t["first_data_row"], column=3)
    c("input cell is Macabacus blue", (v.font.color.rgb or "").upper() == "FF4472C4", v.font.color.rgb)
    c("input cell right-aligned", v.alignment.horizontal == "right")
    c("input cell percent format", "%" in v.number_format, v.number_format)
    return c


# ---------------------------------------------------------------------------
# Eval 4 — SaaS dashboard
# ---------------------------------------------------------------------------

def eval_4(out: Path) -> Checks:
    c = Checks()
    wb = new_workbook(theme="saas", first_sheet="Dashboard")
    ws = wb.active
    r = write_title_banner(ws, "Product Metrics", subtitle="Weekly, trailing 4 weeks")
    r = write_kpi_band(ws, r, [
        ("MRR", 182_400, "currency"),
        ("Active users", 12_480, "number"),
        ("Churn", 0.021, "percent_decimal"),
        ("NPS", 47, "number"),
    ])
    r = write_section_header(ws, "Weekly trend", row=r)
    weeks = [("W01", 410, 95), ("W02", 455, 102), ("W03", 470, 88), ("W04", 512, 91)]
    t = write_table(ws, f"B{r}", ["Week", "New", "Churned", "Net"],
                    [[w, n, ch, f"=C{r + 1 + k}-D{r + 1 + k}"] for k, (w, n, ch) in enumerate(weeks)],
                    num_format="number", total=True)
    path = out / "saas_dash.xlsx"
    _validated(c, wb, path)

    f = load_workbook(path)["Dashboard"]
    pal = get_theme("saas")["palette"]
    c("body font is Aptos", _font_of(f.cell(row=t["first_data_row"], column=2))[0] == "Aptos")
    c("tab color is indigo", (f.sheet_properties.tabColor.rgb or "").upper() == pal["accent"])
    c("header fill is slate ink", f.cell(row=t["header_row"], column=2).fill.start_color.rgb == pal["header_bg"])
    c("net column is a formula", str(f.cell(row=t["first_data_row"], column=5).value).startswith("="))
    if RECALC:
        v = load_workbook(path, data_only=True)["Dashboard"]
        c("net total = 1471", v.cell(row=t["last_data_row"], column=5).value == 1471,
          v.cell(row=t["last_data_row"], column=5).value)
    return c


# ---------------------------------------------------------------------------
# Eval 5 — fill the financial-model skeleton
# ---------------------------------------------------------------------------

def eval_5(out: Path) -> Checks:
    c = Checks()
    tpl = _ROOT / "templates" / "financial-model-skeleton.xlsx"
    orig = load_workbook(tpl)
    wb = load_template("financial-model-skeleton", theme="ib")
    ws = wb["Assumptions"]
    ws["C6"] = 15.0
    for col, v in zip("DEFG", (0.45, 0.35, 0.30, 0.25)):
        ws[f"{col}7"] = v
    for col, v in zip("CDEFG", (0.72, 0.74, 0.75, 0.76, 0.77)):
        ws[f"{col}8"] = v
    for col, v in zip("CDEFG", (0.65, 0.58, 0.52, 0.48, 0.45)):
        ws[f"{col}9"] = v
    path = out / "3stmt.xlsx"
    _validated(c, wb, path)

    f = load_workbook(path)
    c("no new sheets", f.sheetnames == orig.sheetnames, f.sheetnames)
    a = f["Assumptions"]
    c("assumptions are numbers", all(isinstance(a[f"{col}8"].value, float) for col in "CDEFG"))
    c("starting ARR is not a percent format", "%" not in a["C6"].number_format, a["C6"].number_format)
    for name in orig.sheetnames:
        o, n = orig[name], f[name]
        c(f"{name}: tab color kept", (o.sheet_properties.tabColor.rgb == n.sheet_properties.tabColor.rgb))
        c(f"{name}: freeze kept", o.freeze_panes == n.freeze_panes, (o.freeze_panes, n.freeze_panes))
        c(f"{name}: banner kept", o["B1"].value == n["B1"].value)
        if name != "Assumptions":
            same = all(o.cell(row=x.row, column=x.column).value == x.value for row in n.iter_rows() for x in row)
            c(f"{name}: untouched", same)
    return c


# ---------------------------------------------------------------------------
# Eval 6 — engineering parameter sheet
# ---------------------------------------------------------------------------

def eval_6(out: Path) -> Checks:
    c = Checks()
    wb = new_workbook(theme="engineering", first_sheet="E-101")
    ws = wb.active
    r = write_title_banner(ws, "Heat Exchanger E-101", subtitle="Design basis, Rev A")
    r = write_section_header(ws, "Design parameters", row=r)
    rows = [("Design pressure", "bar(g)", 35, 34.6), ("Design temperature", "°C", 150, 148.5),
            ("Flow rate", "kg/h", 12000, 11850)]
    t = write_table(ws, f"B{r}", ["Parameter", "Unit", "Spec", "Actual", "Delta"],
                    [[p, u, s, a, f"=E{r + 1 + k}-D{r + 1 + k}"] for k, (p, u, s, a) in enumerate(rows)],
                    num_format="number_decimal")
    write_note(ws, t["last_data_row"] + 2, "Delta = Actual - Spec. Units are SI except gauge pressure.")
    path = out / "e101.xlsx"
    _validated(c, wb, path)

    f = load_workbook(path)["E-101"]
    d = t["first_data_row"]
    c("body font Consolas", _font_of(f.cell(row=d, column=4))[0] == "Consolas")
    c("header Calibri 11 bold", _font_of(f.cell(row=t["header_row"], column=2)) == ("Calibri", 11.0)
      and f.cell(row=t["header_row"], column=2).font.b)
    c("Unit column between label and Spec", f.cell(row=t["header_row"], column=3).value == "Unit")
    c("Delta formula", f.cell(row=d, column=6).value == f"=E{d}-D{d}", f.cell(row=d, column=6).value)
    return c


# ---------------------------------------------------------------------------
# Theme samples — one workbook per theme, every helper exercised
# ---------------------------------------------------------------------------

def theme_sample(theme: str, out: Path) -> Checks:
    c = Checks()
    nfs = get_theme(theme)["number_formats"]
    wb = new_workbook(theme=theme, first_sheet="Summary")
    ws = wb.active
    r = write_title_banner(ws, f"Theme sample — {THEMES[theme]['meta']['label']}", subtitle="Generated by evals")
    r = write_kpi_band(ws, r, [("Revenue", 1_250_000, "currency"), ("Margin", 0.243, "percent_decimal"),
                               ("Multiple", 7.5, "multiple")])
    r = write_section_header(ws, "Segments", row=r)
    t = write_table(ws, f"B{r}", ["Segment", "FY24", "FY25", "Growth"], [
        ["Alpha", 400, 460, f"=D{r + 1}/C{r + 1}-1"],
        ["Beta", 300, 330, f"=D{r + 2}/C{r + 2}-1"],
    ], num_format="currency", col_formats={"Growth": "percent_decimal"}, total=True)
    tr = t["last_data_row"]
    ws[f"E{tr}"].value = f"=D{tr}/C{tr}-1"
    write_note(ws, tr + 2, "Growth on the total row is computed from totals, not summed.")
    write_source(ws, tr + 3, "Synthetic data for theme regression")
    ws2 = new_sheet(wb, "Detail")
    r2 = write_title_banner(ws2, "Detail")
    write_table(ws2, f"B{r2}", ["Item", "Value"], [["Units", 1200], ["Price", 9.5]], num_format="number_decimal")

    path = out / f"sample_{theme}.xlsx"
    _validated(c, wb, path, max_slop=0)

    f = load_workbook(path)
    body = get_theme(theme)["fonts"]["body"]
    s = f["Summary"]
    c("body font", _font_of(s.cell(row=t["first_data_row"], column=2)) == (body["name"], float(body["size"])),
      _font_of(s.cell(row=t["first_data_row"], column=2)))
    c("growth column percent format", s.cell(row=t["first_data_row"], column=5).number_format == nfs["percent_decimal"]
      or "%" in s.cell(row=t["first_data_row"], column=5).number_format,
      s.cell(row=t["first_data_row"], column=5).number_format)
    c("all theme styles registered", all(style_name(theme, x) in f.named_styles for x in ("title", "header", "total")))
    if RECALC:
        v = load_workbook(path, data_only=True)["Summary"]
        c("total growth = 790/700-1", abs((v[f"E{tr}"].value or 0) - (790 / 700 - 1)) < 1e-9, v[f"E{tr}"].value)
    return c


# ---------------------------------------------------------------------------

def main() -> int:
    global RECALC
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", type=Path, default=None)
    ap.add_argument("--no-recalc", action="store_true")
    args = ap.parse_args()
    RECALC = not args.no_recalc

    spec = json.loads((_ROOT / "evals" / "evals.json").read_text())
    runners = {1: eval_1, 2: eval_2, 3: eval_3, 4: eval_4, 5: eval_5, 6: eval_6}
    missing = {e["id"] for e in spec["evals"]} - set(runners)

    out = args.out or Path(tempfile.mkdtemp(prefix="proxlsx-evals-"))
    out.mkdir(parents=True, exist_ok=True)

    jobs = [(f"eval {e['id']}", (lambda fn: lambda: fn(out))(runners[e["id"]]))
            for e in spec["evals"] if e["id"] in runners]
    jobs += [(f"theme {t}", (lambda t: lambda: theme_sample(t, out))(t)) for t in THEMES]

    results, all_ok = [], not missing
    for name, job in jobs:
        try:
            c = job()
            res = {"name": name, "pass": c.ok, "checks": len(c.items),
                   "failures": [i for i in c.items if not i["pass"]]}
        except Exception:
            res = {"name": name, "pass": False, "error": traceback.format_exc()}
        all_ok &= res["pass"]
        results.append(res)
        print(f"{'PASS' if res['pass'] else 'FAIL'}  {name}", file=sys.stderr)

    print(json.dumps({"status": "pass" if all_ok else "fail", "recalc": RECALC, "out_dir": str(out),
                      "evals_without_runner": sorted(missing), "results": results}, indent=2, default=str))
    return 0 if all_ok else 1


if __name__ == "__main__":
    sys.exit(main())
