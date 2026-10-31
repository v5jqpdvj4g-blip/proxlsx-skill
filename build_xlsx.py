"""
build_xlsx.py — high-level builder API for proxlsx-skill (anti-slop v2).

All content goes through these helpers so Claude never touches raw
Font()/PatternFill()/Border() and never leaves a cell unstyled. Every
helper pulls the active theme's NamedStyles from theme.py.

Typical usage
-------------
    from build_xlsx import (
        new_workbook, new_sheet, write_title_banner, write_section_header,
        write_table, write_total_row, write_kpi_band, set_column_widths,
        apply_freeze, save_and_validate,
    )

    wb = new_workbook(theme='ib')
    ws = new_sheet(wb, 'Income Statement', theme='ib')
    write_title_banner(ws, 'Project Atlas — 3-Statement Model',
                       subtitle='FY2023–FY2028')
    write_section_header(ws, 'Revenue', row=4)
    headers = ['Line Item', 'FY23A', 'FY24A', 'FY25E', 'FY26E', 'FY27E']
    data = [
        ['Subscription',   1000, 1200, 1440, 1728, '=F6*1.2'],
        ['Services',        200,  220,  242,  266,  293],
    ]
    write_table(ws, 'B5', headers, data, num_format='currency', total=True)
    save_and_validate(wb, '/tmp/atlas.xlsx')
"""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path
from typing import Any, Iterable, Sequence

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from theme import (
    THEMES,
    apply_freeze as _apply_freeze,
    apply_theme,
    default_column_widths,
    get_theme,
    set_column_widths as _set_column_widths,
    set_sheet_layout,
    style_name,
)

_SKILL_DIR = Path(__file__).resolve().parent


# ---------------------------------------------------------------------------
# Workbook / sheet construction
# ---------------------------------------------------------------------------

def new_workbook(theme: str = "ib", first_sheet: str = "Cover") -> Workbook:
    """
    Create a fresh workbook with the theme applied, rename the default sheet
    to `first_sheet`, and set sheet-level layout. Never leaves a 'Sheet' tab.
    """
    if theme not in THEMES:
        raise KeyError(f"unknown theme {theme!r}")
    wb = Workbook()
    apply_theme(wb, theme)
    ws = wb.active
    ws.title = first_sheet
    set_sheet_layout(ws, theme)
    _apply_freeze(ws, theme_name=theme)
    _set_column_widths(ws, default_column_widths(theme, "CDEFGHIJ"))
    wb._xlsx_theme = theme  # stash for validator + downstream helpers
    return wb


def new_sheet(wb: Workbook, name: str, theme: str | None = None) -> Worksheet:
    """
    Add a sheet (or return an existing one) with theme layout applied.
    Tab color, gridlines, zoom, freeze pane, and default col widths are set.
    """
    theme = theme or getattr(wb, "_xlsx_theme", "ib")
    if name in wb.sheetnames:
        ws = wb[name]
    else:
        ws = wb.create_sheet(name)
    set_sheet_layout(ws, theme)
    _apply_freeze(ws, theme_name=theme)
    _set_column_widths(ws, default_column_widths(theme, "CDEFGHIJ"))
    return ws


def load_template(template_name: str, theme: str | None = None) -> Workbook:
    """
    Load a skeleton from templates/{template_name}.xlsx. Re-applies the
    theme on top so styles are guaranteed present even if the template was
    built with an older version.
    """
    path = _SKILL_DIR / "templates" / f"{template_name}.xlsx"
    if not path.exists():
        raise FileNotFoundError(f"no template at {path}")
    wb = load_workbook(path)
    if theme is None:
        # Infer the template's own theme from its registered NamedStyles so
        # e.g. the saas dashboard skeleton isn't silently restyled as ib.
        names = set(wb.named_styles)
        theme = next(
            (t for t in THEMES if style_name(t, "header") in names), "ib"
        )
    apply_theme(wb, theme)
    wb._xlsx_theme = theme
    return wb


# ---------------------------------------------------------------------------
# Writing primitives
# ---------------------------------------------------------------------------

def _theme_of(wb: Workbook) -> str:
    return getattr(wb, "_xlsx_theme", "ib")


def _col_letter(n: int) -> str:
    return get_column_letter(n)


def _merge(ws: Worksheet, start: str, end: str) -> None:
    ws.merge_cells(f"{start}:{end}")


def write_title_banner(
    ws: Worksheet,
    title: str,
    subtitle: str | None = None,
    row: int = 1,
    span: int = 10,
) -> int:
    """
    Write a filled banner row with the title (and optional subtitle below).
    Returns the next free row index.
    """
    theme = _theme_of(ws.parent)
    layout = get_theme(theme)["layout"]

    last_col = _col_letter(1 + span)
    _merge(ws, f"B{row}", f"{last_col}{row}")
    cell = ws[f"B{row}"]
    cell.value = title
    cell.style = style_name(theme, "title")
    ws.row_dimensions[row].height = layout["row_height_title"]

    next_row = row + 1
    if subtitle:
        _merge(ws, f"B{next_row}", f"{last_col}{next_row}")
        sc = ws[f"B{next_row}"]
        sc.value = subtitle
        sc.style = style_name(theme, "subtitle")
        ws.row_dimensions[next_row].height = layout["row_height_default"] + 2
        next_row += 1
    return next_row + 1  # leave one blank row after the banner


def write_section_header(ws: Worksheet, label: str, row: int, span: int = 10) -> int:
    theme = _theme_of(ws.parent)
    layout = get_theme(theme)["layout"]
    last_col = _col_letter(1 + span)
    _merge(ws, f"B{row}", f"{last_col}{row}")
    c = ws[f"B{row}"]
    c.value = label
    c.style = style_name(theme, "section")
    ws.row_dimensions[row].height = layout["row_height_header"] - 2
    return row + 1


def _role_for_value(v: Any) -> str:
    """Guess which data role to apply based on the cell value."""
    if isinstance(v, str) and v.startswith("="):
        return "formula"
    if isinstance(v, (int, float)):
        return "input"
    return "label"


def write_table(
    ws: Worksheet,
    start_cell: str,
    headers: Sequence[str],
    data: Sequence[Sequence[Any]],
    *,
    num_format: str = "currency",
    total: bool = False,
    total_label: str = "Total",
    label_col: bool = True,
    col_formats: dict[str, str] | None = None,
) -> dict:
    """
    Write a headered table at start_cell. Headers get the 'header' style,
    the first column (if label_col=True) gets 'label', and the rest get
    data-role styles ('input' for numbers, 'formula' for '=' strings).

    Number format is looked up in theme['number_formats'][num_format] and
    applied to every numeric/formula cell — the one place number formats
    enter the pipeline.

    `col_formats` overrides `num_format` per column, keyed by header text,
    e.g. {'YoY Δ': 'percent_decimal'} so a growth column next to currency
    columns doesn't render 0.25 as "$0".

    Returns {'header_row', 'first_data_row', 'last_data_row', 'last_col'}.
    """
    theme = _theme_of(ws.parent)
    t = get_theme(theme)
    nfs = t["number_formats"]
    unknown = set(col_formats or {}) - set(headers)
    if unknown:
        raise KeyError(f"col_formats keys not in headers: {sorted(unknown)}")
    col_nf = [nfs[(col_formats or {}).get(h, num_format)] for h in headers]
    layout = t["layout"]

    col0 = ws[start_cell].column          # int
    row0 = ws[start_cell].row             # int (header row)
    n_cols = len(headers)
    last_col_i = col0 + n_cols - 1
    last_col = _col_letter(last_col_i)

    # Header row
    for i, h in enumerate(headers):
        c = ws.cell(row=row0, column=col0 + i, value=h)
        c.style = style_name(theme, "header")
    ws.row_dimensions[row0].height = layout["row_height_header"]

    # Data rows
    first_data = row0 + 1
    for ri, row in enumerate(data):
        r = first_data + ri
        for ci, value in enumerate(row):
            c = ws.cell(row=r, column=col0 + ci, value=value)
            if label_col and ci == 0:
                c.style = style_name(theme, "label")
            else:
                role = _role_for_value(value)
                # data-role styles all default to currency; override the fmt
                c.style = style_name(theme, role)
                c.number_format = col_nf[ci]
    last_data = first_data + len(data) - 1

    # Optional total row: SUM columns, label in first column
    if total and data:
        tr = last_data + 1
        for ci in range(n_cols):
            c = ws.cell(row=tr, column=col0 + ci)
            if label_col and ci == 0:
                c.value = total_label
                c.style = style_name(theme, "total")
                c.number_format = "@"  # text for the label
            else:
                col_letter = _col_letter(col0 + ci)
                c.value = f"=SUM({col_letter}{first_data}:{col_letter}{last_data})"
                c.style = style_name(theme, "total")
                c.number_format = col_nf[ci]
        last_data = tr

    return {
        "header_row": row0,
        "first_data_row": first_data,
        "last_data_row": last_data,
        "last_col": last_col,
    }


def write_total_row(
    ws: Worksheet,
    row: int,
    col_range: tuple[str, str],
    label: str = "Total",
    num_format: str = "currency",
    sum_from: tuple[int, int] | None = None,
) -> None:
    """
    Manually place a total row across col_range=(start,end). If sum_from is
    provided as (first_row, last_row), every numeric column SUMs that range;
    otherwise no formula is inserted.
    """
    theme = _theme_of(ws.parent)
    nf = get_theme(theme)["number_formats"][num_format]
    start_col, end_col = col_range
    from openpyxl.utils import column_index_from_string
    c0 = column_index_from_string(start_col)
    c1 = column_index_from_string(end_col)
    for ci in range(c0, c1 + 1):
        c = ws.cell(row=row, column=ci)
        if ci == c0:
            c.value = label
            c.style = style_name(theme, "total")
            c.number_format = "@"
        else:
            if sum_from:
                letter = _col_letter(ci)
                c.value = f"=SUM({letter}{sum_from[0]}:{letter}{sum_from[1]})"
            c.style = style_name(theme, "total")
            c.number_format = nf


def write_kpi_band(
    ws: Worksheet,
    row: int,
    kpis: Sequence[tuple[str, Any, str]],
    cols_per_kpi: int = 2,
) -> int:
    """
    Write a KPI card band: each kpi is (label, value, num_format_name).
    Cards are placed starting at column B and each occupies `cols_per_kpi`
    columns. Returns next free row.
    """
    theme = _theme_of(ws.parent)
    t = get_theme(theme)
    layout = t["layout"]
    nfs = t["number_formats"]

    col = 2  # B
    for label, value, fmt in kpis:
        label_cell = ws.cell(row=row,     column=col, value=label)
        value_cell = ws.cell(row=row + 1, column=col, value=value)
        label_cell.style = style_name(theme, "note")
        value_cell.style = style_name(theme, "header")
        # headers center; force left for KPI value card feel
        from openpyxl.styles import Alignment
        value_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        value_cell.number_format = nfs.get(fmt, nfs["number"])
        end = col + cols_per_kpi - 1
        if cols_per_kpi > 1:
            _merge(ws, f"{_col_letter(col)}{row}",     f"{_col_letter(end)}{row}")
            _merge(ws, f"{_col_letter(col)}{row + 1}", f"{_col_letter(end)}{row + 1}")
        col = end + 1 + 1  # +1 gutter between cards

    ws.row_dimensions[row].height = layout["row_height_default"]
    ws.row_dimensions[row + 1].height = layout["row_height_header"] + 4
    return row + 3


def write_note(ws: Worksheet, row: int, text: str, span: int = 10) -> int:
    theme = _theme_of(ws.parent)
    last_col = _col_letter(1 + span)
    _merge(ws, f"B{row}", f"{last_col}{row}")
    c = ws[f"B{row}"]
    c.value = text
    c.style = style_name(theme, "note")
    return row + 1


def write_source(ws: Worksheet, row: int, text: str, span: int = 10) -> int:
    theme = _theme_of(ws.parent)
    last_col = _col_letter(1 + span)
    _merge(ws, f"B{row}", f"{last_col}{row}")
    c = ws[f"B{row}"]
    c.value = f"Source: {text}"
    c.style = style_name(theme, "source")
    return row + 1


# ---------------------------------------------------------------------------
# Passthroughs
# ---------------------------------------------------------------------------

def apply_freeze(ws: Worksheet, row: int | None = None, col: str | None = None) -> None:
    _apply_freeze(ws, row=row, col=col, theme_name=_theme_of(ws.parent))


def set_column_widths(ws: Worksheet, widths: dict) -> None:
    _set_column_widths(ws, widths)


# ---------------------------------------------------------------------------
# Save + validate (the one enforcement point)
# ---------------------------------------------------------------------------

def save_and_validate(
    wb: Workbook,
    path: str | Path,
    *,
    recalc: bool = True,
    raise_on_fail: bool = True,
) -> dict:
    """
    Save the workbook, run recalc.py (unless disabled), run validate_xlsx.py,
    and return a combined report. Raises RuntimeError on fail if requested.
    """
    path = Path(path)
    wb.save(path)

    report: dict = {"path": str(path), "recalc": None, "validate": None}

    if recalc:
        res = subprocess.run(
            [sys.executable, str(_SKILL_DIR / "recalc.py"), str(path)],
            capture_output=True, text=True,
        )
        try:
            report["recalc"] = json.loads(res.stdout)
        except json.JSONDecodeError:
            report["recalc"] = {"error": res.stdout + res.stderr}

    validate_script = _SKILL_DIR / "validate_xlsx.py"
    if validate_script.exists():
        # recalc already ran above (or was explicitly disabled): don't let the
        # validator launch LibreOffice a second time.
        res = subprocess.run(
            [sys.executable, str(validate_script), str(path), "--no-recalc"],
            capture_output=True, text=True,
        )
        try:
            report["validate"] = json.loads(res.stdout)
        except json.JSONDecodeError:
            report["validate"] = {"error": res.stdout + res.stderr}
    else:
        report["validate"] = {"error": f"validator missing: {validate_script}"}

    # Fail closed: only an explicit success counts. A crashed validator or a
    # recalc that could not run used to fall through as ok=True.
    ok = True
    if recalc and (report["recalc"] or {}).get("status") != "success":
        ok = False
    if (report["validate"] or {}).get("status") != "pass":
        ok = False
    report["ok"] = ok

    if raise_on_fail and not ok:
        raise RuntimeError(f"save_and_validate failed: {json.dumps(report, indent=2)}")

    return report


# ---------------------------------------------------------------------------
# Smoke test
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    wb = new_workbook(theme="ib", first_sheet="Demo")
    ws = wb.active
    next_row = write_title_banner(ws, "Smoke Test", subtitle="ib theme")
    next_row = write_section_header(ws, "Revenue", row=next_row)
    info = write_table(
        ws, f"B{next_row}",
        headers=["Line Item", "FY23A", "FY24A", "FY25E"],
        data=[
            ["Subscription", 1_000, 1_200, 1_440],
            ["Services",       200,   220,   242],
        ],
        num_format="currency",
        total=True,
    )
    write_source(ws, row=info["last_data_row"] + 2, text="Synthetic smoke data")
    wb.save("/tmp/build_xlsx_smoke.xlsx")
    print(f"wrote /tmp/build_xlsx_smoke.xlsx")
    print(f"  table info: {info}")
