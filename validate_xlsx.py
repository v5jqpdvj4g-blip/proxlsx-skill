#!/usr/bin/env python3
"""
validate_xlsx.py — anti-slop style linter for proxlsx-skill (v2).

Loads a saved .xlsx and returns a JSON pass/fail report covering:
  - Sheet-level rules: gridlines hidden (view + print), default sheet
    name removed, tab color set, freeze panes set, A1 non-empty.
  - Cell-level rules: no Calibri-11 default font on content cells,
    numeric cells have an explicit number_format, no rainbow fills,
    strategic (not excessive) borders.
  - Formula health: runs recalc.py as a subprocess and folds its errors
    into the report.
  - Forbidden-term scan: loads FORBIDDEN_TERMS.json and flags any hit
    in cell values or comments.

Usage
-----
    python validate_xlsx.py <file.xlsx>
    → prints a JSON report.

    from validate_xlsx import lint
    report = lint('/tmp/out.xlsx')
"""

from __future__ import annotations

import json
import re
import subprocess
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.cell.cell import Cell
from openpyxl.worksheet.worksheet import Worksheet

_SKILL_DIR = Path(__file__).resolve().parent
_FORBIDDEN_PATH = _SKILL_DIR / "FORBIDDEN_TERMS.json"
_RECALC_PATH = _SKILL_DIR / "recalc.py"

# Cells we consider "content" for font/format checks. A1 and col A (gutter)
# are often intentionally empty or decorative, so we skip them.
_SKIP_COLS = {1}  # column index of A

_DEFAULT_BAD_FONT = ("Calibri", 11.0)
_DEFAULT_SHEET_NAMES = {"Sheet", "Sheet1", "Sheet2", "Sheet3"}

_BORDER_PCT_MAX = 60.0   # above this, borders look like a grid, not structure
_BORDER_PCT_MIN = 0.0    # 0 is fine; we don't require borders
_DISTINCT_FILL_MAX = 8   # rainbow fill threshold


def _load_forbidden() -> tuple[list[str], list[str], bool, bool]:
    if not _FORBIDDEN_PATH.exists():
        return [], [], True, True
    with open(_FORBIDDEN_PATH) as f:
        data = json.load(f)
    return (
        [t for t in data.get("terms", [])],
        [t for t in data.get("soft_terms", [])],
        bool(data.get("case_insensitive", True)),
        bool(data.get("whole_word", True)),
    )


_SML_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def _xfid_to_style_name(path: Path) -> dict[int, str]:
    """
    Map cellStyleXfs index -> NamedStyle name, read straight from styles.xml.

    openpyxl's ``cell.style`` assumes xfId == position in <cellStyles>, which
    only holds for files openpyxl itself wrote. Files re-saved by LibreOffice
    (i.e. every file that went through recalc.py) or Excel carry extra
    cellStyleXfs, so ``cell.style`` returns the wrong name or raises
    IndexError. Resolve names from the real xfId instead.
    """
    try:
        with zipfile.ZipFile(path) as z:
            root = ET.fromstring(z.read("xl/styles.xml"))
    except (KeyError, zipfile.BadZipFile, ET.ParseError):
        return {}
    out: dict[int, str] = {}
    cs = root.find(f"{_SML_NS}cellStyles")
    if cs is not None:
        for el in cs.findall(f"{_SML_NS}cellStyle"):
            try:
                out.setdefault(int(el.get("xfId", "")), el.get("name", ""))
            except ValueError:
                continue
    return out


def _cell_style_name(cell: Cell, xf_names: dict[int, str]) -> str:
    style = getattr(cell, "_style", None)
    xf_id = getattr(style, "xfId", 0) if style is not None else 0
    return xf_names.get(xf_id, "Normal") or "Normal"


def _iter_all_cells(ws: Worksheet):
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None and cell.comment is None:
                continue
            yield cell


def _iter_content_cells(ws: Worksheet):
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            if cell.column in _SKIP_COLS:
                continue
            yield cell


def _is_numeric(cell: Cell) -> bool:
    v = cell.value
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return True
    if isinstance(v, str) and v.startswith("="):
        return True
    return False


def _cell_font_tuple(cell: Cell) -> tuple[str, float]:
    f = cell.font
    return (f.name or "", float(f.sz) if f.sz else 0.0)


def _cell_has_border(cell: Cell) -> bool:
    b = cell.border
    for side in ("left", "right", "top", "bottom"):
        s = getattr(b, side, None)
        if s and s.style:
            return True
    return False


def _cell_fill_argb(cell: Cell) -> str | None:
    f = cell.fill
    if f is None or f.fill_type in (None, "none"):
        return None
    sc = getattr(f, "start_color", None)
    if sc is None:
        return None
    rgb = getattr(sc, "rgb", None)
    if isinstance(rgb, str) and rgb and rgb != "00000000":
        return rgb.upper()
    return None


def _build_forbidden_regex(terms: list[str], case_insensitive: bool, whole_word: bool) -> re.Pattern | None:
    if not terms:
        return None
    parts = []
    for t in terms:
        p = re.escape(t)
        if whole_word:
            # Only demand a boundary on a side that starts/ends with a word
            # character. Terms like "???" or "[INSERT]" must still match when
            # glued to text ("Value???"), which the old (?<=\W)|^ guard missed.
            if re.match(r"\w", t):
                p = r"(?<!\w)" + p
            if re.search(r"\w$", t):
                p = p + r"(?!\w)"
        parts.append(p)
    pat = "|".join(parts)
    flags = re.IGNORECASE if case_insensitive else 0
    return re.compile(pat, flags)


# ---------------------------------------------------------------------------
# Individual checks
# ---------------------------------------------------------------------------

def _check_sheet_names(wb) -> dict:
    bad = [n for n in wb.sheetnames if n in _DEFAULT_SHEET_NAMES]
    return {"pass": not bad, "offenders": bad}


def _check_gridlines(wb) -> dict:
    bad_view, bad_print = [], []
    for ws in wb.worksheets:
        if ws.sheet_view.showGridLines:
            bad_view.append(ws.title)
        if ws.print_options.gridLines:
            bad_print.append(ws.title)
    return {
        "pass": not (bad_view or bad_print),
        "view_visible_on": bad_view,
        "print_visible_on": bad_print,
    }


def _check_tab_colors(wb) -> dict:
    missing = [
        ws.title for ws in wb.worksheets
        if not (ws.sheet_properties.tabColor and getattr(ws.sheet_properties.tabColor, "rgb", None))
    ]
    return {"pass": not missing, "offenders": missing}


def _check_freeze_panes(wb) -> dict:
    missing = [ws.title for ws in wb.worksheets if not ws.freeze_panes]
    return {"pass": not missing, "offenders": missing}


def _check_a1_not_empty(wb) -> dict:
    empty = [ws.title for ws in wb.worksheets if ws["A1"].value is None and ws["B1"].value is None]
    return {"pass": not empty, "offenders": empty}


def _check_default_font(wb, xf_names: dict[int, str]) -> dict:
    """
    Flag Calibri 11 ONLY when the cell uses the built-in 'Normal' style (i.e.
    nothing theme-issued was applied). Calibri 11 inside a custom NamedStyle
    is intentional (e.g. engineering header). The bug we're catching is the
    "user forgot to style the cell" case.
    """
    offenders: list[str] = []
    count = 0
    for ws in wb.worksheets:
        for cell in _iter_content_cells(ws):
            count += 1
            style_name = _cell_style_name(cell, xf_names)
            if style_name != "Normal":
                continue  # trust any non-Normal named style
            name, size = _cell_font_tuple(cell)
            if (name, size) == _DEFAULT_BAD_FONT:
                offenders.append(f"{ws.title}!{cell.coordinate}")
                if len(offenders) > 20:
                    return {"pass": False, "count": len(offenders), "sample": offenders[:20]}
    return {"pass": not offenders, "count": len(offenders), "sample": offenders[:20], "content_cells": count}


def _check_number_format_coverage(wb) -> dict:
    offenders: list[str] = []
    numeric_total = 0
    for ws in wb.worksheets:
        for cell in _iter_content_cells(ws):
            if not _is_numeric(cell):
                continue
            numeric_total += 1
            fmt = cell.number_format or "General"
            if fmt == "General":
                offenders.append(f"{ws.title}!{cell.coordinate}")
                if len(offenders) > 20:
                    return {"pass": False, "numeric_cells": numeric_total,
                            "unformatted": len(offenders), "sample": offenders[:20]}
    return {
        "pass": not offenders,
        "numeric_cells": numeric_total,
        "unformatted": len(offenders),
        "sample": offenders[:20],
    }


def _check_rainbow(wb) -> dict:
    fills: set[str] = set()
    for ws in wb.worksheets:
        for cell in _iter_content_cells(ws):
            argb = _cell_fill_argb(cell)
            if argb:
                fills.add(argb)
    return {
        "pass": len(fills) <= _DISTINCT_FILL_MAX,
        "distinct_fills": len(fills),
        "max_allowed": _DISTINCT_FILL_MAX,
    }


def _check_borders(wb) -> dict:
    bordered = 0
    total = 0
    for ws in wb.worksheets:
        for cell in _iter_content_cells(ws):
            total += 1
            if _cell_has_border(cell):
                bordered += 1
    pct = (bordered / total * 100) if total else 0.0
    return {
        "pass": pct <= _BORDER_PCT_MAX,
        "cells_with_borders_pct": round(pct, 1),
        "bordered": bordered,
        "total": total,
    }


def _check_forbidden_terms(wb) -> dict:
    hard, soft, ci, ww = _load_forbidden()
    hard_re = _build_forbidden_regex(hard, ci, ww)
    soft_re = _build_forbidden_regex(soft, ci, ww)
    hard_hits, soft_hits = [], []

    def scan(where: str, text: str) -> None:
        if hard_re and hard_re.search(text):
            hard_hits.append(f"{where}: {text[:60]}")
        elif soft_re and soft_re.search(text):
            soft_hits.append(f"{where}: {text[:60]}")

    # Scan every cell (column A included), cell comments and sheet names.
    # The style checks skip the gutter column; the term firewall must not.
    for ws in wb.worksheets:
        scan(f"sheet name {ws.title!r}", ws.title)
        for cell in _iter_all_cells(ws):
            if isinstance(cell.value, str):
                scan(f"{ws.title}!{cell.coordinate}", cell.value)
            if cell.comment is not None and cell.comment.text:
                scan(f"{ws.title}!{cell.coordinate} (comment)", cell.comment.text)

    return {
        "pass": not hard_hits,
        "hard_hits": hard_hits[:20],
        "soft_warnings": soft_hits[:20],
        "hard_count": len(hard_hits),
        "soft_count": len(soft_hits),
    }


def _run_recalc(path: Path) -> dict:
    if not _RECALC_PATH.exists():
        return {"skipped": True}
    try:
        res = subprocess.run(
            [sys.executable, str(_RECALC_PATH), str(path)],
            capture_output=True, text=True, timeout=60,
        )
        return json.loads(res.stdout) if res.stdout else {"error": res.stderr}
    except (json.JSONDecodeError, subprocess.TimeoutExpired) as e:
        return {"error": str(e)}


# ---------------------------------------------------------------------------
# Weighted slop score
# ---------------------------------------------------------------------------

_WEIGHTS = {
    "default_font_overridden": 20,
    "gridlines_hidden":        15,
    "number_format_coverage":  15,
    "no_default_sheet_name":   10,
    "freeze_panes_set":        10,
    "tab_colors_set":           8,
    "a1_not_empty":             7,
    "no_rainbow_fills":         7,
    "borders_strategic":        4,
    "forbidden_terms":          4,
}


def _compute_slop_score(checks: dict, forbidden: dict) -> int:
    """Lower is better; 0 = perfect, 100 = worst."""
    score = 0
    for key, weight in _WEIGHTS.items():
        if key == "forbidden_terms":
            passed = forbidden.get("pass", True)
        else:
            passed = checks.get(key, {}).get("pass", True)
        if not passed:
            score += weight
    return score


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def lint(path: str | Path, *, run_recalc: bool = True) -> dict:
    path = Path(path)
    if not path.exists():
        return {"status": "fail", "error": f"no such file: {path}"}

    wb = load_workbook(path)
    xf_names = _xfid_to_style_name(path)

    checks = {
        "gridlines_hidden":        _check_gridlines(wb),
        "no_default_sheet_name":   _check_sheet_names(wb),
        "tab_colors_set":          _check_tab_colors(wb),
        "freeze_panes_set":        _check_freeze_panes(wb),
        "a1_not_empty":            _check_a1_not_empty(wb),
        "default_font_overridden": _check_default_font(wb, xf_names),
        "number_format_coverage":  _check_number_format_coverage(wb),
        "no_rainbow_fills":        _check_rainbow(wb),
        "borders_strategic":       _check_borders(wb),
    }
    forbidden = _check_forbidden_terms(wb)
    recalc = _run_recalc(path) if run_recalc else {"skipped": True}

    slop_score = _compute_slop_score(checks, forbidden)
    style_pass = all(c.get("pass", False) for c in checks.values())
    forbidden_pass = forbidden.get("pass", True)
    # A recalc that could not run (LibreOffice missing, timeout, crash) is not
    # a pass: QA-2 was never checked. Use --no-recalc to skip it explicitly.
    recalc_pass = recalc.get("skipped", False) or recalc.get("status") == "success"

    overall_pass = style_pass and forbidden_pass and recalc_pass

    return {
        "status": "pass" if overall_pass else "fail",
        "slop_score": slop_score,
        "style_checks": checks,
        "forbidden_terms": forbidden,
        "recalc": recalc,
        "summary": {
            "style_pass": style_pass,
            "forbidden_pass": forbidden_pass,
            "recalc_pass": recalc_pass,
        },
    }


def _main() -> int:
    if len(sys.argv) < 2:
        print("Usage: python validate_xlsx.py <file.xlsx> [--no-recalc]", file=sys.stderr)
        return 1
    path = sys.argv[1]
    do_recalc = "--no-recalc" not in sys.argv
    report = lint(path, run_recalc=do_recalc)
    print(json.dumps(report, indent=2))
    return 0 if report["status"] == "pass" else 2


if __name__ == "__main__":
    sys.exit(_main())
