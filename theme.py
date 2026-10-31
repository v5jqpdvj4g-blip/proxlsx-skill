"""
theme.py — locked design system for proxlsx-skill (anti-slop v2).

Four themes register NamedStyles on a workbook so Claude never has to build
Font/Fill/Border objects by hand. Use via build_xlsx.py; do not edit cells
directly with raw openpyxl constructors.

Themes
------
  ib           — IB / Wall Street (default). Arial 10, Macabacus palette.
  engineering  — EPC/HYSYS parameter sheets. Consolas data, tight borders.
  korean_corp  — HD-style corporate. Pretendard body, brand accent.
  saas         — Modern SaaS dashboard. Aptos/Inter, slate palette.

Public API
----------
  THEMES                         dict[str, Theme]
  get_theme(name='ib') -> Theme
  apply_theme(wb, name='ib')     registers named styles + sets normal font
  set_sheet_layout(ws, theme)    gridlines/zoom/margins/tab color/print
  apply_freeze(ws, row, col)     freeze pane shortcut
  set_column_widths(ws, widths)  {'A': 2, 'B': 28, ...}
  format_range(ws, range, name)  apply named style to a range
  style_name(theme, role)        'ib_header' etc.

Color format: all hex strings are 'AARRGGBB' with AA=FF (opaque). Never
pass 6-char hex to openpyxl; it breaks silently on some versions.
"""

from __future__ import annotations

from typing import Iterable

from openpyxl.styles import (
    Alignment,
    Border,
    Font,
    NamedStyle,
    PatternFill,
    Side,
)
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Theme dictionaries (single source of truth)
# ---------------------------------------------------------------------------

THEMES: dict = {
    "ib": {
        "meta": {
            "label": "IB / Wall Street",
            "feel": "Macabacus/BIWS analyst exhibit",
            "default": True,
        },
        "fonts": {
            "body":   {"name": "Arial", "size": 10, "bold": False, "color": "FF000000"},
            "header": {"name": "Arial", "size": 11, "bold": True,  "color": "FFFFFFFF"},
            "title":  {"name": "Arial", "size": 14, "bold": True,  "color": "FFFFFFFF"},
            "note":   {"name": "Arial", "size": 8,  "bold": False, "color": "FF595959"},
        },
        "palette": {
            "ink":        "FF000000",
            "paper":      "FFFFFFFF",
            "input":      "FF4472C4",  # Macabacus blue
            "formula":    "FF000000",
            "link":       "FF70AD47",  # cross-sheet
            "external":   "FFC00000",  # external-file
            "header_bg":  "FF404040",
            "section_bg": "FFD9D9D9",
            "total_bg":   "FFF2F2F2",
            "rule":       "FF000000",
            "accent":     "FF1F4E78",
            "warn":       "FFFFC000",
        },
        "number_formats": {
            "currency":         '$#,##0;($#,##0);"-"',
            "currency_decimal": '$#,##0.0;($#,##0.0);"-"',
            "number":           '#,##0;(#,##0);"-"',
            "number_decimal":   '#,##0.0;(#,##0.0);"-"',
            "percent":          '0%;(0%);"-"',
            "percent_decimal":  '0.0%;(0.0%);"-"',
            "multiple":         '0.0"x"',
            "date":             "mm/dd/yyyy",
            "year_text":        "0",
            "basis_points":     '#,##0" bps"',
            "zero_dash":        '#,##0;(#,##0);"-"',
        },
        "layout": {
            "row_height_default": 15,
            "row_height_header":  20,
            "row_height_title":   26,
            "col_width_gutter":   2,
            "col_width_label":    28,
            "col_width_data":     13,
            "zoom":               100,
            "margin_top":    0.5,
            "margin_bottom": 0.5,
            "margin_left":   0.5,
            "margin_right":  0.5,
        },
        "rules": {
            "gridlines_visible": False,
            "gridlines_print":   False,
            "freeze":            {"row": 2, "col": "B"},
            "tab_color":         "FF1F4E78",
            "orientation":       "landscape",
        },
    },

    "engineering": {
        "meta": {
            "label": "Engineering / Technical",
            "feel": "EPC calculation sheet",
            "default": False,
        },
        "fonts": {
            "body":   {"name": "Consolas",  "size": 10, "bold": False, "color": "FF1A1A1A"},
            "header": {"name": "Calibri",   "size": 11, "bold": True,  "color": "FF1A1A1A"},
            "title":  {"name": "Calibri",   "size": 14, "bold": True,  "color": "FFFFFFFF"},
            "note":   {"name": "Calibri",   "size": 8,  "bold": False, "color": "FF595959"},
        },
        "palette": {
            "ink":        "FF1A1A1A",
            "paper":      "FFFFFFFF",
            "input":      "FF0066CC",
            "formula":    "FF1A1A1A",
            "link":       "FF008233",
            "external":   "FFC8102E",
            "header_bg":  "FF2F5496",
            "section_bg": "FFDCE6F1",
            "total_bg":   "FFEFEFEF",
            "rule":       "FF888888",
            "accent":     "FF2F5496",
            "warn":       "FFC8102E",
        },
        "number_formats": {
            "currency":         '#,##0;(#,##0);"-"',
            "currency_decimal": '#,##0.00;(#,##0.00);"-"',
            "number":           '#,##0;(#,##0);"-"',
            "number_decimal":   '#,##0.000;(#,##0.000);"-"',
            "percent":          '0.0%;(0.0%);"-"',
            "percent_decimal":  '0.00%;(0.00%);"-"',
            "multiple":         '0.000',
            "date":             "yyyy-mm-dd",
            "year_text":        "0",
            "scientific":       '0.000E+00',
            "zero_dash":        '#,##0;(#,##0);"-"',
        },
        "layout": {
            "row_height_default": 14,
            "row_height_header":  18,
            "row_height_title":   24,
            "col_width_gutter":   2,
            "col_width_label":    32,
            "col_width_data":     14,
            "zoom":               100,
            "margin_top":    0.4,
            "margin_bottom": 0.4,
            "margin_left":   0.4,
            "margin_right":  0.4,
        },
        "rules": {
            "gridlines_visible": False,
            "gridlines_print":   False,
            "freeze":            {"row": 3, "col": "C"},
            "tab_color":         "FF2F5496",
            "orientation":       "landscape",
        },
    },

    "korean_corp": {
        "meta": {
            "label": "Korean Corporate (HD-like)",
            "feel": "국내 대기업 결재문서",
            "default": False,
        },
        "fonts": {
            "body":   {"name": "Pretendard",  "size": 10, "bold": False, "color": "FF1A1A1A"},
            "header": {"name": "Pretendard",  "size": 11, "bold": True,  "color": "FFFFFFFF"},
            "title":  {"name": "HD현대체 Bold", "size": 14, "bold": False, "color": "FFFFFFFF"},
            "note":   {"name": "Pretendard",  "size": 9,  "bold": False, "color": "FF595959"},
        },
        "palette": {
            "ink":        "FF1A1A1A",
            "paper":      "FFFFFFFF",
            "input":      "FF003087",  # Discovery Blue
            "formula":    "FF1A1A1A",
            "link":       "FF00AD1D",  # Heritage Green
            "external":   "FFC00000",
            "header_bg":  "FF003087",
            "section_bg": "FFE6ECF3",
            "total_bg":   "FFF5F7FA",
            "rule":       "FF595959",
            "accent":     "FF00AD1D",
            "warn":       "FFFFC000",
        },
        "number_formats": {
            "currency":         '#,##0;(#,##0);"-"',
            "currency_decimal": '#,##0.0;(#,##0.0);"-"',
            "number":           '#,##0;(#,##0);"-"',
            "number_decimal":   '#,##0.0;(#,##0.0);"-"',
            "percent":          '0%;(0%);"-"',
            "percent_decimal":  '0.0%;(0.0%);"-"',
            "multiple":         '0.0"배"',
            "date":             "yyyy-mm-dd",
            "year_text":        "0",
            "won":              '₩#,##0;(₩#,##0);"-"',
            "zero_dash":        '#,##0;(#,##0);"-"',
        },
        "layout": {
            "row_height_default": 17,
            "row_height_header":  22,
            "row_height_title":   28,
            "col_width_gutter":   2,
            "col_width_label":    30,
            "col_width_data":     14,
            "zoom":               100,
            "margin_top":    0.6,
            "margin_bottom": 0.6,
            "margin_left":   0.5,
            "margin_right":  0.5,
        },
        "rules": {
            "gridlines_visible": False,
            "gridlines_print":   False,
            "freeze":            {"row": 3, "col": "B"},
            "tab_color":         "FF003087",
            "orientation":       "landscape",
        },
    },

    "saas": {
        "meta": {
            "label": "Modern SaaS dashboard",
            "feel": "Notion/Linear dashboard export",
            "default": False,
        },
        "fonts": {
            "body":   {"name": "Aptos",        "size": 10, "bold": False, "color": "FF0F172A"},
            "header": {"name": "Aptos Display","size": 12, "bold": True,  "color": "FFFFFFFF"},
            "title":  {"name": "Aptos Display","size": 16, "bold": True,  "color": "FFFFFFFF"},
            "note":   {"name": "Aptos",        "size": 9,  "bold": False, "color": "FF64748B"},
        },
        "palette": {
            "ink":        "FF0F172A",
            "paper":      "FFFFFFFF",
            "input":      "FF6366F1",  # indigo
            "formula":    "FF0F172A",
            "link":       "FF10B981",  # emerald
            "external":   "FFEF4444",  # red
            "header_bg":  "FF0F172A",
            "section_bg": "FFF1F5F9",
            "total_bg":   "FFE2E8F0",
            "rule":       "FFCBD5E1",
            "accent":     "FF6366F1",
            "warn":       "FFF59E0B",
        },
        "number_formats": {
            "currency":         '$#,##0;($#,##0);"-"',
            "currency_decimal": '$#,##0.0;($#,##0.0);"-"',
            "number":           '#,##0;(#,##0);"-"',
            "number_decimal":   '#,##0.0;(#,##0.0);"-"',
            "percent":          '0%;(0%);"-"',
            "percent_decimal":  '0.0%;(0.0%);"-"',
            "multiple":         '0.00"x"',
            "date":             "mmm d, yyyy",
            "year_text":        "0",
            "zero_dash":        '#,##0;(#,##0);"-"',
        },
        "layout": {
            "row_height_default": 18,
            "row_height_header":  24,
            "row_height_title":   32,
            "col_width_gutter":   2,
            "col_width_label":    30,
            "col_width_data":     14,
            "zoom":               100,
            "margin_top":    0.5,
            "margin_bottom": 0.5,
            "margin_left":   0.5,
            "margin_right":  0.5,
        },
        "rules": {
            "gridlines_visible": False,
            "gridlines_print":   False,
            "freeze":            {"row": 3, "col": "B"},
            "tab_color":         "FF6366F1",
            "orientation":       "landscape",
        },
    },
}


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def get_theme(name: str = "ib") -> dict:
    """Return the theme dict or raise KeyError."""
    if name not in THEMES:
        raise KeyError(f"unknown theme {name!r}; pick one of {list(THEMES)}")
    return THEMES[name]


def style_name(theme: str, role: str) -> str:
    """Canonical NamedStyle name for (theme, role)."""
    return f"{theme}_{role}"


# Roles are the canonical set of NamedStyles registered per theme.
ROLES: tuple = (
    "title",
    "subtitle",
    "header",
    "section",
    "label",
    "input",
    "formula",
    "link",
    "external",
    "subtotal",
    "total",
    "note",
    "source",
)


def _font(spec: dict) -> Font:
    return Font(
        name=spec["name"],
        size=spec["size"],
        bold=spec.get("bold", False),
        italic=spec.get("italic", False),
        color=spec["color"],
    )


def _fill(argb: str) -> PatternFill:
    return PatternFill(fill_type="solid", start_color=argb, end_color=argb)


def _side(style: str = "thin", color: str = "FF000000") -> Side:
    return Side(style=style, color=color)


def _border(**kwargs) -> Border:
    return Border(**{k: _side(**v) if isinstance(v, dict) else v for k, v in kwargs.items()})


def _align(horizontal: str = "left", vertical: str = "center", indent: int = 0, wrap: bool = False) -> Alignment:
    return Alignment(horizontal=horizontal, vertical=vertical, indent=indent, wrap_text=wrap)


# ---------------------------------------------------------------------------
# Named-style factory
# ---------------------------------------------------------------------------

def _build_named_styles(theme_name: str) -> list:
    """Return a list of NamedStyle objects for the given theme."""
    t = get_theme(theme_name)
    f = t["fonts"]
    p = t["palette"]
    nf = t["number_formats"]
    rule = p["rule"]

    styles: list = []

    def mk(role: str, **kw) -> NamedStyle:
        ns = NamedStyle(name=style_name(theme_name, role))
        ns.font = kw.get("font", _font(f["body"]))
        if "fill" in kw:
            ns.fill = kw["fill"]
        ns.alignment = kw.get("align", _align())
        if "border" in kw:
            ns.border = kw["border"]
        if "number_format" in kw:
            ns.number_format = kw["number_format"]
        styles.append(ns)
        return ns

    # Banner / titles
    mk("title",
       font=_font(f["title"]),
       fill=_fill(p["header_bg"]),
       align=_align("left", "center", indent=1))

    mk("subtitle",
       font=_font({**f["note"], "color": "FFFFFFFF"}),
       fill=_fill(p["header_bg"]),
       align=_align("left", "center", indent=1))

    # Column headers
    mk("header",
       font=_font(f["header"]),
       fill=_fill(p["header_bg"]),
       align=_align("center", "center"),
       border=Border(top=_side("thin", rule), bottom=_side("thin", rule)))

    # Section band (gray stripe under a section title)
    mk("section",
       font=_font({**f["body"], "bold": True, "color": p["ink"]}),
       fill=_fill(p["section_bg"]),
       align=_align("left", "center", indent=1),
       border=Border(top=_side("thin", rule), bottom=_side("thin", rule)))

    # Row label (first col in a data table)
    mk("label",
       font=_font(f["body"]),
       align=_align("left", "center", indent=1))

    # Data-role named styles: input (blue), formula (black), link (green), external (red)
    def data(role: str, color: str, bold: bool = False) -> None:
        mk(role,
           font=_font({**f["body"], "color": color, "bold": bold}),
           align=_align("right", "center"),
           number_format=nf["currency"])

    data("input",    p["input"])
    data("formula",  p["formula"])
    data("link",     p["link"])
    data("external", p["external"])

    # Subtotals & totals
    mk("subtotal",
       font=_font({**f["body"], "bold": True, "color": p["ink"]}),
       align=_align("right", "center"),
       border=Border(top=_side("thin", rule)),
       number_format=nf["currency"])

    mk("total",
       font=_font({**f["body"], "bold": True, "color": p["ink"]}),
       fill=_fill(p["total_bg"]),
       align=_align("right", "center"),
       border=Border(top=_side("thin", rule), bottom=_side("double", rule)),
       number_format=nf["currency"])

    # Notes & sources
    mk("note",
       font=_font(f["note"]),
       align=_align("left", "top", wrap=True))

    mk("source",
       font=_font({**f["note"], "italic": True}),
       align=_align("left", "top", wrap=True))

    return styles


def apply_theme(wb, name: str = "ib") -> None:
    """
    Register all NamedStyles for the theme on the workbook and set the theme
    body font on the built-in 'Normal' style definition. Call once per
    workbook right after Workbook() / load_workbook().

    Note: this does NOT restyle cells that never get a style assigned; in
    openpyxl those keep the workbook default font (Calibri 11). That is why
    every write must go through build_xlsx helpers, and why validate_xlsx.py
    flags Calibri 11 on Normal-styled content cells.
    """
    get_theme(name)  # raises if unknown
    registered = {s.name for s in wb._named_styles}

    # Override 'Normal' so the default cell font is never Calibri 11.
    body = get_theme(name)["fonts"]["body"]
    normal_font = _font(body)
    if "Normal" in registered:
        wb._named_styles["Normal"].font = normal_font
    else:
        normal = NamedStyle(name="Normal")
        normal.font = normal_font
        wb.add_named_style(normal)

    for ns in _build_named_styles(name):
        if ns.name not in registered:
            wb.add_named_style(ns)


# ---------------------------------------------------------------------------
# Sheet-level helpers
# ---------------------------------------------------------------------------

def set_sheet_layout(ws, theme_name: str = "ib") -> None:
    """
    Apply theme's sheet-level rules: hide gridlines, set zoom, margins,
    tab color, orientation. Does not touch content.
    """
    t = get_theme(theme_name)
    rules = t["rules"]
    layout = t["layout"]

    ws.sheet_view.showGridLines = rules["gridlines_visible"]
    ws.print_options.gridLines = rules["gridlines_print"]
    ws.sheet_view.zoomScale = layout["zoom"]
    ws.sheet_view.zoomScaleNormal = layout["zoom"]

    ws.sheet_format.defaultRowHeight = layout["row_height_default"]

    ws.page_margins.top = layout["margin_top"]
    ws.page_margins.bottom = layout["margin_bottom"]
    ws.page_margins.left = layout["margin_left"]
    ws.page_margins.right = layout["margin_right"]

    ws.page_setup.orientation = rules["orientation"]
    ws.sheet_properties.tabColor = rules["tab_color"]


def apply_freeze(ws, row: int | None = None, col: str | None = None, theme_name: str = "ib") -> None:
    """Freeze pane at the given row/col or the theme default."""
    t = get_theme(theme_name)
    fr = t["rules"]["freeze"]
    r = row if row is not None else fr["row"]
    c = col if col is not None else fr["col"]
    ws.freeze_panes = f"{c}{r}"


def set_column_widths(ws, widths: dict) -> None:
    """widths = {'A': 2, 'B': 28, 'C': 13, ...}"""
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def format_range(ws, cell_range: str, style_name: str) -> None:
    """Apply a NamedStyle to every cell in a range like 'B2:G10'."""
    for row in ws[cell_range]:
        for cell in row:
            cell.style = style_name


def default_column_widths(theme_name: str, data_cols: Iterable[str]) -> dict:
    """
    Build a widths dict: gutter col A, label col B, data cols C.. from the
    theme's layout. `data_cols` is an iterable of column letters.
    """
    t = get_theme(theme_name)["layout"]
    widths: dict = {"A": t["col_width_gutter"], "B": t["col_width_label"]}
    for letter in data_cols:
        widths[letter] = t["col_width_data"]
    return widths


# ---------------------------------------------------------------------------
# Smoke test
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    print(f"themes registered: {len(THEMES)}")
    for name, t in THEMES.items():
        print(f"  {name:12s} — {t['meta']['label']}")
    # Build styles for each theme to catch bad dict entries early.
    for name in THEMES:
        styles = _build_named_styles(name)
        print(f"  {name}: {len(styles)} named styles")
