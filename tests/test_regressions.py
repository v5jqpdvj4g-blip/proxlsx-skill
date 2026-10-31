"""Regression tests for the bugs fixed alongside the CI workflow."""

from __future__ import annotations

import os
import shutil
import sys
from pathlib import Path

import pytest
from openpyxl.comments import Comment

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import recalc  # noqa: E402
from build_xlsx import load_template, new_workbook, save_and_validate, write_table, write_title_banner  # noqa: E402
from validate_xlsx import lint  # noqa: E402

needs_soffice = pytest.mark.skipif(not shutil.which("soffice"), reason="LibreOffice not installed")


def _clean_wb(theme="ib"):
    wb = new_workbook(theme=theme, first_sheet="Demo")
    ws = wb.active
    r = write_title_banner(ws, "Regression")
    write_table(ws, f"B{r}", ["Item", "A", "B"], [["x", 1, f"=C{r + 1}*2"]], total=True)
    return wb, ws


def _hits(tmp_path, mutate):
    wb, ws = _clean_wb()
    mutate(wb, ws)
    p = tmp_path / "f.xlsx"
    wb.save(p)
    return lint(p, run_recalc=False)["forbidden_terms"]


@pytest.mark.parametrize("mutate", [
    lambda wb, ws: ws.cell(row=20, column=1, value="TBD"),                       # gutter column A
    lambda wb, ws: setattr(ws["C20"], "comment", Comment("FIXME later", "x")),   # comment
    lambda wb, ws: setattr(ws, "title", "TBD numbers"),                          # sheet name
    lambda wb, ws: ws.cell(row=20, column=3, value="Value???"),                  # punctuation term glued to text
    lambda wb, ws: ws.cell(row=20, column=3, value="Name: [INSERT]"),
])
def test_forbidden_term_gaps_closed(tmp_path, mutate):
    assert not _hits(tmp_path, mutate)["pass"]


def test_forbidden_whole_word_still_respected(tmp_path):
    # "XXX" must not fire inside a longer token, "TK" not inside "TKO".
    res = _hits(tmp_path, lambda wb, ws: ws.cell(row=20, column=3, value="Ref XXXL / TKO"))
    assert res["pass"], res


def test_save_and_validate_fails_closed_without_libreoffice(tmp_path, monkeypatch):
    wb, _ = _clean_wb()
    monkeypatch.setenv("PATH", "/nonexistent")
    rep = save_and_validate(wb, tmp_path / "x.xlsx", raise_on_fail=False)
    assert rep["ok"] is False
    assert "error" in rep["recalc"]


def test_save_and_validate_catches_forbidden_term_without_recalc(tmp_path):
    wb, ws = _clean_wb()
    ws["C20"] = "TBD"
    ws["C20"].style = "ib_label"
    rep = save_and_validate(wb, tmp_path / "x.xlsx", recalc=False, raise_on_fail=False)
    assert rep["ok"] is False


@needs_soffice
def test_validator_survives_libreoffice_rewrite(tmp_path):
    """recalc.py re-saves via LibreOffice; cell.style used to raise IndexError
    there, the validator crashed, and save_and_validate reported ok=True."""
    wb, ws = _clean_wb()
    ws["C20"] = "TBD"
    ws["C20"].style = "ib_label"
    rep = save_and_validate(wb, tmp_path / "x.xlsx", raise_on_fail=False)
    assert rep["recalc"]["status"] == "success"
    assert "error" not in rep["validate"], rep["validate"]
    assert rep["validate"]["forbidden_terms"]["hard_count"] == 1
    assert rep["ok"] is False


@needs_soffice
@pytest.mark.parametrize("theme", ["ib", "engineering", "korean_corp", "saas"])
def test_clean_workbook_passes_after_recalc(tmp_path, theme):
    wb, _ = _clean_wb(theme)
    rep = save_and_validate(wb, tmp_path / f"{theme}.xlsx", raise_on_fail=False)
    assert rep["ok"], rep


@needs_soffice
def test_recalc_ignores_error_text_inside_notes(tmp_path):
    wb, ws = _clean_wb()
    ws["C20"] = "Lookups return #N/A when the key is missing"
    ws["C20"].style = "ib_note"
    p = tmp_path / "n.xlsx"
    wb.save(p)
    assert recalc.recalc(str(p))["status"] == "success"


def test_load_template_infers_theme():
    assert load_template("dashboard-skeleton")._xlsx_theme == "saas"
    assert load_template("engineering-sheet-skeleton")._xlsx_theme == "engineering"
    assert load_template("financial-model-skeleton")._xlsx_theme == "ib"


def test_write_table_col_formats(tmp_path):
    wb, ws = _clean_wb()
    info = write_table(ws, "B10", ["Seg", "FY24", "Growth"], [["a", 10, 0.2]],
                       num_format="currency", col_formats={"Growth": "percent_decimal"})
    assert "%" in ws.cell(row=info["first_data_row"], column=4).number_format
    assert "$" in ws.cell(row=info["first_data_row"], column=3).number_format
    with pytest.raises(KeyError):
        write_table(ws, "B20", ["Seg", "FY24"], [["a", 1]], col_formats={"Nope": "percent"})


def test_financial_template_arr_not_percent():
    wb = load_template("financial-model-skeleton")
    assert "%" not in wb["Assumptions"]["C6"].number_format
