# Anti-slop checklist for xlsx outputs

Run this mentally (or with `validate_xlsx.py`) before declaring any xlsx done. Every item below is checked automatically by the validator unless marked 👁️ (requires a human look).

## Sheet-level rules
- [ ] All sheet tabs have descriptive names (no `Sheet`, `Sheet1`, `Sheet2`)
- [ ] Every sheet has a tab color set (theme picks one automatically)
- [ ] Every sheet has freeze panes set at the data-start row/col
- [ ] Gridlines are hidden (both on screen and on print)
- [ ] Every sheet has a title banner or non-empty A1/B1 — no "blank cover"
- [ ] Page orientation is set (landscape is the default for financial work)
- [ ] Print margins are set to 0.5" (or theme default)
- [ ] Zoom level is 100% (fixed widths + readable)

## Typography rules
- [ ] No cell uses the built-in Calibri 11 (unless the theme explicitly uses it for headers)
- [ ] Every content cell has a NamedStyle from the theme
- [ ] Body font matches the theme: Arial 10 (ib), Consolas 10 (engineering), Pretendard 10 (korean_corp), Aptos 10 (saas)
- [ ] Header rows are bold and filled with `header_bg`
- [ ] Title banner exists on every sheet (or at minimum every visible-on-open sheet)

## Number formatting
- [ ] Every numeric cell has a non-`General` number format
- [ ] Currency cells use `$#,##0;($#,##0);"-"` (or theme variant)
- [ ] Percent cells use `0.0%;(0.0%);"-"` (one decimal for ratios)
- [ ] Multiple cells use `0.0"x"` (valuation multiples)
- [ ] Negative numbers display with parentheses, not minus signs
- [ ] Zero values collapse to `"-"` (the third argument of the number format)
- [ ] 👁️ Currency symbols only at the top and total row of a column, never every row
- [ ] 👁️ Year columns use text format (`"2024"`, not `2,024`)

## Color rules
- [ ] Distinct fill colors on the sheet ≤ 8 (no rainbow)
- [ ] Input cells use the theme's `input` color (Macabacus blue for ib)
- [ ] Formula cells use the theme's `formula` color (black for ib)
- [ ] Cross-sheet links use `link` color (green for ib)
- [ ] External-file references use `external` color (red for ib)
- [ ] 👁️ Colors carry semantic meaning, not decoration

## Border rules
- [ ] Borders only where they mark structure (section headers, subtotals, totals)
- [ ] Total rows have a top thin border and bottom double border
- [ ] Subtotal rows have a top thin border
- [ ] Cells-with-borders pct ≤ 60% (above that it looks like a re-drawn grid)
- [ ] 👁️ No mix of thin/medium/thick/dashed borders in the same region

## Formula rules (existing v1 rules, still enforced)
- [ ] Zero `#REF!` / `#DIV/0!` / `#VALUE!` / `#N/A` / `#NAME?` / `#NULL!` / `#NUM!` errors
- [ ] No Python-computed values hardcoded into cells (use formulas referencing Assumptions)
- [ ] All hardcoded numeric inputs are in a dedicated Assumptions sheet or a clearly-labeled band
- [ ] Every hardcoded input has a source note or comment
- [ ] No circular references

## Text hygiene
- [ ] No hard forbidden terms in any cell: `TBD`, `TK`, `FIXME`, `XXX`, `[INSERT]`, `???`, `Lorem ipsum`
- [ ] No vague AI-slop adjectives: `revolutionary`, `cutting-edge`, `state-of-the-art`, `world-class`, `seamlessly`, `paradigm shift`
- [ ] "Synergy" / "leverage" only when quantified
- [ ] Numbers are not written as text ("15" vs 15); use formulas for derived values
- [ ] Source notes use the pattern `Source: [System], [Date], [Reference]`

## How to run the automated portion
```bash
python validate_xlsx.py path/to/file.xlsx
```
Target: `status: pass`, `slop_score <= 10`.
