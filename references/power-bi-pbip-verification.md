# Power BI (PBIP) verification & packaging

Companion to the xlsx pipeline for when a workbook has a Power BI counterpart authored
as a **PBIP** (Power BI Project — the text/source format: TMDL semantic model + JSON
report). `validate_pbip.py` is the PBIP analogue of `validate_xlsx.py`: it cannot execute
DAX (no Analysis Services on Linux), but it catches the defects hand-authored PBIP
actually ships with.

## Why a static checker

A `.pbix` is a compiled binary only Power BI Desktop can write. The reviewable master is
the **PBIP** folder: `<Name>.SemanticModel/` (TMDL) + `<Name>.Report/` (per-visual JSON).
Because visuals reference model objects by name in JSON, the classic failure mode is a
**dangling reference** — a card bound to a measure that was renamed or deleted. It opens
without a build error and simply renders blank. That is exactly what a name-level checker
finds without a running engine.

## What `validate_pbip.py` checks

Run it against one or more PBIP roots (`python validate_pbip.py "<project folder>" ...`).
Exit code is non-zero if any `E*` fires.

| Code | Check |
|------|-------|
| `E1` | a visual/page/report references a table `Entity` the model has no table for |
| `E2` | a reference names a `Property` (measure/column) that `Entity` does not expose |
| `E3` | a measure's DAX calls `[Name]` that is neither a measure nor a column (string literals stripped first) |
| `E4` | a measure's DAX calls `Table[Column]` where the table or column does not exist |
| `E5` | `pages.json` pageOrder / activePageName is out of sync with the page folders on disk |
| `E6` | `report.json` / `page.json` points at a RegisteredResources file missing on disk |

It parses TMDL structurally (`table` / `measure` / `column` declarations, quoted or bare)
and walks every report JSON for `{ "Expression": { "SourceRef": { "Entity" } }, "Property" }`
leaves — which covers measures, columns, and aggregations alike.

## Excel → DAX narrative fidelity

When the Power BI build is meant to reproduce an Excel narrative/auto-analysis engine,
verify the DAX against the workbook's own formulas rather than trusting comments:

1. Pull the driving formulas from the workbook (the narrative/helper/control sheets) with
   openpyxl (`data_only=False` for the formula, `data_only=True` for threshold constants).
2. Confirm the DAX thresholds equal the workbook constants (e.g. a status band's watch /
   critical cut-offs should match the Control sheet, not be re-guessed).
3. Preserve the **returned token set** of any status/classification measure. Report visuals
   colour-map those exact strings in conditional formatting, so changing the strings breaks
   the colours silently. Change thresholds and add branches; do not rename the outputs.
4. Prefer additive, monotonic branches (e.g. an Excel materiality floor that only reclassifies
   immaterial rows) over rewrites you cannot runtime-test here.

## Packaging

The PBIP folder is the master. To hand over a drop-in copy, zip the project folder(s)
preserving structure and confirm **file-level** parity against the source (ignore directory
entries — `zip -r` adds them). A true `.pbix` is produced only by opening the `.pbip` in
Power BI Desktop and choosing *File → Save As → .pbix*; it cannot be compiled off Desktop.
