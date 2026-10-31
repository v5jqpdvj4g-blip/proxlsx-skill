# Templates

Pre-built skeletons that load via `build_xlsx.load_template(name, theme='ib')`. All four pass `validate_xlsx.py` with `slop_score=0` at build time. Regenerate them any time with:

```bash
python templates/_build_templates.py
```

## financial-model-skeleton.xlsx  (theme: ib)

Six-sheet scaffold for a 3-statement model.

| Sheet            | What's in it                              | Editable region                          |
|------------------|-------------------------------------------|------------------------------------------|
| `Cover`          | Title banner, prepared-by/date/version    | C{N..N+2}                                |
| `Assumptions`    | Revenue drivers table (5-year), sources   | C{6..9} through G{6..9}                  |
| `Income Statement` | Stub with title + note                  | (you build it)                           |
| `Balance Sheet`  | Stub with title + note                    | (you build it)                           |
| `Cash Flow`      | Stub with title + note                    | (you build it)                           |
| `Summary`        | Stub with title + note                    | (you build it)                           |

Rule: **never** write values into IS/BS/CF — use formulas referencing `Assumptions`.

## data-report-skeleton.xlsx  (theme: ib)

Two-sheet layout for a data-pull deliverable.

| Sheet      | What's in it                                                  |
|------------|---------------------------------------------------------------|
| `Report`   | Title banner + 3-row summary table (FY23–FY25, YoY Δ and 2-yr CAGR as % formulas) with total row |
| `Raw Data` | Blank (paste raw data starting at row 2)                      |

Edit the table body cells directly; totals recalc on save.

## dashboard-skeleton.xlsx  (theme: saas)

Product-metrics dashboard layout.

| Sheet          | What's in it                                              |
|----------------|-----------------------------------------------------------|
| `Dashboard`    | Title banner, 4-card KPI band, 4-week trend table         |
| `Data source`  | Single description cell                                   |

The KPI band lives in B{r}..G{r+1}; replace the zero placeholders with formulas that pull from your data source sheet.

## engineering-sheet-skeleton.xlsx  (theme: engineering)

Single-sheet parameter table for EPC workflows.

| Sheet         | What's in it                                                    |
|---------------|-----------------------------------------------------------------|
| `Parameters`  | Title banner + Parameter / Unit / Spec / Actual / Δ / Notes     |

Extend rows for additional parameters; the `Δ` column ships with `=IF(E{i}="","",E{i}-D{i})` so it stays blank until an Actual is entered.

## Adding a new template

1. Add a `build_xxx()` function in `_build_templates.py`.
2. Call `save_and_validate(wb, _TEMPLATES_DIR / 'xxx-skeleton.xlsx', recalc=False, raise_on_fail=False)` at the end.
3. Add an entry to the `builders` list in `main()`.
4. Rerun `python templates/_build_templates.py` (it lints every skeleton and exits non-zero on failure).
5. Confirm `python validate_xlsx.py templates/xxx-skeleton.xlsx --no-recalc` reports `status=pass`.
