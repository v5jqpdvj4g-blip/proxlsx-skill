# Pro XLSX Design System v1 — the four themes

`theme.py` is the authoritative source. This file is a human-readable extract for anyone who wants to eyeball the design system without reading Python.

## Axes every theme sets

- **fonts**: body / header / title / note — `{name, size, bold, color}`
- **palette**: ink / paper / input / formula / link / external / header_bg / section_bg / total_bg / rule / accent / warn
- **number_formats**: currency / currency_decimal / number / number_decimal / percent / percent_decimal / multiple / date / year_text / zero_dash (+ theme extras)
- **layout**: row heights (default/header/title), col widths (gutter/label/data), zoom, margins
- **rules**: gridlines_visible (false), gridlines_print (false), freeze `{row,col}`, tab_color, orientation

## Theme snapshots

### `ib` — IB / Wall Street (default)
Feel: Macabacus/BIWS analyst deck. Dense, muted, parenthesized negatives.

- Body font: **Arial 10**, black
- Header: Arial 11 bold, white on dark gray `#404040`
- Title: Arial 14 bold, white on `#404040`, row height 26pt
- Palette: ink `#000000` · input `#4472C4` · link `#70AD47` · external `#C00000` · section_bg `#D9D9D9` · total_bg `#F2F2F2`
- Currency: `$#,##0;($#,##0);"-"`
- Freeze: B2 · Gutter col A width 2 · Data col width 13 · Zoom 100
- Tab color: `#1F4E78` (navy)

### `engineering` — EPC calculation sheet
Feel: Consolas for numeric columns so decimals align. Calibri bold headers for Western readability.

- Body font: **Consolas 10**, charcoal `#1A1A1A`
- Header: Calibri 11 bold, `#1A1A1A` on `#DCE6F1`
- Title: Calibri 14 bold, white on `#2F5496`, row height 24pt
- Palette: ink `#1A1A1A` · input `#0066CC` · link `#008233` · external `#C8102E` · rule `#888888`
- Number formats lean decimal-heavy: `currency_decimal` → `#,##0.00`, `number_decimal` → `#,##0.000`
- Extra: `scientific` → `0.000E+00`
- Freeze: C3 (wider gutter for parameter labels) · Zoom 100
- Tab color: `#2F5496` (steel blue)

### `korean_corp` — 국내 대기업 결재문서
Feel: Matches sister skill `001_hd-ksoe-pptx-toolkit`. Pretendard for body, HD현대체 for titles, Discovery Blue accent.

- Body font: **Pretendard 10**, charcoal `#1A1A1A`
- Header: Pretendard 11 bold, white on `#003087`
- Title: HD현대체 Bold 14, white on `#003087`, row height 28pt
- Palette: input `#003087` (Discovery Blue) · link `#00AD1D` (Heritage Green) · external `#C00000`
- Extra formats: `won` → `₩#,##0;(₩#,##0);"-"`, `multiple` → `0.0"배"`
- Freeze: B3 · Zoom 100 · A4-friendly margins
- Tab color: `#003087`

### `saas` — Modern SaaS dashboard
Feel: Aptos/Aptos Display. Generous whitespace. Slate + indigo + emerald.

- Body font: **Aptos 10**, slate `#0F172A`
- Header: Aptos Display 12 bold, white on `#0F172A`
- Title: Aptos Display 16 bold, white on `#0F172A`, row height 32pt
- Palette: input `#6366F1` (indigo) · link `#10B981` (emerald) · external `#EF4444` (red) · rule `#CBD5E1`
- Date format: `mmm d, yyyy`
- Freeze: B3 · Generous row heights (default 18pt, header 24pt)
- Tab color: `#6366F1`

## Registered NamedStyles per theme

Each theme registers these 13 roles (prefixed with the theme key, e.g. `ib_header`):

`title`, `subtitle`, `header`, `section`, `label`, `input`, `formula`, `link`, `external`, `subtotal`, `total`, `note`, `source`

Data roles (input/formula/link/external/subtotal/total) pre-set `number_format` to `currency`. Override per cell with `cell.number_format = theme['number_formats'][key]` when you need percent/multiple/etc.

## How to add a new theme

1. Add a key to `THEMES` in `theme.py` with `meta`, `fonts`, `palette`, `number_formats`, `layout`, `rules`.
2. Run `python theme.py` — it prints theme count and builds styles as a smoke check.
3. Add a template builder in `templates/_build_templates.py` and run it.
4. Run `validate_xlsx.py` against the new template — `slop_score` must be ≤ 10.
5. Add a row to the theme picker table in `SKILL.md`.
