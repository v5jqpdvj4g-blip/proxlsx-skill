# Number format cookbook

Named formats live in `theme.py['{theme}']['number_formats']`. Pass the key (not the string) to `write_table(..., num_format='currency')` etc. The builder pulls the right string from the active theme so `'percent'` is `0%;(0%);"-"` in `ib` and `0.00%;(0.00%);"-"` in `engineering` without changing your code.

## The core set (every theme defines these)

| Key                 | ib string                       | Renders                         | Use for                          |
|---------------------|----------------------------------|----------------------------------|----------------------------------|
| `currency`          | `$#,##0;($#,##0);"-"`           | `$1,250` · `(250)` · `-`         | Revenue, costs, cash flows (whole dollars) |
| `currency_decimal`  | `$#,##0.0;($#,##0.0);"-"`       | `$1,250.5`                       | Per-unit pricing, small ratios   |
| `number`            | `#,##0;(#,##0);"-"`              | `1,250` · `(250)` · `-`          | Units, counts, headcount         |
| `number_decimal`    | `#,##0.0;(#,##0.0);"-"`         | `1,250.5`                        | Ratios, densities, engineering   |
| `percent`           | `0%;(0%);"-"`                    | `15%` · `(3%)` · `-`             | Growth, shares                   |
| `percent_decimal`   | `0.0%;(0.0%);"-"`                | `15.0%` · `(2.5%)`               | Rates, margins                   |
| `multiple`          | `0.0"x"`                         | `5.0x`                           | Valuation multiples (EV/EBITDA)  |
| `date`              | `mm/dd/yyyy`                     | `04/12/2026`                     | Dates                            |
| `year_text`         | `0`                              | `2024` (no comma)                | Year headers — never `2,024`     |
| `zero_dash`         | `#,##0;(#,##0);"-"`              | `-` for 0                        | When you want em-dash on zero    |

## Theme-specific extras

- `ib`: `basis_points` → `#,##0" bps"`
- `engineering`: `scientific` → `0.000E+00`, `number_decimal` → 3 decimals
- `korean_corp`: `won` → `₩#,##0;(₩#,##0);"-"`, `multiple` → `0.0"배"`
- `saas`: `date` → `mmm d, yyyy`

## Anti-patterns

- ❌ `General` format on numeric cells → validator flags it.
- ❌ `0.00` on a percentage column → should use `percent_decimal`.
- ❌ Mixing `#,##0` and `#,##0.00` in the same column → pick one.
- ❌ Currency symbol in every cell of a column → symbol belongs only on the first data row and the total.
- ❌ `-123` instead of `(123)` for negatives.

## How to apply directly

If you must override a specific cell's format (e.g., a percentage cell inside an otherwise-currency table):

```python
from theme import get_theme
t = get_theme('ib')
cell.number_format = t['number_formats']['percent_decimal']
```

This is safe after `cell.style = '...'` because `number_format` mutates only that one attribute.
