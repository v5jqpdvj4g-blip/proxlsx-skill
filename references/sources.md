# Sources used to build the xlsx design system

The four themes in `theme.py` were not invented — they transcribe conventions from these sources. Follow a link when you need to defend a choice or grow the system.

## IB / Wall Street (the default)

- Macabacus — Improving Model Readability with Color Formatting: https://macabacus.com/blog/improving-model-readability-with-color-formatting
- Macabacus — Formatting Alignment in Excel: https://www.macabacus.com/blog/formatting-alignment-in-excel
- Macabacus Help — Color Palettes: https://help.macabacus.com/article/200-color-palettes
- Breaking Into Wall Street — Excel Tutorials: https://breakingintowallstreet.com/kb/excel/
- Wall Street Prep — Financial Modeling Knowledge Base: https://www.wallstreetprep.com/knowledge/financial-modeling/
- Financial Edge — Financial Model Formatting: https://www.fe.training/free-resources/financial-modeling/financial-model-formatting/
- Training The Street — Formatting a Financial Model: https://trainingthestreet.com/fmaamoc-article-2-formatting-a-financial-model/
- Wall Street Oasis — Financial Model Formatting: https://www.wallstreetoasis.com/resources/financial-modeling/financial-model-formatting
- Wall Street Oasis — Calibri vs. Arial thread: https://www.wallstreetoasis.com/forum/investment-banking/what-font-and-size-do-you-model-in
- Forecastr — Color Coding Guide: https://www.forecastr.co/blog/financial-model-color-coding-guide
- excel4finance — Custom Number Formatting: http://excel4finance.blogspot.com/2010/08/custom-number-formatting.html

## Python libraries referenced

- openpyxl (foundation): https://openpyxl.readthedocs.io/
- XlsxWriter (alternative with richer chart support): https://xlsxwriter.readthedocs.io/
- StyleFrame (pandas + openpyxl styling wrapper): https://github.com/DeepSpace2/StyleFrame
- Great Tables (reference for tabular beauty): https://github.com/posit-dev/great-tables

## Engineering theme

- Training-domain convention: Consolas / monospace for numeric data tables to align decimals visually
- SI unit conventions follow ISO 80000-1

## Korean corporate theme

- Pretendard fallback chain: https://github.com/orioncactus/pretendard
- HD현대 CI palette was derived by matching the HD KSOE pptx sister skill in this collection (`001_hd-ksoe-pptx-toolkit`)

## SaaS dashboard theme

- Aptos font family (Microsoft 2023+ default): https://www.wps.com/blog/aptos-vs-calibri-comparison-of-microsoft-s-old-and-new-default-fonts/
- Slate / indigo / emerald palette borrowed from Tailwind's default palette
