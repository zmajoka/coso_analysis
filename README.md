# COSO Youth Employment Program Analysis
Baseline survey data analysis and do files.

## Do files

| File | What it does |
|------|--------------|
| `do/01_import_clean.do` | Imports `data_exploration.xlsx`, cleans and labels it following `docs/Original Questionnaire_COSO_V5.pdf`, saves the clean dataset |

## Folders (globals set at the top of each do file)

- `$data` – raw data (e.g. `data_exploration.xlsx`)
- `$cleandata` – clean datasets only
- `$out` – everything else: logs, tables, figures and other outputs

## Coding conventions

- Do not use locals unless absolutely necessary. Short `foreach` loops are fine.
- Use globals only for file paths.
- Keep the code simple and straightforward so it is easy to spot-check:
  minimal locals, no user-written programs, write variable lists out in full.
- Export analytical output with `tabout`.
