# COSO Youth Employment Program Analysis
Baseline survey data analysis and do files.

## Do files

| File | What it does |
|------|--------------|
| `do/01_import_clean.do` | Imports `data_exploration.xlsx`, cleans and labels it following `docs/Original Questionnaire_COSO_V5.pdf`, saves the clean dataset |
| `do/02_analysis.do` | All analysis, one numbered section per analysis. All tables go to `$out/Results.xls` (`tabout`, first table `replace`, the rest `append`). Section 2: Analysis 1, provisional programme track and track sizes. Section 3: Analysis 2, sector concentration by region and urban/rural (saturation risk). Section 4: Analysis 3, skills baseline by track. Section 5: Analysis 4, financial constraint profile by track. Section 6: Analysis 5, gender-differentiated constraints |

Data problems found during cleaning, and questions for the survey firm, are in `docs/data_issues.md`.

**Analysis sample:** completed surveys only (`keep if complete == 1`).

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
