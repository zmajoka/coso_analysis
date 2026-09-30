# COSO baseline – data issues and questions for the survey firm

Working notes from cleaning `data_exploration.xlsx` with `do/01_import_clean.do`
(log: `output/01_cleaning.log`). To be sent to the firm after the first round of
analysis.

- Export used: 1,032 rows (call attempts), latest interview 4 Sep 2026.
- Fieldwork is still ongoing. The firm keeps calling back respondents where there
  were problems, so all counts below will change with the next export.
- No names or phone numbers are recorded here. Look up respondents by `cover_id` /
  `interview__key` in the restricted file `coso_baseline_pii.dta`.

## Analysis sample

- Analysis uses **completed surveys only**: `complete == 1`, i.e. call result
  A4 = "Entretien réalisé", consent given (B6D = Oui), and the interview reached
  the end of section R. In this export that is 628 interviews.
- This is based on the survey content, not on the Survey Solutions status. Of the
  628, 414 are "Completed", 209 "RejectedBySupervisor" and 5
  "InterviewerAssigned".
- The 17 `cover_id`s in issue 1 below have more than one completed interview and
  are not resolved yet.

---

## A. Data errors to raise with the firm

### 1. Two or more completed interviews under the same `cover_id`, with different people

18 extra completed interviews across 17 `cover_id`s. The recorded name and sex
are the same, but birth year, sub-prefecture and locality differ, often across
different districts. For example:

| cover_id | Interview 1 (sub-prefecture, birth year) | Interview 2 |
|---|---|---|
| H-526 | Odienné, 2006 | Fadiadougou, **2017** |
| H-194 | Séguélon, 1990 | Mankono, 1991 |
| H-206 | Koumbala, 1993 | Sarhala, 2002 |
| H-265 | Kalamon, 1989 | Ouangolodougou, 1992 |

`cover_id`s affected: F-100, F-29, H-194, H-206, H-261, H-265, H-35, H-526,
H-550, H-592, H-612, H-760, H-830, H-866, H-926, H-927. Only H-760 (Katogo both
times, born 2002 / 2004) could be one person.

What it looks like: interviewers recorded a respondent under the wrong
assignment (wrong `cover_id`), and the name B7 was taken from the preloaded name
instead of what the respondent said. Several of the second interviews come from
the same villages (Sandala, Nafoungolo, Zanaplédougou, Djakaridjavogo).

Still to check on our side (locally, restricted file): whether the phone numbers
(R4, R6) also match within each pair; whether `localite` is preloaded or typed by
the interviewer.

### 2. Excel export: header row misaligned

From B1 onward, every value sits one column to the right of its header. The
column under "B1" is empty, and the last column (`assignment__id`) has no header.
The do file works around this (section 1), but a corrected export would be
safer.

### 3. Excel export: answers exported as French text instead of codes

Categorical answers come as text ("Oui", "Homme", ...), special answers in
numeric questions as text ("Ne sait pas", "Ne veut pas dire"), and dates as
month/day/year. The do file converts all of these. A Stata export with codes
(Survey Solutions "Stata" format) would avoid it.

### 4. Interviewer and supervisor not in the export

`nom_agent` and `nom_sup` are in the questionnaire but not in the file. They are
needed to check data quality by interviewer.

### 5. Supervisor rejections

365 interviews are "RejectedBySupervisor", 209 of them completed. In some cases a
new interview was opened instead of correcting the rejected one (see issue 1).

### 6. Timing problems

- Interviews left open for a long time: F-100 (44 hours), H-866 (23 hours),
  H-830 (21 hours), F-29 (25 hours).
- 20 completed interviews shorter than 10 minutes (`flag_short`), against about
  30 minutes planned.
- H-592: one interview dated 08/07/2026, before fieldwork started.

### 7. Implausible or inconsistent answers (flags in the clean data)

| Flag | Rows | Issue |
|---|---|---|
| `flag_age` | 32 | Age outside 15–40 (includes a birth year of 2017) |
| `flag_short` | 20 | Completed interview under 10 minutes |
| `flag_geo` | 9 | Reported district/region differs from the cover page |
| `flag_hours` | 6 | More than 112 hours worked last week |
| `flag_first_job` | 6 | Age at first job below 5 or above current age |
| `flag_income` | 1 | Earnings/savings/debt not positive |
| `flag_errors` | 1 | Survey Solutions validation errors |

### 8. Questionnaire / CAPI design issues

- **Q2** (times volunteered, last 30 days) is a text question. 76 answers are
  words, not numbers ("AUCUN", "NON", "OUI", "UNE FOIS", "PLUSIEURS FOIS",
  "1H30", ...). It should be a numeric question.
- **`consentement`** (computed in CAPI) is 1 for anyone who did not refuse a
  callback, including people never reached: 90 rows say consent without consent
  in B6D. The do file rebuilds consent from B6D.
- **R5** (agrees to be recontacted): everyone answered Yes.
- **H3** (wage): 194 of about 600 working respondents refused; H5 (earnings): 127
  refused. Brackets (H3A/H5A) are available.
- **`localite`** has an encoding problem for at least one value ("FERK?").

---

## B. Questions for the firm

1. For each `cover_id` in issue 1: who was actually interviewed in each
   interview? Can this be checked against the call recordings?
2. Can you send the sample/preload file (`cover_id`, name, phone number,
   locality) loaded into the tablets, so interviews can be matched to applicants?
3. Is `localite` on the cover page preloaded, or typed by the interviewer? Is B7
   typed from what the respondent says, or copied from the preloaded name?
4. Can you add `nom_agent` and `nom_sup` to the export?
5. Can you export in Survey Solutions Stata format (codes and labels), or fix
   the Excel header row?
6. Are rejected interviews being corrected (same interview) or redone as new
   interviews? Which one should we use when both exist?
7. Why were some interviews open for more than 20 hours, and some completed in
   under 10 minutes?
8. How are you planning the call-backs, and how will you report response rates
   and attrition (by district, sex, and treatment status if relevant)?

---

## C. Pending changes to the do file (after the firm replies)

- Q2: recode the text answers ("AUCUN"/"NON" → 0, "UNE FOIS" → 1,
  "2 FOIS" → 2) and add `Q2_any` (volunteered at all).
- Remove `flag_age_c2`: C2 is only asked when birth year is unknown, so it can
  never be checked.
- Add `flag_dup_conflict` for issue 1. Choose one row per respondent using:
  completed first, then Completed ahead of Rejected ahead of Assigned, then the
  latest start time.
- Add a flag for interviews open more than 3 hours.
- Confirm the programme's age limits (currently 15–40 as a placeholder).
- Re-run on each new export. If the firm fixes the header row, the file will have
  214 columns instead of 215: the `rename (_all)` step in section 1 will then
  stop, and the `B_blank` placeholder has to be removed.
