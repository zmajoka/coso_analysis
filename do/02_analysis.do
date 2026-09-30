/*==============================================================================
  COSO Youth Employment Program - Côte d'Ivoire
  Baseline phone survey

  File    : 02_analysis.do
  Purpose : All analysis of the applicant baseline, one section per analysis.

  Input   : $cleandata/coso_baseline_clean.dta   (from 01_import_clean.do)
  Outputs : $out/Results.xls    (all tables, one after the other)
            $out/02_analysis.log

  Sample  : completed surveys only (complete == 1).

  Sections
    0. Paths
    1. Data and sample (common to all analyses)
    2. Analysis 1 - Track size (provisional track assignment)

  ------------------------------------------------------------------------------
  Analysis 1 - Track size
  Provisional track rule
    Track A - Start-up
        K1 = No (never started an income-generating activity)
        AND K2 = Yes (would like to start one)
    Track B - Existing enterprise growth
        K1 = Yes AND main job F1 = employer, own-account worker or member of
        a producers' cooperative
    Track C - Apprenticeship to work
        Not employed (E1-E3A) AND aspired sector P4 is a skilled trade:
        manufacturing, construction, transport, accommodation/food,
        ICT/digital, other services (repair, personal, arts)
        (no condition on prior training)
    Tracks A and C can overlap (not working, never started an activity,
    wants to start one, aspires to a trade): reported as a separate category.
    B cannot overlap with A (K1) or C (employment).
    Undetermined - everyone else, split by reason.
==============================================================================*/

version 17
clear all
set more off
set varabbrev off

*------------------------------------------------------------------------------*
* 0. Paths
*------------------------------------------------------------------------------*

global data   "C:\Users\WB461621\OneDrive - WBG\SPJ\Cote dIvoire\ImpactEval\Data"
global cleandata "C:\Users\WB461621\OneDrive - WBG\SPJ\Cote dIvoire\ImpactEval\Data\Clean"
global out   "C:\Users\WB461621\OneDrive - WBG\SPJ\Cote dIvoire\ImpactEval\Output"

capture log close
log using "$out/02_analysis.log", replace text

*------------------------------------------------------------------------------*
* 1. Data and sample (common to all analyses)
*------------------------------------------------------------------------------*
use "$cleandata/coso_baseline_clean.dta", clear

keep if complete == 1
di as txt "Completed interviews: " as res _N

* Respondents with more than one completed interview under the same cover_id
* (see docs/data_issues.md, issue 1) - used for a sensitivity table only
bysort cover_id: egen n_complete = total(complete)
gen byte dup_id = (n_complete > 1)
label var dup_id "cover_id has more than one completed interview"

* Accepted by supervisor (status Completed) vs rejected / still assigned
gen byte status_ok = (interview__status == "Completed")
label var status_ok "Interview status: Completed (not rejected)"
label values dup_id status_ok yesno

* Column variable for one-way tables
gen byte total = 1
label define total 1 "Total"
label values total total
label var total "Total"

*==============================================================================*
* 2. ANALYSIS 1 - TRACK SIZE
*==============================================================================*

*------------------------------------------------------------------------------*
* 2.1 Track criteria
*------------------------------------------------------------------------------*
* Inputs: spot check
tab K1, missing
tab K2 if K1 == 0, missing
tab F1 if K1 == 1, missing
tab employed, missing
tab P4 if employed == 0, missing

* Track A: never started an activity and wants to start one
gen byte crit_A = (K1 == 0 & K2 == 1)
label var crit_A "Meets Track A criteria (K1 = No, K2 = Yes)"

* Track B: has started an activity and is employer / own-account / cooperative
gen byte crit_B = (K1 == 1 & inlist(F1, 2, 3, 4))
label var crit_B "Meets Track B criteria (K1 = Yes, F1 = employer/own-account/coop)"

* Track C: not employed and aspires to a skilled-trade sector
*   P4: 3 Manufacturing, 4 Construction, 6 Transport, 7 Accommodation/food,
*       8 ICT/digital, 14 Other services (repair, personal, arts)
gen byte crit_C = (employed == 0 & inlist(P4, 3, 4, 6, 7, 8, 14))
label var crit_C "Meets Track C criteria (not employed, aspires to a skilled trade)"

label values crit_A crit_B crit_C yesno

* Check: B never overlaps with A or C (should both be 0)
count if crit_B == 1 & crit_A == 1
count if crit_B == 1 & crit_C == 1

*------------------------------------------------------------------------------*
* 2.2 Provisional track
*------------------------------------------------------------------------------*
gen byte track = 5
replace track = 1 if crit_A == 1 & crit_C == 0
replace track = 2 if crit_B == 1
replace track = 3 if crit_C == 1 & crit_A == 0
replace track = 4 if crit_A == 1 & crit_C == 1

label define track 1 "Track A - Start-up"                    ///
                   2 "Track B - Existing enterprise growth"  ///
                   3 "Track C - Apprenticeship to work"      ///
                   4 "Tracks A and C (both apply)"           ///
                   5 "Undetermined"
label values track track
label var track "Provisional programme track"

* Reason for undetermined
*   1 Employed, but not employer/own-account/coop with an own activity
*     (wage worker, family worker, apprentice, or never started an activity
*     and not wanting to)
*   2 Not employed, started an activity before (no longer running it)
*   3 Not employed, never started, no wish to start, aspired sector is not
*     a skilled trade
*   4 Missing information (K1, employment or P4)
gen byte undet_reason = .
replace undet_reason = 4 if track == 5 & (missing(K1) | missing(employed))
replace undet_reason = 1 if track == 5 & missing(undet_reason) & employed == 1
replace undet_reason = 2 if track == 5 & missing(undet_reason) & employed == 0 & K1 == 1
replace undet_reason = 3 if track == 5 & missing(undet_reason) & employed == 0 & K1 == 0
replace undet_reason = 4 if track == 5 & missing(undet_reason)

label define undet_reason                                                    ///
    1 "Employed, no own enterprise (wage/family worker, apprentice, or no wish to start)" ///
    2 "Not employed, had an activity before (closed)"                         ///
    3 "Not employed, no wish to start, aspiration not a skilled trade"        ///
    4 "Missing information"
label values undet_reason undet_reason
label var undet_reason "Reason track is undetermined"

* Spot checks
tab track, missing
tab undet_reason if track == 5, missing
tab track employed, missing
tab track K1, missing

*------------------------------------------------------------------------------*
* 2.3 Tables
*------------------------------------------------------------------------------*

* 2.3.1 Track size
tabout track total using "$out/Results.xls", ///
    replace c(freq col) format(0c 1p) layout(cb) style(xls) ///
    h1("Analysis 1 - Provisional programme track - completed interviews")

* 2.3.2 Criteria met (tracks A and C can overlap)
tabout crit_A total using "$out/Results.xls", ///
    append c(freq col) format(0c 1p) layout(cb) style(xls) ///
    h1("Analysis 1 - Meets Track A criteria: never started an activity and wants to start one")

tabout crit_B total using "$out/Results.xls", ///
    append c(freq col) format(0c 1p) layout(cb) style(xls) ///
    h1("Analysis 1 - Meets Track B criteria: started an activity, now employer/own-account/cooperative")

tabout crit_C total using "$out/Results.xls", ///
    append c(freq col) format(0c 1p) layout(cb) style(xls) ///
    h1("Analysis 1 - Meets Track C criteria: not employed, aspires to a skilled trade")

* 2.3.3 Why undetermined
tabout undet_reason total if track == 5 using "$out/Results.xls", ///
    append c(freq col) format(0c 1p) layout(cb) style(xls) ///
    h1("Analysis 1 - Undetermined track - reason")

* 2.3.4 Track by sex and by district
tabout track C1 using "$out/Results.xls", ///
    append c(freq col) format(0c 1p) layout(cb) style(xls) ///
    h1("Analysis 1 - Provisional track by sex")

tabout track B1 using "$out/Results.xls", ///
    append c(freq col) format(0c 1p) layout(cb) style(xls) ///
    h1("Analysis 1 - Provisional track by district")

* 2.3.5 Sensitivity: interview status and duplicate IDs
tabout track status_ok using "$out/Results.xls", ///
    append c(freq col) format(0c 1p) layout(cb) style(xls) ///
    h1("Analysis 1 - Provisional track by interview status (Yes = Completed, No = rejected/assigned)")

tabout track total if dup_id == 0 using "$out/Results.xls", ///
    append c(freq col) format(0c 1p) layout(cb) style(xls) ///
    h1("Analysis 1 - Provisional track - excluding cover_ids with more than one completed interview")

*==============================================================================*
* 3. ANALYSIS 2 - (next analysis goes here; tables use -append-)
*==============================================================================*

log close
