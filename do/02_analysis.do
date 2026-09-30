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
    3. Analysis 2 - Sector concentration by geography (saturation risk)

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
tabout track using "$out/Results.xls", ///
    replace oneway c(freq col) format(0c 1p) layout(cb) style(xls) ///
    h1("Analysis 1 - Provisional programme track - completed interviews")

* 2.3.2 Criteria met (tracks A and C can overlap)
tabout crit_A using "$out/Results.xls", ///
    append oneway c(freq col) format(0c 1p) layout(cb) style(xls) ///
    h1("Analysis 1 - Meets Track A criteria: never started an activity and wants to start one")

tabout crit_B using "$out/Results.xls", ///
    append oneway c(freq col) format(0c 1p) layout(cb) style(xls) ///
    h1("Analysis 1 - Meets Track B criteria: started an activity, now employer/own-account/cooperative")

tabout crit_C using "$out/Results.xls", ///
    append oneway c(freq col) format(0c 1p) layout(cb) style(xls) ///
    h1("Analysis 1 - Meets Track C criteria: not employed, aspires to a skilled trade")

* 2.3.3 Why undetermined
tabout undet_reason if track == 5 using "$out/Results.xls", ///
    append oneway c(freq col) format(0c 1p) layout(cb) style(xls) ///
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

tabout track if dup_id == 0 using "$out/Results.xls", ///
    append oneway c(freq col) format(0c 1p) layout(cb) style(xls) ///
    h1("Analysis 1 - Provisional track - excluding cover_ids with more than one completed interview")

*==============================================================================*
* 3. ANALYSIS 2 - SECTOR CONCENTRATION BY GEOGRAPHY
*==============================================================================*
* Market saturation risk: are many applicants aiming for (or already in) the
* same sector in the same area?
*   Aspired sector : P4, all completed interviews
*   Current sector : F2, employed applicants only (F asked if employed)
*   Geography      : region (B2) and urban/rural (B5)
* A region x sector cell where more than 30% of the region's applicants share
* the same sector is flagged as a saturation risk zone.
* Note: P4/F2 are broad sectors ("Other services" groups tailors, mechanics,
* hairdressers...). Regions are small (see 3.1): read shares with the counts.

*------------------------------------------------------------------------------*
* 3.1 Sample by region and urban/rural
*------------------------------------------------------------------------------*
tabout B2 using "$out/Results.xls", ///
    append oneway c(freq col) format(0c 1p) layout(cb) style(xls) ///
    h1("Analysis 2 - Completed interviews by region")

tabout B2 B5 using "$out/Results.xls", ///
    append c(freq row) format(0c 1p) layout(cb) style(xls) ///
    h1("Analysis 2 - Completed interviews by region and urban/rural")

*------------------------------------------------------------------------------*
* 3.2 Aspired sector (P4) by region and by urban/rural
*------------------------------------------------------------------------------*
tabout P4 B2 using "$out/Results.xls", ///
    append c(freq col) format(0c 1p) layout(cb) style(xls) ///
    h1("Analysis 2 - Aspired sector (P4) by region: number and % of the region's applicants")

tabout P4 B5 using "$out/Results.xls", ///
    append c(freq col) format(0c 1p) layout(cb) style(xls) ///
    h1("Analysis 2 - Aspired sector (P4) by urban/rural")

*------------------------------------------------------------------------------*
* 3.3 Aspired sector: top 3 per region and saturation risk (> 30%)
*------------------------------------------------------------------------------*
preserve
    keep if !missing(P4, B2)
    gen byte one = 1
    collapse (sum) n = one, by(B2 P4)

    bysort B2: egen n_region = total(n)
    gen share = 100 * n / n_region
    egen rank = rank(-n), by(B2) field          // 1 = most common; ties share a rank
    gen byte saturation = (share > 30)

    label var n          "Applicants aspiring to the sector"
    label var n_region   "Applicants in the region"
    label var share      "Share of the region's applicants (%)"
    label var rank       "Rank of the sector in the region"
    label var saturation "More than 30% of the region's applicants"

    * Spot check
    sort B2 rank
    list B2 P4 n n_region share rank saturation if rank <= 3, sepby(B2) noobs

    tabout P4 B2 if rank <= 3 using "$out/Results.xls", ///
        append sum c(mean share) format(1) layout(cb) style(xls) ///
        h1("Analysis 2 - Top 3 aspired sectors in each region: % of the region's applicants (blank = not in top 3)")

    tabout P4 B2 if rank <= 3 using "$out/Results.xls", ///
        append sum c(mean rank) format(0) layout(cb) style(xls) ///
        h1("Analysis 2 - Top 3 aspired sectors in each region: rank (1 = most common)")

    count if saturation == 1
    if r(N) > 0 {
        tabout P4 B2 if saturation == 1 using "$out/Results.xls", ///
            append sum c(mean share) format(1) layout(cb) style(xls) ///
            h1("Analysis 2 - SATURATION RISK: aspired sector chosen by more than 30% of the region's applicants (%)")
    }
restore

*------------------------------------------------------------------------------*
* 3.4 Current sector (F2, employed) by region and by urban/rural
*------------------------------------------------------------------------------*
tabout F2 B2 if employed == 1 using "$out/Results.xls", ///
    append c(freq col) format(0c 1p) layout(cb) style(xls) ///
    h1("Analysis 2 - Current sector (F2) by region, employed applicants: number and % of the region's employed")

tabout F2 B5 if employed == 1 using "$out/Results.xls", ///
    append c(freq col) format(0c 1p) layout(cb) style(xls) ///
    h1("Analysis 2 - Current sector (F2) by urban/rural, employed applicants")

*------------------------------------------------------------------------------*
* 3.5 Current sector: top 3 per region and saturation risk (> 30%)
*------------------------------------------------------------------------------*
preserve
    keep if employed == 1 & !missing(F2, B2)
    gen byte one = 1
    collapse (sum) n = one, by(B2 F2)

    bysort B2: egen n_region = total(n)
    gen share = 100 * n / n_region
    egen rank = rank(-n), by(B2) field
    gen byte saturation = (share > 30)

    label var n          "Employed applicants in the sector"
    label var n_region   "Employed applicants in the region"
    label var share      "Share of the region's employed applicants (%)"
    label var rank       "Rank of the sector in the region"
    label var saturation "More than 30% of the region's employed applicants"

    * Spot check
    sort B2 rank
    list B2 F2 n n_region share rank saturation if rank <= 3, sepby(B2) noobs

    tabout F2 B2 if rank <= 3 using "$out/Results.xls", ///
        append sum c(mean share) format(1) layout(cb) style(xls) ///
        h1("Analysis 2 - Top 3 current sectors in each region: % of the region's employed applicants (blank = not in top 3)")

    tabout F2 B2 if rank <= 3 using "$out/Results.xls", ///
        append sum c(mean rank) format(0) layout(cb) style(xls) ///
        h1("Analysis 2 - Top 3 current sectors in each region: rank (1 = most common)")

    count if saturation == 1
    if r(N) > 0 {
        tabout F2 B2 if saturation == 1 using "$out/Results.xls", ///
            append sum c(mean share) format(1) layout(cb) style(xls) ///
            h1("Analysis 2 - SATURATION RISK: current sector of more than 30% of the region's employed applicants (%)")
    }
restore

*------------------------------------------------------------------------------*
* 3.6 Aspired vs current sector (employed): moving into or staying in a sector
*------------------------------------------------------------------------------*
tabout F2 P4 if employed == 1 using "$out/Results.xls", ///
    append c(freq row) format(0c 1p) layout(cb) style(xls) ///
    h1("Analysis 2 - Current sector (rows) by aspired sector (columns), employed applicants: row %")

*==============================================================================*
* 4. ANALYSIS 3 - (next analysis goes here; tables use -append-)
*==============================================================================*

log close
