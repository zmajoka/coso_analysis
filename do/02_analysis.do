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
    4. Analysis 3 - Skills baseline by track
    5. Analysis 4 - Financial constraint profile by track

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
        Not employed (E1-E3A) AND aspired sector (P4_sector) is a skilled trade:
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
tab P4_sector if employed == 0, missing

* Track A: never started an activity and wants to start one
gen byte crit_A = (K1 == 0 & K2 == 1)
label var crit_A "Meets Track A criteria (K1 = No, K2 = Yes)"

* Track B: has started an activity and is employer / own-account / cooperative
gen byte crit_B = (K1 == 1 & inlist(F1, 2, 3, 4))
label var crit_B "Meets Track B criteria (K1 = Yes, F1 = employer/own-account/coop)"

* Track C: not employed and aspires to a skilled-trade sector
*   P4_sector: 3 Manufacturing, 4 Construction, 6 Transport, 7 Accommodation/food,
*       8 ICT/digital, 14 Other services (repair, personal, arts)
gen byte crit_C = (employed == 0 & inlist(P4_sector, 3, 4, 6, 7, 8, 14))
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
*   4 Missing information (K1, employment or P4_sector)
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
*   Aspired sector : P4_sector (P4 with 'Other' recoded from P4X), all completed
*   Current sector : F2, employed applicants only (F asked if employed)
*   Geography      : region (B2) and urban/rural (B5)
* A region x sector cell where more than 30% of the region's applicants share
* the same sector is flagged as a saturation risk zone.
* Note: P4_sector/F2 are broad sectors ("Other services" groups tailors, mechanics,
* hairdressers...). Regions are small: read shares with the counts (3.1, 3.3).

*------------------------------------------------------------------------------*
* 3.1 Aspired sector (P4, "Other" recoded) by region and by urban/rural
*------------------------------------------------------------------------------*
tabout P4_sector B2 using "$out/Results.xls", ///
    append c(freq col) format(0c 1p) layout(cb) style(xls) ///
    h1("Analysis 2 - Aspired sector (P4, "Other" recoded) by region: number and % of the region's applicants")

tabout P4_sector B5 using "$out/Results.xls", ///
    append c(freq col) format(0c 1p) layout(cb) style(xls) ///
    h1("Analysis 2 - Aspired sector (P4, "Other" recoded) by urban/rural")

*------------------------------------------------------------------------------*
* 3.2 Aspired sector: top 3 per region and saturation risk (> 30%)
*------------------------------------------------------------------------------*
preserve
    keep if !missing(P4_sector, B2)
    gen byte one = 1
    collapse (sum) n = one, by(B2 P4_sector)

    bysort B2: egen n_region = total(n)
    gen share = 100 * n / n_region
    egen rank = rank(n), by(B2) field           // 1 = most common; ties share a rank
    gen byte saturation = (share > 30)

    label var n          "Applicants aspiring to the sector"
    label var n_region   "Applicants in the region"
    label var share      "Share of the region's applicants (%)"
    label var rank       "Rank of the sector in the region"
    label var saturation "More than 30% of the region's applicants"

    * Spot check
    sort B2 rank
    list B2 P4_sector n n_region share rank saturation if rank <= 3, sepby(B2) noobs

    tabout P4_sector B2 if rank <= 3 using "$out/Results.xls", ///
        append sum c(mean share) format(1) layout(cb) style(xls) ///
        h1("Analysis 2 - Top 3 aspired sectors in each region: % of the region's applicants (blank = not in top 3)")

    tabout P4_sector B2 if rank <= 3 using "$out/Results.xls", ///
        append sum c(mean rank) format(0) layout(cb) style(xls) ///
        h1("Analysis 2 - Top 3 aspired sectors in each region: rank (1 = most common)")

    count if saturation == 1
    if r(N) > 0 {
        tabout P4_sector B2 if saturation == 1 using "$out/Results.xls", ///
            append sum c(mean share) format(1) layout(cb) style(xls) ///
            h1("Analysis 2 - SATURATION RISK: aspired sector chosen by more than 30% of the region's applicants (%)")
    }
restore

*------------------------------------------------------------------------------*
* 3.3 Current sector (F2, employed) by region and by urban/rural
*------------------------------------------------------------------------------*
tabout F2 B2 if employed == 1 using "$out/Results.xls", ///
    append c(freq col) format(0c 1p) layout(cb) style(xls) ///
    h1("Analysis 2 - Current sector (F2) by region, employed applicants: number and % of the region's employed")

tabout F2 B5 if employed == 1 using "$out/Results.xls", ///
    append c(freq col) format(0c 1p) layout(cb) style(xls) ///
    h1("Analysis 2 - Current sector (F2) by urban/rural, employed applicants")

*------------------------------------------------------------------------------*
* 3.4 Current sector: top 3 per region and saturation risk (> 30%)
*------------------------------------------------------------------------------*
preserve
    keep if employed == 1 & !missing(F2, B2)
    gen byte one = 1
    collapse (sum) n = one, by(B2 F2)

    bysort B2: egen n_region = total(n)
    gen share = 100 * n / n_region
    egen rank = rank(n), by(B2) field
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

*==============================================================================*
* 4. ANALYSIS 3 - SKILLS BASELINE BY TRACK
*==============================================================================*
* Uses the provisional track from Analysis 1 (section 2).
*   K7     : self-rated business management skills
*   O3     : vocational/technical training or apprenticeship, last 12 months
*   O4     : field of that training (free text)
*   O5, O6 : training provider, usefulness
*   P4_sector : aspired sector (P4 with "Other" recoded from P4X)
*   O4_sector : sector of the training field (O4 coded, cleaning 6.10)
*   J2     : time without work and searching (asked only if J0 = searching)

*------------------------------------------------------------------------------*
* 4.1 Self-rated business skills (K7) by provisional track
*------------------------------------------------------------------------------*
tabout K7 track using "$out/Results.xls", ///
    append c(freq col) format(0c 1p) layout(cb) style(xls) ///
    h1("Analysis 3 - Self-rated business management skills (K7) by provisional track")

*------------------------------------------------------------------------------*
* 4.2 Prior training (O module) by aspired sector
*------------------------------------------------------------------------------*
tabout O3 P4_sector using "$out/Results.xls", ///
    append c(freq col) format(0c 1p) layout(cb) style(xls) ///
    h1("Analysis 3 - Vocational/technical training in the last 12 months (O3) by aspired sector")

tabout O5 P4_sector if O3 == 1 using "$out/Results.xls", ///
    append c(freq col) format(0c 1p) layout(cb) style(xls) ///
    h1("Analysis 3 - Training provider (O5) by aspired sector, trained applicants")

tabout O6 P4_sector if O3 == 1 using "$out/Results.xls", ///
    append c(freq col) format(0c 1p) layout(cb) style(xls) ///
    h1("Analysis 3 - Usefulness of the training (O6) by aspired sector, trained applicants")

* Does the training field match the aspired sector?
*   1 = same sector, 0 = different sector, missing = not trained, or training
*   not occupation-specific / not classified (O4_sector 97, 98)
gen byte train_match = (O4_sector == P4_sector) if O3 == 1 & inrange(O4_sector, 1, 14)
label var train_match "Training field matches aspired sector"
label values train_match yesno

* Spot check: training field, its coded sector and the aspired sector
preserve
    keep if O3 == 1
    sort P4_sector O4_sector
    list P4_sector O4 O4_sector train_match, sepby(P4_sector) noobs
restore

tabout O4_sector P4_sector if O3 == 1 using "$out/Results.xls", ///
    append c(freq col) format(0c 1p) layout(cb) style(xls) ///
    h1("Analysis 3 - Sector of the training field (O4 coded) by aspired sector, trained applicants")

tabout train_match P4_sector using "$out/Results.xls", ///
    append c(freq col) format(0c 1p) layout(cb) style(xls) ///
    h1("Analysis 3 - Training field matches aspired sector, by aspired sector (trained, occupation-specific training)")

*------------------------------------------------------------------------------*
* 4.3 Track C: unemployment duration (J2) by prior training and aspired sector
*------------------------------------------------------------------------------*
* Track C criteria met (crit_C: Track C alone or Tracks A and C).
* J2 is blank for those not searching for work (J0 = No).
tab J2 if crit_C == 1, missing

tabout J2 O3 if crit_C == 1 using "$out/Results.xls", ///
    append mi c(freq col) format(0c 1p) layout(cb) style(xls) ///
    h1("Analysis 3 - Track C: time without work and searching (J2) by training in the last 12 months (O3); Missing = not searching")

tabout J2 P4_sector if crit_C == 1 using "$out/Results.xls", ///
    append mi c(freq col) format(0c 1p) layout(cb) style(xls) ///
    h1("Analysis 3 - Track C: time without work and searching (J2) by aspired sector; Missing = not searching")

*==============================================================================*
* 5. ANALYSIS 4 - FINANCIAL CONSTRAINT PROFILE BY TRACK
*==============================================================================*
* Track groups use the criteria from Analysis 1 (section 2), so people meeting
* both A and C criteria are in both groups:
*   Track A: crit_A == 1    Track B: crit_B == 1    Track C: crit_C == 1
* Capital need for start-up is not asked in this questionnaire: to be added
* when the source data (e.g. application forms) is merged by cover_id.

*------------------------------------------------------------------------------*
* 5.1 Track A (start-up): savings and credit access, by aspired sector
*------------------------------------------------------------------------------*
* Savings behaviour (L3) and amount saved in a typical month (L4, savers only)
tabout L3 P4_sector if crit_A == 1 using "$out/Results.xls", ///
    append c(freq col) format(0c 1p) layout(cb) style(xls) ///
    h1("Analysis 4 - Track A: saves money regularly (L3), by aspired sector")

tabout P4_sector if crit_A == 1 using "$out/Results.xls", ///
    append sum oneway c(N L4 p25 L4 median L4 p75 L4) format(0c 0c 0c 0c) style(xls) ///
    h1("Analysis 4 - Track A: amount saved in a typical month (L4, FCFA), savers, by aspired sector")

* Credit access: accounts (L1, L2), current loan (L6), source of loan (L7)
tabout L1 P4_sector if crit_A == 1 using "$out/Results.xls", ///
    append c(freq col) format(0c 1p) layout(cb) style(xls) ///
    h1("Analysis 4 - Track A: has a mobile money account (L1), by aspired sector")

tabout L2 P4_sector if crit_A == 1 using "$out/Results.xls", ///
    append c(freq col) format(0c 1p) layout(cb) style(xls) ///
    h1("Analysis 4 - Track A: has a bank account (L2), by aspired sector")

tabout L6 P4_sector if crit_A == 1 using "$out/Results.xls", ///
    append c(freq col) format(0c 1p) layout(cb) style(xls) ///
    h1("Analysis 4 - Track A: currently has a loan or debt (L6), by aspired sector")

tabout L7 if crit_A == 1 & L6 == 1 using "$out/Results.xls", ///
    append oneway c(freq col) format(0c 1p) layout(cb) style(xls) ///
    h1("Analysis 4 - Track A: main source of the loan (L7), those with a loan")

*------------------------------------------------------------------------------*
* 5.2 Track B (existing enterprise): registration, employees, earnings
*------------------------------------------------------------------------------*
* Registration (H4; Don't know shown)
tabout H4 if crit_B == 1 using "$out/Results.xls", ///
    append oneway mi c(freq col) format(0c 1p) layout(cb) style(xls) ///
    h1("Analysis 4 - Track B: activity is registered (H4)")

* Paid workers excluding self (H4A; asked of employers and own-account workers)
gen byte workers_cat = 0 if H4A == 0
replace workers_cat = 1 if H4A == 1
replace workers_cat = 2 if inrange(H4A, 2, 4)
replace workers_cat = 3 if H4A >= 5 & !missing(H4A)
label define workers_cat 0 "None" 1 "1" 2 "2 to 4" 3 "5 or more"
label values workers_cat workers_cat
label var workers_cat "Number of paid workers (H4A, grouped)"

tabout workers_cat if crit_B == 1 using "$out/Results.xls", ///
    append oneway c(freq col) format(0c 1p) layout(cb) style(xls) ///
    h1("Analysis 4 - Track B: number of paid workers, excluding self (H4A)")

* Earnings from the activity last month (H5, FCFA); refusals (.r) excluded
tabout track if crit_B == 1 using "$out/Results.xls", ///
    append sum oneway c(N H5 p25 H5 median H5 p75 H5 mean H5) format(0c 0c 0c 0c 0c) style(xls) ///
    h1("Analysis 4 - Track B: earnings from the activity last month (H5, FCFA)")

* Bracket for those who would not give the amount (H5A)
tabout H5A if crit_B == 1 & H5 == .r using "$out/Results.xls", ///
    append oneway c(freq col) format(0c 1p) layout(cb) style(xls) ///
    h1("Analysis 4 - Track B: earnings bracket (H5A, FCFA), those who would not give the amount")

*------------------------------------------------------------------------------*
* 5.3 Track C (apprenticeship): shocks, resilience and household vulnerability
*------------------------------------------------------------------------------*
* Shock and resilience module (N)
tabout N1 if crit_C == 1 using "$out/Results.xls", ///
    append oneway c(freq col) format(0c 1p) layout(cb) style(xls) ///
    h1("Analysis 4 - Track C: household affected by a shock, last 12 months (N1)")

tabout N3 if crit_C == 1 & N1 == 1 using "$out/Results.xls", ///
    append oneway c(freq col) format(0c 1p) layout(cb) style(xls) ///
    h1("Analysis 4 - Track C: impact of the shock (N3), households with a shock")

tabout N5 if crit_C == 1 & N1 == 1 using "$out/Results.xls", ///
    append oneway c(freq col) format(0c 1p) layout(cb) style(xls) ///
    h1("Analysis 4 - Track C: time to recover from the shock (N5), households with a shock")

tabout N6 if crit_C == 1 using "$out/Results.xls", ///
    append oneway c(freq col) format(0c 1p) layout(cb) style(xls) ///
    h1("Analysis 4 - Track C: could raise 250,000 FCFA for an emergency (N6)")

tabout track if crit_C == 1 using "$out/Results.xls", ///
    append sum oneway c(N N7 median N7 mean N7) format(0c 0c 1c) style(xls) ///
    h1("Analysis 4 - Track C: ability to withstand an economic shock (N7, 1-10)")

* Household vulnerability score: count of 5 risk signs (0-5)
*   1 Could not raise 250,000 FCFA in an emergency (N6 = No, impossible)
*   2 Regularly relies on outside help for basic needs (M2 = Yes, regularly)
*   3 Nobody in the household earns (D4 = 0)
*   4 At least half the household is under 15 (D2 / D1 >= 0.5)
*   5 Few assets: 2 or fewer of the 10 in D5 (n_assets <= 2)
* Missing if any of the five is missing. No cutoff is set here: the full
* distribution is shown so the stipend threshold can be chosen.
gen byte vuln_1 = (N6 == 3)              if !missing(N6)
gen byte vuln_2 = (M2 == 1)              if !missing(M2)
gen byte vuln_3 = (D4 == 0)              if !missing(D4)
gen byte vuln_4 = (D2 / D1 >= 0.5)       if !missing(D1, D2) & D1 > 0
gen byte vuln_5 = (n_assets <= 2)        if !missing(n_assets)
label var vuln_1 "Could not raise 250,000 FCFA in an emergency"
label var vuln_2 "Regularly relies on outside help"
label var vuln_3 "Nobody in the household earns"
label var vuln_4 "At least half the household under 15"
label var vuln_5 "2 or fewer household assets"

gen byte vuln_score = vuln_1 + vuln_2 + vuln_3 + vuln_4 + vuln_5
label var vuln_score "Household vulnerability score (0-5)"

* Spot check
tab vuln_score if crit_C == 1, missing

tabout vuln_score if crit_C == 1 using "$out/Results.xls", ///
    append oneway c(freq col cum) format(0c 1p 1p) layout(cb) style(xls) ///
    h1("Analysis 4 - Track C: household vulnerability score (0-5): number and cumulative % at each score")

tabout track if crit_C == 1 using "$out/Results.xls", ///
    append sum oneway c(mean vuln_1 mean vuln_2 mean vuln_3 mean vuln_4 mean vuln_5) ///
    format(2 2 2 2 2) style(xls) ///
    h1("Analysis 4 - Track C: share with each vulnerability sign (N6 emergency funds, M2 outside help, D4 no earner, D2/D1 half under 15, D5 few assets)")

*==============================================================================*
* 6. ANALYSIS 5 - (next analysis goes here; tables use -append-)
*==============================================================================*

log close
