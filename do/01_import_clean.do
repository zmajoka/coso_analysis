/*==============================================================================
  COSO Youth Employment Program - Côte d'Ivoire
  Baseline phone survey (Questionnaire_COSO_V5, Survey Solutions)

  File    : 01_import_clean.do
  Purpose : Import the raw Excel export, clean it according to the
            questionnaire, label it (English), and save an analysis-ready
            dataset.

  Input   : $data/data_exploration.xlsx
  Outputs : $cleandata/coso_baseline_clean.dta  (analysis file, no PII)
            $cleandata/coso_baseline_pii.dta    (names / phone numbers only,
                                                 linked by interview__key)
            $out/01_import_clean.log

  Cleaning conventions
    - Yes/No questions (1=Oui, 2=Non) are recoded to 1=Yes, 0=No.
    - Special codes become extended missing values:
        .d  Don't know          (98, 9998, 998, 999998, 99998, "Ne sait pas")
        .r  Refused / won't say (99999, 999999, H2==10)
        .h  Respondent hung up  (-95, "Répondant a raccroché")
      Survey Solutions "not asked" codes (-999999999, ##N/A##) become
      system missing (.) or "".
    - All rows (call attempts) are kept. Use the flags consent, complete
      and final_obs to select the analysis sample.
    - Data quality problems are flagged (flag_*), never dropped or edited.
==============================================================================*/

version 17
clear all
set more off
set varabbrev off

*------------------------------------------------------------------------------*
* 0. Paths - set the project root once here
*------------------------------------------------------------------------------*

global data   "C:\Users\WB461621\OneDrive - WBG\SPJ\Cote dIvoire\ImpactEval\Data"
global cleandata "C:\Users\WB461621\OneDrive - WBG\SPJ\Cote dIvoire\ImpactEval\Data\Clean"
global out   "C:\Users\WB461621\OneDrive - WBG\SPJ\Cote dIvoire\ImpactEval\Output"

capture log close
log using "$out/01_cleaning.log", replace

*------------------------------------------------------------------------------*
* 1. Import
*------------------------------------------------------------------------------*
* The header row of the Excel file is misaligned: from B1 onward every value
* sits one column to the right of its header (the column under "B1" is empty
* and the last column, assignment__id, has no header). So the header row is
* not used: columns are read as A, B, C, ... and renamed below in the true
* order. Everything is read as text, so every conversion below is explicit.
import excel using "$data/data_exploration.xlsx", allstring clear
drop in 1                                  // the misaligned header row

* True column order (215 columns). Stops if the file has a different number
* of columns. B_blank is the empty column under the "B1" header.
rename (_all) ( ///
    interview__key interview__id cover_id cover_district cover_region localite ///
    HHA_debut A1 A2_date A2_heure A4 A4X A5_date A5_heure HHA_fin              ///
    HHB_debut B_blank B1 B2 B3 B4 B5 B6 B6A B6B B6C B6D B6E B6F consentement   ///
        B7 HHB_fin                                                              ///
    HHC_debut C1 C2A_Jour C2A_Mois C2A_annee C2 C3 C4 C5 C6 C7 C8 C9 C10 C11  ///
        age HHC_fin                                                             ///
    HHD_debut D1 D2 D3 D3X D4 D5__1 D5__2 D5__3 D5__4 D5__5 D5__6 D5__7       ///
        D5__8 D5__9 D5__10 D5__11 HHD_fin                                       ///
    HHE_debut E1 E2 E3 E3A EMPLOYE HHE_fin                                      ///
    HHF_debut F1 F1X F2 F2X F3 F3X F4 HHF_fin                                   ///
    HHG_debut G1 G2 G3 G4 G4A G5 G5A G6 G7 G7A HHG_fin                          ///
    HHH_debut H1 H2 H3 H3A H4 H4A H5 H5A HHH_fin                                ///
    HHI_debut I1 I2 I2X I3 I3X HHI_fin                                          ///
    HHJ_debut J0 J0A J0B J1 J1X J2 J3 HHJ_fin                                   ///
    HHK_debut K1 K2 K3 K3X K4 K4X K5__1 K5__2 K5__3 K5__4 K5__5 K5__6 K5__7   ///
        K5__9 K5X K6 K7 HHK_fin                                                 ///
    HHL_debut L1 L2 L3 L4 L5 L5X L6 L7 L7X L8 L9 HHL_fin                        ///
    HHM_debut M1 M2 M3 M4 HHM_fin                                               ///
    HHN_debut N1 N2 N2X N3 N4 N4X N5 N6 N7 HHN_fin                              ///
    HHO_debut O1 O2 O3 O4 O5 O6 HHO_fin                                         ///
    HHP_debut P1 P2 P3 P4 P4X P5 HHP_fin                                        ///
    HHQ_debut Q1 Q2 Q3 Q4 Q5 Q6 Q7 Q8 Q9 Q10 HHQ_fin                            ///
    HHR_debut R1 R1A R2 R2A R3 R3A R3B R4 R5 R6 R7 HHR_fin                      ///
    sssys_irnd has__errors interview__status assignment__id )
* (P4X is called P44X in the questionnaire; renamed here)

* The empty column must really be empty
count if B_blank != ""                     // should be 0
drop B_blank

* Drop fully empty rows that Excel sometimes carries
drop if strtrim(interview__key) == ""
isid interview__key

di as txt "Observations imported: " as res _N

*------------------------------------------------------------------------------*
* 2. Text clean-up (all variables are text here)
*------------------------------------------------------------------------------*
* Survey Solutions writes "." (not asked) for categorical questions and
* "##N/A##" for empty timestamps: both become "".
gen byte hungup_any = 0

foreach v of varlist interview__key-assignment__id {
    quietly replace `v' = stritrim(strtrim(`v'))
    quietly replace `v' = "" if `v' == "##N/A##" | `v' == "."
    quietly replace hungup_any = 1 if `v' == "-95"
}

* "Répondant a raccroché" typed in an open-text question -> hung up, blank.
* (Only open-text questions: A4 has a category "Le répondant a raccroché".)
foreach v in A4X C9 C10 C11 D3X F1X F2X F3X G5A I2X I3X J1X K3X K4X K5X ///
             L5X L7X N2X N4X O4 P1 P4X P5 {
    quietly replace hungup_any = 1 if strpos(lower(`v'), "raccroch") > 0
    quietly replace `v' = "" if strpos(lower(`v'), "raccroch") > 0
}
* (-95 in other questions becomes .h in section 4)

tab hungup_any

*------------------------------------------------------------------------------*
* 3. Variable types
*------------------------------------------------------------------------------*

* 3.1 Numeric questions
*     Special answers are exported as text: put back the questionnaire codes
*     (they become .d / .r in section 4).
replace C2A_Jour  = "98"     if C2A_Jour  == "Ne sait pas"
replace C2A_Mois  = "98"     if C2A_Mois  == "Ne sait pas"
replace C2A_annee = "9998"   if C2A_annee == "Ne sait pas"
replace G3        = "998"    if G3        == "Ne sait pas"
replace H3        = "99999"  if H3        == "Ne veut pas dire"
replace H5        = "999999" if H5        == "Ne veut pas dire"
replace L4        = "999998" if L4        == "Ne sait pas / ça varie"
replace L8        = "99998"  if L8        == "Ne sait pas"

*     Check the output: a line "contains nonnumeric characters; no replace"
*     means that variable has text values that must be looked at.
destring consentement C2A_Jour C2A_Mois C2A_annee C2 age D1 D2 D4       ///
    D5__1 D5__2 D5__3 D5__4 D5__5 D5__6 D5__7 D5__8 D5__9 D5__10 D5__11 ///
    EMPLOYE G3 G4A G6 H3 H4A H5                                         ///
    K5__1 K5__2 K5__3 K5__4 K5__5 K5__6 K5__7 K5__9                     ///
    L4 L8 N7 O2 P3 sssys_irnd has__errors, replace

* Q2 (times volunteered) is a TEXT question in CAPI but asks for a count.
* Values that are not numbers are listed, then set to missing.
tab Q2 if missing(real(Q2)) & Q2 != ""
destring Q2, replace force

* interview__status is exported as text (e.g. "Completed"); it stays text.

* 3.2 Categorical questions: exported as French answer text -> codes.
*     Each question gets a French label with the exact questionnaire wording
*     and codes; -noextend- stops the do-file if an answer text does not
*     match. If it stops, run -tab <variable>- and compare with the label.
*     (English labels replace the French ones in section 7.)

* Geography (official codes of the questionnaire)
label define district ///
    105 "DENGUELE" ///
    111 "SAVANES" ///
    113 "WOROBA" ///
    114 "ZANZAN"

label define region ///
    10524 "FOLON" ///
    10510 "KABADOUGOU" ///
    11120 "BAGOUE" ///
    11103 "PORO" ///
    11132 "TCHOLOGO" ///
    11319 "BAFING" ///
    11322 "BERE" ///
    11314 "WORODOUGOU" ///
    11423 "BOUNKANI" ///
    11408 "GONTOUGO"

label define dept ///
    10510034 "ODIENNE" ///
    10510067 "MADINANI" ///
    10510088 "SAMATIGUILA" ///
    10510100 "GBELEBAN" ///
    10510105 "SEGUELON" ///
    10524068 "MINIGNAN" ///
    10524081 "KANIASSO" ///
    11103029 "KORHOGO" ///
    11103075 "DIKODOUGOU" ///
    11103090 "SINEMATIALI" ///
    11103103 "M'BENGUE" ///
    11120015 "BOUNDIALI" ///
    11120044 "TENGRELA" ///
    11120083 "KOUTO" ///
    11132023 "FERKESSEDOUGOU" ///
    11132086 "OUANGOLODOUGOU" ///
    11132101 "KONG" ///
    11314039 "SEGUELA" ///
    11314080 "KANI" ///
    11319046 "TOUBA" ///
    11319082 "KORO" ///
    11319087 "OUANINOU" ///
    11322032 "MANKONO" ///
    11322069 "KOUNAHIRI" ///
    11322097 "DIANRA" ///
    11408010 "BONDOUKOU" ///
    11408043 "TANDA" ///
    11408070 "KOUN-FAO" ///
    11408089 "SANDEGUE" ///
    11408093 "TRANSUA" ///
    11423014 "BOUNA" ///
    11423063 "NASSIAN" ///
    11423076 "DOROPO" ///
    11423091 "TEHINI"

label define souspref ///
    1051003401 "BAKO" ///
    1051003402 "BOUGOUSSO" ///
    1051003403 "DIOULATIEDOUGOU" ///
    1051003404 "ODIENNE" ///
    1051003405 "TIEME" ///
    1051006701 "FENGOLO" ///
    1051006702 "MADINANI" ///
    1051006703 "N'GOLOBLASSO" ///
    1051008801 "KIMBIRILA-SUD" ///
    1051008802 "SAMATIGUILA" ///
    1051010001 "GBELEBAN" ///
    1051010002 "SEYDOUGOU" ///
    1051010003 "SAMANGO" ///
    1051010501 "GBONGAHA" ///
    1051010502 "SEGUELON" ///
    1052406801 "KIMBIRILA-NORD" ///
    1052406802 "MINIGNAN" ///
    1052406803 "SOKORO" ///
    1052406804 "TIENKO" ///
    1052408101 "GOULIA" ///
    1052408102 "KANIASSO" ///
    1052408103 "MAHANDIANA-SOKOURANI" ///
    1110302901 "DASSOUNGBOHO" ///
    1110302902 "KANOROBA" ///
    1110302903 "KARAKORO" ///
    1110302904 "KIEMOU" ///
    1110302905 "KONI" ///
    1110302906 "KORHOGO" ///
    1110302907 "KOMBOLOKOURA" ///
    1110302908 "KOMBORODOUGOU" ///
    1110302909 "LATAHA" ///
    1110302910 "NAFOUN" ///
    1110302911 "NAPIE" ///
    1110302912 "N'GANON" ///
    1110302913 "NIOFOIN" ///
    1110302914 "SIRASSO" ///
    1110302915 "SOHOUO" ///
    1110302916 "TIORONIARADOUGOU" ///
    1110307501 "BORON" ///
    1110307502 "DIKODOUGOU" ///
    1110307503 "GUIEMBE" ///
    1110309001 "BAHOUAKAHA" ///
    1110309002 "KAGBOLODOUGOU" ///
    1110309003 "SEDIOGO" ///
    1110309004 "SINEMATIALI" ///
    1110310301 "BOUGOU" ///
    1110310302 "KATIALI" ///
    1110310303 "KATOGO" ///
    1110310304 "M'BENGUE" ///
    1112001501 "BAYA" ///
    1112001502 "BOUNDIALI" ///
    1112001503 "GANAONI" ///
    1112001504 "KASSERE" ///
    1112001505 "SIEMPURGO" ///
    1112004401 "DEBETE" ///
    1112004402 "KANAKONO" ///
    1112004403 "PAPARA" ///
    1112004404 "TENGRELA" ///
    1112008301 "BLESSEGUE" ///
    1112008302 "GBON" ///
    1112008303 "KOLIA" ///
    1112008304 "KOUTO" ///
    1112008305 "SIANHALA" ///
    1113202301 "FERKESSEDOUGOU" ///
    1113202302 "KOUMBALA" ///
    1113202303 "TOGONIERE" ///
    1113208601 "DIAWALA" ///
    1113208602 "KAOUARA" ///
    1113208603 "NIELLE" ///
    1113208604 "OUANGOLODOUGOU" ///
    1113208605 "TOUMOUKORO" ///
    1113210101 "BILIMONO" ///
    1113210102 "KONG" ///
    1113210103 "NAFANA(DE KONG)" ///
    1113210104 "SIKOLO" ///
    1131403901 "BOBI" ///
    1131403902 "DIARABANA" ///
    1131403903 "DUALLA" ///
    1131403904 "KAMALO" ///
    1131403905 "MASSALA" ///
    1131403906 "SEGUELA" ///
    1131403907 "SIFIE" ///
    1131403908 "WOROFLA" ///
    1131408001 "DJIBROSSO" ///
    1131408002 "FADIADOUGOU" ///
    1131408003 "KANI" ///
    1131408004 "MORONDO" ///
    1131904601 "DIOMAN" ///
    1131904602 "FOUNGBESSO" ///
    1131904603 "GUINTEGUELA" ///
    1131904604 "TOUBA" ///
    1131908201 "BOOKO" ///
    1131908202 "BOROTOU" ///
    1131908203 "KORO" ///
    1131908204 "MAHANDOUGOU" ///
    1131908205 "NIOKOSSO" ///
    1131908701 "GBELO" ///
    1131908702 "GOUEKAN" ///
    1131908703 "KOONAN" ///
    1131908704 "OUANINOU" ///
    1131908705 "SABOUDOUGOU" ///
    1131908706 "SANTA (DE OUANINOU)" ///
    1132203201 "BOUANDOUGOU" ///
    1132203202 "MANKONO" ///
    1132203203 "MARHANDALLAH" ///
    1132203204 "SARHALA" ///
    1132203205 "TIENINGBOUE" ///
    1132206901 "KONGASSO" ///
    1132206902 "KOUNAHIRI" ///
    1132209701 "DIANRA" ///
    1132209702 "DIANRA VILLAGE" ///
    1140801001 "APPIMANDOUM" ///
    1140801002 "PINDA-BOROKO" ///
    1140801003 "BONDO" ///
    1140801004 "BONDOUKOU" ///
    1140801005 "GOUMERE" ///
    1140801006 "LAOUDI BA" ///
    1140801007 "SAPLI-SEPINGO" ///
    1140801008 "SOROBANGO" ///
    1140801009 "TABAGNE" ///
    1140801010 "TAGADI" ///
    1140801011 "TAOUDI" ///
    1140801012 "YEZIMALA" ///
    1140804301 "AMANVI" ///
    1140804302 "DIAMBA" ///
    1140804303 "TANDA" ///
    1140804304 "TIEDIO" ///
    1140807001 "BOAHIA" ///
    1140807002 "KOKOMIAN" ///
    1140807003 "KOUASSI-DATEKRO" ///
    1140807004 "KOUN-FAO" ///
    1140807005 "TANKESSE" ///
    1140807006 "TIENKOIKRO" ///
    1140808901 "BANDAKAGNI-TOMORA" ///
    1140808902 "DIMANDOUGOU" ///
    1140808903 "SANDEGUE" ///
    1140808904 "YOROBODI" ///
    1140809301 "ASSUEFRY" ///
    1140809302 "KOUASSIA-NIAGUINI" ///
    1140809303 "TRANSUA" ///
    1142301401 "BOUKO" ///
    1142301402 "BOUNA" ///
    1142301403 "ONDEFIDOUO" ///
    1142301404 "YOUNDOUO" ///
    1142306301 "BOGOFA" ///
    1142306302 "KOTOUBA" ///
    1142306303 "KAKPIN" ///
    1142306304 "NASSIAN" ///
    1142306305 "SOMINASSE" ///
    1142307601 "DANOA" ///
    1142307602 "DOROPO" ///
    1142307603 "KALAMON" ///
    1142307604 "NIAMOUE" ///
    1142309101 "GOGO" ///
    1142309102 "TEHINI" ///
    1142309103 "TOUGBO"
label copy district fr_B1
label copy region   fr_B2
label copy dept     fr_B3
label copy souspref fr_B4
label define fr_B5 1 "Urbain" 2 "Rural"

* Yes/No
label define fr_yesno 1 "Oui" 2 "Non"

* Other questions
label define fr_A1 1 "Première tentative" 2 "Deuxième tentative"             ///
    3 "Troisième tentative ou plus"
label define fr_A4 1 "Entretien réalisé (compléter le questionnaire)"        ///
    2 "Pas de réponse" 3 "Numéro invalide ou incorrect" 4 "Ligne occupée"    ///
    5 "Rendez-vous fixé pour un rappel" 6 "Le répondant a refusé de participer" ///
    7 "Le répondant a raccroché / appel interrompu en cours d’entretien"      ///
    8 "Autre (préciser)"
label define fr_C1 1 "Homme" 2 "Femme"
label define fr_C3 1 "Célibataire" 2 "Marié(e) monogame" 3 "Marié(e) polygame" ///
    4 "Union libre/Concubinage" 5 "Divorcé(e)" 6 "Veuf(ve)"
label define fr_C4 0 "Aucun niveau" 1 "Préscolaire" 2 "Primaire"             ///
    3 "Secondaire général cycle I"                                           ///
    4 "Secondaire technique/professionnel cycle I"                           ///
    5 "Secondaire général cycle II"                                          ///
    6 "Secondaire technique/professionnel cycle II"                          ///
    7 "Supérieur cycle court (BAC+ 1; BP; BTS1; BAC+2; BTS2; DUT)"           ///
    8 "Licence" 9 "Maitrise/ Master 1" 10 "Master 2/DEA/DESS" 11 "Doctorat"  ///
    12 "Post Doctorat"
label define fr_C5 1 "AUCUN" 2 "CEPE" 3 "BEPC/BEP/CAP" 4 "BAC/BT" 5 "BTS/DEUG 2" ///
    6 "LICENCE/BACHELOR" 7 "MASTER/DESS" 8 "DOCTORATS"
label define fr_C8 1 "Côte d’Ivoire" 2 "Hors Côte d'Ivoire"
label define fr_D3 1 "Chef de ménage" 2 "Conjoint du CM" 3 "Enfant du CM"    ///
    4 "Autre parent du CM" 5 "Sans lien de parenté avec le CM" 6 "Amis du CM" ///
    7 "Autre à préciser"
label define fr_E3A 1 "Principalement à la vente"                           ///
    2 "Principalement à l'usage du ménage"
label define fr_F1 1 "Salarié" 2 "Employeur (avec employés)"                ///
    3 "Travailleur indépendant / à compte propre (sans employés)"            ///
    4 "Membre de coopératives de producteurs"                                ///
    5 "Travailleurs familiaux collaborant à l’entreprise familiale"          ///
    7 "Apprenti/Stagiaire" 9 "Autre (Précisez)"
label define fr_F2 1 "Agriculture, élevage, pêche, sylviculture"            ///
    2 "Industries extractives" 3 "Industrie manufacturière"                  ///
    4 "Construction/BTP" 5 "Commerce de gros ou de détail"                   ///
    6 "Transport, entreposage, logistique" 7 "Hébergement et restauration"   ///
    8 "Information, communication, services numériques"                      ///
    9 "Activités financières et d’assurance" 10 "Administration publique"   ///
    11 "Éducation" 12 "Santé humaine et action sociale"                       ///
    13 "Travail domestique chez un particulier"                               ///
    14 "Autres services (réparation, services personnels, arts/spectacles)"  ///
    99 "Autre (préciser)"
label copy fr_F2 fr_P4
label define fr_F3 1 "À domicile" 2 "Dans une structure attenante/proche du domicile" ///
    3 "Au domicile de l’employeur"                                           ///
    4 "Lieu fixe hors domicile, en intérieur (bureau, magasin, usine)"       ///
    5 "Étal/emplacement fixe dans un espace public (stand de marché, kiosque, emplacement de vente autorisé)" ///
    6 "Chantier de construction" 7 "Champ/exploitation agricole"             ///
    8 "Mobile/sans emplacement fixe (vente itinérante, porte-à-porte)"       ///
    9 "Autre (préciser)"
label define fr_F4 1 "Moins de 6 mois" 2 "6 mois à moins d’un an"           ///
    3 "1 an à moins de 2 ans" 4 "2 ans à moins de 5 ans" 5 "5 ans ou plus"
label define fr_G1 1 "Oui, à durée indéterminée (CDI)"                      ///
    2 "Oui, à durée déterminée (CDD)" 3 "Oui, consultance"                   ///
    4 "Accord verbal seulement" 5 "Aucun accord"
label define fr_H1 1 "En espèces uniquement" 2 "En nature uniquement" 3 "Les deux"
label define fr_H2 1 "Quotidienne/Jour" 2 "Hebdomadaire/Semaine"             ///
    3 "Quinzaine/Toutes les deux semaines" 4 "Mensuelle/Mois"                ///
    5 "Bimestre/Tous les deux mois" 6 "Trimestre/Tous les trois mois"        ///
    7 "Semestre/Tous les six mois" 8 "Année" 9 "Irrégulière" 10 "Ne veut pas dire"
label define fr_H3A 1 "Moins de 30 000" 2 "[30 000 - 75 000[" 3 "[75 000 - 150 000[" ///
    4 "[150 000 - 300 000[" 5 "[300 000 - 500 000[" 6 "[500 000 - 800 000[" ///
    7 "[800 000 - 1 000 000[" 8 "Plus d'un million"
label copy fr_H3A fr_H5A
label define fr_H4 1 "Oui" 2 "Non" 3 "Ne sait pas"
label define fr_I2 1 "Revenu insuffisant dans l’emploi actuel"              ///
    2 "Emploi actuel instable/temporaire"                                    ///
    3 "Je veux un emploi qui correspond mieux à mes compétences"             ///
    4 "Je veux simplement un meilleur emploi" 5 "Autre"
label define fr_I3 1 "Candidature spontanée auprès d’employeurs"            ///
    2 "Réponse à une offre d’emploi (en ligne, journal, affiche)"            ///
    3 "Inscription auprès d’une agence (Agence Emploi Jeunes, etc.)"         ///
    4 "Réseau personnel (famille, amis)" 5 "Concours/test pour le secteur public" ///
    6 "Démarches pour créer sa propre entreprise" 9 "Autre"
label copy fr_I3 fr_J1
label define fr_J0A 1 "Vous ne pensiez pas en trouver, vous étiez découragé(e)" ///
    2 "Vous pensiez qu'il n'y avait pas d'emploi disponible correspondant à vos compétences dans votre zone" ///
    3 "Vous suiviez (ou aviez le projet de suivre) des études ou une formation" ///
    4 "Problèmes de santé ou situation de handicap"                          ///
    5 "Obligations familiales (enfants, proches)"                            ///
    6 "Vous attendiez le début de la saison agricole"                        ///
    7 "Vous attendiez le résultat de démarches antérieures (emploi, concours, programme)" ///
    8 "Vous avez déjà trouvé un emploi ou une activité qui commencera dans les trois prochains mois" ///
    9 "Autre (préciser)"
label define fr_J2 1 "Moins d’1 mois" 2 "1 mois à moins de 3 mois"           ///
    3 "3 mois à moins de 6 mois" 4 "6 mois à moins de 12 mois"               ///
    5 "1 an à moins de 2 ans" 6 "2 ans ou plus"
label define fr_K3 1 "Épargne personnelle" 2 "Aide familiale"                ///
    3 "Microfinance/coopérative d’épargne (COOPEC)"                          ///
    4 "Programme gouvernemental (AGR/Agence Emploi Jeunes)" 5 "Prêt bancaire" ///
    6 "ONG/organisation humanitaire" 9 "Autre"
label copy fr_K3 fr_K4
label define fr_K4 7 "Ne sait pas", add
label define fr_K7 1 "Très faibles" 2 "Faibles" 3 "Moyennes" 4 "Bonnes" 5 "Très bonnes"
label define fr_L3 1 "Oui, régulièrement" 2 "Oui, occasionnellement" 3 "Non, jamais"
label define fr_L5 1 "Domicile (cash)" 2 "Tontine/AVEC" 3 "Compte bancaire"  ///
    4 "Mobile money" 5 "Autres (à préciser)"
label define fr_L7 1 "Banque" 2 "Institution de microfinance/COOPEC"         ///
    3 "Crédit mobile money" 4 "Tontine/AVEC" 5 "Famille ou amis" 6 "Employeur" ///
    7 "Prêteur informel" 9 "Autre"
label define fr_L9 1 "Démarrer ou développer une activité"                  ///
    2 "Couvrir des besoins quotidiens/du ménage" 3 "Éducation/frais de scolarité" ///
    4 "Santé/dépenses médicales" 5 "Logement" 9 "Autre"
label define fr_M1 1 "Je dépends principalement de l’aide d’autrui"         ///
    2 "Je me prends en charge pour certains de mes besoins"                  ///
    3 "Je suis capable de me prendre en charge complètement"
label copy fr_M1 fr_M4
label define fr_M2 1 "Oui, régulièrement" 2 "Oui, ponctuellement" 3 "Non, jamais"
label define fr_M3 1 "Pas du tout d’accord" 2 "Plutôt pas d’accord" 3 "Neutre" ///
    4 "Plutôt d’accord" 5 "Tout à fait d’accord"
label define fr_N2 1 "Maladie/accident" 2 "Choc climatique (sécheresse, inondation)" ///
    3 "Perte d'activité/emploi" 4 "Décès d'un membre du ménage"              ///
    5 "Catastrophe naturelle" 9 "Autre (préciser)"
label define fr_N3 1 "Aucun impact" 2 "Impact limité, vite surmonté"          ///
    3 "Impact important, encore en cours de redressement"                    ///
    4 "Impact très important, activité arrêtée ou fortement réduite"
label define fr_N4 1 "Épargne personnelle" 2 "Vente d'actifs/biens"          ///
    3 "Emprunt/crédit" 4 "Aide familiale ou communautaire"                   ///
    5 "Aide externe (ONG, État)" 6 "Réduction de la consommation" 9 "Autre"
label define fr_N5 1 "Moins d'1 mois" 2 "1 à 3 mois" 3 "3 à 6 mois"          ///
    4 "Plus de 6 mois" 5 "Pas encore rétabli"
label define fr_N6 1 "Oui, facilement" 2 "Oui, avec difficulté" 3 "Non, impossible"
label define fr_O5 1 "École/centre de formation professionnelle formel"     ///
    2 "Employeur/entreprise privée"                                          ///
    3 "Programme gouvernemental (ex. Agence Emploi Jeunes)"                  ///
    4 "ONG/programme d’alphabétisation ou de compétences de vie"             ///
    5 "Apprentissage informel (famille/membre de la communauté)"             ///
    6 "Apprentissage sur le tas/au travail" 9 "Autre"
label define fr_O6 1 "Pas utile" 2 "Plutôt utilie" 3 "Très utile"
label define fr_P2 1 "Trouver un emploi salarié" 2 "Créer une entreprise"   ///
    3 "Développer une activité existante"
label define fr_Q3 1 "On peut faire confiance à la plupart des gens"        ///
    2 "Il faut être prudent"
label define fr_Q4 1 "Pas du tout confiance" 2 "Juste un peu confiance"      ///
    3 "Partiellement confiance" 4 "Beaucoup confiance" 9 "Ne sait pas"
label copy fr_Q4 fr_Q5
label copy fr_Q4 fr_Q6
label copy fr_Q4 fr_Q7
label copy fr_Q4 fr_Q8
label copy fr_Q4 fr_Q9
label copy fr_Q4 fr_Q10
label define fr_R7 1 "Vous-même" 2 "Un parent" 3 "Un ami/connaissance"

* Convert: yes/no questions
foreach v in B6 B6A B6B B6D B6E C6 C7 E1 E2 E3 G2 G4 G5 G7 G7A I1 J0 ///
             J0B J3 K1 K2 K6 L1 L2 L6 N1 O1 O3 Q1 R1 R2 R3 R3A R5 {
    encode `v', gen(`v'_n) label(fr_yesno) noextend
    order `v'_n, after(`v')
    drop `v'
    rename `v'_n `v'
}

* Convert: all other categorical questions (label fr_<variable>)
foreach v in B1 B2 B3 B4 B5 A1 A4 C1 C3 C4 C5 C8 D3 E3A F1 F2 F3 F4 G1 ///
             H1 H2 H3A H4 H5A I2 I3 J0A J1 J2 K3 K4 K7 L3 L5 L7 L9      ///
             M1 M2 M3 M4 N2 N3 N4 N5 N6 O5 O6 P2 P4 Q3 Q4 Q5 Q6 Q7 Q8  ///
             Q9 Q10 R7 {
    encode `v', gen(`v'_n) label(fr_`v') noextend
    order `v'_n, after(`v')
    drop `v'
    rename `v'_n `v'
}

* 3.3 Timestamps (Survey Solutions "current time", e.g. 2026-08-13T11:20:08)
*     -> Stata %tc
foreach v in HHA_debut HHA_fin HHB_debut HHB_fin HHC_debut HHC_fin            ///
             HHD_debut HHD_fin HHE_debut HHE_fin HHF_debut HHF_fin            ///
             HHG_debut HHG_fin HHH_debut HHH_fin HHI_debut HHI_fin            ///
             HHJ_debut HHJ_fin HHK_debut HHK_fin HHL_debut HHL_fin            ///
             HHM_debut HHM_fin HHN_debut HHN_fin HHO_debut HHO_fin            ///
             HHP_debut HHP_fin HHQ_debut HHQ_fin HHR_debut HHR_fin {
    gen double `v'_c = clock(substr(subinstr(`v', "T", " ", 1), 1, 19), "YMDhms")
    count if missing(`v'_c) & `v' != ""        // should be 0
    format `v'_c %tcDD/NN/CCYY_HH:MM:SS
    order `v'_c, after(`v')
    drop `v'
    rename `v'_c `v'
}

* 3.4 Dates (exported month/day/year, e.g. 8/13/2026) -> Stata %td
*     A5_heure is a DATE question in CAPI (despite its name): read as a date.
foreach v in A2_date A5_date A5_heure B6C B6F {
    gen `v'_c = date(substr(`v', 1, 10), "MDY")
    count if missing(`v'_c) & `v' != ""        // should be 0
    format `v'_c %tdDD/NN/CCYY
    order `v'_c, after(`v')
    drop `v'
    rename `v'_c `v'
}

* 3.5 What is still text should be only the open-ended / ID variables
*     (plus A2_heure, typed by the interviewer, and interview__status)
ds, has(type string)

*------------------------------------------------------------------------------*
* 4. Missing-value codes
*------------------------------------------------------------------------------*

* Survey Solutions "not asked" (-999999999) -> .   and   hung up (-95) -> .h
ds, has(type numeric)
mvdecode `r(varlist)', mv(-999999999 = . \ -95 = .h)

* Hung up (-95) typed in a text question -> blank
ds, has(type string)
foreach v in `r(varlist)' {
    replace `v' = "" if `v' == "-95"
}

* Don't know (.d)
replace C2A_Jour  = .d if C2A_Jour  == 98
replace C2A_Mois  = .d if C2A_Mois  == 98
replace C2A_annee = .d if C2A_annee == 9998
replace G3        = .d if G3        == 998
replace L4        = .d if L4        == 999998    // "Ne sait pas / ça varie"
replace L8        = .d if L8        == 99998
replace H4        = .d if H4        == 3
replace K4        = .d if K4        == 7
replace Q4        = .d if Q4        == 9
replace Q5        = .d if Q5        == 9
replace Q6        = .d if Q6        == 9
replace Q7        = .d if Q7        == 9
replace Q8        = .d if Q8        == 9
replace Q9        = .d if Q9        == 9
replace Q10       = .d if Q10       == 9

* Refused / prefers not to say (.r)
replace H2 = .r if H2 == 10
replace H3 = .r if H3 == 99999
replace H5 = .r if H5 == 999999

*------------------------------------------------------------------------------*
* 5. Yes/No questions: 1=Oui, 2=Non  ->  1=Yes, 0=No
*------------------------------------------------------------------------------*
* Spot check before recoding: every variable should have min 1 and max 2
summarize B6 B6A B6B B6D B6E C6 C7 E1 E2 E3 G2 G4 G5 G7 G7A H4 I1 J0 ///
    J0B J3 K1 K2 K6 L1 L2 L6 N1 O1 O3 Q1 R1 R2 R3 R3A R5

recode B6 B6A B6B B6D B6E C6 C7 E1 E2 E3 G2 G4 G5 G7 G7A H4 I1 J0 ///
    J0B J3 K1 K2 K6 L1 L2 L6 N1 O1 O3 Q1 R1 R2 R3 R3A R5 (2 = 0)

* Multi-select dummies (D5, K5) are exported as 1/0: min 0, max 1 expected
summarize D5__* K5__*

*------------------------------------------------------------------------------*
* 6. Constructed variables
*------------------------------------------------------------------------------*

* 6.1 Consent
* The CAPI variable -consentement- equals 1 whenever B6B and B6E are not "No",
* so it is 1 for unreached respondents and for callbacks. Rebuild it from B6D.
rename consentement consentement_capi
gen byte consent = B6D if inlist(B6D, 0, 1)
label var consent "Consented to the interview (B6D)"

* 6.2 Completion and one row per respondent
gen byte complete = (A4 == 1 & consent == 1 & !missing(HHR_fin))
label var complete "Completed interview (A4 = done, consented, reached section R)"

gen byte partial = (consent == 1 & missing(HHR_fin))
label var partial "Consented but interview not finished"

bysort cover_id: gen n_rows_id = _N if cover_id != ""
label var n_rows_id "Number of rows (attempts) with this cover_id"

* final_obs: best row per cover_id = completed first, then latest attempt
gen double _s_att  = cond(missing(A1), -1, A1)
gen double _s_date = cond(missing(HHA_debut), -1, HHA_debut)
gsort cover_id -complete -_s_att -_s_date
by cover_id: gen byte final_obs = (_n == 1) if cover_id != ""
replace final_obs = 1 if cover_id == ""
drop _s_att _s_date
label var final_obs "Selected row for this respondent (completed, else latest attempt)"

* 6.3 Age: use full date of birth where known, CAPI fallback otherwise
*     (CAPI -age- = C2, else 2026 - birth year)
rename age age_capi
gen dob = mdy(C2A_Mois, C2A_Jour, C2A_annee)
format dob %tdDD/NN/CCYY
label var dob "Date of birth (C2A)"

gen int_date = A2_date
replace int_date = dofc(HHA_debut) if missing(int_date)
format int_date %tdDD/NN/CCYY
label var int_date "Interview date (A2, else section A start time)"

gen age_clean = floor((int_date - dob) / 365.25) if !missing(dob, int_date)
replace age_clean = year(int_date) - C2A_annee if missing(age_clean) & ///
    !missing(C2A_annee, int_date)
replace age_clean = C2 if missing(age_clean) & !missing(C2)
label var age_clean "Age in completed years (constructed)"

* 6.4 Sex
gen byte female = (C1 == 2) if inlist(C1, 1, 2)
label var female "Female"

* 6.5 Employment status (recomputes CAPI -EMPLOYE-)
gen byte employed = (E1 == 1 | E2 == 1 | (E3 == 1 & E3A == 1)) if !missing(E1)
label var employed "Employed last 7 days (E1/E2/E3, recomputed)"

* 6.6 Assets
egen n_assets = rowtotal(D5__1 D5__2 D5__3 D5__4 D5__5 D5__6 D5__7 D5__8 ///
    D5__9 D5__10) if !missing(D5__1)
label var n_assets "Number of household assets owned (D5, out of 10)"

* 6.7 Hours
gen hours_total = G3 + cond(G5 == 1 & !missing(G6), G6, 0) if !missing(G3)
label var hours_total "Hours worked last week, all jobs (G3 + G6)"

* 6.8 Monthly labour income (FCFA)
*     Wage: H3 converted to monthly by pay frequency H2.
*     Daily pay assumes 26 working days a month; irregular pay is left missing.
gen wage_month = .
replace wage_month = H3 * 26        if H2 == 1
replace wage_month = H3 * 52 / 12   if H2 == 2
replace wage_month = H3 * 26 / 12   if H2 == 3
replace wage_month = H3             if H2 == 4
replace wage_month = H3 / 2         if H2 == 5
replace wage_month = H3 / 3         if H2 == 6
replace wage_month = H3 / 6         if H2 == 7
replace wage_month = H3 / 12        if H2 == 8
replace wage_month = .r             if H3 == .r
label var wage_month "Monthly wage, FCFA (H3 by H2 frequency)"

gen inc_month = wage_month if F1 == 1
replace inc_month = H5 if inlist(F1, 2, 3, 4)
label var inc_month "Monthly labour income, FCFA (wage or self-employment earnings)"

* 6.9 Durations (minutes)
gen dur_total = (HHR_fin - HHB_debut) / 60000
label var dur_total "Interview length, minutes (section B start to R end)"
foreach s in A B C D E F G H I J K L M N O P Q R {
    gen dur_`s' = (HH`s'_fin - HH`s'_debut) / 60000
    label var dur_`s' "Section `s' duration, minutes"
}

*------------------------------------------------------------------------------*
* 7. Value labels (English)
*------------------------------------------------------------------------------*
label define yesno 0 "No" 1 "Yes"

label define A1 1 "First attempt" 2 "Second attempt" 3 "Third attempt or more"
label define A4 1 "Interview completed" 2 "No answer" 3 "Invalid/incorrect number" ///
    4 "Line busy" 5 "Callback appointment set" 6 "Respondent refused"          ///
    7 "Respondent hung up / call interrupted" 8 "Other"

label define C1 1 "Male" 2 "Female"
label define C3 1 "Single" 2 "Married, monogamous" 3 "Married, polygamous"   ///
    4 "Cohabiting / free union" 5 "Divorced" 6 "Widowed"
label define C4 0 "None" 1 "Preschool" 2 "Primary"                          ///
    3 "Lower secondary, general" 4 "Lower secondary, technical/vocational"  ///
    5 "Upper secondary, general" 6 "Upper secondary, technical/vocational"  ///
    7 "Short tertiary (BAC+1/+2, BP, BTS, DUT)" 8 "Licence (Bachelor)"      ///
    9 "Maitrise / Master 1" 10 "Master 2 / DEA / DESS" 11 "Doctorate"       ///
    12 "Post-doctorate"
label define C5 1 "None" 2 "CEPE" 3 "BEPC/BEP/CAP" 4 "BAC/BT" 5 "BTS/DEUG" ///
    6 "Licence/Bachelor" 7 "Master/DESS" 8 "Doctorate"
label define C8 1 "Côte d'Ivoire" 2 "Outside Côte d'Ivoire"
label define D3 1 "Household head" 2 "Spouse of head" 3 "Child of head"    ///
    4 "Other relative of head" 5 "Not related to head" 6 "Friend of head" 7 "Other"
label define E3A 1 "Mainly for sale" 2 "Mainly for household use"
label define F1 1 "Employee" 2 "Employer (with employees)"                  ///
    3 "Own-account worker (no employees)" 4 "Member of producers' cooperative" ///
    5 "Contributing family worker" 7 "Apprentice / intern" 9 "Other"
label define sector 1 "Agriculture, livestock, fishing, forestry"          ///
    2 "Mining and quarrying" 3 "Manufacturing" 4 "Construction"             ///
    5 "Wholesale / retail trade" 6 "Transport, storage, logistics"          ///
    7 "Accommodation and food services" 8 "Information, communication, digital" ///
    9 "Finance and insurance" 10 "Public administration" 11 "Education"     ///
    12 "Health and social work" 13 "Domestic work in private households"   ///
    14 "Other services (repair, personal, arts)" 99 "Other"
label define F3 1 "At home" 2 "Structure attached to / near home"          ///
    3 "Employer's home" 4 "Fixed premises outside home, indoors"            ///
    5 "Fixed stall in public space" 6 "Construction site"                   ///
    7 "Field / farm" 8 "Mobile / no fixed location" 9 "Other"
label define F4 1 "Less than 6 months" 2 "6 months to <1 year"             ///
    3 "1 to <2 years" 4 "2 to <5 years" 5 "5 years or more"
label define G1 1 "Yes, permanent (CDI)" 2 "Yes, fixed-term (CDD)"         ///
    3 "Yes, consultancy" 4 "Verbal agreement only" 5 "No agreement"
label define H1 1 "Cash only" 2 "In kind only" 3 "Both"
label define H2 1 "Daily" 2 "Weekly" 3 "Every two weeks" 4 "Monthly"       ///
    5 "Every two months" 6 "Quarterly" 7 "Every six months" 8 "Yearly"     ///
    9 "Irregular"
label define bracket 1 "Less than 30,000" 2 "30,000 - <75,000"             ///
    3 "75,000 - <150,000" 4 "150,000 - <300,000" 5 "300,000 - <500,000"    ///
    6 "500,000 - <800,000" 7 "800,000 - <1,000,000" 8 "1,000,000 or more"
label define I2 1 "Insufficient income in current job"                     ///
    2 "Current job unstable / temporary" 3 "Want a job that better fits my skills" ///
    4 "Simply want a better job" 5 "Other"
label define search 1 "Unsolicited application to employers"              ///
    2 "Answered a job ad (online, newspaper, poster)"                       ///
    3 "Registered with an agency (e.g. Agence Emploi Jeunes)"               ///
    4 "Personal network (family, friends)" 5 "Public-sector exam/test"      ///
    6 "Steps to start own business" 9 "Other"
label define J0A 1 "Discouraged, did not think would find"                 ///
    2 "No jobs matching skills in area" 3 "In (or planning) education/training" ///
    4 "Health problems / disability" 5 "Family obligations"                 ///
    6 "Waiting for agricultural season" 7 "Waiting for results of earlier steps" ///
    8 "Found a job starting within 3 months" 9 "Other"
label define J2 1 "Less than 1 month" 2 "1 to <3 months" 3 "3 to <6 months" ///
    4 "6 to <12 months" 5 "1 to <2 years" 6 "2 years or more"
label define funding 1 "Personal savings" 2 "Family support"               ///
    3 "Microfinance / savings cooperative (COOPEC)"                         ///
    4 "Government programme (AGR / Agence Emploi Jeunes)" 5 "Bank loan"     ///
    6 "NGO / humanitarian organisation" 9 "Other"
label define K7 1 "Very weak" 2 "Weak" 3 "Average" 4 "Good" 5 "Very good"
label define freq3 1 "Yes, regularly" 2 "Yes, occasionally" 3 "No, never"
label define L5 1 "At home (cash)" 2 "Tontine / VSLA (AVEC)" 3 "Bank account" ///
    4 "Mobile money" 5 "Other"
label define L7 1 "Bank" 2 "Microfinance institution / COOPEC"             ///
    3 "Mobile money credit" 4 "Tontine / VSLA (AVEC)" 5 "Family or friends" ///
    6 "Employer" 7 "Informal lender" 9 "Other"
label define L9 1 "Start or grow a business" 2 "Daily / household needs"   ///
    3 "Education / school fees" 4 "Health / medical costs" 5 "Housing" 9 "Other"
label define selfsuff 1 "Mainly dependent on help from others"             ///
    2 "Self-supporting for some needs" 3 "Fully self-supporting"
label define agree5 1 "Strongly disagree" 2 "Somewhat disagree" 3 "Neutral" ///
    4 "Somewhat agree" 5 "Strongly agree"
label define N2 1 "Illness / accident" 2 "Climate shock (drought, flood)"  ///
    3 "Loss of business / job" 4 "Death of a household member"             ///
    5 "Natural disaster" 9 "Other"
label define N3 1 "No impact" 2 "Limited impact, quickly overcome"         ///
    3 "Major impact, still recovering" 4 "Very major impact, activity stopped or sharply reduced"
label define N4 1 "Personal savings" 2 "Sale of assets" 3 "Borrowing / credit" ///
    4 "Family or community help" 5 "External aid (NGO, State)"              ///
    6 "Reduced consumption" 9 "Other"
label define N5 1 "Less than 1 month" 2 "1 to 3 months" 3 "3 to 6 months"  ///
    4 "More than 6 months" 5 "Not yet recovered"
label define N6 1 "Yes, easily" 2 "Yes, with difficulty" 3 "No, impossible"
label define N7 1 "1 (lowest)" 10 "10 (highest)"
label define O5 1 "Formal vocational school / training centre"            ///
    2 "Employer / private firm" 3 "Government programme (e.g. Agence Emploi Jeunes)" ///
    4 "NGO / literacy or life-skills programme"                             ///
    5 "Informal apprenticeship (family / community)" 6 "On-the-job learning" 9 "Other"
label define O6 1 "Not useful" 2 "Somewhat useful" 3 "Very useful"
label define P2 1 "Find a wage job" 2 "Start a business" 3 "Grow an existing activity"
label define Q3 1 "Most people can be trusted" 2 "Need to be very careful"
label define trust 1 "Not at all" 2 "Just a little" 3 "Somewhat" 4 "A lot"
label define R7 1 "Self" 2 "A relative" 3 "A friend / acquaintance"
label define missing_only .d "Don't know" .r "Refused / prefers not to say" ///
    .h "Respondent hung up"
label define status -1 "Deleted" 0 "Restored" 20 "Created"                 ///
    40 "SupervisorAssigned" 60 "InterviewerAssigned"                         ///
    65 "RejectedBySupervisor" 80 "ReadyForInterview" 85 "SentToCapi"         ///
    95 "Restarted" 100 "Completed" 120 "ApprovedBySupervisor"                ///
    125 "RejectedByHeadquarters" 130 "ApprovedByHeadquarters"

label define B5 1 "Urban" 2 "Rural"

* Drop the French labels used for importing (section 3.2)
label dir
foreach l in `r(names)' {
    if substr("`l'", 1, 3) == "fr_" label drop `l'
}

* Add the extended-missing labels to every value label
label dir
foreach l in `r(names)' {
    label define `l' .d "Don't know" .r "Refused / prefers not to say" ///
        .h "Respondent hung up", modify
}

* Attach value labels
label values B1 district
label values B2 region
label values B3 dept
label values B4 souspref
label values B5 B5
label values B6 B6A B6B B6D B6E C6 C7 E1 E2 E3 G2 G4 G5 G7 G7A H4 I1 J0 ///
    J0B J3 K1 K2 K6 L1 L2 L6 N1 O1 O3 Q1 R1 R2 R3 R3A R5 yesno
label values D5__* K5__* yesno
label values consent complete partial final_obs female employed hungup_any yesno
label values EMPLOYE consentement_capi yesno
label values A1 A1
label values A4 A4
label values C1 C1
label values C3 C3
label values C4 C4
label values C5 C5
label values C8 C8
label values D3 D3
label values E3A E3A
label values F1 F1
label values F2 P4 sector
label values F3 F3
label values F4 F4
label values G1 G1
label values H1 H1
label values H2 H2
label values H3A H5A bracket
label values I2 I2
label values I3 J1 search
label values J0A J0A
label values J2 J2
label values K3 K4 funding
label values K7 K7
label values L3 M2 freq3
label values L5 L5
label values L7 L7
label values L9 L9
label values M1 M4 selfsuff
label values M3 agree5
label values N2 N2
label values N3 N3
label values N4 N4
label values N5 N5
label values N6 N6
label values N7 N7
label values O5 O5
label values O6 O6
label values P2 P2
label values Q3 Q3
label values Q4 Q5 Q6 Q7 Q8 Q9 Q10 trust
label values R7 R7
label values C2A_Jour C2A_Mois C2A_annee C2 D1 D2 D4 G3 G4A G6 H3 H4A H5  ///
    L4 L8 O2 P3 Q2 wage_month inc_month missing_only
capture confirm numeric variable interview__status
if !_rc label values interview__status status

*------------------------------------------------------------------------------*
* 8. Variable labels (English)
*------------------------------------------------------------------------------*
* Identification / cover
label var interview__key     "Survey Solutions interview key"
label var interview__id      "Survey Solutions interview ID"
label var cover_id           "Respondent ID (preloaded)"
label var cover_district     "District (preloaded)"
label var cover_region       "Region (preloaded)"
label var localite           "Locality (preloaded)"
label var sssys_irnd         "Survey Solutions random number"
label var has__errors        "Number of validation errors"
label var interview__status  "Survey Solutions interview status"
label var assignment__id     "Survey Solutions assignment ID"

* Section timestamps
foreach s in A B C D E F G H I J K L M N O P Q R {
    label var HH`s'_debut "Section `s' start time"
    label var HH`s'_fin   "Section `s' end time"
}

* A. Call attempts
label var A1       "A1. Call attempt number"
label var A2_date  "A2. Date of the call"
label var A2_heure "A2. Time of the call"
label var A4       "A4. Result of this call attempt"
label var A4X      "A4X. Other call result (specify)"
label var A5_date  "A5. Planned date of next attempt"
label var A5_heure "A5. Planned time of next attempt"

* B. Identification and consent
label var B1  "B1. District"
label var B2  "B2. Region"
label var B3  "B3. Department"
label var B4  "B4. Sub-prefecture / commune"
label var B5  "B5. Area of residence"
label var B6  "B6. Speaking to the intended respondent"
label var B6A "B6A. Can speak to the respondent"
label var B6B "B6B. Agrees to a callback (respondent unavailable)"
label var B6C "B6C. Preferred callback date (respondent unavailable)"
label var B6D "B6D. Agrees to answer the questions (consent)"
label var B6E "B6E. Agrees to a callback (not a good time)"
label var B6F "B6F. Preferred callback date (not a good time)"
label var consentement_capi "Consent as computed in CAPI (unreliable, see consent)"
label var B7  "B7. Respondent full name"

* C. Socio-demographic profile
label var C1        "C1. Sex"
label var C2A_Jour  "C2A. Day of birth"
label var C2A_Mois  "C2A. Month of birth"
label var C2A_annee "C2A. Year of birth"
label var C2        "C2. Age in completed years (if birth year unknown)"
label var C3        "C3. Current marital status"
label var C4        "C4. Highest level of education attended"
label var C5        "C5. Highest diploma obtained"
label var C6        "C6. Can read and write in at least one language"
label var C7        "C7. Moved to another locality/country for 6+ months, last 12 months"
label var C8        "C8. Previous place of residence"
label var C9        "C9. Previous locality"
label var C10       "C10. Locality moved to"
label var C11       "C11. Country of origin"
label var age_capi  "Age as computed in CAPI (C2, else 2026 - birth year)"

* D. Household
label var D1  "D1. Number of usual household members"
label var D2  "D2. Number of household members under 15"
label var D3  "D3. Relationship to household head"
label var D3X "D3X. Other relationship to head (specify)"
label var D4  "D4. Number of household members in paid work"
label var D5__1  "D5. Household owns: television"
label var D5__2  "D5. Household owns: mobile phone"
label var D5__3  "D5. Household owns: motorcycle"
label var D5__4  "D5. Household owns: bicycle"
label var D5__5  "D5. Household owns: wheelbarrow"
label var D5__6  "D5. Household owns: refrigerator/freezer"
label var D5__7  "D5. Household owns: computer"
label var D5__8  "D5. Household owns: radio"
label var D5__9  "D5. Household owns: fan"
label var D5__10 "D5. Household owns: TNT decoder / Canal+ / Easy TV / StarTimes"
label var D5__11 "D5. Household owns: none of these"

* E. Activity status
label var E1      "E1. Worked for pay/profit last 7 days (even 1 hour)"
label var E2      "E2. Temporarily absent from a job/business"
label var E3      "E3. Helped unpaid in a family business/farm"
label var E3A     "E3A. Output of family activity mainly for sale or own use"
label var EMPLOYE "Employed (CAPI computed)"

* F. Main job
label var F1  "F1. Status in main job"
label var F1X "F1X. Other job status (specify)"
label var F2  "F2. Sector of main activity"
label var F2X "F2X. Other sector (specify)"
label var F3  "F3. Main place of work"
label var F3X "F3X. Other place of work (specify)"
label var F4  "F4. Time in this job without interruption"

* G. Job quality
label var G1  "G1. Has an employment contract"
label var G2  "G2. Has job-related social protection"
label var G3  "G3. Hours worked in main job last week"
label var G4  "G4. Last week's hours are the usual hours"
label var G4A "G4A. Hours usually worked (if G4 = No)"
label var G5  "G5. Has another job/income-generating activity"
label var G5A "G5A. Other activity (description)"
label var G6  "G6. Hours per week in other job(s)"
label var G7  "G7. Wanted to work more paid hours, last 30 days"
label var G7A "G7A. Available to work more hours in next 2 weeks"

* H. Income and business characteristics
label var H1  "H1. Form of payment (employees)"
label var H2  "H2. Pay frequency (employees)"
label var H3  "H3. Net pay for last pay period, FCFA"
label var H3A "H3A. Net pay bracket, FCFA (if H3 refused)"
label var H4  "H4. Activity/business is registered"
label var H4A "H4A. Number of paid workers (excluding self)"
label var H5  "H5. Earnings from the activity last month, FCFA"
label var H5A "H5A. Earnings bracket, FCFA (if H5 refused)"

* I. Search for another job (employed)
label var I1  "I1. Actively looking for another job"
label var I2  "I2. Main reason for looking"
label var I2X "I2X. Other reason (specify)"
label var I3  "I3. Main search method"
label var I3X "I3X. Other search method (specify)"

* J. Unemployment (not employed)
label var J0  "J0. Looked for work/tried to start a business, last 30 days"
label var J0A "J0A. Main reason for not looking"
label var J0B "J0B. Would like to work now"
label var J1  "J1. Main search method, last 30 days"
label var J1X "J1X. Other search method (specify)"
label var J2  "J2. Time without work and searching"
label var J3  "J3. Could start work within two weeks"

* K. Entrepreneurship
label var K1  "K1. Has ever started own income-generating activity"
label var K2  "K2. Would like to start own activity in the future"
label var K3  "K3. Main source of start-up funds"
label var K3X "K3X. Other source (specify)"
label var K4  "K4. Expected main source of funding"
label var K4X "K4X. Other source (specify)"
label var K5__1 "K5. Not yet started: lack of capital/financing"
label var K5__2 "K5. Not yet started: lack of entrepreneurial skills"
label var K5__3 "K5. Not yet started: lack of equipment/premises"
label var K5__4 "K5. Not yet started: lack of market/demand"
label var K5__5 "K5. Not yet started: family obligations"
label var K5__6 "K5. Not yet started: fear of failure/risk"
label var K5__7 "K5. Not yet started: administrative/legal obstacles"
label var K5__9 "K5. Not yet started: other"
label var K5X "K5X. Other reason (specify)"
label var K6  "K6. Ever received entrepreneurship/business training"
label var K7  "K7. Self-assessed business management skills"

* L. Financial inclusion and savings
label var L1  "L1. Has a mobile money account"
label var L2  "L2. Has a bank account"
label var L3  "L3. Saves money regularly"
label var L4  "L4. Amount saved in a typical month, FCFA"
label var L5  "L5. Main place savings are kept"
label var L5X "L5X. Other place (specify)"
label var L6  "L6. Currently has a loan or debt"
label var L7  "L7. Main source of loan/debt"
label var L7X "L7X. Other source (specify)"
label var L8  "L8. Total amount currently owed, FCFA"
label var L9  "L9. Main purpose of the loan"

* M. Change in status
label var M1 "M1. How you currently make a living"
label var M2 "M2. Received external help for basic needs, last 12 months"
label var M3 "M3. Seen by family/community as contributing economically"
label var M4 "M4. In 12 months, expect to be self-supporting or dependent"

* N. Resilience to shocks
label var N1  "N1. Household affected by a shock, last 12 months"
label var N2  "N2. Most important shock"
label var N2X "N2X. Other shock (specify)"
label var N3  "N3. Impact of the shock on income/activity"
label var N4  "N4. Main way of coping with the shock"
label var N4X "N4X. Other coping strategy (specify)"
label var N5  "N5. Time to recover pre-shock income/activity"
label var N6  "N6. Could raise 250,000 FCFA for an emergency"
label var N7  "N7. Ability to withstand an economic shock (1-10)"

* O. Work experience and training
label var O1 "O1. Ever had another job/income-generating activity"
label var O2 "O2. Age at first job/activity"
label var O3 "O3. Vocational/technical training or apprenticeship, last 12 months"
label var O4 "O4. Trade or field of training"
label var O5 "O5. Main training provider"
label var O6 "O6. Usefulness of the training"

* P. Career aspirations
label var P1  "P1. Desired occupation in the next five years"
label var P2  "P2. Wants wage job, new business or grow existing activity"
label var P3  "P3. Desired monthly income, FCFA"
label var P4  "P4. Sector of greatest interest"
label var P4X "P4X. Other sector (specify)"
label var P5  "P5. Main obstacle to career goals"

* Q. Social cohesion and engagement
label var Q1  "Q1. Member of an association/group/community organisation"
label var Q2  "Q2. Times volunteered for community activity, last 30 days"
label var Q3  "Q3. Generalised trust"
label var Q4  "Q4. Trust in youth of different origin"
label var Q5  "Q5. Trust in the police"
label var Q6  "Q6. Trust in the armed forces"
label var Q7  "Q7. Trust in traditional leaders"
label var Q8  "Q8. Trust in religious leaders"
label var Q9  "Q9. Trust in local government authorities"
label var Q10 "Q10. Trust in central government authorities"

* R. Panel follow-up
label var R1  "R1. Current number is best contact"
label var R1A "R1A. Best number for future contact"
label var R2  "R2. Has an alternative number"
label var R2A "R2A. Alternative number"
label var R3  "R3. Uses WhatsApp"
label var R3A "R3A. WhatsApp on the same number"
label var R3B "R3B. WhatsApp number"
label var R4  "R4. Contact of a person who can reach the respondent"
label var R5  "R5. Agrees to be recontacted for future waves"
label var R6  "R6. Number for Wave mobile payment"
label var R7  "R7. Owner of the Wave number"

label var hungup_any "Respondent hung up during interview (-95 recorded)"

*------------------------------------------------------------------------------*
* 9. Data quality flags (nothing is changed or dropped)
*------------------------------------------------------------------------------*
gen byte flag_consent_capi = (consentement_capi == 1 & consent != 1)
label var flag_consent_capi "CAPI consent = 1 but B6D consent not given"

gen byte flag_employe = (employed != EMPLOYE) if !missing(employed, EMPLOYE)
label var flag_employe "Recomputed employment differs from CAPI EMPLOYE"

* Age bounds: edit to match program eligibility
gen byte flag_age = (age_clean < 15 | age_clean > 40) if !missing(age_clean)
label var flag_age "Age outside 15-40"

gen byte flag_age_c2 = (abs(age_clean - C2) > 1) if !missing(age_clean, C2) ///
    & !missing(C2A_annee)
label var flag_age_c2 "Reported age C2 inconsistent with birth year"

gen byte flag_hh_u15 = (D2 >= D1) if !missing(D1, D2)
label var flag_hh_u15 "Members under 15 (D2) not below household size (D1)"

gen byte flag_hh_workers = (D4 > D1) if !missing(D1, D4)
label var flag_hh_workers "Earning members (D4) exceed household size (D1)"

gen byte flag_hh_size = (D1 < 1 | D1 > 50) if !missing(D1)
label var flag_hh_size "Household size (D1) below 1 or above 50"

gen byte flag_assets_none = (D5__11 == 1 & n_assets > 0) if !missing(D5__11)
label var flag_assets_none "'None of the above' selected together with an asset"

gen byte flag_hours = (hours_total > 112) if !missing(hours_total)
label var flag_hours "More than 112 hours worked last week (16h/day)"

gen byte flag_income = (H3 < 0 | H5 <= 0 | L4 < 0 | L8 < 0 | P3 < 0) ///
    if !missing(H3) | !missing(H5) | !missing(L4) | !missing(L8) | !missing(P3)
label var flag_income "Negative income/savings/debt, or H5 not positive"

gen byte flag_first_job = (O2 < 5 | O2 > age_clean) if !missing(O2, age_clean)
label var flag_first_job "Age at first job (O2) below 5 or above current age"

gen byte flag_skip_emp = (employed == 1 & !missing(J0)) | ///
                         (employed == 0 & !missing(F1)) if !missing(employed)
label var flag_skip_emp "Employed answered section J, or non-employed answered F"

decode B1, gen(_dist)
decode B2, gen(_reg)
gen byte flag_geo = 0 if !missing(B1)
replace flag_geo = 1 if !missing(B1) & cover_district != "" & ///
    upper(cover_district) != upper(strtrim(_dist))
replace flag_geo = 1 if !missing(B2) & cover_region != "" & ///
    upper(cover_region) != upper(strtrim(_reg))
drop _dist _reg
label var flag_geo "Reported district/region differs from cover"

* Official codes are nested: region 11103 is in district 111,
* department 11103029 is in region 11103, sub-prefecture 1110302906 is in
* department 11103029
gen byte flag_geo_nest = (floor(B2 / 100) != B1) if !missing(B1, B2)
replace flag_geo_nest = 1 if floor(B3 / 1000) != B2 & !missing(B2, B3)
replace flag_geo_nest = 1 if floor(B4 / 100) != B3 & !missing(B3, B4)
label var flag_geo_nest "Region/department/sub-prefecture not nested in the level above"

gen byte flag_short = (dur_total < 10) if complete == 1 & !missing(dur_total)
label var flag_short "Completed interview shorter than 10 minutes"

gen byte flag_errors = (has__errors > 0) if !missing(has__errors)
label var flag_errors "Survey Solutions reports validation errors"

* Summary: sum = number of rows flagged, mean = share of rows checked
tabstat flag_*, statistics(sum mean n) columns(statistics) varwidth(20)

*------------------------------------------------------------------------------*
* 10. Save: PII file and analysis file
*------------------------------------------------------------------------------*
order flag_*, last
sort cover_id interview__key
compress
label data "COSO baseline phone survey - clean, all call attempts"

* Personal identifiers go to a separate restricted file
preserve
    keep interview__key cover_id B7 R1A R2A R3B R4 R6
    label data "COSO baseline - personal identifiers (restricted)"
    save "$cleandata/coso_baseline_pii.dta", replace
restore

drop B7 R1A R2A R3B R4 R6
save "$cleandata/coso_baseline_clean.dta", replace

di as txt _n "Rows: " as res _N
tab complete final_obs, missing

log close
