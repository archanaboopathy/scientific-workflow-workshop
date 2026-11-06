# Extraction protocol: HIV-1 Phase 3 trial review

Work dir: `hiv_review/` (run commands from here). Inputs: `master_rows.json`, `products.json` (built by `read_master.py` from tab 1 of `NIH_ARV_Phase3_Master_Week48_Scatter_X_Offset.xlsx`), `groups.json` (7 product groups).
Helper: `python3 lit.py ...` (stdlib only; responses cached in `cache/`, not committed). Commands:
- `abstract PMID [PMID...]`, `search "pubmed query" [N]`, `fulltext PMID [regex]` (Europe PMC open access only),
- `ctgov NCTID`, `ctresults NCTID [regex]` (posted CT.gov outcome results), `ctsearch DRUGNAME` (Phase 3 HIV studies with that intervention),
- `label BRANDNAME [regex]` (current US label from DailyMed; sec 14 Clinical Studies, 12.4 Microbiology; Highlights "Initial U.S. Approval").
Every command is run with `timeout 300`. No package installs. Sources: PubMed, Europe PMC, ClinicalTrials.gov, DailyMed/FDA, CROI/IAS abstract pages.

## Scope decisions (user-approved; do not re-ask)
- All 41 products in column C. Empty cells get "Not applicable" / "Not reported".
- Every Phase 3/3b adult HIV-1 treatment trial, registrational and post-approval. Exclude pediatric, PrEP, HBV, PK/bioequivalence, Phase 2.
- New column: Treatment population (Treatment-naive / Switch / Treatment-experienced).
- Week 48 column: FDA Snapshot <50 c/mL; if unavailable, the endpoint used with its definition labelled.
- Resistance %: two denominators (randomized/treated; resistance-tested) plus raw counts; % only from stated counts.
- Source priority: peer-reviewed paper > FDA label/review > CROI/IAS abstract > CT.gov results; source named per value.
- Existing rows re-checked; every change logged. Output: Excel with original + new columns and a Sources & Changes sheet.
- Component rows without standalone HIV product (BIC, ISL, TAF) keep their placeholder.

## Inclusion rule (per product in column C)
1. Phase 3 or 3b interventional trial.
2. Adults with HIV-1, treatment indication. Trials in specific adult groups are included (note it).
3. Product is the investigational intervention in >=1 arm. Coformulations: that exact combination (FDC or same components co-administered as the studied regimen). Single agents: in the experimental arm and being evaluated (absent from comparator, or dose/formulation comparison); not background-only.
4. Registrational and post-approval trials in scope.
5. Has reported primary efficacy results; otherwise `excluded` with "no results reported".
One record per product x trial. Screen systematically: PubMed searches (acronyms, "phase 3b", pre-2005 trials) and `ctsearch` per drug.

## Field rules (NO extrapolation)
- Only explicitly stated values; never compute complements, convert time points or derive values. "Not reported" / "Not applicable".
- Every value names its source (e.g. "PMID 28867497 abstract", "CT.gov NCT02607930 posted results", "Biktarvy US label sec 14.2").
- `week48_snapshot`: per trial, e.g. "92.4% (290/314) B/F/TAF vs 93.0% (293/315) DTG/ABC/3TC" (n/N only if stated). If none: "Not reported (FDA Snapshot)" and the used endpoint in `week48_alt_endpoint`.
- Resistance in the product arm: `te_res_any_component` (per arm text), `te_res_new_drug_text` (to the product/any component of it), integers `res_new_drug_n`, `res_N_randomized`, `res_N_tested` only if stated. No percentages. `resistance_timepoint` = primary analysis; later analyses in `res_final_text`.
- `publication_citation` "First author et al. Title. Journal. Year;Vol:pages."; `pmid`; `publication_year`.
- `population`: "Treatment-naive" | "Switch" | "Treatment-experienced".
- Existing rows: re-check columns G-R; log each change in `changes`. Verify approval year via label "Initial U.S. Approval". Placeholder rows: confirm in `product_notes` if no in-scope trial.
- `flags`: items for human review.

## Output
`out/<group>.json`:
{"group": "", "records": [{"product": "", "existing_row": null, "trial_name": "", "nct_id": "", "population": "", "design": "", "experimental_arm": "", "comparator_arm": "", "publication_citation": "", "pmid": "", "publication_year": 0, "approval_year_checked": "", "approval_year_source": "", "week48_snapshot": "", "week48_snapshot_source": "", "week48_alt_endpoint": "", "week48_alt_source": "", "te_res_any_component": "", "te_res_new_drug_text": "", "res_new_drug_n": null, "res_N_randomized": null, "res_N_tested": null, "resistance_timepoint": "", "resistance_source": "", "res_final_text": "", "notes": "", "flags": [], "changes": [{"column": "", "field": "", "previous": "", "new": "", "reason": "", "source": ""}]}], "excluded": [{"product": "", "trial": "", "nct_id": "", "pmid": "", "reason": ""}], "product_notes": {}}
Validate with `python3 -c "import json;json.load(open('out/<group>.json'))"`.
