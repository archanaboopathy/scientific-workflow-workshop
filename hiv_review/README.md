# HIV-1 Phase 3 trial review: how `HIV_Phase3_Trial_Review.xlsx` was built

This folder contains the inputs, scripts and extracted results behind
`HIV_Phase3_Trial_Review.xlsx` in the repository root. The workbook was compiled with AI
assistance (GitHub Copilot agents). **It is not scientific validation.** A qualified reviewer
must check every value against its primary source before use.

## 1. Request

Identify Phase 3 trials for the HIV drugs in column C of the NIH antiretroviral table. Build a
table with the same columns plus the treatment-emergent resistance % for the treatment drug.
The focus is Week 48 efficacy by FDA Snapshot, treatment-emergent resistance, publication year
and reference manuscript. No extrapolation.

## 2. Input

- `NIH_ARV_Phase3_Master_Week48_Scatter_X_Offset.xlsx` (repo root), **first tab only**
  ("Phase 3 Master"): 63 rows, 41 products in column C. The other tabs were not used.
- `read_master.py` turns that tab into `master_rows.json` (original rows with row numbers) and
  `products.json` (the 41 products with their PubMed search terms).

## 3. Scope decisions (confirmed by the user before research started)

| Topic | Decision |
|---|---|
| Products | All 41 in column C. Empty cells get "Not applicable" or "Not reported". |
| Trials | Every Phase 3/3b adult HIV-1 treatment trial, registrational and post-approval. Pediatric-only, PrEP, HBV, PK/bioequivalence and Phase 2 trials are excluded. |
| Inclusion | The product must be the investigational intervention in at least one arm. A drug that is only background therapy in every arm does not count. |
| Population | New column: Treatment-naive / Switch / Treatment-experienced. |
| Week 48 | FDA Snapshot HIV-1 RNA <50 copies/mL. If unavailable, the endpoint the trial used, labelled with its definition. |
| Resistance | % emergent resistance to the product using two denominators (randomized/treated; resistance-tested) plus raw counts. A % is calculated only when the source states the counts. |
| Sources | Peer-reviewed paper > FDA label/review > CROI/IAS abstract > ClinicalTrials.gov posted results. Each value names its source. |
| Existing rows | Every value re-checked against primary sources; every change logged. |
| BIC, ISL, TAF | No standalone HIV product, so these rows keep their placeholder. Their trials appear under the coformulated product rows. |
| DOR/ISL 017–020 | These used islatravir 0.75 mg, which is not the approved dose (approved Idvynso is DOR 100 mg/ISL 0.25 mg). They are included and labelled "[unapproved ISL 0.75 mg dose]". |

`PROTOCOL.md` holds the full extraction rules every agent followed.

## 4. Sources and tools

Only public sources were used, through `lit.py` (Python standard library only):

| Command | Source |
|---|---|
| `abstract`, `search` | PubMed E-utilities |
| `fulltext` | Europe PMC open-access full text |
| `ctgov`, `ctresults`, `ctsearch` | ClinicalTrials.gov API v2 (registry records and posted results) |
| `label` | DailyMed current US labels (sec. 14 Clinical Studies, 12.4 Microbiology, "Initial U.S. Approval") |

Some agents also read FDA documents and CROI/IAS abstract pages directly. HTTP responses are
cached in `cache/`, which is git-ignored. `openpyxl` was used only to write the Excel file, in
a local virtual environment approved by the user. No other packages were installed.

## 5. Process

1. **Candidate search (session 2).** PubMed was searched for each product, restricted to HIV
   and Phase III publication type or "phase 3/III" in title/abstract. Pediatric, PrEP and
   HBV titles were excluded. This returned 836 unique records. Removing reviews, news,
   comments and Phase II-only records left 627. 480 of those looked like primary reports and
   147 like secondary analyses.
2. **Grouping.** The 41 products were split into 7 groups (`groups.json`). Each group was given
   to a separate research agent.
3. **Per-group extraction.** For each product, each agent did the following, following
   `PROTOCOL.md`:
   - screened the candidates, and searched ClinicalTrials.gov (`ctsearch`) and PubMed again
     (trial acronyms, "phase 3b", older trials without phase tags);
   - applied the inclusion rule;
   - extracted values from the primary publication, FDA label and ClinicalTrials.gov;
   - re-checked every existing master row;
   - logged every changed value with its source;
   - recorded excluded trials with a reason;
   - flagged anything a human must check.

   Results are one JSON file per group.
4. **Two research runs.** The work was done twice because agent sessions were interrupted:
   - **Base run:** `out/g1_DTG.json` … `out/g7_ENTRY.json`, covering all 41 products and used
     as the base.
   - **Session 2:** `session2/g2.json` (BIC/LEN, LEN, FTR, IBA), `session2/g5.json` (DOR/ISL,
     DOR, EFV products) and `session2/g6.json` (DRV/c/TAF/FTC, DRV, ATV, RTV, ETR, MVC). The
     other session-2 groups were interrupted and are not used.
5. **Merge (`merge_supplement.py` → `out/g8_supplement.json`).** At the user's direction, the
   base run is the base. Trials were matched on product plus NCT ID, trial name/acronym or
   PMID. The supplement adds:
   - **11 trials** that session 2 found and the base run neither included nor excluded:
     - DOR: Opti-DOR, DORA
     - EFV: Study 020, ACTG 388
     - EFV/TDF/FTC: SWEET, ANRS 129 BKVIR
     - ATV: SWAN, Honda 2011, OREY
     - ETR: Etraswitch, Ruane 2010 pilot
   - **DOR/ISL 017, 018, 019, 020**, which the base run excluded. They are included per the
     user decision and removed from the excluded list.

   For the other **9 disagreements** where the base run excluded a trial session 2 included,
   the base decision is kept. All 13 disagreements, including the 4 overrides, are on the
   **Reconciliation** sheet. Every added row says where it came from in its Notes and Flags.
6. **Build (`build_final.py`).** The script reads `out/g*.json` and writes the workbook:
   - Columns A–R keep the original layout. The new columns have blue headers.
   - Resistance % = n emergent / N, calculated only from integers stated in the source.
   - Original rows not matched to an in-scope trial are listed on "Sources & Changes".
7. **Checks run:**
   - All JSON files parse.
   - All 169 rows of the base-run workbook appear unchanged in the merged workbook.
   - `python scripts/validate.py` passes (3 tests). These tests cover the repository's
     workshop code only, not this workbook.

## 6. Output: `HIV_Phase3_Trial_Review.xlsx`

| Sheet | Contents |
|---|---|
| Read Me | Method summary and limitations |
| Phase 3 Master | 184 rows (one per product × trial): 126 new trial rows, 53 original rows with corrected values, 5 placeholder rows. Columns A–R as original, plus 12 new columns (population, two resistance %, raw counts, resistance timepoint, end-of-trial resistance, efficacy/resistance sources, PMID, design/arms, row status, flags). |
| Sources & Changes | Each value changed in an original row (previous → new, reason, source), dropped original rows, and per-row sources |
| Excluded Trials | 203 screened trials and the reason each was excluded |
| Product Notes | Per-product notes: approval-year source, coverage and placeholder confirmations |
| Reconciliation | 13 trials where the two research runs disagreed, and which decision was kept |

Five original rows were not carried forward:
- Rows 28–30 (ABC/3TC, TAF/FTC, TDF/FTC placeholders) were replaced by the in-scope trials
  found for those products.
- Rows 50–51 (DRV POWER-1/POWER-2) are Phase 2b trials and are now on Excluded Trials.

## 7. Reproduce

```bash
cd hiv_review
python3 read_master.py          # stdlib only; regenerates master_rows.json / products.json
python3 merge_supplement.py     # stdlib only; regenerates out/g8_supplement.json
python3 build_final.py          # requires openpyxl; writes ../HIV_Phase3_Trial_Review.xlsx
```

The research step (agents using `lit.py`) cannot be replayed automatically. Its results are the
committed `out/*.json` and `session2/*.json` files.

## 8. Limitations and what still needs human review

- **Full text access:** many papers are paywalled. Values often come from abstracts,
  open-access full text, FDA labels or ClinicalTrials.gov results instead of the full paper.
  The source is named in each row.
- **Pooled data:** some trials report only pooled efficacy or resistance (for example
  DUET-1/2, MOTIVATE, and EVG/c/TAF/FTC Studies 104/111). These are marked as pooled.
- **Conflicting sources:** some sources disagree. For example, Idvynso label vs publications
  on 051/052 resistance, the ARTISTRY-1 paper vs label, and CAPELLA Week 156 counts. These are
  flagged, not resolved.
- **Not FDA Snapshot:** older trials (pre-2008) and several investigator-led trials have no
  Snapshot result. The endpoint they used is shown and labelled.
- **Inclusion calls to confirm:** these are listed in the Flags column and on the
  Reconciliation sheet. Examples:
  - trial phase not stated;
  - small pilot trials;
  - expanded access (TMB-311);
  - HIV/HBV coinfection (ALLIANCE);
  - lenacapavir counted via coformulation trials;
  - DOR/ISL 019, which enrolled adolescents ≥12 years.
- **Agent extraction:** values were extracted by AI agents and spot-checked only through the
  automated checks above. No human has verified them against the sources.
