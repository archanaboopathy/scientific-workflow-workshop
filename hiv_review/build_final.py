import json, glob, datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

import os
W = os.path.dirname(os.path.abspath(__file__)) + '/'
OUT = os.path.join(W, '..', 'HIV_Phase3_Trial_Review.xlsx')
master = json.load(open(W + 'master_rows.json'))
order = [p for p, _ in json.load(open(W + 'products.json'))]
HDR = list(k for k in master[0] if k != '_row')
prod_info = {}
for m in master:
    prod_info.setdefault(m['NIH Abbreviation/Components'], m)

recs, excl, pnotes = [], [], {}
for f in sorted(glob.glob(W + 'out/g*.json')):
    d = json.load(open(f))
    recs += d.get('records', []); excl += d.get('excluded', [])
    pnotes.update(d.get('product_notes', {}))

# User decision: component rows without a standalone HIV product keep their placeholder.
PLACEHOLDER_ONLY = {'BIC', 'ISL', 'TAF'}
for r in [r for r in recs if r.get('product') in PLACEHOLDER_ONLY]:
    excl.append({'product': r['product'], 'trial': r.get('trial_name'), 'nct_id': r.get('nct_id'), 'pmid': r.get('pmid'),
                 'reason': 'Covered under coformulated product row(s); component placeholder kept per user decision'})
recs = [r for r in recs if r.get('product') not in PLACEHOLDER_ONLY]

def s(v):
    if v is None: return ''
    if isinstance(v, list): return '; '.join(map(str, v))
    return str(v)

def pct(n, N):
    if isinstance(n, int) and isinstance(N, int) and N > 0:
        return f"{100 * n / N:.1f}% ({n}/{N})"
    return 'Not reported' if n is None or N is None else 'Not calculable'

NEW = ['Treatment Population', 'TE Resistance to Product - % of Randomized/Treated (n/N)',
       'TE Resistance to Product - % of Resistance-Tested (n/N)', 'TE Resistance - Raw Counts',
       'Resistance Timepoint', 'End-of-Trial Resistance (later timepoint)', 'Efficacy Source',
       'Resistance Source', 'PMID', 'Study Design / Arms', 'Row Status', 'Flags for Human Review']
rows = []
matched_existing = set()
for r in recs:
    p = r.get('product'); base = prod_info.get(p, {})
    ex = r.get('existing_row')
    if ex: matched_existing.add(ex)
    snap = s(r.get('week48_snapshot'))
    alt = s(r.get('week48_alt_endpoint'))
    if alt.strip().lower().rstrip('.') in ('', 'not applicable', 'n/a', 'na', 'none', 'not reported'): alt = ''
    L = snap if not alt else (f"{snap} | Endpoint used: {alt}" if snap else f"Endpoint used: {alt}")
    nct = s(r.get('nct_id'))
    n, Nr, Nt = r.get('res_new_drug_n'), r.get('res_N_randomized'), r.get('res_N_tested')
    raw = ', '.join(f"{k}={v}" for k, v in [('n emergent', n), ('N randomized/treated', Nr), ('N tested', Nt)] if v is not None) or 'Not reported'
    status = ('New row' if not ex else ('Existing row - values changed' if r.get('changes') else 'Existing row - verified, no change'))
    eff_src = '; '.join(x for x in [s(r.get('week48_snapshot_source')), s(r.get('week48_alt_source'))] if x)
    vals = {
        'Source Table': base.get('Source Table'), 'Drug Class': base.get('Drug Class'),
        'NIH Abbreviation/Components': p, 'Generic Name': base.get('Generic Name'),
        'Drug/Brand Name': base.get('Drug/Brand Name'), 'Recommended/Common Status': base.get('Recommended/Common Status'),
        'Registrational Phase 3 Trial Name': s(r.get('trial_name')), 'ClinicalTrials.gov ID': nct or 'Not reported',
        'Peer-Reviewed Phase 3 Publication': s(r.get('publication_citation')), 'Peer-Reviewed Publication Year': r.get('publication_year'),
        'Original FDA HIV Approval Year': r.get('approval_year_checked') or base.get('Original FDA HIV Approval Year'),
        'Week 48 FDA Snapshot: HIV-1 RNA <50 copies/mL': L,
        'Treatment-Emergent Resistance - Any Regimen Component': s(r.get('te_res_any_component')),
        'Treatment-Emergent Resistance - Newly Approved Drug/Product': s(r.get('te_res_new_drug_text')),
        'Trial/Publication Notes': s(r.get('notes')), 'NIH Source URL': base.get('NIH Source URL'),
        'ClinicalTrials.gov URL': f"https://clinicaltrials.gov/study/{nct}" if nct.startswith('NCT') else 'Not applicable',
        'Regulatory/Publication Source Note': f"Efficacy: {eff_src or 'n/a'} | Resistance: {s(r.get('resistance_source')) or 'n/a'} | Approval year: {s(r.get('approval_year_source')) or 'n/a'}",
        NEW[0]: s(r.get('population')), NEW[1]: pct(n, Nr), NEW[2]: pct(n, Nt), NEW[3]: raw,
        NEW[4]: s(r.get('resistance_timepoint')), NEW[5]: s(r.get('res_final_text')) or 'Not reported',
        NEW[6]: eff_src, NEW[7]: s(r.get('resistance_source')), NEW[8]: s(r.get('pmid')),
        NEW[9]: ' | '.join(x for x in [s(r.get('design')), 'Exp: ' + s(r.get('experimental_arm')), 'Comp: ' + s(r.get('comparator_arm'))] if x.strip(': ')),
        NEW[10]: status, NEW[11]: s(r.get('flags')),
    }
    rows.append((order.index(p) if p in order else 99, r.get('publication_year') or 9999, vals, r))

# Existing rows not represented by any record (placeholders or excluded)
dropped = []
for m in master:
    if m['_row'] in matched_existing: continue
    p = m['NIH Abbreviation/Components']
    has_new = any(r.get('product') == p for r in recs)
    trial = s(m['Registrational Phase 3 Trial Name'])
    is_placeholder = trial.startswith('No ') or 'Historical' in trial
    if is_placeholder and not has_new:
        vals = {k: m[k] for k in HDR}
        vals.update({NEW[0]: 'Not applicable', NEW[1]: 'Not applicable', NEW[2]: 'Not applicable', NEW[3]: 'Not applicable',
                     NEW[4]: 'Not applicable', NEW[5]: 'Not applicable', NEW[6]: '', NEW[7]: '', NEW[8]: '',
                     NEW[9]: '', NEW[10]: 'Existing placeholder - no in-scope Phase 3 trial confirmed',
                     NEW[11]: s(pnotes.get(p, 'No product note returned; confirm placeholder'))})
        rows.append((order.index(p), 0, vals, {}))
    else:
        dropped.append(m)

rows.sort(key=lambda x: (x[0], x[1], s(x[2]['Registrational Phase 3 Trial Name'])))

wb = Workbook(); ws = wb.active; ws.title = 'Phase 3 Master'
cols = HDR + NEW
ws.append(cols)
bold = Font(bold=True); blue = PatternFill('solid', fgColor='DDEBF7')
orange = PatternFill('solid', fgColor='FCE4D6'); green = PatternFill('solid', fgColor='E2EFDA')
for i, c in enumerate(ws[1], 1):
    c.font = bold; c.alignment = Alignment(wrap_text=True, vertical='top')
    if i > len(HDR): c.fill = blue
    ws.column_dimensions[get_column_letter(i)].width = 14 if i <= 6 else 32
for _, _, vals, r in rows:
    ws.append([vals.get(k) for k in cols])
    rr = ws.max_row
    for c in ws[rr]: c.alignment = Alignment(wrap_text=True, vertical='top')
    st = vals[NEW[10]]
    ws.cell(rr, cols.index(NEW[10]) + 1).fill = green if st == 'New row' else (orange if 'changed' in st else PatternFill())
    if vals[NEW[11]]: ws.cell(rr, cols.index(NEW[11]) + 1).fill = orange
ws.freeze_panes = 'H2'; ws.auto_filter.ref = ws.dimensions

sc = wb.create_sheet('Sources & Changes')
sc.append(['Product', 'Trial', 'Master row (original)', 'Column', 'Field', 'Previous value', 'New value', 'Reason', 'Source'])
for r in recs:
    for ch in r.get('changes', []) or []:
        sc.append([r.get('product'), r.get('trial_name'), r.get('existing_row'), ch.get('column'), ch.get('field'),
                   s(ch.get('previous')), s(ch.get('new')), s(ch.get('reason')), s(ch.get('source'))])
for m in dropped:
    sc.append([m['NIH Abbreviation/Components'], m['Registrational Phase 3 Trial Name'], m['_row'], 'ALL', 'Row',
               'Row present in original', 'Row not carried forward', 'Not matched to an in-scope trial record (see Excluded Trials sheet)', ''])
sc.append([]); sc.append(['PER-ROW SOURCES']); sc.append(['Product', 'Trial', 'PMID', 'Primary publication', 'Efficacy source', 'Resistance source', 'Approval year source'])
for _, _, vals, r in rows:
    if r: sc.append([r.get('product'), r.get('trial_name'), s(r.get('pmid')), s(r.get('publication_citation')),
                     vals[NEW[6]], vals[NEW[7]], s(r.get('approval_year_source'))])
for c in sc[1]: c.font = bold
for col in 'ABCDEFGHI': sc.column_dimensions[col].width = 30

ex = wb.create_sheet('Excluded Trials')
ex.append(['Product', 'Trial', 'NCT ID', 'PMID', 'Reason excluded'])
for e in excl: ex.append([s(e.get('product')), s(e.get('trial')), s(e.get('nct_id')), s(e.get('pmid')), s(e.get('reason'))])
for c in ex[1]: c.font = bold
for col in 'ABCDE': ex.column_dimensions[col].width = 35

pn = wb.create_sheet('Product Notes')
pn.append(['Product', 'Notes (approval year source, coverage, placeholder confirmation)'])
for p in order: pn.append([p, s(pnotes.get(p, ''))])
for c in pn[1]: c.font = bold
pn.column_dimensions['A'].width = 16; pn.column_dimensions['B'].width = 140

rm = wb.create_sheet('Read Me', 0)
for l in [
    'HIV-1 Phase 3 Trial Review - generated ' + datetime.date.today().isoformat(),
    'Input: first tab ("Phase 3 Master") of NIH_ARV_Phase3_Master_Week48_Scatter_X_Offset.xlsx; products = column C (41 products).',
    'Scope: Phase 3/3b adult HIV-1 treatment trials (registrational and post-approval) where the product is the investigational intervention in at least one arm. Pediatric, PrEP, HBV, PK/bioequivalence and Phase 2 trials excluded.',
    'Row grain: one row per product x trial. Columns A-R keep the original layout; blue columns are new.',
    'Week 48 column: FDA Snapshot HIV-1 RNA <50 copies/mL where reported; otherwise the endpoint the trial used is shown with its definition ("Endpoint used: ...").',
    'Resistance %: calculated only from counts stated in the source: n (participants in the product arm with emergent resistance to the product) / N randomized-treated and / N with resistance testing. "Not reported" = counts not stated.',
    'No extrapolation: no complements, cross-timepoint conversions or derived values other than the two resistance percentages above.',
    'Source priority: peer-reviewed paper > FDA label/review > CROI/IAS abstract > ClinicalTrials.gov posted results. Each row names its sources.',
    'Existing rows were re-checked against primary sources; every change is listed on "Sources & Changes". Row Status column: New row / Existing verified / Existing changed / Placeholder.',
    'Limitations: many full texts are paywalled; values were taken from abstracts, open-access full text, FDA labels and ClinicalTrials.gov results where full text was unavailable. Orange "Flags" cells need human review.',
    'This compilation was produced with AI assistance and is NOT scientific validation. A qualified reviewer must verify values against the primary sources before any use.',
]: rm.append([l])
rm['A1'].font = bold; rm.column_dimensions['A'].width = 160
wb.save(OUT)
print('records', len(recs), 'rows', len(rows), 'changes', sc.max_row, 'excluded', len(excl), 'dropped existing', len(dropped))
