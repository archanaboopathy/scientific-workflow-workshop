"""Build out/g8_supplement.json from the session-2 results in session2/.

The hiv_review/out/g1-g7 results are the base. This adds only:
  1. Trials found in session 2 that the base run neither included nor excluded.
  2. DOR/ISL trials 017-020 (ISL 0.75 mg), which the base run excluded but the user chose to
     include, labelled as an unapproved dose.
Where the two runs disagree on any other trial, the base decision is kept and the disagreement is
listed in `reconciliation` for human review.
"""
import glob, json, os, re

W = os.path.dirname(os.path.abspath(__file__)) + '/'
PLACEHOLDER_ONLY = {'BIC', 'ISL', 'TAF'}
USER_INCLUDE = {('DOR/ISL', 'MK-8591A-017'), ('DOR/ISL', 'MK-8591A-018'),
                ('DOR/ISL', 'MK-8591A-019'), ('DOR/ISL', 'MK-8591A-020')}
UNAPPROVED_FLAG = ('Unapproved dose: ISL 0.75 mg (approved Idvynso is DOR 100 mg/ISL 0.25 mg). '
                   'Included per user decision (2026-11-05), labelled as unapproved dose.')


def norm(s):
    return re.sub(r'[^a-z0-9]', '', (s or '').lower())


def keys(r):
    name = r.get('trial_name') or r.get('trial') or ''
    ks = set(re.findall(r'NCT\d{8}', str(r.get('nct_id') or '')))
    if norm(name):
        ks.add(norm(name))
    ks |= {norm(t) for t in re.findall(r'\(([^)]+)\)', name)}
    pm = str(r.get('pmid') or '').strip()
    if pm.isdigit():
        ks.add('pmid' + pm)
    return ks


base_recs, base_excl = [], []
for f in sorted(glob.glob(W + 'out/g[1-7]_*.json')):
    d = json.load(open(f))
    base_recs += d['records']
    base_excl += d.get('excluded', [])

s2 = []
for f in sorted(glob.glob(W + 'session2/g*.json')):
    s2 += json.load(open(f))['records']

added, drop_excluded, recon = [], [], []
for r in s2:
    p, k = r['product'], keys(r)
    if p in PLACEHOLDER_ONLY:
        continue
    if any(a['product'] == p and keys(a) & k for a in base_recs):
        continue
    hits = [x for x in base_excl if x.get('product') == p and keys(x) & k]
    if hits and (p, r['trial_name']) not in USER_INCLUDE:
        recon.append({'product': p, 'trial': r['trial_name'], 'nct_id': r.get('nct_id'),
                      'base_decision': 'Excluded: ' + str(hits[0].get('reason')),
                      'session2_decision': 'Included', 'kept': 'Base decision (excluded)',
                      'session2_flags': '; '.join(r.get('flags') or [])})
        continue
    rec = dict(r)
    rec['existing_row'] = None
    rec['changes'] = []
    origin = 'Added from session-2 results (not covered by base hiv_review run).'
    if hits:
        drop_excluded += [{'product': p, 'trial': h.get('trial'), 'nct_id': h.get('nct_id')} for h in hits]
        rec['trial_name'] = r['trial_name'] + ' [unapproved ISL 0.75 mg dose]'
        rec['flags'] = [UNAPPROVED_FLAG] + [f for f in (r.get('flags') or []) if 'approved Idvynso dose' not in f]
        origin = 'Added from session-2 results; base run excluded it (unapproved dose), overridden per user decision.'
        recon.append({'product': p, 'trial': r['trial_name'], 'nct_id': r.get('nct_id'),
                      'base_decision': 'Excluded: ' + str(hits[0].get('reason')),
                      'session2_decision': 'Included', 'kept': 'Included per user decision (unapproved dose label)',
                      'session2_flags': ''})
    else:
        rec['flags'] = list(r.get('flags') or [])
    rec['flags'].append(origin)
    rec['notes'] = (str(rec.get('notes') or '') + ' ' + origin).strip()
    added.append(rec)

json.dump({'group': 'g8_supplement', 'records': added, 'excluded': [], 'product_notes': {},
           'drop_excluded': drop_excluded, 'reconciliation': recon},
          open(W + 'out/g8_supplement.json', 'w'), indent=1, ensure_ascii=False)
print(f'added {len(added)} records, overrode {len(drop_excluded)} base exclusions, {len(recon)} reconciliation rows')
