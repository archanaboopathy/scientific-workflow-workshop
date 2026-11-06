"""Figures for HIV_Phase3_Trial_Review.xlsx (run after build_final.py).

Reads the Phase 3 Master sheet and writes:
  figures/fig1_efficacy_over_time.png
  figures/fig2_resistance_over_time.png
  figures/fig3_efficacy_vs_resistance.png
  figures/fig*_trendline.png (same figures with an unweighted OLS line of best fit)
and adds "Plot Data" and "Figures" sheets to the workbook.

Rules (no extrapolation):
- Efficacy = the product-arm Week 48 FDA Snapshot HIV-1 RNA <50 copies/mL % exactly as stated at
  the start of column L. Rows whose column L is "Not reported", "Not applicable" or pooled are not
  plotted; alternative endpoints (TLOVR, Week 24, etc.) are never plotted.
- Resistance = "TE Resistance to Product - % of Randomized/Treated", which build_final.py
  calculates only from counts stated in the source. Rows without stated counts are not plotted.
- Filled markers: resistance assessed at Week 48-52 (incl. Month 11/12); hollow markers: other
  timepoints (e.g. Week 96/144, Week 25/26, end of study).
- X axis = peer-reviewed publication year (column J), with small horizontal offsets so points from
  the same year don't overlap.
- A trial listed under several products (e.g. CAB/RPV and CAB) is plotted once, using the
  coformulated (complete-regimen) row. Rows from the same trial with different arms (e.g. ACTG
  A5257 ATV/r and DRV/r) are kept.
"""
import os, re
from collections import defaultdict
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.lines import Line2D
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font

HERE = os.path.dirname(os.path.abspath(__file__))
WB = os.path.join(HERE, '..', 'HIV_Phase3_Trial_Review.xlsx')
FIG = os.path.join(HERE, 'figures')
os.makedirs(FIG, exist_ok=True)

wb = load_workbook(WB)
ws = wb['Phase 3 Master']
h = [c.value for c in ws[1]]
col = {k: h.index(k) for k in h}
EFF = 'Week 48 FDA Snapshot: HIV-1 RNA <50 copies/mL'
RES = 'TE Resistance to Product - % of Randomized/Treated (n/N)'


def eff_value(cell):
    t = str(cell or '').split(' | Endpoint used')[0].strip()
    if not t or t.lower().startswith(('not ', 'pooled')) or 'pooled' in t.lower()[:40]:
        return None
    m = re.match(r'^(\d+(?:\.\d+)?)%', t)
    return float(m.group(1)) if m else None


def res_value(cell):
    m = re.match(r'^(\d+(?:\.\d+)?)% \((\d+)/(\d+)\)', str(cell or ''))
    return (float(m.group(1)), int(m.group(2)), int(m.group(3))) if m else (None, None, None)


def wk48(tp):
    t = str(tp or '').lower()
    if re.search(r'96|144|213|123 weeks|week 2[56]|end of treatment|not specified', t):
        return False
    return bool(re.search(r'48|52|month 1[12]', t))


rows = []
for i, r in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
    year = r[col['Peer-Reviewed Publication Year']]
    try:
        year = int(year)
    except (TypeError, ValueError):
        year = 0
    e = eff_value(r[col[EFF]])
    rp, rn, rN = res_value(r[col[RES]])
    if year < 1990 or (e is None and rp is None):
        continue
    rows.append(dict(row=i, product=r[col['NIH Abbreviation/Components']], cls=r[col['Drug Class']],
                     trial=r[col['Registrational Phase 3 Trial Name']], nct=r[col['ClinicalTrials.gov ID']],
                     year=year, pop=r[col['Treatment Population']], eff=e, res=rp, res_n=rn, res_N=rN,
                     res_tp=r[col['Resistance Timepoint']], res_w48=wk48(r[col['Resistance Timepoint']])))

# One point per trial (prefer coformulated row); keep distinct arms of the same trial.
groups = defaultdict(list)
for d in rows:
    key = d['nct'] if str(d['nct']).startswith('NCT') else d['trial']
    groups[key].append(d)
plot = []
for key, g in groups.items():
    Ns = [d['res_N'] for d in g]
    if len(g) > 1 and None not in Ns and len(set(Ns)) == len(Ns):
        plot += g
        continue
    g.sort(key=lambda d: ('+' not in str(d['cls']), d['row']))
    keep = g[0]
    keep['also_listed_under'] = ', '.join(d['product'] for d in g[1:])
    plot.append(keep)
for d in plot:
    d.setdefault('also_listed_under', '')

classes = sorted({d['cls'] for d in plot}, key=lambda c: (c.split()[0], c))
palette = plt.get_cmap('tab20').colors
markers = ['o', 's', '^', 'D', 'v', 'P', 'X', '<', '>', 'h', '*', 'p', 'd', '8', 'H']
STY = {c: (palette[(2 * k) % 20 if k < 10 else (2 * k - 19) % 20], markers[k % len(markers)]) for k, c in enumerate(classes)}

# Deterministic horizontal offsets within a year.
def offsets(items, key):
    by_year = defaultdict(list)
    for d in items:
        if d[key] is not None:
            by_year[d['year']].append(d)
    out = {}
    for y, ds in by_year.items():
        ds.sort(key=lambda d: (d['cls'], d['trial']))
        n = len(ds)
        for j, d in enumerate(ds):
            out[id(d)] = y + (0 if n == 1 else (-0.35 + 0.7 * j / (n - 1)))
    return out


def class_legend(ax, used, loc='lower left', extra=None):
    handles = [Line2D([], [], marker=STY[c][1], color=STY[c][0], linestyle='', markersize=8, label=c) for c in classes if c in used]
    if extra:
        handles += extra
    ax.legend(handles=handles, loc='upper left', bbox_to_anchor=(1.01, 1), fontsize=8, frameon=True,
              title='Drug class (column B)', title_fontsize=8)


def scatter(ax, items, xk, yk, xoff=None, filled_by=None, size=55):
    for d in items:
        if d[yk] is None or (xk != 'year' and d[xk] is None):
            continue
        x = xoff[id(d)] if xoff else d[xk]
        c, m = STY[d['cls']]
        filled = True if filled_by is None else d[filled_by]
        ax.scatter(x, d[yk], s=size if not callable(size) else size(d), marker=m,
                   facecolors=c if filled else 'none', edgecolors=c if not filled else 'black',
                   linewidths=1.3 if not filled else 0.5, alpha=0.9, zorder=3)


NOTE = 'AI-assisted compilation; not scientifically validated. Source: HIV_Phase3_Trial_Review.xlsx (Phase 3 Master).'

def fit_line(ax, items, xk, yk, unit):
    """Ordinary least-squares line over the observed x range only (unweighted, one point per trial)."""
    pts = [(d[xk], d[yk]) for d in items if d[xk] is not None and d[yk] is not None]
    x = np.array([p[0] for p in pts], float); y = np.array([p[1] for p in pts], float)
    slope, icpt = np.polyfit(x, y, 1)
    r2 = 1 - ((y - (slope * x + icpt)) ** 2).sum() / ((y - y.mean()) ** 2).sum()
    xs = np.array([x.min(), x.max()])
    ax.plot(xs, slope * xs + icpt, color='black', linestyle='--', linewidth=1.5, zorder=4,
            label=f'OLS fit: slope {slope:+.2f} {unit}, R² = {r2:.2f}, n = {len(x)}')
    ax.set_title(ax.get_title() + f'\nDashed line (OLS fit): slope {slope:+.2f} {unit}; R² = {r2:.2f}; n = {len(x)} trials', fontsize=10)
    FITS.append((ax.get_title().split(chr(10))[0], xk, yk, slope, icpt, r2, len(x)))


FITS = []
OUT = {}
for TREND, SUF in ((False, ''), (True, '_trendline')):
    TITLE_SUF = 'With line of best fit: ' if TREND else ''
    FIT_NOTE = '\nDashed line: unweighted OLS across observed range only; descriptive, not a statistical test.' if TREND else ''
    # Figure 1: efficacy over time
    E = [d for d in plot if d['eff'] is not None]
    fig, ax = plt.subplots(figsize=(13, 7))
    scatter(ax, E, 'year', 'eff', offsets(E, 'eff'))
    for d in E:
        if d['eff'] < 80:
            ax.annotate(re.split(r' \(|,| \[', d['trial'])[0], (offsets(E, 'eff')[id(d)], d['eff']), fontsize=7,
                        xytext=(4, -3), textcoords='offset points')
    ax.set_xlabel('Publication year (peer-reviewed primary publication)')
    ax.set_ylabel('Week 48 FDA Snapshot HIV-1 RNA <50 copies/mL (%), product arm')
    ax.set_title(TITLE_SUF + f'i) Week 48 efficacy (FDA Snapshot) over time - {len(E)} trials')
    ax.grid(alpha=0.3)
    if TREND:
        fit_line(ax, E, 'year', 'eff', '%-points/year')
    class_legend(ax, {d['cls'] for d in E})
    fig.text(0.01, 0.005, NOTE + FIT_NOTE + ' Only trials with a stated Week 48 Snapshot value are shown (pre-2008 trials used other endpoints).', fontsize=7)
    fig.tight_layout(rect=(0, 0.04, 1, 1))
    f1 = os.path.join(FIG, f'fig1_efficacy_over_time{SUF}.png'); fig.savefig(f1, dpi=150); plt.close(fig)

    # Figure 2: resistance over time
    Rr = [d for d in plot if d['res'] is not None]
    fig, ax = plt.subplots(figsize=(13, 7))
    xo = offsets(Rr, 'res')
    scatter(ax, Rr, 'year', 'res', xo, filled_by='res_w48')
    for d in Rr:
        if d['res'] >= 3:
            ax.annotate(re.split(r' \(|,| \[', d['trial'])[0], (xo[id(d)], d['res']), fontsize=7, xytext=(5, 2), textcoords='offset points')
    ax.set_xlabel('Publication year (peer-reviewed primary publication)')
    ax.set_ylabel('Treatment-emergent resistance to product (% of randomized/treated)')
    ax.set_title(TITLE_SUF + f'ii) Treatment-emergent resistance to the treatment drug over time - {len(Rr)} trials')
    ax.grid(alpha=0.3)
    if TREND:
        fit_line(ax, Rr, 'year', 'res', '%-points/year')
    tp = [Line2D([], [], marker='o', color='grey', markerfacecolor='grey', linestyle='', label='Resistance at Week 48-52'),
          Line2D([], [], marker='o', color='grey', markerfacecolor='none', linestyle='', label='Other timepoint (e.g. Wk 25/26, 96, 144, end of study)')]
    class_legend(ax, {d['cls'] for d in Rr}, loc='upper left', extra=tp)
    fig.text(0.01, 0.005, NOTE + FIT_NOTE + ' % = n/N calculated only from counts stated in the source.', fontsize=7)
    fig.tight_layout(rect=(0, 0.04, 1, 1))
    f2 = os.path.join(FIG, f'fig2_resistance_over_time{SUF}.png'); fig.savefig(f2, dpi=150); plt.close(fig)

    # Figure 3: efficacy vs resistance (trials with both values)
    B = [d for d in plot if d['eff'] is not None and d['res'] is not None]
    fig, (a1, a2) = plt.subplots(1, 2, figsize=(18, 7.5), gridspec_kw={'width_ratios': [1.25, 1]})
    xo = offsets(B, 'eff')
    size = lambda d: 40 + 120 * d['res']
    scatter(a1, B, 'year', 'eff', xo, filled_by='res_w48', size=size)
    a1.set_xlabel('Publication year'); a1.set_ylabel('Week 48 FDA Snapshot <50 copies/mL (%)')
    a1.set_title('iii-a) Efficacy over time; marker size = % emergent resistance')
    a1.grid(alpha=0.3)
    if TREND:
        fit_line(a1, B, 'year', 'eff', '%-points/year')
    sz = [Line2D([], [], marker='o', color='grey', markerfacecolor='none', linestyle='', markersize=(40 + 120 * v) ** 0.5, label=f'{v}% resistance') for v in (0, 1, 2, 4)]
    class_legend(a1, {d['cls'] for d in B}, loc='lower left', extra=sz)
    scatter(a2, B, 'eff', 'res', filled_by='res_w48', size=70)
    for d in B:
        if d['res'] > 0:
            a2.annotate(str(d['year']), (d['eff'], d['res']), fontsize=6.5, xytext=(4, 3), textcoords='offset points', color='dimgray')
    a2.text(0.02, 0.06, f"{sum(d['res'] == 0 for d in B)} trials at 0% resistance (along x axis)", transform=a2.transAxes, fontsize=8, color='dimgray')
    a2.set_xlabel('Week 48 FDA Snapshot <50 copies/mL (%)'); a2.set_ylabel('Treatment-emergent resistance to product (% of randomized/treated)')
    a2.set_title('iii-b) Efficacy vs resistance (labels = publication year)')
    a2.grid(alpha=0.3)
    if TREND:
        fit_line(a2, B, 'eff', 'res', 'per efficacy %-point')
    fig.suptitle(TITLE_SUF + f'iii) Efficacy vs treatment-emergent resistance - {len(B)} trials with both values stated', fontsize=12)
    fig.text(0.01, 0.005, NOTE + FIT_NOTE + ' Hollow markers: resistance not at Week 48-52.', fontsize=7)
    fig.tight_layout(rect=(0, 0.04, 1, 0.96))
    f3 = os.path.join(FIG, f'fig3_efficacy_vs_resistance{SUF}.png'); fig.savefig(f3, dpi=150); plt.close(fig)

    OUT[SUF] = (f1, f2, f3)

# Workbook sheets
for name in ('Plot Data', 'Figures'):
    if name in wb.sheetnames:
        del wb[name]
pd_ = wb.create_sheet('Plot Data')
cols = ['Master row', 'Product', 'Drug class', 'Trial', 'NCT ID', 'Publication year', 'Population',
        'Week 48 Snapshot <50 c/mL (%) plotted', 'Resistance % (randomized/treated) plotted', 'Resistance n', 'Resistance N',
        'Resistance timepoint', 'Resistance at Week 48-52?', 'Also listed under (not plotted twice)', 'In fig 1', 'In fig 2', 'In fig 3']
pd_.append(cols)
for d in sorted(plot, key=lambda d: (d['year'], d['cls'], d['trial'])):
    pd_.append([d['row'], d['product'], d['cls'], d['trial'], d['nct'], d['year'], d['pop'], d['eff'], d['res'], d['res_n'], d['res_N'],
                d['res_tp'], 'Yes' if d['res_w48'] else 'No', d['also_listed_under'],
                'Yes' if d['eff'] is not None else '', 'Yes' if d['res'] is not None else '',
                'Yes' if d['eff'] is not None and d['res'] is not None else ''])
for c in pd_[1]:
    c.font = Font(bold=True)
pd_.freeze_panes = 'A2'
fg = wb.create_sheet('Figures')
fg['A1'] = 'Figures generated by hiv_review/make_figures.py from the Phase 3 Master sheet (data on "Plot Data"). ' + NOTE
fg['A1'].font = Font(bold=True)
for k, f in enumerate(OUT[''] + OUT['_trendline']):
    img = XLImage(f)
    img.width, img.height = img.width * 0.55, img.height * 0.55
    fg.add_image(img, f'A{3 + k * 46}')
wb.save(WB)
for t in FITS:
    print('fit', t)
print(f'plotted: fig1 {len(E)}, fig2 {len(Rr)}, fig3 {len(B)} trials; deduplicated {len(rows) - len(plot)} repeat rows')
