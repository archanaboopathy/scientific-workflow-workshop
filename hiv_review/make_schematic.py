"""Schematic of how the HIV Phase 3 master workbook, efficacy/resistance figures and network plot were made.

Run with matplotlib: python hiv_review/make_schematic.py -> hiv_review/figures/schematic_pipeline.png
"""
import os

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "figures", "schematic_pipeline.png")

fig, ax = plt.subplots(figsize=(17, 23))
ax.set_xlim(0, 100)
ax.set_ylim(0, 150)
ax.axis("off")


def box(x, y, w, h, title, body="", fc="#f4f6f8", ec="#455a64", tsize=11, bsize=9, lw=1.4, boff=3.0):
    ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.25,rounding_size=0.8",
                                fc=fc, ec=ec, lw=lw))
    ax.text(x + w / 2, y + h - 0.9, title, ha="center", va="top", fontsize=tsize, fontweight="bold")
    if body:
        ax.text(x + 0.9, y + h - boff, body, ha="left", va="top", fontsize=bsize, linespacing=1.35)


def arrow(x1, y1, x2, y2, style="-|>", ls="-", color="#37474f", rad=0.0):
    ax.add_patch(FancyArrowPatch((x1, y1), (x2, y2), arrowstyle=style, mutation_scale=16, lw=1.6,
                                 color=color, linestyle=ls, connectionstyle=f"arc3,rad={rad}"))


def section(y, text, color):
    ax.add_patch(FancyBboxPatch((0.5, y), 99, 2.6, boxstyle="round,pad=0.1", fc=color, ec="none"))
    ax.text(1.5, y + 1.3, text, ha="left", va="center", fontsize=13, fontweight="bold", color="white")


ax.text(50, 149.2, "How the HIV-1 Phase 3 trial review workbook and figures were generated",
        ha="center", va="top", fontsize=17, fontweight="bold")
ax.text(50, 146.9, "AI-assisted (GitHub Copilot agents) compilation from public sources. Not scientific validation: "
        "every value needs human review against its primary source.",
        ha="center", va="top", fontsize=10, style="italic", color="#b71c1c")

# ---------------------------------------------------------------- 1. master workbook
section(141.5, "1  Build the master workbook", "#1565c0")

box(2, 129, 30, 11, "Input workbook",
    "NIH_ARV_Phase3_Master_Week48_\nScatter_X_Offset.xlsx\n1st tab only ('Phase 3 Master'):\n63 rows, 41 products (column C)",
    fc="#e3f2fd")
box(35, 129, 30, 11, "read_master.py",
    "Parses the first tab into\nmaster_rows.json (original rows)\nproducts.json (41 products +\nPubMed search terms)", fc="#e3f2fd")
box(68, 129, 30, 11, "Scope decisions + PROTOCOL.md",
    "Adult HIV-1 Phase 3/3b treatment trials;\nproduct investigational in >=1 arm.\nWeek 48 FDA Snapshot <50 c/mL.\n"
    "Emergent resistance n/N (stated only).\nSource hierarchy; no extrapolation.", fc="#e3f2fd")
arrow(32.3, 134.5, 34.7, 134.5)
arrow(65.3, 134.5, 67.7, 134.5)

box(2, 116, 46, 10.5, "Candidate literature search (lit.py)",
    "PubMed: each product + HIV + Phase III -> 836 unique records\n"
    "Remove reviews/news/Phase II-only -> 627 candidates\n"
    "(480 likely primary reports, 147 secondary analyses)", fc="#e8eaf6")
box(51, 116, 47, 10.5, "Public sources used by every agent (lit.py)",
    "PubMed E-utilities  |  Europe PMC open-access full text\n"
    "ClinicalTrials.gov API v2 (records + posted results)\n"
    "DailyMed labels (sec. 14, 12.4)  |  FDA reviews, CROI/IAS abstracts", fc="#e8eaf6")
arrow(50, 128.7, 25, 126.8)
arrow(83, 128.7, 75, 126.8)

ax.text(50, 114.4, "7 parallel research sub-agents, one per ARV class group (groups.json)",
        ha="center", va="center", fontsize=12, fontweight="bold")
agents = [
    ("g1  INSTI: DTG", "DTG/3TC\nDTG/ABC/3TC\nDTG/RPV\nDTG", "#1f77b4"),
    ("g2  INSTI: EVG, RAL", "EVG/c/TAF/FTC\nEVG/c/TDF/FTC\nEVG\nRAL", "#17becf"),
    ("g3  LA INSTI +\nNNRTI: CAB, RPV", "CAB/RPV, CAB\nRPV/TAF/FTC\nRPV/TDF/FTC\nRPV", "#9467bd"),
    ("g4  NRTI /\nNRTTI backbones", "ABC/3TC, TAF/FTC\nTDF/FTC, TDF/3TC\nABC, FTC, 3TC\nTAF, TDF, ISL", "#7f7f7f"),
    ("g5  BIC, capsid,\nNNRTI", "BIC/TAF/FTC, BIC/LEN\nBIC, LEN, DOR/ISL\nDOR/TDF/3TC, DOR\nEFV products, ETR", "#2ca02c"),
    ("g6  Protease\ninhibitors", "DRV/c/TAF/FTC\nATV\nDRV\nRTV", "#d62728"),
    ("g7  Entry /\nattachment", "FTR\nIBA\nMVC", "#8c564b"),
]
w, gap = 13.1, 0.95
for i, (t, b, c) in enumerate(agents):
    x = 1.5 + i * (w + gap)
    ax.add_patch(FancyBboxPatch((x, 97), w, 14, boxstyle="round,pad=0.2,rounding_size=0.8",
                                fc=c, ec="black", lw=1.0, alpha=0.18))
    ax.add_patch(FancyBboxPatch((x, 107.2), w, 3.8, boxstyle="round,pad=0.2,rounding_size=0.8",
                                fc=c, ec="black", lw=1.0))
    ax.text(x + w / 2, 109.1, t, ha="center", va="center", fontsize=9, fontweight="bold", color="white")
    ax.text(x + w / 2, 106.2, b, ha="center", va="top", fontsize=8.3, linespacing=1.35)
    arrow(x + w / 2, 96.7, 25, 93.3, color="#78909c")
    if t.startswith(("g2", "g5", "g6")):
        arrow(x + w / 2 + 2, 96.7, 75, 93.3, ls="--", color="#c62828")
ax.text(50, 112.4, "Per product: screen candidates, extra ClinicalTrials.gov/PubMed searches, apply inclusion rule, "
        "extract Week 48 efficacy + resistance, re-check every existing row, log changes and exclusions, flag issues",
        ha="center", va="center", fontsize=8.5, style="italic", color="#37474f")
ax.text(99, 94.8, "dashed: g2, g5, g6 re-run", ha="right", va="center", fontsize=8, color="#c62828")

box(2, 83, 46, 10, "Base research run (used as the base)",
    "out/g1_DTG.json ... out/g7_ENTRY.json\nAll 41 products; one JSON per agent\n"
    "(agents interrupted once; partial results saved)", fc="#fff8e1")
box(52, 83, 46, 10, "Session-2 re-run (supplement only)",
    "session2/g2.json, g5.json, g6.json\nOther session-2 groups interrupted,\nnot used", fc="#fff8e1")

box(2, 71, 96, 8, "merge_supplement.py  ->  out/g8_supplement.json",
    "Match trials on product + NCT ID / acronym / PMID.  Add 11 trials found only in session 2 "
    "(e.g. Opti-DOR, DORA, SWEET, SWAN) and DOR/ISL 017-020 (user decision).\n"
    "Base-run decision kept for the other 9 disagreements; all 13 listed on the Reconciliation sheet.",
    fc="#fff3e0")
arrow(25, 82.7, 35, 79.3)
arrow(75, 82.7, 65, 79.3)

box(2, 57.5, 96, 9.5, "build_final.py  ->  HIV_Phase3_Trial_Review.xlsx  (master workbook)",
    "Phase 3 Master: one row per product x trial; original columns A-R + 12 new columns (population, resistance % "
    "on 2 denominators, raw counts,\ntimepoints, sources, PMID, design/arms, flags).  Resistance % = n/N only "
    "from integers stated in the source.\nOther sheets: Read Me, Sources & Changes, Excluded Trials, Product "
    "Notes, Reconciliation.   Checks: JSON parse, original rows preserved, scripts/validate.py.",
    fc="#e8f5e9", ec="#2e7d32", lw=2)
arrow(50, 70.7, 50, 67.3)

# ---------------------------------------------------------------- 2. efficacy / resistance figures
section(51, "", "#2e7d32")
ax.add_patch(FancyBboxPatch((50.5, 51), 49, 2.6, boxstyle="round,pad=0.1", fc="#ad1457", ec="none"))
ax.text(1.5, 52.3, "2  Efficacy and resistance figures (make_figures.py)", ha="left", va="center",
        fontsize=13, fontweight="bold", color="white")
ax.text(51.5, 52.3, "3  Network plot, Figure 1 (network/build_network.py)", ha="left", va="center",
        fontsize=13, fontweight="bold", color="white")

arrow(30, 57.2, 25, 50.0)
arrow(70, 57.2, 75, 50.0, ls="--")

box(2, 37, 46, 12, "Select plottable values from Phase 3 Master",
    "x = peer-reviewed publication year (column J)\n"
    "Efficacy = numeric Week 48 FDA Snapshot <50 c/mL (%)\n"
    "Resistance = emergent resistance to the product,\n   % of randomized/treated, only from stated counts\n"
    "One point per trial (coformulated row); not-reported,\n   pooled, non-Snapshot values dropped",
    fc="#e8f5e9", bsize=8.8)
arrow(25, 36.7, 25, 34.3)

figs = [
    ("i) Efficacy vs time", "fig1_efficacy_over_time\nn=64; OLS +0.55 pts/yr"),
    ("ii) Resistance vs time", "fig2_resistance_over_time\nn=62; OLS -0.09 pts/yr"),
    ("iii) Efficacy vs\nresistance", "fig3 (a: by year,\nb: scatter); n=41"),
]
for i, (t, b) in enumerate(figs):
    x = 2 + i * 15.6
    box(x, 22, 14.6, 12, t, b, fc="#c8e6c9", tsize=9.5, bsize=8, boff=5.2)
box(2, 11, 46, 9, "Variants and outputs",
    "*_trendline.png: unweighted OLS line of best fit (descriptive)\n"
    "fig4 iii-b: B/F/TAF trials (GS-US-380-1489/1490/1844/1878/\n4030/4449/4458) as filled maroon INSTI+NRTI symbols\n"
    "Symbols coloured by drug class; Plot Data + Figures sheets", fc="#e8f5e9", bsize=8.5)
arrow(25, 21.7, 25, 20.3)

# ---------------------------------------------------------------- 3. network plot
box(52, 41, 46, 8, "Input: final master workbook (user copy)",
    "HIV_Phase3_Trial_Review_from_GitHub_AB_110526.xlsx\n'Phase 3 Master' (169 rows): col AB Study Design/Arms,\n"
    "arm Ns, population, notes", fc="#fce4ec", bsize=8.5)
arrow(75, 40.7, 75, 39.3)
box(52, 29, 46, 10, "Curate arm table (one row per randomized arm)",
    "122 randomized trials with >=2 arms, de-duplicated across\nproduct rows.  Excluded: single-arm studies, VOLITION,\n"
    "WAVES open-label extension, placeholder rows.\nArm N as stated; N not stated -> 0 and node marked *",
    fc="#fce4ec", bsize=8.5)
arrow(75, 28.7, 75, 27.3)
box(52, 18, 46, 9, "Build network",
    "Node = full regimen (96), coloured by regimen class;\narea proportional to total participants\n"
    "Edge = trials directly comparing 2 nodes (width + label)\n'Continued baseline ART' hub at centre",
    fc="#fce4ec", bsize=8.5)
arrow(75, 17.7, 75, 16.3)
box(52, 7.5, 46, 8.5, "Figure 1: network_1_regimen.png",
    "96 nodes, 122 edges.  Also: anchor-agent (Fig 2, 36 nodes)\nand drug-class (Fig 3, 14 nodes) versions;\n"
    "HIV_Phase3_network_data.xlsx (arms, nodes, edges)",
    fc="#f8bbd0", ec="#ad1457", lw=2, bsize=8.5)

ax.add_patch(FancyBboxPatch((2, 1.0), 96, 4.5, boxstyle="round,pad=0.2", fc="#ffebee", ec="#b71c1c", lw=1.4))
ax.text(50, 3.25, "Human review still required: scientific accuracy of every extracted value, inclusion calls (Flags column, "
        "Reconciliation sheet), regimen/anchor/class mapping,\nand interpretation across populations "
        "(naive vs switch vs experienced). Automated tests cover repository code only.",
        ha="center", va="center", fontsize=9, color="#b71c1c")

os.makedirs(os.path.dirname(OUT), exist_ok=True)
fig.savefig(OUT, dpi=150, bbox_inches="tight", pad_inches=0.3)
print("wrote", OUT)
