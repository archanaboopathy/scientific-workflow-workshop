"""Figure-2-style network plots of Phase 3 HIV-1 trials (regimen, anchor-agent and class level).

Arm-level data were transcribed from the 'Phase 3 Master' sheet of
HIV_Phase3_Trial_Review_from_GitHub_AB_110526.xlsx (efficacy, design and notes columns).
Only randomized trials with >=2 randomized arms are included. Arm N is taken as stated in the
workbook; 'NR' = not stated per arm (contributes 0 to node size); 'derived' = stated total minus
stated arm N (flagged).
"""
import math, itertools, os, textwrap
from collections import defaultdict
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D
import openpyxl

OUT = os.path.dirname(os.path.abspath(__file__))
SOC = "Continued baseline ART"
PBO = "Placebo + background"

# regimen -> (anchor, class)
REG = {
    "DTG/3TC": ("DTG+3TC", "INSTI 2DR (oral)"),
    "DTG/RPV": ("DTG+RPV", "INSTI 2DR (oral)"),
    "DTG+TDF/FTC": ("DTG", "INSTI 3DR"),
    "DTG+FTC/TAF": ("DTG", "INSTI 3DR"),
    "DTG/ABC/3TC": ("DTG", "INSTI 3DR"),
    "DTG+TDF/3TC": ("DTG", "INSTI 3DR"),
    "DTG+2 NRTIs": ("DTG", "INSTI 3DR"),
    "DTG+background (TE)": ("DTG", "INSTI 3DR"),
    "B/F/TAF": ("BIC", "INSTI 3DR"),
    "E/C/F/TAF": ("EVG", "INSTI 3DR"),
    "E/C/F/TAF+DRV": ("EVG", "INSTI 3DR"),
    "E/C/F/TDF": ("EVG", "INSTI 3DR"),
    "EVG+PI/r+agent (TE)": ("EVG", "INSTI 3DR"),
    "RAL+TDF/FTC": ("RAL", "INSTI 3DR"),
    "RAL QD 800 mg+TDF/FTC": ("RAL", "INSTI 3DR"),
    "RAL HD 1200 mg QD+TDF/FTC": ("RAL", "INSTI 3DR"),
    "RAL+TDF/3TC": ("RAL", "INSTI 3DR"),
    "RAL+2 NRTIs": ("RAL", "INSTI 3DR"),
    "RAL+background (TE)": ("RAL", "INSTI 3DR"),
    "CAB+RPV LA Q4W": ("CAB+RPV LA", "LA INSTI+NNRTI"),
    "CAB+RPV LA Q8W": ("CAB+RPV LA", "LA INSTI+NNRTI"),
    "BIC/LEN": ("BIC+LEN", "Capsid inhibitor-containing"),
    "ISL/LEN weekly": ("ISL+LEN", "Capsid inhibitor-containing"),
    "LEN+background (TE)": ("LEN", "Capsid inhibitor-containing"),
    "DOR/ISL": ("DOR+ISL", "NNRTI+NRTTI 2DR"),
    "DOR/3TC/TDF": ("DOR", "NNRTI 3DR"),
    "DOR+2 NRTIs": ("DOR", "NNRTI 3DR"),
    "EFV/FTC/TDF": ("EFV", "NNRTI 3DR"),
    "EFV+TDF/FTC": ("EFV", "NNRTI 3DR"),
    "EFV 400 mg+TDF/FTC": ("EFV", "NNRTI 3DR"),
    "EFV 400 mg+TDF/3TC": ("EFV", "NNRTI 3DR"),
    "EFV+TDF/3TC": ("EFV", "NNRTI 3DR"),
    "EFV+2 NRTIs": ("EFV", "NNRTI 3DR"),
    "EFV+ZDV/3TC": ("EFV", "NNRTI 3DR"),
    "EFV+ABC/3TC": ("EFV", "NNRTI 3DR"),
    "EFV+TDF+3TC": ("EFV", "NNRTI 3DR"),
    "EFV+d4T+3TC": ("EFV", "NNRTI 3DR"),
    "EFV+ABC QD+3TC": ("EFV", "NNRTI 3DR"),
    "EFV+ABC BID+3TC": ("EFV", "NNRTI 3DR"),
    "EFV+3TC QD+ZDV": ("EFV", "NNRTI 3DR"),
    "EFV+3TC BID+ZDV": ("EFV", "NNRTI 3DR"),
    "EFV+FTC+ddI": ("EFV", "NNRTI 3DR"),
    "EFV+d4T+ddI": ("EFV", "NNRTI 3DR"),
    "RPV/FTC/TDF": ("RPV", "NNRTI 3DR"),
    "RPV/FTC/TAF": ("RPV", "NNRTI 3DR"),
    "RPV+TDF/FTC": ("RPV", "NNRTI 3DR"),
    "RPV+2 NRTIs": ("RPV", "NNRTI 3DR"),
    "ETR+background (TE)": ("ETR", "NNRTI 3DR"),
    "Loviride+3TC+ZDV": ("Loviride (investigational NNRTI)", "NNRTI 3DR"),
    "D/C/F/TAF": ("DRV/c", "Boosted PI 3DR"),
    "DRV/c+F/TDF": ("DRV/c", "Boosted PI 3DR"),
    "DRV/r+2 NRTIs": ("DRV/r", "Boosted PI 3DR"),
    "DRV/r+TDF/FTC": ("DRV/r", "Boosted PI 3DR"),
    "DRV/r 400/100+2 NRTIs": ("DRV/r", "Boosted PI 3DR"),
    "DRV/r BID+OBR (TE)": ("DRV/r", "Boosted PI 3DR"),
    "ATV/r+TDF/FTC": ("ATV/r", "Boosted PI 3DR"),
    "ATV/r+ABC/3TC": ("ATV/r", "Boosted PI 3DR"),
    "ATV/r+TDF+NRTI (TE)": ("ATV/r", "Boosted PI 3DR"),
    "LPV/r+2 NRTIs": ("LPV/r", "Boosted PI 3DR"),
    "LPV/r+TDF/FTC": ("LPV/r", "Boosted PI 3DR"),
    "LPV/r+ABC+ddI": ("LPV/r", "Boosted PI 3DR"),
    "LPV/r+TDF+NRTI (TE)": ("LPV/r", "Boosted PI 3DR"),
    "LPV/r+OBR (TE)": ("LPV/r", "Boosted PI 3DR"),
    "ATV+ZDV/3TC": ("ATV (unboosted)", "Unboosted / early PI"),
    "ATV+2 NRTIs (TE)": ("ATV (unboosted)", "Unboosted / early PI"),
    "ATV+SQV+TDF+NRTI (TE)": ("ATV+SQV", "Unboosted / early PI"),
    "IDV+ZDV/3TC": ("IDV", "Unboosted / early PI"),
    "RTV 600 mg BID+NRTIs": ("RTV (full-dose PI)", "Unboosted / early PI"),
    "EFV+IDV": ("EFV+IDV", "NRTI-sparing / PI monotherapy"),
    "LPV/r+RAL": ("LPV/r+RAL", "NRTI-sparing / PI monotherapy"),
    "DRV/r+RAL": ("DRV/r+RAL", "NRTI-sparing / PI monotherapy"),
    "MVC+DRV/r": ("MVC+DRV/r", "NRTI-sparing / PI monotherapy"),
    "LPV/r monotherapy": ("LPV/r monotherapy", "NRTI-sparing / PI monotherapy"),
    "DRV/r monotherapy": ("DRV/r monotherapy", "NRTI-sparing / PI monotherapy"),
    "MVC QD+OBT": ("MVC", "CCR5 antagonist"),
    "MVC BID+OBT": ("MVC", "CCR5 antagonist"),
    "MVC BID+ZDV/3TC": ("MVC", "CCR5 antagonist"),
    "MVC QD+ZDV/3TC": ("MVC", "CCR5 antagonist"),
    "FTR+background (TE)": ("FTR", "Attachment inhibitor"),
    "TDF+background (TE)": ("TDF add-on", "NRTI-only / NRTI add-on"),
    "ABC+ZDV/3TC (triple NRTI)": ("NRTI-only", "NRTI-only / NRTI add-on"),
    "3TC+ZDV": ("NRTI-only", "NRTI-only / NRTI add-on"),
    "3TC monotherapy": ("NRTI-only", "NRTI-only / NRTI add-on"),
    "ZDV monotherapy": ("NRTI-only", "NRTI-only / NRTI add-on"),
    "ddC+ZDV": ("NRTI-only", "NRTI-only / NRTI add-on"),
    "ABC/3TC FDC QD+PI/NNRTI": ("Backbone-only comparison", "Backbone-only comparison"),
    "ABC+3TC BID+PI/NNRTI": ("Backbone-only comparison", "Backbone-only comparison"),
    "ABC/3TC FDC+TDF+PI/NNRTI (TE)": ("Backbone-only comparison", "Backbone-only comparison"),
    "ABC+3TC separate+TDF+PI/NNRTI (TE)": ("Backbone-only comparison", "Backbone-only comparison"),
    "F/TAF+third agent": ("Backbone-only comparison", "Backbone-only comparison"),
    "F/TDF+third agent": ("Backbone-only comparison", "Backbone-only comparison"),
    "ABC/3TC+third agent": ("Backbone-only comparison", "Backbone-only comparison"),
    "FTC+ZDV/d4T+PI/NNRTI": ("Backbone-only comparison", "Backbone-only comparison"),
    "3TC+ZDV/d4T+PI/NNRTI": ("Backbone-only comparison", "Backbone-only comparison"),
    SOC: (SOC, SOC),
    PBO: (PBO, PBO),
}

# trial, NCT, workbook row(s), [(regimen, N or 'NR', N-source flag)]
S, D, NR = "stated", "derived", "NR"
TRIALS = [
    ("ASPIRE", "NCT02263326", "2", [("DTG/3TC", 44, S), (SOC, 45, S)]),
    ("GEMINI-1", "NCT02831673", "3", [("DTG/3TC", 356, S), ("DTG+TDF/FTC", 358, S)]),
    ("GEMINI-2", "NCT02831764", "4", [("DTG/3TC", 360, S), ("DTG+TDF/FTC", 359, S)]),
    ("TANGO", "NCT03446573", "5", [("DTG/3TC", 369, S), (SOC, 372, S)]),
    ("SALSA", "NCT04021290", "7", [("DTG/3TC", 246, S), (SOC, 247, S)]),
    ("GS-US-380-1489", "NCT02607930", "10", [("B/F/TAF", 314, S), ("DTG/ABC/3TC", 315, S)]),
    ("GS-US-380-1490", "NCT02607956", "11", [("B/F/TAF", 320, S), ("DTG+FTC/TAF", 325, S)]),
    ("GS-US-380-1844", "NCT02603120", "12", [("B/F/TAF", 282, S), ("DTG/ABC/3TC", 281, S)]),
    ("GS-US-380-1878", "NCT02603107", "13", [("B/F/TAF", 290, S), (SOC, 287, S)]),
    ("GS-US-380-1961", "NCT02652624", "14", [("B/F/TAF", 234, S), (SOC, 236, S)]),
    ("BRAAVE 2020", "NCT03631732", "15", [("B/F/TAF", 330, S), (SOC, 165, S)]),
    ("GS-US-380-4030", "NCT03110380", "16", [("B/F/TAF", 284, S), ("DTG+FTC/TAF", 281, S)]),
    ("ALLIANCE", "NCT03547908", "18", [("B/F/TAF", 119, S), ("DTG+TDF/FTC", 122, S)]),
    ("LAPTOP", "NCT03696160", "19", [("B/F/TAF", 220, S), ("D/C/F/TAF", 222, S)]),
    ("SINGLE", "NCT01263015", "20,105", [("DTG/ABC/3TC", 414, S), ("EFV/FTC/TDF", 419, S)]),
    ("ARIA", "NCT01910402", "21", [("DTG/ABC/3TC", 248, S), ("ATV/r+TDF/FTC", 247, S)]),
    ("STRIIVING", "NCT02105987", "22", [("DTG/ABC/3TC", 275, S), (SOC, 278, S)]),
    ("GS-US-292-0104", "NCT01780506", "23", [("E/C/F/TAF", 435, S), ("E/C/F/TDF", 432, S)]),
    ("GS-US-292-0111", "NCT01797445", "24", [("E/C/F/TAF", 431, S), ("E/C/F/TDF", 435, S)]),
    ("GS-US-292-0109", "NCT01815736", "25", [("E/C/F/TAF", 959, S), (SOC, 477, S)]),
    ("GS-US-292-0119", "NCT01968551", "28", [("E/C/F/TAF+DRV", 89, S), (SOC, 46, S)]),
    ("GS-US-292-1824", "NCT02605954", "31", [("E/C/F/TAF", 183, S), (SOC, 91, S)]),
    ("GS-US-292-1826", "NCT02616783", "32", [("E/C/F/TAF", 109, S), (SOC, 55, S)]),
    ("GS-US-236-0102", "NCT01095796", "34", [("E/C/F/TDF", 348, S), ("EFV/FTC/TDF", 352, S)]),
    ("GS-US-236-0103", "NCT01106586", "35", [("E/C/F/TDF", 353, S), ("ATV/r+TDF/FTC", 355, S)]),
    ("STRATEGY-NNRTI", "NCT01495702", "37", [("E/C/F/TDF", 290, S), (SOC, 143, S)]),
    ("STRATEGY-PI", "NCT01475838", "38", [("E/C/F/TDF", 290, S), (SOC, 139, S)]),
    ("WAVES", "NCT01705574", "40", [("E/C/F/TDF", 289, S), ("ATV/r+TDF/FTC", 286, S)]),
    ("ATLAS", "NCT02951052", "42,96", [("CAB+RPV LA Q4W", 308, S), (SOC, 308, S)]),
    ("ATLAS-2M", "NCT03299049", "43,97", [("CAB+RPV LA Q8W", 522, S), ("CAB+RPV LA Q4W", 523, S)]),
    ("FLAIR", "NCT02938520", "44,98", [("CAB+RPV LA Q4W", 283, S), ("DTG/ABC/3TC", 283, S)]),
    ("SOLAR", "NCT04542070", "45,99", [("CAB+RPV LA Q8W", 447, S), ("B/F/TAF", 223, S)]),
    ("CARES", "PACTR202104874490818", "46,100", [("CAB+RPV LA Q8W", 255, S), (SOC, 257, S)]),
    ("IMPALA", "NCT05546242", "47,101", [("CAB+RPV LA Q8W", 268, S), ("DTG+2 NRTIs", 269, S)]),
    ("LATITUDE (A5359)", "NCT03635788", "48,102", [("CAB+RPV LA Q4W", 152, S), (SOC, 154, S)]),
    ("SWORD-1", "NCT02429791", "50", [("DTG/RPV", 252, S), (SOC, 256, S)]),
    ("SWORD-2", "NCT02422797", "51", [("DTG/RPV", 261, S), (SOC, 255, S)]),
    ("ARTISTRY-1", "NCT05502341", "52,159", [("BIC/LEN", 371, S), (SOC, 186, S)]),
    ("ARTISTRY-2", "NCT06333808", "53,160", [("BIC/LEN", 383, S), ("B/F/TAF", 191, S)]),
    ("MK-8591A-051", "NCT05631093", "54", [("DOR/ISL", 366, S), (SOC, 185, S)]),
    ("MK-8591A-052", "NCT05630755", "55", [("DOR/ISL", 342, S), ("B/F/TAF", 171, S)]),
    ("MK-8591A-053", "NCT05705349", "56", [("DOR/ISL", 269, S), ("B/F/TAF", 267, S)]),
    ("DRIVE-AHEAD", "NCT02403674", "57,150", [("DOR/3TC/TDF", 364, S), ("EFV/FTC/TDF", 364, S)]),
    ("DRIVE-SHIFT", "NCT02397096", "58,151", [("DOR/3TC/TDF", 447, S), (SOC, 223, S)]),
    ("Study 934", "NCT00112047", "59,75,83,94", [("EFV+TDF/FTC", 257, S), ("EFV+ZDV/3TC", 254, S)]),
    ("Study 073", "-", "60", [("EFV/FTC/TDF", 203, S), (SOC, 97, S)]),
    ("GS-US-366-1160", "NCT02345226", "61", [("RPV/FTC/TAF", 438, S), ("EFV/FTC/TDF", 437, S)]),
    ("GS-US-366-1216", "NCT02345252", "62", [("RPV/FTC/TAF", 316, S), ("RPV/FTC/TDF", 313, S)]),
    ("SPIRIT", "NCT01252940", "63", [("RPV/FTC/TDF", 317, S), (SOC, NR, NR)]),
    ("STaR", "NCT01309243", "64", [("RPV/FTC/TDF", 394, S), ("EFV/FTC/TDF", 392, S)]),
    ("SALIF", "NCT01709084", "65", [("RPV/FTC/TDF", 213, S), ("EFV/FTC/TDF", 211, S)]),
    ("AMBER", "NCT02431247", "66", [("D/C/F/TAF", 362, S), ("DRV/c+F/TDF", 363, S)]),
    ("EMERALD", "NCT02269917", "67", [("D/C/F/TAF", 763, S), (SOC, 378, S)]),
    ("ESS30008 (SEAL)", "NCT00046176", "70", [("ABC/3TC FDC QD+PI/NNRTI", NR, NR), ("ABC+3TC BID+PI/NNRTI", NR, NR)]),
    ("CAL30001", "NCT00044577", "71", [("ABC/3TC FDC+TDF+PI/NNRTI (TE)", NR, NR), ("ABC+3TC separate+TDF+PI/NNRTI (TE)", NR, NR)]),
    ("ACTG A5202", "NCT00118898", "72,76,136", [("EFV+ABC/3TC", 465, S), ("ATV/r+ABC/3TC", 463, S), ("EFV+TDF/FTC", 464, S), ("ATV/r+TDF/FTC", 465, S)]),
    ("GS-US-311-1089", "NCT02121795", "73", [("F/TAF+third agent", 333, S), ("F/TDF+third agent", 330, S)]),
    ("GS-US-311-1717", "NCT02469246", "74", [("F/TAF+third agent", 253, S), ("ABC/3TC+third agent", 248, S)]),
    ("CNA3005", "-", "78", [("ABC+ZDV/3TC (triple NRTI)", 262, S), ("IDV+ZDV/3TC", 265, S)]),
    ("CNA30024", "-", "79", [("EFV+ABC/3TC", 324, S), ("EFV+ZDV/3TC", 325, S)]),
    ("CNA30021 (ZODIAC)", "-", "80", [("EFV+ABC QD+3TC", 384, S), ("EFV+ABC BID+3TC", 386, S)]),
    ("FTC-301A", "NCT00006208", "81", [("EFV+FTC+ddI", 286, S), ("EFV+d4T+ddI", 285, S)]),
    ("FTC-303", "NCT00002416", "82", [("FTC+ZDV/d4T+PI/NNRTI", 294, S), ("3TC+ZDV/d4T+PI/NNRTI", 146, S)]),
    ("NUCA3001", "-", "85", [("3TC+ZDV", NR, NR), ("3TC monotherapy", NR, NR), ("ZDV monotherapy", NR, NR)]),
    ("NUCA3002", "-", "86", [("3TC+ZDV", NR, NR), ("ddC+ZDV", NR, NR)]),
    ("NUCB3001", "-", "87", [("3TC+ZDV", NR, NR), ("ZDV monotherapy", NR, NR)]),
    ("NUCB3002", "-", "88", [("3TC+ZDV", NR, NR), ("ZDV monotherapy", NR, NR)]),
    ("CAESAR", "-", "89", [("3TC+ZDV", 907, S), ("Loviride+3TC+ZDV", 462, S), (PBO, 471, S)]),
    ("EPV20001", "-", "90", [("EFV+3TC QD+ZDV", 278, S), ("EFV+3TC BID+ZDV", 276, S)]),
    ("GS-00-907", "NCT00002450", "92", [("TDF+background (TE)", 368, S), (PBO, 182, S)]),
    ("GS-99-903", "NCT00158821", "93", [("EFV+TDF+3TC", 299, S), ("EFV+d4T+3TC", 301, S)]),
    ("SAILING", "NCT01231516", "104", [("DTG+background (TE)", 354, S), ("RAL+background (TE)", 361, S)]),
    ("SPRING-2", "NCT01227824", "106", [("DTG+2 NRTIs", 411, S), ("RAL+2 NRTIs", 411, S)]),
    ("FLAMINGO", "NCT01449929", "107", [("DTG+2 NRTIs", 242, S), ("DRV/r+2 NRTIs", 242, S)]),
    ("VIKING-4", "NCT01568892", "109", [("DTG+background (TE)", NR, NR), (PBO, NR, NR)]),
    ("ADVANCE", "NCT03122262", "110", [("DTG+FTC/TAF", NR, NR), ("DTG+TDF/FTC", NR, NR), ("EFV/FTC/TDF", NR, NR)]),
    ("DAWNING", "NCT02227238", "111", [("DTG+2 NRTIs", 312, S), ("LPV/r+2 NRTIs", 312, S)]),
    ("NAMSAL", "NCT02777229", "112", [("DTG+TDF/3TC", 310, S), ("EFV 400 mg+TDF/3TC", 303, S)]),
    ("DolPHIN-2", "NCT03249181", "113", [("DTG+2 NRTIs", 120, S), ("EFV+2 NRTIs", 117, S)]),
    ("INSPIRING", "NCT02178592", "114", [("DTG+2 NRTIs", 69, S), ("EFV+2 NRTIs", 44, S)]),
    ("VESTED", "NCT03048422", "115", [("DTG+FTC/TAF", NR, NR), ("DTG+TDF/FTC", NR, NR), ("EFV/FTC/TDF", 200, S)]),
    ("Study 145", "NCT00708162", "116", [("EVG+PI/r+agent (TE)", 351, S), ("RAL+background (TE)", 351, S)]),
    ("BENCHMRK-1", "NCT00293267", "117", [("RAL+background (TE)", 231, S), (PBO, 118, S)]),
    ("BENCHMRK-2", "NCT00293254", "118", [("RAL+background (TE)", 228, S), (PBO, 119, S)]),
    ("STARTMRK", "NCT00369941", "119", [("RAL+TDF/FTC", 241, S), ("EFV+TDF/FTC", 230, S)]),
    ("SPIRAL", "NCT00528892", "120", [("RAL+2 NRTIs", 139, S), (SOC, 134, S)]),
    ("SWITCHMRK-1", "NCT00443703", "121", [("RAL+2 NRTIs", 172, S), ("LPV/r+2 NRTIs", 174, S)]),
    ("SWITCHMRK-2", "NCT00443729", "122", [("RAL+2 NRTIs", 175, S), ("LPV/r+2 NRTIs", 178, S)]),
    ("EASIER ANRS 138", "NCT00454337", "123", [("RAL+background (TE)", NR, NR), (SOC, NR, NR)]),
    ("PROGRESS", "NCT00711009", "124", [("LPV/r+RAL", 101, S), ("LPV/r+TDF/FTC", 105, S)]),
    ("QDMRK", "NCT00745823", "125", [("RAL QD 800 mg+TDF/FTC", 382, S), ("RAL+TDF/FTC", 386, S)]),
    ("ACTG A5257", "NCT00811954", "127,137,142", [("RAL+TDF/FTC", 603, S), ("ATV/r+TDF/FTC", 605, S), ("DRV/r+TDF/FTC", 601, S)]),
    ("EARNEST", "NCT00988039", "128", [("LPV/r+RAL", 433, S), ("LPV/r+2 NRTIs", 426, S), ("LPV/r monotherapy", 418, S)]),
    ("NEAT001/ANRS143", "NCT01066962", "129", [("DRV/r+RAL", 401, S), ("DRV/r+TDF/FTC", 404, S)]),
    ("ONCEMRK", "NCT02131233", "130", [("RAL HD 1200 mg QD+TDF/FTC", 531, S), ("RAL+TDF/FTC", 266, S)]),
    ("Reflate TB 2", "NCT02273765", "131", [("RAL+TDF/3TC", 230, S), ("EFV+TDF/3TC", 227, S)]),
    ("AI424-034", "NCT00013897", "132", [("ATV+ZDV/3TC", 405, S), ("EFV+ZDV/3TC", 405, S)]),
    ("AI424-043", "NCT00028301", "133", [("ATV+2 NRTIs (TE)", 144, S), ("LPV/r+2 NRTIs", 146, S)]),
    ("AI424-045", "NCT00035932", "134", [("ATV/r+TDF+NRTI (TE)", NR, NR), ("ATV+SQV+TDF+NRTI (TE)", NR, NR), ("LPV/r+TDF+NRTI (TE)", NR, NR)]),
    ("CASTLE", "NCT00272779", "135", [("ATV/r+TDF/FTC", 440, S), ("LPV/r+TDF/FTC", 443, S)]),
    ("TITAN", "NCT00110877", "138", [("DRV/r BID+OBR (TE)", 298, S), ("LPV/r+OBR (TE)", 297, D)]),
    ("ARTEMIS", "NCT00258557", "139", [("DRV/r+TDF/FTC", 343, S), ("LPV/r+TDF/FTC", 346, D)]),
    ("ODIN", "NCT00524368", "141", [("DRV/r+2 NRTIs", 294, S), ("DRV/r BID+OBR (TE)", 296, S)]),
    ("2LADY", "NCT00928187", "144", [("DRV/r+TDF/FTC", 154, S), ("LPV/r+TDF/FTC", 152, S), ("LPV/r+ABC+ddI", 145, S)]),
    ("WRHI 052", "NCT02671383", "145", [("DRV/r 400/100+2 NRTIs", 148, S), ("LPV/r+2 NRTIs", 152, S)]),
    ("MANET", "NCT02155101", "146", [("DRV/r monotherapy", 81, S), (SOC, 39, S)]),
    ("Study 247 (RTV)", "-", "148", [("RTV 600 mg BID+NRTIs", 543, S), (PBO, 547, S)]),
    ("DRIVE-FORWARD", "NCT02275780", "149", [("DOR+2 NRTIs", 383, S), ("DRV/r+2 NRTIs", 383, S)]),
    ("Study 006 (DMP 266-006)", "-", "152", [("EFV+ZDV/3TC", NR, NR), ("EFV+IDV", NR, NR), ("IDV+ZDV/3TC", NR, NR)]),
    ("ENCORE1", "NCT01011413", "153", [("EFV 400 mg+TDF/FTC", 321, S), ("EFV+TDF/FTC", 309, S)]),
    ("DUET-1", "NCT00254046", "154", [("ETR+background (TE)", 304, S), (PBO, 308, S)]),
    ("DUET-2", "NCT00255099", "155", [("ETR+background (TE)", 295, D), (PBO, 296, D)]),
    ("ECHO", "NCT00540449", "156", [("RPV+TDF/FTC", 346, S), ("EFV+TDF/FTC", 344, S)]),
    ("THRIVE", "NCT00543725", "157", [("RPV+2 NRTIs", 340, S), ("EFV+2 NRTIs", 338, S)]),
    ("CAPELLA (cohort 1)", "NCT04150068", "158", [("LEN+background (TE)", 24, S), (PBO, 12, S)]),
    ("ISLEND-1", "NCT06630286", "161", [("ISL/LEN weekly", 304, S), ("B/F/TAF", 303, S)]),
    ("ISLEND-2", "NCT06630299", "162", [("ISL/LEN weekly", 314, S), (SOC, 312, S)]),
    ("BRIGHTE (randomized cohort)", "NCT02362503", "163", [("FTR+background (TE)", NR, NR), (PBO, NR, NR)]),
    ("MOTIVATE-1", "NCT00098306", "166", [("MVC QD+OBT", 232, S), ("MVC BID+OBT", 235, S), (PBO, 118, S)]),
    ("MOTIVATE-2", "NCT00098722", "167", [("MVC QD+OBT", 182, S), ("MVC BID+OBT", 191, S), (PBO, 91, S)]),
    ("MERIT", "NCT00098293", "168", [("MVC BID+ZDV/3TC", 360, S), ("MVC QD+ZDV/3TC", 174, S), ("EFV+ZDV/3TC", 361, S)]),
    ("MODERN", "NCT01345630", "169", [("MVC+DRV/r", 396, S), ("DRV/r+TDF/FTC", 401, S)]),
]

EXCLUDED = [
    ("STAT, EYEWITNESS, GS-US-380-4449, GS-US-292-0112/1249/1825, NCT02616029, GS-US-236-0123/0118, PHI04, DIAMOND, "
     "VIKING-3, REALMRK, GRACE, GS-US-216-0130, NCT00460746, TMB-301, TMB-311, NCT01896921", "single-arm"),
    ("VOLITION", "non-randomized participant choice"),
    ("WAVES open-label extension switch", "re-randomization of WAVES participants (same NCT; avoid double counting)"),
    ("EFV/TDF/3TC, TDF/3TC, ISL, TAF, BIC rows", "no Phase 3 trial assigned (placeholders)"),
]

CLASS_COLORS = {
    "INSTI 3DR": "#1f77b4", "INSTI 2DR (oral)": "#17becf", "LA INSTI+NNRTI": "#9467bd",
    "Capsid inhibitor-containing": "#e377c2", "NNRTI+NRTTI 2DR": "#bcbd22", "NNRTI 3DR": "#2ca02c",
    "Boosted PI 3DR": "#d62728", "Unboosted / early PI": "#ff9896", "NRTI-sparing / PI monotherapy": "#ff7f0e",
    "CCR5 antagonist": "#8c564b", "Attachment inhibitor": "#c49c94", "NRTI-only / NRTI add-on": "#7f7f7f",
    "Backbone-only comparison": "#c7c7c7", SOC: "#aec7e8", PBO: "#dbdb8d",
}
CLASS_ORDER = list(CLASS_COLORS)


def check_against_workbook(xlsx):
    wb = openpyxl.load_workbook(xlsx, read_only=True)
    names = {i: r[6] for i, r in enumerate(wb["Phase 3 Master"].iter_rows(values_only=True), 1)}
    for t, _, rows, arms in TRIALS:
        for r in rows.split(","):
            assert int(r) in names, (t, r)
        for reg, _, _ in arms:
            assert reg in REG, reg


def build(level):
    node_n, node_trials, node_cls, node_nr = defaultdict(int), defaultdict(set), {}, defaultdict(set)
    edges, within = defaultdict(set), defaultdict(set)
    for t, _, _, arms in TRIALS:
        mapped = []
        for reg, n, _ in arms:
            a, c = REG[reg]
            node = {"regimen": reg, "anchor": a, "class": c}[level]
            node_cls[node] = c
            mapped.append(node)
            node_trials[node].add(t)
            if isinstance(n, int):
                node_n[node] += n
            else:
                node_nr[node].add(t)
        for u, v in itertools.combinations(sorted(set(mapped)), 2):
            edges[(u, v)].add(t)
        for node in {m for m in mapped if mapped.count(m) > 1}:
            within[node].add(t)
    connected = {u for e in edges for u in e}
    return node_n, node_trials, node_cls, node_nr, edges, within, connected


def draw(level, title, fname, label_min=1, figsize=(16, 16), font=10, smax=4200):
    res = build(level)
    node_n, node_trials, node_cls, node_nr, edges, within, connected = res
    nodes = sorted(connected, key=lambda x: (CLASS_ORDER.index(node_cls[x]), x))
    ring = [n for n in nodes if n != SOC]
    k = len(ring)
    pos = {nd: (math.cos(math.pi / 2 - 2 * math.pi * (i + 0.5) / k), math.sin(math.pi / 2 - 2 * math.pi * (i + 0.5) / k))
           for i, nd in enumerate(ring)}
    pos[SOC] = (0.0, 0.0)  # hub: stay-on-baseline comparator drawn at centre
    fig, ax = plt.subplots(figsize=figsize)
    maxw = max(len(v) for v in edges.values())
    for (u, v), ts in sorted(edges.items(), key=lambda kv: len(kv[1])):
        (x1, y1), (x2, y2) = pos[u], pos[v]
        w = 0.6 + 7.0 * (len(ts) - 1) / max(maxw - 1, 1)
        ax.plot([x1, x2], [y1, y2], color="black", lw=w, alpha=0.7, zorder=1, solid_capstyle="round")
        if len(ts) >= label_min:
            ax.text((x1 + x2) / 2, (y1 + y2) / 2, str(len(ts)), fontsize=font, ha="center", va="center", zorder=3,
                    bbox=dict(boxstyle="round,pad=0.15", fc="white", ec="none", alpha=0.85))
    maxn = max(node_n[n] for n in nodes) or 1
    for nd in nodes:
        x, y = pos[nd]
        ax.scatter([x], [y], s=60 + smax * node_n[nd] / maxn, color=CLASS_COLORS[node_cls[nd]],
                   edgecolor="black", lw=0.8, zorder=2)
        lbl = f"{nd} (n={node_n[nd]:,}{'*' if node_nr[nd] else ''})"
        if nd == SOC:
            ax.text(0, -0.07 - 0.05 * math.sqrt(node_n[nd] / maxn), lbl, fontsize=font, ha="center", va="top",
                    fontweight="bold", zorder=4, bbox=dict(boxstyle="round,pad=0.2", fc="white", ec="none", alpha=0.85))
            continue
        ang = math.atan2(y, x)
        r = 1.08 + 0.07 * math.sqrt(node_n[nd] / maxn)
        lx, ly = r * math.cos(ang), r * math.sin(ang)
        if level == "regimen":
            rot = math.degrees(ang)
            ha = "left"
            if rot > 90 or rot < -90:
                rot += 180
                ha = "right"
            ax.text(lx, ly, lbl, fontsize=font, ha=ha, va="center", rotation=rot, rotation_mode="anchor")
        else:
            ha = "center" if abs(lx) < 0.02 else ("left" if lx > 0 else "right")
            va = "center" if abs(lx) >= 0.02 else ("bottom" if ly > 0 else "top")
            ax.text(lx, ly, lbl, fontsize=font, ha=ha, va=va)
    lim = 1.75 if level == "regimen" else 1.45
    ax.set_xlim(-lim, lim); ax.set_ylim(-lim, lim); ax.set_aspect("equal"); ax.axis("off")
    used = sorted({node_cls[n] for n in nodes}, key=CLASS_ORDER.index)
    handles = [Line2D([0], [0], marker="o", color="w", markerfacecolor=CLASS_COLORS[c], markeredgecolor="black",
                      markersize=11, label=c) for c in used]
    ax.legend(handles=handles, title="Regimen class", loc="upper left", bbox_to_anchor=(1.0, 1.0),
              fontsize=9, title_fontsize=10, frameon=False)
    n_trials = len({t for v in edges.values() for t in v})
    ax.set_title(title, fontsize=15, fontweight="bold", pad=10)
    iso = sorted(set(node_trials) - connected)
    wtxt = "; ".join(f"{k} {len(v)}" for k, v in sorted(within.items(), key=lambda kv: -len(kv[1])) if k in connected)
    cap = (f"Node area proportional to total participants in included randomized trials (n, arm Ns as stated in the master "
           f"workbook). Edge width and label = number of randomized Phase 3 trials directly comparing the two nodes "
           f"({n_trials} of {len(TRIALS)} trials contribute edges; multi-arm trials give one edge per arm pair"
           f"{'; edges with 1 trial unlabelled' if label_min > 1 else ''}). "
           f"* node includes arm(s) with N not stated per arm (counted as 0).")
    if level != "regimen" and wtxt:
        cap += f" Within-node comparisons, not drawn (trials): {wtxt}."
    if iso:
        cap += f" Not shown (only within-node comparisons): {', '.join(iso)}."
    ax.text(-lim, -lim, "\n".join(textwrap.wrap(cap, int(figsize[0] * 11))), fontsize=9, ha="left", va="top")
    fig.subplots_adjust(left=0.02, right=0.98, top=0.95, bottom=0.02)
    fig.savefig(os.path.join(OUT, fname), dpi=170, bbox_inches="tight", pad_inches=0.3)
    plt.close(fig)
    return res


def main():
    import sys
    # Pass the final master workbook (HIV_Phase3_Trial_Review_from_GitHub_AB_110526.xlsx) as argv[1].
    xlsx = sys.argv[1] if len(sys.argv) > 1 else (
        "/workspaces/scientific-workflow-workshop/.copilot/uploads/2026-11-16-bc1c29a5/"
        "HIV_Phase3_Trial_Review_from_GitHub_AB_110526.xlsx")
    check_against_workbook(xlsx)
    results = {
        "regimen": draw("regimen", "Network of randomized Phase 3 HIV-1 trials: full regimen",
                        "network_1_regimen.png", label_min=2, figsize=(22, 22), font=8, smax=2600),
        "anchor": draw("anchor", "Network of randomized Phase 3 HIV-1 trials: anchor (third) agent or 2-drug regimen",
                       "network_2_anchor.png", figsize=(18, 18), font=10, smax=2600),
        "class": draw("class", "Network of randomized Phase 3 HIV-1 trials: regimen class",
                      "network_3_class.png", figsize=(14, 14), font=11, smax=4500),
    }
    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = "Read Me"
    for line in [
        "Network plot data (Figure-2 style) built from HIV_Phase3_Trial_Review_from_GitHub_AB_110526.xlsx, sheet 'Phase 3 Master'.",
        "Included: randomized trials with >=2 randomized arms. Trials recorded under several product rows counted once (by NCT/trial).",
        "Arm N = as stated in workbook (randomized/treated/analysed, whichever the cell gives). NR = not stated per arm (0 in node size). derived = stated total minus stated arm N.",
        "Edges: one per pair of distinct nodes per trial; weight = number of trials. Same-node arms (dose/backbone comparisons) are 'within-node' and not drawn.",
        "Anchor = third agent of 3-drug regimens; 2-drug regimens and NRTI-sparing combos are their own anchor. Mixed stay-on-baseline arms = 'Continued baseline ART'.",
        "Requires human scientific review: regimen/anchor/class mapping, N choices, inclusion of strategy and non-registrational trials.",
    ] + [f"Excluded: {a} - {b}" for a, b in EXCLUDED]:
        ws.append([line])
    ws = wb.create_sheet("Arms")
    ws.append(["Trial", "NCT/registry", "Workbook row(s)", "Regimen", "Anchor", "Class", "Arm N", "N source"])
    for t, nct, rows, arms in TRIALS:
        for reg, n, src in arms:
            ws.append([t, nct, rows, reg, REG[reg][0], REG[reg][1], n if isinstance(n, int) else None, src])
    for lvl, (node_n, node_trials, node_cls, node_nr, edges, within, connected) in results.items():
        ws = wb.create_sheet(f"Nodes_{lvl}")
        ws.append(["Node", "Class", "Total N (stated)", "Trials", "Trials with N not stated", "Shown in plot",
                   "Within-node trials"])
        for nd in sorted(node_trials, key=lambda x: (CLASS_ORDER.index(node_cls[x]), x)):
            ws.append([nd, node_cls[nd], node_n[nd], len(node_trials[nd]), "; ".join(sorted(node_nr[nd])),
                       "Yes" if nd in connected else "No", "; ".join(sorted(within.get(nd, [])))])
        ws = wb.create_sheet(f"Edges_{lvl}")
        ws.append(["Node A", "Node B", "No. trials", "Trials"])
        for (u, v), ts in sorted(edges.items(), key=lambda kv: (-len(kv[1]), kv[0])):
            ws.append([u, v, len(ts), "; ".join(sorted(ts))])
    for s in wb.worksheets:
        if s.title == "Read Me":
            s.column_dimensions["A"].width = 150
            continue
        for col in s.columns:
            s.column_dimensions[col[0].column_letter].width = min(60, max(10, max(len(str(c.value or "")) for c in col) + 2))
    wb.save(os.path.join(OUT, "HIV_Phase3_network_data.xlsx"))
    for lvl, r in results.items():
        print(lvl, "nodes shown", len(r[6]), "of", len(r[1]), "edges", len(r[4]))


if __name__ == "__main__":
    main()
