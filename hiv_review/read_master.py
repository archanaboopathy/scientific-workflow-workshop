"""Read the first sheet of the master workbook with the standard library only."""
import json, re, zipfile, xml.etree.ElementTree as ET
from collections import OrderedDict

SRC = '../NIH_ARV_Phase3_Master_Week48_Scatter_X_Offset.xlsx'
M = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS = {'m': M}

def col_idx(ref):
    n = 0
    for ch in re.match(r'[A-Z]+', ref).group():
        n = n * 26 + ord(ch) - 64
    return n - 1

z = zipfile.ZipFile(SRC)
shared = []
if 'xl/sharedStrings.xml' in z.namelist():
    for si in ET.fromstring(z.read('xl/sharedStrings.xml')).findall('m:si', NS):
        shared.append(''.join(t.text or '' for t in si.iter('{%s}t' % M)))
first = ET.fromstring(z.read('xl/workbook.xml')).find('m:sheets/m:sheet', NS)
rid = first.get('{%s}id' % R)
rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
target = next(r.get('Target') for r in rels if r.get('Id') == rid)
path = target.lstrip('/') if target.startswith('/') else 'xl/' + target
rows = {}
for row in ET.fromstring(z.read(path)).iter('{%s}row' % M):
    vals = {}
    for c in row.findall('m:c', NS):
        t, v = c.get('t'), c.find('m:v', NS)
        if t == 'inlineStr':
            val = ''.join(x.text or '' for x in c.iter('{%s}t' % M))
        elif v is None:
            val = None
        elif t == 's':
            val = shared[int(v.text)]
        elif t in ('str', 'b', 'e'):
            val = v.text
        else:
            f = float(v.text); val = int(f) if f.is_integer() else f
        vals[col_idx(c.get('r'))] = val
    rows[int(row.get('r'))] = vals
hdr_row = min(rows)
ncol = max(rows[hdr_row]) + 1
HDR = [rows[hdr_row].get(i) for i in range(ncol)]
master = []
for rn in sorted(rows):
    if rn == hdr_row or not any(v not in (None, '') for v in rows[rn].values()):
        continue
    d = OrderedDict((HDR[i], rows[rn].get(i)) for i in range(ncol) if HDR[i])
    d['_row'] = rn
    master.append(d)
pc = HDR[2]
products = list(OrderedDict.fromkeys(m[pc] for m in master if m[pc]))
json.dump(master, open('master_rows.json', 'w'), indent=1)
json.dump([[p, [m['_row'] for m in master if m[pc] == p]] for p in products], open('products.json', 'w'), indent=1)
print(first.get('name'), '| header:', HDR)
print(len(master), 'rows;', len(products), 'products')
for m in master:
    print(m['_row'], m[pc], '|', m.get(HDR[1]), '|', m.get(HDR[6]))
