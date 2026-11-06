"""Literature helper (stdlib only). Responses cached in cache/. See PROTOCOL.md."""
import hashlib, html, json, os, re, sys, time, urllib.parse, urllib.request

CACHE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'cache')
os.makedirs(CACHE, exist_ok=True)
EU = 'https://eutils.ncbi.nlm.nih.gov/entrez/eutils/'
EPMC = 'https://www.ebi.ac.uk/europepmc/webservices/rest/'
CTG = 'https://clinicaltrials.gov/api/v2/studies'
DM = 'https://dailymed.nlm.nih.gov/dailymed/services/v2/'


def get(url, binary=False):
    key = os.path.join(CACHE, hashlib.sha1(url.encode()).hexdigest())
    if os.path.exists(key):
        data = open(key, 'rb').read()
    else:
        for attempt in range(3):
            try:
                req = urllib.request.Request(url, headers={'User-Agent': 'hiv-review/1.0'})
                data = urllib.request.urlopen(req, timeout=60).read()
                break
            except Exception as e:
                if attempt == 2:
                    raise
                time.sleep(2 + attempt * 3)
        open(key, 'wb').write(data)
        time.sleep(0.35)
    return data if binary else data.decode('utf-8', 'replace')


def strip(x):
    x = re.sub(r'<[^>]+>', ' ', x)
    return re.sub(r'\s+', ' ', html.unescape(x)).strip()


def grep(text, rx, width=300):
    out = []
    for m in re.finditer(rx, text, re.I):
        s, e = max(0, m.start() - width), min(len(text), m.end() + width)
        out.append('...' + text[s:e] + '...')
    return out


def abstract(*pmids):
    x = get(EU + 'efetch.fcgi?db=pubmed&retmode=xml&id=' + ','.join(pmids))
    for art in re.findall(r'<PubmedArticle>.*?</PubmedArticle>', x, re.S):
        g = lambda rx: (re.search(rx, art, re.S) or [None, ''])[1]
        authors = re.findall(r'<LastName>(.*?)</LastName>', art)
        print('PMID', g(r'<PMID[^>]*>(\d+)'), '|', strip(g(r'<ArticleTitle>(.*?)</ArticleTitle>')))
        print('  ', (authors[0] + ' et al.') if authors else '', strip(g(r'<ISOAbbreviation>(.*?)</ISOAbbreviation>')),
              strip(g(r'<PubDate>(.*?)</PubDate>')), 'Vol', g(r'<Volume>(.*?)</Volume>'), 'pp', g(r'<MedlinePgn>(.*?)</MedlinePgn>'))
        print('   PMCID', g(r'<ArticleId IdType="pmc">(.*?)</ArticleId>'), 'DOI', g(r'<ArticleId IdType="doi">(.*?)</ArticleId>'))
        print('  ', strip(' '.join(re.findall(r'<AbstractText[^>]*>(.*?)</AbstractText>', art, re.S))))
        print()


def search(q, n=40):
    j = json.loads(get(EU + 'esearch.fcgi?db=pubmed&retmode=json&retmax=%d&term=%s' % (int(n), urllib.parse.quote(q))))
    ids = j['esearchresult']['idlist']
    print('count', j['esearchresult']['count'])
    if not ids:
        return
    s = json.loads(get(EU + 'esummary.fcgi?db=pubmed&retmode=json&id=' + ','.join(ids)))['result']
    for i in ids:
        r = s.get(i, {})
        print(i, r.get('pubdate', '')[:4], (r.get('authors') or [{}])[0].get('name', ''), '|', r.get('title', ''), '|', r.get('source', ''))


def fulltext(pmid, rx=None):
    j = json.loads(get(EPMC + 'search?format=json&query=EXT_ID:%s%%20AND%%20SRC:MED&resultType=core' % pmid))
    res = j.get('resultList', {}).get('result', [])
    pmc = res[0].get('pmcid') if res else None
    if not pmc or res[0].get('isOpenAccess') != 'Y' and res[0].get('inEPMC') != 'Y':
        print('No open-access full text in Europe PMC for', pmid, '(pmcid=%s)' % pmc)
        return
    try:
        t = strip(get(EPMC + '%s/fullTextXML' % pmc))
    except Exception as e:
        print('Full text fetch failed', pmc, e); return
    print('Full text', pmc, len(t), 'chars')
    if rx:
        for h in grep(t, rx)[:25]:
            print(h, '\n')
    else:
        print(t[:6000])


def _ct(nct):
    return json.loads(get(CTG + '/' + nct))


def ctgov(nct):
    d = _ct(nct)['protocolSection']
    idm, st, des = d['identificationModule'], d['statusModule'], d.get('designModule', {})
    print(nct, '|', idm.get('briefTitle'), '|', idm.get('acronym', ''))
    print(' phase', des.get('phases'), '| status', st.get('overallStatus'), '| start', st.get('startDateStruct', {}).get('date'),
          '| enrollment', des.get('enrollmentInfo', {}).get('count'))
    for a in d.get('armsInterventionsModule', {}).get('armGroups', []):
        print(' ARM', a.get('label'), '|', a.get('type'), '|', a.get('interventionNames'))
    for o in d.get('outcomesModule', {}).get('primaryOutcomes', []):
        print(' PRIMARY', o.get('measure'), '|', o.get('timeFrame'))
    print(' elig', d.get('eligibilityModule', {}).get('minimumAge'), d.get('eligibilityModule', {}).get('maximumAge'))
    refs = d.get('referencesModule', {}).get('references', [])
    for r in refs:
        if r.get('type') == 'RESULT':
            print(' REF', r.get('pmid'), r.get('citation', '')[:200])


def ctresults(nct, rx=None):
    d = _ct(nct)
    rs = d.get('resultsSection')
    if not rs:
        print('No posted results for', nct); return
    for o in rs.get('outcomeMeasuresModule', {}).get('outcomeMeasures', []):
        title = o.get('title', '')
        if rx and not re.search(rx, title + ' ' + o.get('description', ''), re.I):
            continue
        grp = {g['id']: g.get('title') for g in o.get('groups', [])}
        print('==', o.get('type'), '|', title, '|', o.get('timeFrame'), '|', o.get('unitOfMeasure'))
        print('   desc:', o.get('description', '')[:500])
        for dn in o.get('denoms', []):
            print('   N:', {grp.get(c['groupId']): c['value'] for c in dn.get('counts', [])})
        for cl in o.get('classes', []):
            for cat in cl.get('categories', []):
                print('  ', cl.get('title', ''), cat.get('title', ''), {grp.get(m['groupId']): m.get('value') for m in cat.get('measurements', [])})


def ctsearch(drug):
    q = urllib.parse.urlencode({'query.cond': 'HIV', 'query.intr': drug, 'filter.advanced': 'AREA[Phase](PHASE3)',
                                'pageSize': 200, 'fields': 'NCTId,BriefTitle,Acronym,OverallStatus,HasResults,MinimumAge,StartDate'})
    j = json.loads(get(CTG + '?' + q))
    print('studies', len(j.get('studies', [])))
    for s in j.get('studies', []):
        p = s['protocolSection']; i = p['identificationModule']
        print(i['nctId'], p['statusModule'].get('overallStatus'), 'results' if s.get('hasResults') else '-',
              p['statusModule'].get('startDateStruct', {}).get('date'), p.get('eligibilityModule', {}).get('minimumAge'),
              '|', i.get('acronym', ''), '|', i.get('briefTitle'))


def label(brand, rx=None):
    j = json.loads(get(DM + 'spls.json?drug_name=%s&pagesize=20' % urllib.parse.quote(brand)))
    data = j.get('data', [])
    if not data:
        print('No SPL for', brand); return
    for d in data[:5]:
        print('SPL', d['setid'], d.get('title', '')[:150])
    sid = data[0]['setid']
    t = strip(get('https://dailymed.nlm.nih.gov/dailymed/services/v2/spls/%s.xml' % sid))
    m = re.search(r'Initial U\.?S\.? Approval:?\s*\d{4}', t)
    print('Using', sid, '|', m.group() if m else 'Initial U.S. Approval not found')
    for h in grep(t, rx or r'14\.\d? ?[A-Z]|Resistance', 400)[:30]:
        print(h, '\n')


if __name__ == '__main__':
    cmd, args = sys.argv[1], sys.argv[2:]
    {'abstract': abstract, 'search': search, 'fulltext': fulltext, 'ctgov': ctgov,
     'ctresults': ctresults, 'ctsearch': ctsearch, 'label': label}[cmd](*args)
