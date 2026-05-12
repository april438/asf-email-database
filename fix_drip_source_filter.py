import sys, re, json
sys.stdout.reconfigure(encoding='utf-8')
import openpyxl
from collections import defaultdict

APRISM  = r"D:\Documents\April the Legendary\ASF\ASF 2026\Visual Builder Projects\DataViz\asf-email-database\aprism.html"
SF_BASE = r"D:\Documents\April the Legendary\ASF\ASF 2026\Visual Builder Projects\Email report\all emails"
SF_FILES = {
    2021: SF_BASE + r"\2021 sf email report.xlsx",
    2022: SF_BASE + r"\2022 sf email report.xlsx",
    2023: SF_BASE + r"\2023 sf email report.xlsx",
    2024: SF_BASE + r"\2024 sf email report.xlsx",
    2025: SF_BASE + r"\2025 sf email report.xlsx",
}

MONTHS_STR = ['','January','February','March','April','May','June',
              'July','August','September','October','November','December']
SF_ONLY_MONTHS_2021 = {1, 3, 5}   # No Drip broadcast data for these months
FIX_YEARS = {2021, 2022, 2023, 2024, 2025}

def safe_int(val):
    try: return int(round(float(val)))
    except: return 0

def parse_term_date(term):
    parts = str(term).split('_')
    try: return (int(parts[0]), int(parts[1]), int(parts[2]))
    except: return None

def keywords(text):
    stop = {'the','a','an','is','in','on','at','to','of','for','and','or','not',
            'it','be','as','are','was','were','with','from','your','our','you',
            'we','i','this','that','have','has','had','do','will','can','get',
            'all','my','by','up','so','but','if','no','go'}
    words = re.sub(r'[^a-z0-9\s]', ' ', text.lower()).split()
    return {w for w in words if len(w) > 2 and w not in stop}

def slug_keywords(term):
    parts = term.split('_')
    slug = '_'.join(parts[3:]) if len(parts) > 3 else term
    return keywords(slug)

def score(kw_a, kw_b):
    return len(kw_a & kw_b)

def parse_email_date(date_str):
    if not date_str: return None
    s = str(date_str).strip()
    m = re.match(r'^(\d{1,2})/(\d{1,2})/(\d{4})$', s)
    if m: return (int(m.group(3)), int(m.group(1)), int(m.group(2)))
    m = re.match(r'^(\d{4})-(\d{2})-(\d{2})$', s)
    if m: return (int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return None

# ── Load SF terms — drip source ONLY ─────────────────────────────────────────
print("Loading SF terms (drip source only)...")
# Index: (year, month, day) -> list of term dicts
sf_index = defaultdict(list)

for yr, path in SF_FILES.items():
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    cur_src = ''
    loaded = 0
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i < 15: continue
        src  = str(row[1]).strip() if row[1] is not None else ''
        term = str(row[2]).strip() if row[2] is not None else ''

        # Track which source group we're in
        if src and src.lower() != 'subtotal':
            cur_src = src

        # Only use drip-source rows
        if cur_src != 'drip': continue
        if not term or term.lower() in ('subtotal', ''): continue

        dt = parse_term_date(term)
        if not dt or dt[0] != yr: continue

        ok   = safe_int(row[5]) if len(row) > 5 else 0
        sets = safe_int(row[8]) if len(row) > 8 else 0
        cls  = safe_int(row[9]) if len(row) > 9 else 0
        sf_index[(dt[0], dt[1], dt[2])].append({
            'term': term,
            'slug_kw': slug_keywords(term),
            'ok': ok, 'sets': sets, 'cls': cls
        })
        loaded += 1
    wb.close()
    print(f"  {yr}: {loaded} drip-source terms loaded")

# ── Load seed ─────────────────────────────────────────────────────────────────
print("\nLoading seed...")
with open(APRISM, 'r', encoding='utf-8') as f:
    html = f.read()

pat = re.compile(r'(window\.ASF_SEED\s*=\s*)(\{.*?\})(;\s*</script>)', re.DOTALL)
m = pat.search(html)
assert m, "window.ASF_SEED not found"
seed = json.loads(m.group(2))

# ── Remove Jan/Mar/May 2021 SF-only entries (will be rebuilt below) ───────────
before = len(seed['emails'])
seed['emails'] = [e for e in seed['emails']
                  if not (e.get('year') == 2021 and e.get('monthNum') in SF_ONLY_MONTHS_2021)]
seed['months']  = [mo for mo in seed['months']
                   if mo.get('monthKey') not in ('2021-01','2021-03','2021-05')]
removed = before - len(seed['emails'])
print(f"Removed {removed} old Jan/Mar/May 2021 entries")

# ── Reset attribution for all affected years ──────────────────────────────────
reset = 0
for e in seed['emails']:
    if e.get('year') in FIX_YEARS:
        e['okupdrs'] = None
        e['sets'] = None
        e['closes'] = None
        reset += 1
print(f"Reset attribution for {reset} emails (years {min(FIX_YEARS)}–{max(FIX_YEARS)})")

# ── Re-match attribution using drip-only terms ────────────────────────────────
matched = 0
for e in seed['emails']:
    if e.get('year') not in FIX_YEARS: continue
    dt = parse_email_date(e.get('date'))
    if not dt: continue
    candidates = sf_index.get((dt[0], dt[1], dt[2]), [])
    if not candidates: continue
    subj_kw = keywords(e.get('subject', ''))
    best = max(candidates, key=lambda t: score(t['slug_kw'], subj_kw))
    if score(best['slug_kw'], subj_kw) == 0 and len(candidates) > 1:
        continue  # multiple candidates but no overlap — skip to avoid wrong match
    e['okupdrs'] = best['ok']   if best['ok']   else None
    e['sets']    = best['sets'] if best['sets']  else None
    e['closes']  = best['cls']  if best['cls']   else None
    matched += 1
print(f"Re-matched {matched} emails with drip-only attribution")

# ── Rebuild Jan/Mar/May 2021 from drip-only SF terms ─────────────────────────
print("\nRebuilding Jan/Mar/May 2021 (drip-only)...")
for mo in sorted(SF_ONLY_MONTHS_2021):
    month_name = MONTHS_STR[mo]
    month_key  = f'2021-{mo:02d}'

    # Collect all drip terms for this year/month
    month_terms = []
    for (yr, mth, dy), terms in sf_index.items():
        if yr == 2021 and mth == mo:
            for t in terms:
                month_terms.append({**t, 'dy': dy})
    month_terms.sort(key=lambda t: t['dy'])

    new_emails = []
    for t in month_terms:
        new_emails.append({
            'month':    f'{month_name} 2021',
            'monthKey': month_key,
            'year':     2021,
            'monthNum': mo,
            'date':     f"{mo}/{t['dy']}/2021",
            'url':      None,
            'subject':  t['term'],
            'openRate': None,
            'clickRate':None,
            'okupdrs':  t['ok']   if t['ok']   else None,
            'sets':     t['sets'] if t['sets']  else None,
            'closes':   t['cls']  if t['cls']   else None,
        })
    seed['emails'].extend(new_emails)

    ok  = sum(t['ok']   for t in month_terms)
    s   = sum(t['sets'] for t in month_terms)
    c   = sum(t['cls']  for t in month_terms)
    seed['months'].append({
        'monthKey': month_key,
        'totalEmails': len(new_emails),
        'totalOkupdrs': ok,
        'totalSets': s,
        'totalCloses': c,
        'avgOpenPct': 0
    })
    print(f"  {month_name} 2021: {len(new_emails)} entries | OKUPDRs={ok} Sets={s} Closes={c}")

# ── Rebuild month meta for all affected months ────────────────────────────────
print("\nRebuilding month meta...")
by_mk = defaultdict(list)
for e in seed['emails']:
    mk = e.get('monthKey')
    if mk: by_mk[mk].append(e)

# Identify affected month keys (any 2021-2025 month)
affected_keys = {mk for mk in by_mk if mk[:4] in ('2021','2022','2023','2024','2025')}

# Remove and replace affected month meta
seed['months'] = [mo for mo in seed['months'] if mo.get('monthKey') not in affected_keys]

for mk in sorted(affected_keys):
    emails = by_mk[mk]
    ok  = sum(e.get('okupdrs') or 0 for e in emails)
    s   = sum(e.get('sets')    or 0 for e in emails)
    c   = sum(e.get('closes')  or 0 for e in emails)
    rates = []
    for e in emails:
        or_ = e.get('openRate')
        if or_:
            try: rates.append(float(str(or_).replace('%','')))
            except: pass
    avg_open = round(sum(rates)/len(rates), 1) if rates else 0
    seed['months'].append({
        'monthKey': mk,
        'totalEmails': len(emails),
        'totalOkupdrs': round(ok, 1),
        'totalSets': round(s, 1),
        'totalCloses': round(c, 1),
        'avgOpenPct': avg_open
    })

seed['months'].sort(key=lambda mo: mo.get('monthKey', ''))
print(f"Total seed emails: {len(seed['emails'])}")
print(f"Total seed months: {len(seed['months'])}")

# ── Bump DB_VER ───────────────────────────────────────────────────────────────
html = re.sub(r"const DB_NAME = 'aprismDB', DB_VER = \d+;",
              "const DB_NAME = 'aprismDB', DB_VER = 5;", html)
print("DB_VER bumped to 5")

# ── Write back ────────────────────────────────────────────────────────────────
new_seed = json.dumps(seed, separators=(',', ':'), ensure_ascii=False)
new_html  = html[:m.start()] + m.group(1) + new_seed + m.group(3) + html[m.end():]

with open(APRISM, 'w', encoding='utf-8') as f:
    f.write(new_html)
print("Saved.")
