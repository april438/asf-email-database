import sys, re, json
sys.stdout.reconfigure(encoding='utf-8')
import openpyxl
from datetime import datetime
from collections import defaultdict

APRISM = r"D:\Documents\April the Legendary\ASF\ASF 2026\Visual Builder Projects\DataViz\asf-email-database\aprism.html"
FILE   = r"D:\Documents\April the Legendary\ASF\ASF 2026\Visual Builder Projects\Email report\all emails\2017 - 2018 - 2019 BEST Months.xlsx"

MONTHS_STR = ['','January','February','March','April','May','June',
              'July','August','September','October','November','December']
TARGET_YEARS = {2017, 2018, 2019}

def fmt_date(dt):
    if isinstance(dt, datetime):
        return f'{dt.month}/{dt.day}/{dt.year}'
    s = str(dt).strip()
    m = re.match(r'^(\d{4})-(\d{2})-(\d{2})', s)
    if m: return f'{int(m.group(2))}/{int(m.group(3))}/{m.group(1)}'
    return s

def safe_attr(val):
    if val is None: return None
    try:
        f = float(val)
        return int(round(f)) if f else None
    except: return None

# ── Read emails from detail sheet ────────────────────────────────────────────
print("Reading Email_Report_Detail_by_Month...")
wb = openpyxl.load_workbook(FILE, read_only=True, data_only=True)
ws = wb['Email_Report_Detail_by_Month']

emails_by_month = {}   # monthKey → list of email dicts
cur_year = None
cur_month_num = None

for row in ws.iter_rows(values_only=True):
    col0 = str(row[0]).strip() if row[0] is not None else ''
    col1 = str(row[1]).strip() if row[1] is not None else ''

    # Year header: "YEAR 2017"
    if col0.startswith('YEAR '):
        try: cur_year = int(col0.split()[1])
        except: pass
        continue

    # Month header: "--- January 2018 ---"
    if col0.startswith('---'):
        inner = col0.replace('-','').strip()   # "January 2018"
        for i, mn in enumerate(MONTHS_STR):
            if mn and col0.find(mn) >= 0:
                cur_month_num = i
                break
        # Year may also be in the header
        ym = re.search(r'(\d{4})', col0)
        if ym: cur_year = int(ym.group(1))
        continue

    # Skip non-data rows
    if col0 in ('Date Sent', '') or 'TOTAL' in col1 or 'GRAND' in col1:
        continue

    # Data row: col[0] must be a datetime
    date_val = row[0]
    if not isinstance(date_val, datetime):
        continue

    term = col1
    if not term:
        continue

    date_str = fmt_date(date_val)
    month_key  = f'{cur_year}-{cur_month_num:02d}'
    month_name = f'{MONTHS_STR[cur_month_num]} {cur_year}'

    entry = {
        'month':    month_name,
        'monthKey': month_key,
        'year':     cur_year,
        'monthNum': cur_month_num,
        'date':     date_str,
        'url':      None,
        'subject':  term,
        'openRate': None,
        'clickRate':None,
        'okupdrs':  safe_attr(row[2]),
        'sets':     safe_attr(row[3]),
        'closes':   safe_attr(row[4]),
    }

    if month_key not in emails_by_month:
        emails_by_month[month_key] = []
    emails_by_month[month_key].append(entry)

wb.close()

# Report
total_emails = 0
for mk in sorted(emails_by_month):
    emails = emails_by_month[mk]
    ok  = sum(e.get('okupdrs') or 0 for e in emails)
    s   = sum(e.get('sets')    or 0 for e in emails)
    c   = sum(e.get('closes')  or 0 for e in emails)
    print(f"  {mk}: {len(emails)} emails | OKUPDRs={ok} Sets={s} Closes={c}")
    total_emails += len(emails)
print(f"Total: {total_emails} emails across {len(emails_by_month)} months")

# ── Load seed ────────────────────────────────────────────────────────────────
print("\nLoading seed...")
with open(APRISM, 'r', encoding='utf-8') as f:
    html = f.read()

pat = re.compile(r'(window\.ASF_SEED\s*=\s*)(\{.*?\})(;\s*</script>)', re.DOTALL)
m = pat.search(html)
assert m, "window.ASF_SEED not found"
seed = json.loads(m.group(2))

# Remove any existing 2017/2018/2019 data
before = len(seed['emails'])
seed['emails'] = [e for e in seed['emails'] if e.get('year') not in TARGET_YEARS]
seed['months']  = [mo for mo in seed['months']
                   if not mo.get('monthKey','').startswith(('2017-','2018-','2019-'))]
removed = before - len(seed['emails'])
if removed:
    print(f"Removed {removed} existing 2017–2019 entries")

# ── Add new emails ────────────────────────────────────────────────────────────
for mk in sorted(emails_by_month):
    emails = emails_by_month[mk]
    seed['emails'].extend(emails)

    ok  = sum(e.get('okupdrs') or 0 for e in emails)
    s   = sum(e.get('sets')    or 0 for e in emails)
    c   = sum(e.get('closes')  or 0 for e in emails)
    rates = [float(e['openRate'].replace('%','')) for e in emails if e.get('openRate')]
    avg_open = round(sum(rates)/len(rates), 1) if rates else 0

    seed['months'].append({
        'monthKey':      mk,
        'totalEmails':   len(emails),
        'totalOkupdrs':  round(ok, 1),
        'totalSets':     round(s, 1),
        'totalCloses':   round(c, 1),
        'avgOpenPct':    avg_open,
    })

seed['months'].sort(key=lambda mo: mo.get('monthKey', ''))
print(f"\nTotal seed emails:  {len(seed['emails'])}")
print(f"Total seed months:  {len(seed['months'])}")

# ── Bump DB_VER ───────────────────────────────────────────────────────────────
html = re.sub(r"const DB_NAME = 'aprismDB', DB_VER = \d+;",
              "const DB_NAME = 'aprismDB', DB_VER = 6;", html)
print("DB_VER bumped to 6")

# ── Write back ────────────────────────────────────────────────────────────────
new_seed = json.dumps(seed, separators=(',', ':'), ensure_ascii=False)
new_html  = html[:m.start()] + m.group(1) + new_seed + m.group(3) + html[m.end():]

with open(APRISM, 'w', encoding='utf-8') as f:
    f.write(new_html)
print("Saved.")
