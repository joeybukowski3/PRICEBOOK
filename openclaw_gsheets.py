"""
OPENCLAW Price Check — Google Sheets version
Writes directly to Google Sheets instead of Excel.
Tabs: OPENCLAW | PriceBook Refrigerators | PriceBook TVs
"""
import requests, re, time, os, json
from datetime import date
import gspread
from google.oauth2.service_account import Credentials

BESTBUY_API_KEY = 'twmOxg13zZdgLikzAfGATIXl'
SHEET_ID        = '1-i5CuSAXKFrvULPn_9bI5g5yCI3djy_AwC3c0apgkLY'
TODAY_LABEL     = f"{date.today().month}/{date.today().day}"
REPO_URL        = "https://github.com/joeybukowski3/PRICEBOOK"
SHEET_URL       = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}"

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
]

# ── AUTH ──────────────────────────────────────────────────────────────────────
def get_client():
    # In GitHub Actions, credentials come from a secret as JSON string
    creds_json = os.environ.get('GSHEET_CREDENTIALS')
    if creds_json:
        info = json.loads(creds_json)
    else:
        # Local fallback
        with open('gsheet_credentials.json') as f:
            info = json.load(f)
    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    return gspread.authorize(creds)

# ── BB API ────────────────────────────────────────────────────────────────────
def check_sku(sku):
    url    = f'https://api.bestbuy.com/v1/products/{sku}.json'
    params = {'apiKey': BESTBUY_API_KEY,
              'show':   'sku,name,salePrice,regularPrice,onlineAvailability'}
    try:
        r = requests.get(url, params=params, timeout=10)
        if r.status_code == 200:
            d = r.json()
            return (d.get('salePrice') or d.get('regularPrice'),
                    d.get('onlineAvailability', False), d.get('name',''))
        elif r.status_code == 404: return None, False, 'NOT FOUND'
        elif r.status_code == 403: return None, False, 'RATE LIMITED'
        else: return None, False, f'API {r.status_code}'
    except Exception as e:
        return None, False, str(e)[:50]

def clean_sku(raw):
    if raw is None: return None
    s = str(raw).replace('.0','').strip()
    if s.isdigit(): return s
    m = re.search(r'skuId=(\d+)|/(\d{7,8})\.p', s)
    return (m.group(1) or m.group(2)) if m else None

def bb_search_url(sku):
    return f"https://www.bestbuy.com/site/searchpage.jsp?st={sku}"

# ── EMAIL ─────────────────────────────────────────────────────────────────────
def send_email(checked, unchanged, n_listed, oos, not_found,
               vs_listed, vs_prev, today):
    resend_key   = os.environ.get('RESEND_API_KEY','')
    notify_email = os.environ.get('NOTIFY_EMAIL','')
    if not resend_key or not notify_email:
        print("No email credentials — skipping")
        return

    def table_rows(data):
        if not data:
            return '<tr><td colspan="6" style="padding:14px;text-align:center;color:#6B7280;font-size:12px">&#x2705; No changes</td></tr>'
        html = ""
        for itype, name, sku, pa, pb, pct in sorted(data, key=lambda x: abs(x[5]), reverse=True):
            color = "#DC2626" if pb > pa else "#15803D"
            bg    = "#FFF7ED" if pb > pa else "#DCFCE7"
            arrow = "&#8599;" if pb > pa else "&#8600;"
            sku_s = str(sku) if sku else "—"
            link  = bb_search_url(sku_s) if sku_s.isdigit() else "#"
            html += f"""<tr style="background:{bg}">
<td style="padding:6px 10px;font-size:12px;color:#374151">{itype}</td>
<td style="padding:6px 10px;font-size:12px;color:#374151">{(name or '')[:45]}</td>
<td style="padding:6px 10px;font-size:12px;text-align:center"><a href="{link}" style="color:#2563EB;font-size:11px">{sku_s}</a></td>
<td style="padding:6px 10px;font-size:12px;text-align:center">${pa:.2f}</td>
<td style="padding:6px 10px;font-size:12px;text-align:center;font-weight:bold;color:{color}">${pb:.2f}</td>
<td style="padding:6px 10px;font-size:12px;text-align:center;font-weight:bold;color:{color}">{arrow} {pct:+.1f}%</td>
</tr>"""
        return html

    def make_table(title, h1, h2, rows):
        return f"""<h3 style="margin:20px 0 6px;font-size:12px;font-weight:bold;color:#1F2937">{title}</h3>
<table style="width:100%;border-collapse:collapse;background:white;border:1px solid #E5E7EB;border-radius:6px;overflow:hidden">
<thead><tr style="background:#1F2937">
<th style="padding:8px 10px;text-align:left;color:white;font-size:11px">Type</th>
<th style="padding:8px 10px;text-align:left;color:white;font-size:11px">Item</th>
<th style="padding:8px 10px;text-align:center;color:white;font-size:11px">SKU</th>
<th style="padding:8px 10px;text-align:center;color:white;font-size:11px">{h1}</th>
<th style="padding:8px 10px;text-align:center;color:white;font-size:11px">{h2}</th>
<th style="padding:8px 10px;text-align:center;color:white;font-size:11px">Change</th>
</tr></thead>
<tbody>{rows}</tbody></table>"""

    t1 = make_table("CHANGES VS LISTED PRICE", "Listed", "Today", table_rows(vs_listed))
    t2 = make_table("CHANGES SINCE LAST RUN",  "Prev Run", "Today", table_rows(vs_prev))
    n_prev = len(vs_prev)

    html = f"""<div style="font-family:Arial,sans-serif;max-width:720px;margin:0 auto">
<div style="background:#0D1B2A;padding:20px;border-radius:8px 8px 0 0">
<h2 style="color:white;margin:0;font-size:18px">&#x1F4CA; OPENCLAW Daily Price Check</h2>
<p style="color:#9CA3AF;margin:4px 0 0;font-size:13px">{today}</p>
</div>
<div style="background:#F9FAFB;padding:20px;border:1px solid #E5E7EB;border-top:none;border-radius:0 0 8px 8px">
<div style="display:flex;gap:10px;margin-bottom:20px">
<div style="flex:1;background:white;border:1px solid #E5E7EB;border-radius:6px;padding:10px;text-align:center">
<div style="font-size:20px;font-weight:bold;color:#0D1B2A">{checked}</div>
<div style="font-size:11px;color:#6B7280">Checked</div></div>
<div style="flex:1;background:#DCFCE7;border:1px solid #BBF7D0;border-radius:6px;padding:10px;text-align:center">
<div style="font-size:20px;font-weight:bold;color:#15803D">{unchanged}</div>
<div style="font-size:11px;color:#15803D">Unchanged</div></div>
<div style="flex:1;background:{'#FFF7ED' if n_listed else '#F9FAFB'};border:1px solid {'#FED7AA' if n_listed else '#E5E7EB'};border-radius:6px;padding:10px;text-align:center">
<div style="font-size:20px;font-weight:bold;color:#{'EA580C' if n_listed else '6B7280'}">{n_listed}</div>
<div style="font-size:11px;color:#{'EA580C' if n_listed else '6B7280'}">vs Listed</div></div>
<div style="flex:1;background:{'#FFF7ED' if n_prev else '#F9FAFB'};border:1px solid {'#FED7AA' if n_prev else '#E5E7EB'};border-radius:6px;padding:10px;text-align:center">
<div style="font-size:20px;font-weight:bold;color:#{'EA580C' if n_prev else '6B7280'}">{n_prev}</div>
<div style="font-size:11px;color:#{'EA580C' if n_prev else '6B7280'}">New Changes</div></div>
<div style="flex:1;background:#FEE2E2;border:1px solid #FECACA;border-radius:6px;padding:10px;text-align:center">
<div style="font-size:20px;font-weight:bold;color:#DC2626">{not_found+oos}</div>
<div style="font-size:11px;color:#DC2626">OOS/404</div></div>
</div>
{t1}
{t2}
<p style="font-size:11px;color:#9CA3AF;margin-top:16px;text-align:center">
<a href="{SHEET_URL}" style="color:#2563EB;font-weight:bold;text-decoration:none">&#x1F4CA; Open Google Sheet</a>
&nbsp;&nbsp;|&nbsp;&nbsp;
<a href="{REPO_URL}" style="color:#2563EB;text-decoration:none">View Repo</a>
</p></div></div>"""

    payload = {"from": "PriceCheck <onboarding@resend.dev>", "to": [notify_email],
               "subject": f"OPENCLAW {today} — {n_listed} vs listed | {n_prev} new changes",
               "html": html}
    try:
        resp = requests.post("https://api.resend.com/emails",
            headers={"Authorization": f"Bearer {resend_key}",
                     "Content-Type": "application/json"},
            data=json.dumps(payload), timeout=15)
        print(f"Email {'sent' if resp.status_code in (200,201) else 'failed: '+str(resp.status_code)}")
    except Exception as e:
        print(f"Email error: {e}")

# ── GOOGLE SHEETS HELPERS ─────────────────────────────────────────────────────
def get_or_create_tab(spreadsheet, title):
    try:
        return spreadsheet.worksheet(title)
    except gspread.exceptions.WorksheetNotFound:
        return spreadsheet.add_worksheet(title=title, rows=200, cols=50)

def find_col(ws, header_val, header_row=1):
    """Find column index (1-based) of a header value."""
    row = ws.row_values(header_row)
    for i, v in enumerate(row, start=1):
        if str(v) == str(header_val):
            return i
    return None

def col_letter(n):
    """Convert 1-based column number to letter."""
    result = ""
    while n:
        n, r = divmod(n - 1, 26)
        result = chr(65 + r) + result
    return result

# ── MAIN ──────────────────────────────────────────────────────────────────────
print(f"\nConnecting to Google Sheets...")
gc          = get_client()
spreadsheet = gc.open_by_key(SHEET_ID)
print(f"Connected: {spreadsheet.title}")

# ── OPENCLAW TAB ──────────────────────────────────────────────────────────────
ws_oc = get_or_create_tab(spreadsheet, "OPENCLAW")
all_values = ws_oc.get_all_values()

if not all_values:
    print("OPENCLAW tab is empty — needs initial data load from GITHUBOPENCLAWPRICECHECK.xlsx")
    print("Run setup_gsheets.py first to migrate existing data.")
    exit(1)

headers   = all_values[0]
data_rows = all_values[1:]

# Find column positions
col_item_type = 3   # D = Item Type (0-indexed: 3)
col_name      = 4   # E = Name
col_sku       = 6   # G = SKU
col_listed    = 8   # I = Listed Price

# Find today's column and previous column
today_col_idx = None
prev_col_idx  = None
date_col_idxs = []

for i, h in enumerate(headers):
    s = str(h)
    if '/' in s and len(s) <= 5:
        date_col_idxs.append(i)
        if s == TODAY_LABEL:
            today_col_idx = i

if date_col_idxs:
    prev_col_idx = date_col_idxs[-2] if len(date_col_idxs) >= 2 else (date_col_idxs[-1] if today_col_idx is None else None)
    if today_col_idx is not None and len(date_col_idxs) >= 2:
        prev_col_idx = date_col_idxs[-2]

# Add today column if not exists
if today_col_idx is None:
    today_col_idx = len(headers)
    headers.append(TODAY_LABEL)
    # Write header
    ws_oc.update_cell(1, today_col_idx + 1, TODAY_LABEL)
    print(f"Added column {TODAY_LABEL} at position {today_col_idx + 1}")
else:
    print(f"Updating existing column {TODAY_LABEL}")

print(f"Prev column: {prev_col_idx} | Today column: {today_col_idx}")
print(f"\nOPENCLAW PRICE CHECK — {TODAY_LABEL}")
print(f"{'='*55}")

checked = unchanged = changed = oos = not_found = errors = 0
vs_listed = []
vs_prev   = []
updates   = []  # (row_num, col_num, value) 1-based

for i, row in enumerate(data_rows, start=2):  # start=2 for sheet row number
    # Pad row if needed
    while len(row) <= max(col_item_type, col_name, col_sku, col_listed):
        row.append('')

    row_num   = row[2] if len(row) > 2 else ''   # col C
    item_type = row[col_item_type] if len(row) > col_item_type else ''
    name      = row[col_name]      if len(row) > col_name      else ''
    sku_raw   = row[col_sku]       if len(row) > col_sku       else ''
    listed_v  = row[col_listed]    if len(row) > col_listed     else ''

    if not row_num: continue

    sku = clean_sku(sku_raw)
    if not sku:
        updates.append((i, today_col_idx + 1, 'No SKU'))
        errors += 1
        continue

    price, avail, bb_name = check_sku(sku)
    time.sleep(0.35)
    checked += 1

    if bb_name == 'RATE LIMITED':
        time.sleep(3)
        price, avail, bb_name = check_sku(sku)
        time.sleep(0.5)

    if price is None:
        val = 'OOS' if not avail else bb_name[:12]
        updates.append((i, today_col_idx + 1, val))
        not_found += 1
        print(f"  ❌ {sku} [{item_type}] — {bb_name}")
        continue

    if not avail:
        oos += 1

    # vs listed
    try:
        listed_f = float(str(listed_v).replace('$','').replace(',','')) if listed_v else None
    except: listed_f = None

    if listed_f and abs(price - listed_f) / listed_f > 0.01:
        pct = (price - listed_f) / listed_f * 100
        changed += 1
        vs_listed.append((item_type, name, sku, listed_f, price, pct))
        print(f"  {'📈' if price > listed_f else '📉'} {sku} listed:${listed_f:.2f} today:${price:.2f} ({pct:+.1f}%)")
    else:
        unchanged += 1

    # vs prev run
    if prev_col_idx is not None and prev_col_idx < len(row):
        try:
            prev_f = float(row[prev_col_idx]) if row[prev_col_idx] else None
        except: prev_f = None
        if prev_f and abs(price - prev_f) / prev_f > 0.005:
            pct2 = (price - prev_f) / prev_f * 100
            vs_prev.append((item_type, name, sku, prev_f, price, pct2))
            print(f"  🔄 {sku} prev:${prev_f:.2f} today:${price:.2f} ({pct2:+.1f}%)")

    updates.append((i, today_col_idx + 1, price))

# Write all updates in batches
print(f"\nWriting {len(updates)} price updates to Google Sheets...")
batch = []
for sheet_row, sheet_col, val in updates:
    cell = f"{col_letter(sheet_col)}{sheet_row}"
    batch.append({'range': f"OPENCLAW!{cell}", 'values': [[val]]})

if batch:
    # Write in chunks of 50 to avoid quota limits
    for i in range(0, len(batch), 50):
        chunk = batch[i:i+50]
        spreadsheet.values_batch_update({'valueInputOption': 'RAW', 'data': chunk})
        time.sleep(0.5)

print(f"\nDONE — checked:{checked} unchanged:{unchanged} vs_listed:{changed} new_changes:{len(vs_prev)} oos:{oos} not_found:{not_found}")
send_email(checked, unchanged, changed, oos, not_found, vs_listed, vs_prev, TODAY_LABEL)
print(f"Google Sheet updated: {SHEET_URL}")
