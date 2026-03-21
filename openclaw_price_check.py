"""
OPENCLAW Price Check
Checks all 53 SKUs via BB API and fills in today's price column.
Run from your OPENCLAWPRICES folder

Usage: python openclaw_price_check.py
"""
import requests, re, time, os, sys
import openpyxl
sys.stdout.reconfigure(encoding='utf-8')
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import date

BESTBUY_API_KEY = 'twmOxg13zZdgLikzAfGATIXl'
EXCEL_FILE      = 'GITHUBOPENCLAWPRICECHECK.xlsx'
TODAY_LABEL     = f"{date.today().month}/{date.today().day}"

thin = Side(style='thin', color='D1D5DB')
T    = Border(left=thin, right=thin, top=thin, bottom=thin)
TEAL_L  = "E0FAF5"; GREEN_BG = "DCFCE7"; GREEN_FG = "15803D"
AMBER_L = "FFFBEB"; AMBER    = "F59E0B"
LGRAY   = "F3F4F6"; DGRAY    = "374151"
NAVY    = "0D1B2A"; WHITE    = "FFFFFF"
RED_BG  = "FEE2E2"; RED_FG   = "DC2626"

def check_sku(sku):
    """Look up current price and availability by SKU."""
    url    = f'https://api.bestbuy.com/v1/products/{sku}.json'
    params = {'apiKey': BESTBUY_API_KEY,
              'show':   'sku,name,salePrice,regularPrice,onlineAvailability'}
    try:
        r = requests.get(url, params=params, timeout=10)
        if r.status_code == 200:
            d     = r.json()
            price = d.get('salePrice') or d.get('regularPrice')
            avail = d.get('onlineAvailability', False)
            name  = d.get('name', '')
            return price, avail, name
        elif r.status_code == 404:
            return None, False, 'NOT FOUND'
        elif r.status_code == 403:
            return None, False, 'RATE LIMITED'
        else:
            return None, False, f'API ERROR {r.status_code}'
    except Exception as e:
        return None, False, str(e)[:60]

def clean_sku(raw):
    """Extract numeric SKU from cell value."""
    if raw is None: return None
    s = str(raw).replace('.0','').strip()
    if s.isdigit(): return s
    # Try extracting from URL
    m = re.search(r'skuId=(\d+)|/(\d{7,8})\.p', s)
    if m: return m.group(1) or m.group(2)
    return None

# ── LOAD WORKBOOK ─────────────────────────────────────────────────────────────
wb = openpyxl.load_workbook(EXCEL_FILE)
ws = wb.active

# Find last real data row
last_row = 1
for row in range(ws.max_row, 1, -1):
    if ws.cell(row=row, column=3).value:
        last_row = row
        break

# Find next empty column for today's prices
# Header is in row 1, check for existing today column
header_row = 1
next_col   = None
today_col  = None

for col in range(1, ws.max_column + 2):
    val = ws.cell(row=header_row, column=col).value
    if val is None and next_col is None:
        next_col = col
    if str(val) == TODAY_LABEL:
        today_col = col
        break

if today_col:
    write_col = today_col
    print(f"Updating existing column for {TODAY_LABEL} (col {write_col})")
else:
    write_col = next_col
    # Write header
    hdr = ws.cell(row=header_row, column=write_col, value=TODAY_LABEL)
    hdr.font      = Font(color=WHITE, bold=True, size=9, name='Arial')
    hdr.fill      = PatternFill('solid', start_color=NAVY)
    hdr.alignment = Alignment(horizontal='center', vertical='center')
    hdr.border    = T
    ws.column_dimensions[hdr.column_letter].width = 12
    print(f"Adding new column {TODAY_LABEL} (col {write_col})")

# ── CHECK ALL SKUs ─────────────────────────────────────────────────────────────
print(f"\n{'='*55}")
print(f"OPENCLAW PRICE CHECK — {TODAY_LABEL}")
print(f"{'='*55}")

checked = unchanged = changed = oos = not_found = errors = 0
changes = []

for row in range(2, last_row + 1):
    row_num   = ws.cell(row=row, column=3).value  # col C = #
    item_type = ws.cell(row=row, column=4).value  # col D = Item Type
    name      = ws.cell(row=row, column=5).value  # col E = Name
    sku_raw   = ws.cell(row=row, column=7).value  # col G = SKU
    listed    = ws.cell(row=row, column=9).value  # col I = Listed Price

    if not row_num: continue

    sku = clean_sku(sku_raw)
    if not sku:
        cell = ws.cell(row=row, column=write_col, value='No SKU')
        cell.font      = Font(color=DGRAY, size=9, name='Arial')
        cell.fill      = PatternFill('solid', start_color=LGRAY)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border    = T
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
        status_text = 'OOS' if not avail else bb_name[:12]
        cell = ws.cell(row=row, column=write_col, value=status_text)
        cell.font      = Font(color=RED_FG, bold=True, size=9, name='Arial')
        cell.fill      = PatternFill('solid', start_color=RED_BG)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border    = T
        not_found += 1
        print(f"  ❌ {sku} [{item_type}] {(name or '')[:35]} — {bb_name}")
        continue

    if not avail:
        oos += 1
        print(f"  📦 {sku} [{item_type}] {(name or '')[:35]} — OOS ${price:.2f}")

    # Compare to listed price
    try:
        listed_f = float(str(listed).replace('$','').replace(',','')) if listed else None
    except: listed_f = None

    if listed_f and abs(price - listed_f) / listed_f > 0.01:
        diff = price - listed_f
        pct  = (diff / listed_f) * 100
        bg   = AMBER_L if price > listed_f else GREEN_BG
        fg   = AMBER   if price > listed_f else GREEN_FG
        changed += 1
        changes.append((item_type, name, listed_f, price, pct))
        print(f"  {'📈' if price > listed_f else '📉'} {sku} [{item_type}] {(name or '')[:30]} | ${listed_f:.2f}→${price:.2f} ({pct:+.1f}%)")
    else:
        bg, fg = TEAL_L, "00897B"
        unchanged += 1

    cell = ws.cell(row=row, column=write_col, value=price)
    cell.font         = Font(color=fg, bold=(bg != TEAL_L), size=10, name='Arial')
    cell.fill         = PatternFill('solid', start_color=bg)
    cell.number_format = '$#,##0.00'
    cell.alignment    = Alignment(horizontal='center', vertical='center')
    cell.border       = T

# ── SAVE ─────────────────────────────────────────────────────────────────────
wb.save(EXCEL_FILE)

print(f"\n{'='*55}")
print(f"COMPLETE — {TODAY_LABEL}")
print(f"  ✅ Checked:   {checked}")
print(f"  — Unchanged:  {unchanged}")
print(f"  📊 Changed:   {changed}")
print(f"  📦 OOS:       {oos}")
print(f"  ❌ Not found: {not_found}")
print(f"  ⚠️  Errors:    {errors}")
if changes:
    print(f"\nPrice changes:")
    for itype, name, old, new, pct in sorted(changes, key=lambda x: abs(x[4]), reverse=True):
        print(f"  {itype} | {(name or '')[:35]} | ${old:.2f}→${new:.2f} ({pct:+.1f}%)")
print(f"{'='*55}")
print(f"Saved: {EXCEL_FILE}")
