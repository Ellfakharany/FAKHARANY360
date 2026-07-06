"""
Converts Closure files (Mobile/FBB/Fixed workbook + WE Pay workbook) into the
JSON row schema used by FAKHARANY360's dashboard.

Two modes:

1) Single pair (manual):
   python convert_closure.py <mobile.xlsb> <wepay.xlsb> <MM-YYYY> <output.json>

2) Batch/folder mode (used by the GitHub Action):
   python convert_closure.py --scan <raw_dir> <data_dir>
   Scans <raw_dir> for *.xlsb files, auto-detects each file's month and
   whether it's the Mobile or WE Pay workbook from its filename (matching
   the naming convention already used, e.g. "...Mobile...May-2026...xlsb"
   and "We_Pay...May-2026...xlsb"), pairs them up, converts every complete
   month found, writes <data_dir>/<YYYY-MM>.json for each, and refreshes
   <data_dir>/manifest.json with the sorted list of available months.
"""
import sys, os, json, re, glob
import pandas as pd

MONTH_ABBR = {1:'Jan',2:'Feb',3:'Mar',4:'Apr',5:'May',6:'Jun',
              7:'Jul',8:'Aug',9:'Sep',10:'Oct',11:'Nov',12:'Dec'}
MONTH_NAME_TO_NUM = {
    'jan':1,'january':1, 'feb':2,'february':2, 'mar':3,'march':3, 'apr':4,'april':4,
    'may':5, 'jun':6,'june':6, 'jul':7,'july':7, 'aug':8,'august':8,
    'sep':9,'sept':9,'september':9, 'oct':10,'october':10, 'nov':11,'november':11,
    'dec':12,'december':12,
}

LOW_SET = {25, 29, 32, 37}
MID_SET = {40, 45, 46, 52}


def classify_tier(n):
    if n in LOW_SET: return 'low'
    if n in MID_SET: return 'mid'
    return 'high'

def norm_code(x):
    return str(x).strip().upper()

def parse_database_sheet(path):
    """Database sheet: header row is row index 7 (0-based), data starts row 8."""
    df = pd.read_excel(path, sheet_name='Database', engine='pyxlsb', header=None)
    rows = {}
    for i in range(8, len(df)):
        r = df.iloc[i]
        code = r[0]
        if pd.isna(code): continue
        code = norm_code(code)

        # Skip summary/total rows — they have something in the code column
        # (e.g. "Grand Total") but no real store name, which produced NaN
        # fields all the way through and broke the JSON output.
        if pd.isna(r[1]) or 'grand total' in code.lower() or 'total' == code.lower():
            continue

        def num(col):
            v = r[col]
            return 0 if pd.isna(v) else float(v)

        rows[code] = {
            'storeCode': code,
            'store': r[1],
            'partner': r[2],
            'classification': r[3],
            'region': r[4],
            'accountManager': r[5],
            'channelManager': r[6],
            'areaManager': r[7],
            'supervisor': r[8],
            'regionalManager': r[10],
            # Sub-product families (Database columns, 0-indexed)
            'kixTarget': num(12), 'kixSubs': num(13),
            'tazbeetTarget': num(15), 'tazbeetSubs': num(16),
            'dataSimMifiTarget': num(18), 'dataSimMifiSubs': num(19),
            'paygTarget': num(21), 'paygSubs': num(22),
            'prepaidTarget': num(24), 'prepaidSubs': num(25),
            'weClubTarget': num(27), 'weClubSubs': num(28),
            'goldTarget': num(30), 'goldSubs': num(31),
            'weMixTarget': num(33), 'weMixSubs': num(34),
            # Aggregate totals
            'mobileTarget': num(36), 'mobileSubs': num(37),
            'fwaTarget': num(42), 'fwaSubs': num(43),
            'fixedTarget': num(48), 'fixedSubs': num(49),
            'fbbTarget': num(54), 'fbbSubs': num(55),
        }
    return rows

def parse_tariffs_sheet(path):
    """Tariffs Per Stoers sheet: header row index 6, data from row 7."""
    df = pd.read_excel(path, sheet_name='Tariffs Per Stoers', engine='pyxlsb', header=None)
    header = df.iloc[6]
    code_col = None
    col_labels = {}
    for j, label in header.items():
        if pd.isna(label): continue
        label = str(label).strip()
        if label == 'StoreCodeBSS':
            code_col = j
        else:
            col_labels[j] = label

    out = {}
    for i in range(7, len(df)):
        r = df.iloc[i]
        code = r[code_col]
        if pd.isna(code): continue
        code = norm_code(code)

        low = mid = high = pt12 = 0.0
        kix_fields, taz_fields = {}, {}

        for j, label in col_labels.items():
            v = r[j]
            if pd.isna(v) or v == 0: continue
            low_label = label.lower()

            if low_label == '12 pt':
                pt12 += float(v)
                continue
            if 'grand total' in low_label:
                continue

            is_kix = 'kix' in low_label
            is_taz = 'tazbeet' in low_label
            if not (is_kix or is_taz):
                continue  # Gold/Wallet/Wifi/etc packages don't factor into tariff tiers
            v = float(v)

            m = re.search(r'(\d+)(?!.*\d)', label)
            if not m:
                # Non-numeric variant (e.g. "Kix Fn" flexible plan) — still count it,
                # bucketed as High tier since it doesn't fit a low/mid price point.
                high += v
                key = ('kix' if is_kix else 'taz') + 'Other'
                if is_kix: kix_fields[key] = kix_fields.get(key, 0) + v
                else: taz_fields[key] = taz_fields.get(key, 0) + v
                continue
            n = int(m.group(1))
            tier = classify_tier(n)
            if tier == 'low': low += v
            elif tier == 'mid': mid += v
            else: high += v

            if is_kix:
                kix_fields['kix' + str(n)] = kix_fields.get('kix' + str(n), 0) + v
            else:
                taz_fields['taz' + str(n)] = taz_fields.get('taz' + str(n), 0) + v

        out[code] = {'lowT': low, 'midT': mid, 'highT': high, 'pt12': pt12,
                     **kix_fields, **taz_fields}
    return out

def parse_wallet_sheet(path):
    df = pd.read_excel(path, sheet_name='Sales VS Target', engine='pyxlsb', header=None)
    header = df.iloc[0]
    col = {str(v).strip(): j for j, v in header.items() if pd.notna(v)}
    out = {}
    for i in range(1, len(df)):
        r = df.iloc[i]
        code = r[col['StoreCodeBSS']]
        if pd.isna(code): continue
        code = norm_code(code)
        out[code] = {
            'walletTarget': 0 if pd.isna(r[col['Wallet Target']]) else float(r[col['Wallet Target']]),
            'walletSales':  0 if pd.isna(r[col['Sales']])         else float(r[col['Sales']]),
        }
    return out

def detect_month(filename):
    """Finds a Month-Year pattern in a filename, e.g. 'May-2026', 'May_2026', 'May 2026'."""
    m = re.search(r'([A-Za-z]{3,9})[\s_.\-]+(\d{4})', filename)
    if not m:
        return None
    name = m.group(1).lower()
    year = int(m.group(2))
    if name not in MONTH_NAME_TO_NUM:
        return None
    mm = MONTH_NAME_TO_NUM[name]
    return f'{mm:02d}-{year}'

def detect_kind(filename):
    """Mobile Closure workbook vs WE Pay Closure workbook, from filename keywords."""
    low = filename.lower()
    if 'we_pay' in low or 'wepay' in low or 'we pay' in low or 'wallet' in low:
        return 'wallet'
    if 'mobile' in low:
        return 'mobile'
    return None

def scan_and_convert(raw_dir, data_dir):
    files = glob.glob(os.path.join(raw_dir, '*.xlsb'))
    pairs = {}  # month_str -> {'mobile': path, 'wallet': path}
    for f in files:
        name = os.path.basename(f)
        month_str = detect_month(name)
        kind = detect_kind(name)
        if not month_str or not kind:
            print(f'  ⚠️  Skipping (could not detect month/type): {name}')
            continue
        pairs.setdefault(month_str, {})[kind] = f

    os.makedirs(data_dir, exist_ok=True)
    processed = []
    for month_str, pair in sorted(pairs.items()):
        if 'mobile' not in pair or 'wallet' not in pair:
            missing = 'WE Pay' if 'mobile' in pair else 'Mobile'
            print(f'  ⚠️  {month_str}: missing the {missing} file — skipped')
            continue
        yyyy, mm = month_str.split('-')[1], month_str.split('-')[0]
        out_name = f'{yyyy}-{mm}.json'
        out_path = os.path.join(data_dir, out_name)
        rows = build_month(pair['mobile'], pair['wallet'], month_str)
        with open(out_path, 'w', encoding='utf-8') as fh:
            json.dump(sanitize_rows(rows), fh, ensure_ascii=False, allow_nan=False)
        print(f'  ✅ {month_str} → {out_name} ({len(rows)} stores)')
        processed.append(f'{yyyy}-{mm}')

    # Refresh manifest.json with every JSON file present in data_dir
    all_months = sorted(set(
        os.path.splitext(os.path.basename(p))[0]
        for p in glob.glob(os.path.join(data_dir, '*.json'))
        if re.match(r'^\d{4}-\d{2}$', os.path.splitext(os.path.basename(p))[0])
    ))
    manifest_path = os.path.join(data_dir, 'manifest.json')
    with open(manifest_path, 'w', encoding='utf-8') as fh:
        json.dump({'months': all_months}, fh, ensure_ascii=False, indent=2)
    print(f'  📄 manifest.json updated — {len(all_months)} month(s) total: {all_months}')
    return processed


def sanitize_rows(rows):
    """Belt-and-suspenders: replace any NaN/Infinity that slipped through
    (e.g. from an unexpected blank cell) with a safe default, so we never
    write invalid JSON again. Numbers -> 0, strings -> ''."""
    import math
    for row in rows:
        for k, v in list(row.items()):
            if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
                row[k] = 0
    return rows


def build_month(mobile_path, wepay_path, month_str):
    mm, yyyy = month_str.split('-')
    month2 = MONTH_ABBR[int(mm)]

    db = parse_database_sheet(mobile_path)
    tariffs = parse_tariffs_sheet(mobile_path)
    wallet = parse_wallet_sheet(wepay_path)

    rows = []
    for code, base in db.items():
        t = tariffs.get(code, {})
        w = wallet.get(code, {'walletTarget': 0, 'walletSales': 0})
        row = {
            'month': month_str, 'month2': month2,
            'store': base['store'], 'partner': base['partner'],
            'classification': base['classification'], 'region': base['region'],
            'accountManager': base['accountManager'], 'channelManager': base['channelManager'],
            'areaManager': base['areaManager'], 'supervisor': base['supervisor'],
            'regionalManager': base['regionalManager'], 'storeCode': code,
            'mobileTarget': base['mobileTarget'], 'mobileSubs': base['mobileSubs'],
            'goldTarget': base['goldTarget'], 'goldSubs': base['goldSubs'],
            'fbbTarget': base['fbbTarget'], 'fbbSubs': base['fbbSubs'],
            'fixedTarget': base['fixedTarget'], 'fixedSubs': base['fixedSubs'],
            'walletTarget': w['walletTarget'], 'walletSales': w['walletSales'],
            'fwaTarget': base['fwaTarget'], 'fwaSubs': base['fwaSubs'],
            'lowT': t.get('lowT', 0), 'midT': t.get('midT', 0),
            'highT': t.get('highT', 0), 'pt12': t.get('pt12', 0),
            'kixTarget': base['kixTarget'], 'kixSubs': base['kixSubs'],
            'tazbeetTarget': base['tazbeetTarget'], 'tazbeetSubs': base['tazbeetSubs'],
            'dataSimTarget': base['dataSimMifiTarget'], 'dataSimSubs': base['dataSimMifiSubs'],
            'weMixTarget': base['weMixTarget'], 'weMixSubs': base['weMixSubs'],
            'weClubTarget': base['weClubTarget'], 'weClubSubs': base['weClubSubs'],
            'paygTarget': base['paygTarget'], 'paygSubs': base['paygSubs'],
        }
        for k, v in t.items():
            if k.startswith('kix') or k.startswith('taz'):
                row[k] = v
        rows.append(row)
    return rows

if __name__ == '__main__':
    if len(sys.argv) >= 2 and sys.argv[1] == '--scan':
        raw_dir, data_dir = sys.argv[2:4]
        scan_and_convert(raw_dir, data_dir)
    else:
        mobile_path, wepay_path, month_str, out_path = sys.argv[1:5]
        rows = build_month(mobile_path, wepay_path, month_str)
        with open(out_path, 'w', encoding='utf-8') as f:
            json.dump(sanitize_rows(rows), f, ensure_ascii=False, allow_nan=False)
        print(f'Wrote {len(rows)} store rows to {out_path}')
