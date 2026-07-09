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

def normalize_name(s):
    """Strips non-breaking spaces / collapses stray whitespace so the same
    store never ends up as two different-looking entries across months.
    Also treats Excel PivotTable export artifacts like the literal text
    '(blank)' as truly empty, so it doesn't show up as a fake filter option
    in the dashboard (this happens when the source workbook's PivotTable had
    no value for that cell and Excel wrote the placeholder text instead of
    leaving it empty)."""
    if s is None:
        return s
    s = str(s).replace('\xa0', ' ')
    s = re.sub(r'\s+', ' ', s).strip()
    if s.lower() in ('(blank)', 'blank', 'n/a', 'na', '#n/a'):
        return ''
    return s

def find_total_col(header_row, product_label, field_label, required=True):
    """Locate a 'grand total' column (e.g. 'FBB Target') by its exact header
    text in the given header row, instead of a hardcoded column index.
    Package columns get added/removed over time (e.g. WE added a batch of
    new FBB speed tiers in June-2026), which shifts every column after them —
    a fixed index silently starts reading the wrong package's numbers. Since
    the WORKBOOK ALWAYS repeats the product name in the total's header
    (e.g. 'FBB Target', 'FBB Subscriptions', 'FBB %'), searching by that text
    is immune to columns moving around.

    Some products (e.g. FWA) simply didn't exist yet in older months' workbooks,
    so their columns are legitimately absent — not a layout error. When
    required=False, a missing column just returns None (with a warning printed)
    instead of raising, so the caller can treat that product as 0/unavailable
    for that month rather than failing the whole conversion.
    """
    target_text = f'{product_label} {field_label}'.strip().lower()
    for j, label in header_row.items():
        if pd.isna(label):
            continue
        if str(label).strip().lower() == target_text:
            return j
    if not required:
        print(f"  ⚠️  Column '{target_text}' not found — treating as unavailable for this month.")
        return None
    raise ValueError(
        f"Could not find column '{target_text}' in the Database sheet header row — "
        f"the workbook layout may have changed again. Check row 5 of the 'Database' sheet."
    )


def find_col_by_keyword(header_row, keyword, exclude=None, max_col=20):
    """Locate a per-store header column (Account Manager, Area Manager, etc.)
    by searching for a keyword rather than trusting a fixed column number.
    These columns have been reordered/duplicated between months before (e.g.
    March-2026 inserted a second 'Channel Manager' column and moved 'Regional
    Manager' one slot earlier), which silently fed the wrong manager's name
    into the wrong field under a fixed-index scheme. Only searches the first
    `max_col` columns, since the keyword could otherwise coincidentally match
    something far off in the huge package-total section of the sheet.
    `exclude`, if given, skips any header containing that substring too —
    used to tell 'Channel Manager' apart from 'Regional Manager' having both
    literally contain neither word, but keeps the two "Channel Manager"-ish
    columns (Partner Channel Manager vs plain Channel Manager) distinguishable
    if ever needed.
    """
    keyword = keyword.lower()
    for j, label in header_row.items():
        if j >= max_col:
            break
        if pd.isna(label):
            continue
        low = str(label).strip().lower()
        if keyword in low and (exclude is None or exclude not in low):
            return j
    return None


def parse_database_sheet(path):
    """Database sheet: header row is row index 7 (0-based) for per-store
    columns, and row index 5 holds the repeated 'X Target'/'X Subscriptions'/
    'X %' labels for each product's grand-total triple. Data starts row 8."""
    df = pd.read_excel(path, sheet_name='Database', engine='pyxlsb', header=None)
    store_header = df.iloc[7]
    total_header = df.iloc[5]

    # Locate the per-store manager/partner columns by header keyword, not by
    # a fixed index — see find_col_by_keyword docstring for why. Falls back
    # to the historical fixed index only if the keyword search comes up
    # empty, so this stays compatible with older months whose headers may be
    # blank/differently worded.
    store_col      = find_col_by_keyword(store_header, 'branch name')
    partner_col    = find_col_by_keyword(store_header, 'partner')
    classif_col    = find_col_by_keyword(store_header, 'classification')
    region_col     = find_col_by_keyword(store_header, 'region')
    account_col    = find_col_by_keyword(store_header, 'account manager')
    channel_col    = find_col_by_keyword(store_header, 'channel manager')
    area_col       = find_col_by_keyword(store_header, 'area man')  # tolerates "Area Manger" typo
    supervisor_col = find_col_by_keyword(store_header, 'supervisor')
    regional_col   = find_col_by_keyword(store_header, 'regional manager')

    # Fall back to the original fixed positions for anything the keyword
    # search didn't find, so a month with unusually blank/odd headers still
    # processes instead of erroring out.
    if store_col is None: store_col = 1
    if partner_col is None: partner_col = 2
    if classif_col is None: classif_col = 3
    if region_col is None: region_col = 4
    if account_col is None: account_col = 5
    if channel_col is None: channel_col = 6
    if area_col is None: area_col = 7
    if supervisor_col is None: supervisor_col = 8
    if regional_col is None: regional_col = 10

    # Resolve each product's grand-total columns by header text (see
    # find_total_col docstring) — NOT by a fixed column index, since new
    # package/tariff columns inserted upstream shift everything after them.
    # required=False everywhere: a product that hasn't launched yet in an
    # older month (e.g. FWA before it existed) just means that column is
    # absent from the sheet — not a broken layout — so we tolerate it and
    # fall back to 0 for that product/month instead of aborting the whole run.
    mobile_t_col = find_total_col(total_header, 'Mobile', 'Target', required=False)
    mobile_a_col = find_total_col(total_header, 'Mobile', 'Subscriptions', required=False)
    fwa_t_col    = find_total_col(total_header, 'FWA', 'Target', required=False)
    fwa_a_col    = find_total_col(total_header, 'FWA', 'Subscriptions', required=False)
    fixed_t_col  = find_total_col(total_header, 'Fixed', 'Target', required=False)
    fixed_a_col  = find_total_col(total_header, 'Fixed', 'Subscriptions', required=False)
    fbb_t_col    = find_total_col(total_header, 'FBB', 'Target', required=False)
    fbb_a_col    = find_total_col(total_header, 'FBB', 'Subscriptions', required=False)

    rows = {}
    for i in range(8, len(df)):
        r = df.iloc[i]
        code = r[0]
        if pd.isna(code): continue
        code = norm_code(code)

        # Skip summary/total rows — they have something in the code column
        # (e.g. "Grand Total") but no real store name, which produced NaN
        # fields all the way through and broke the JSON output.
        if pd.isna(r[store_col]) or 'grand total' in code.lower() or 'total' == code.lower():
            continue

        def num(col):
            if col is None:
                return 0  # this product's column didn't exist in this month's workbook
            v = r[col]
            return 0 if pd.isna(v) else float(v)

        rows[code] = {
            'storeCode': code,
            'store': normalize_name(r[store_col]),
            'partner': normalize_name(r[partner_col]),
            'classification': r[classif_col],
            'region': r[region_col],
            'accountManager': normalize_name(r[account_col]),
            'channelManager': normalize_name(r[channel_col]),
            'areaManager': normalize_name(r[area_col]),
            'supervisor': normalize_name(r[supervisor_col]),
            'regionalManager': normalize_name(r[regional_col]),
            # Sub-product families (Database columns, 0-indexed) — these are
            # near the front of the sheet and haven't moved historically, but
            # if WE ever restructures Mobile's own package breakdown too,
            # these should get the same header-lookup treatment.
            'kixTarget': num(12), 'kixSubs': num(13),
            'tazbeetTarget': num(15), 'tazbeetSubs': num(16),
            'dataSimMifiTarget': num(18), 'dataSimMifiSubs': num(19),
            'paygTarget': num(21), 'paygSubs': num(22),
            'prepaidTarget': num(24), 'prepaidSubs': num(25),
            'weClubTarget': num(27), 'weClubSubs': num(28),
            'goldTarget': num(30), 'goldSubs': num(31),
            'weMixTarget': num(33), 'weMixSubs': num(34),
            # Aggregate totals — located by header text, see above
            'mobileTarget': num(mobile_t_col), 'mobileSubs': num(mobile_a_col),
            'fwaTarget': num(fwa_t_col), 'fwaSubs': num(fwa_a_col),
            'fixedTarget': num(fixed_t_col), 'fixedSubs': num(fixed_a_col),
            'fbbTarget': num(fbb_t_col), 'fbbSubs': num(fbb_a_col),
        }
    return rows

def parse_idle_sheet(path):
    """IDLE sheet: holds a separate Activation-vs-Sales split for Mobile,
    Fixed and FBB ('<Product> All Sales' / '<Product> All Activation'),
    same header-name-lookup approach as the Database totals — columns move
    around whenever WE adds/removes package columns upstream. Returns {} on
    any problem (missing sheet, missing columns, etc.) so older workbooks
    without this split just fall back to Sales-only everywhere, instead of
    breaking the whole conversion.
    """
    try:
        df = pd.read_excel(path, sheet_name='IDLE', engine='pyxlsb', header=None)
    except Exception as e:
        print(f"  ⚠️  IDLE sheet: could not read ({e}) — falling back to Sales-only for this month.")
        return {}

    header = df.iloc[5]
    cols = {}
    try:
        for prod in ('Mobile', 'Fixed', 'FBB'):
            cols[prod+'Sales'] = find_total_col(header, prod, 'All Sales')
            cols[prod+'Act']   = find_total_col(header, prod, 'All Activation')
    except ValueError as e:
        # this month's IDLE sheet doesn't have the split — skip quietly, but
        # log *why* so we can tell "sheet missing" apart from "headers renamed"
        print(f"  ⚠️  IDLE sheet: {e} — falling back to Sales-only for this month.")
        return {}

    out = {}
    for i in range(8, len(df)):
        r = df.iloc[i]
        code = r[0]
        if pd.isna(code) or pd.isna(r[1]): continue
        code = norm_code(code)

        def num(col):
            v = r[col]
            return 0 if pd.isna(v) else float(v)

        out[code] = {
            'mobileSubsAct': num(cols['MobileAct']),
            'fixedSubsAct':  num(cols['FixedAct']),
            'fbbSubsAct':    num(cols['FBBAct']),
        }
    return out


def parse_tariffs_sheet(path):
    """Tariffs Per Stoers sheet: header row usually at index 6, data right
    after it. That row index has shifted before (e.g. March-2026 workbook),
    so we search nearby rows for the 'StoreCodeBSS' marker instead of
    assuming a fixed position — the same problem the Database/IDLE totals
    had, just for the header *row* instead of a header *column*.
    """
    df = pd.read_excel(path, sheet_name='Tariffs Per Stoers', engine='pyxlsb', header=None)

    header_row_idx = None
    code_col = None
    col_labels = {}
    for candidate in range(0, min(15, len(df))):
        row = df.iloc[candidate]
        found_code_col = None
        labels = {}
        for j, label in row.items():
            if pd.isna(label): continue
            label = str(label).strip()
            if label == 'StoreCodeBSS':
                found_code_col = j
            else:
                labels[j] = label
        if found_code_col is not None:
            header_row_idx, code_col, col_labels = candidate, found_code_col, labels
            break

    if code_col is None:
        print("  ⚠️  Tariffs Per Stoers: could not find 'StoreCodeBSS' header in the first "
              "15 rows — skipping tariff-mix data for this month (targets/subs are unaffected).")
        return {}

    out = {}
    for i in range(header_row_idx + 1, len(df)):
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

    header_row_idx = None
    col = {}
    for candidate in range(0, min(15, len(df))):
        row = df.iloc[candidate]
        labels = {str(v).strip(): j for j, v in row.items() if pd.notna(v)}
        if 'StoreCodeBSS' in labels:
            header_row_idx, col = candidate, labels
            break

    if header_row_idx is None or 'Wallet Target' not in col or 'Sales' not in col:
        print("  ⚠️  Sales VS Target (Wallet): could not find expected headers "
              "('StoreCodeBSS' / 'Wallet Target' / 'Sales') — skipping Wallet data for this month.")
        return {}

    out = {}
    for i in range(header_row_idx + 1, len(df)):
        r = df.iloc[i]
        code = r[col['StoreCodeBSS']]
        if pd.isna(code): continue
        code = norm_code(code)
        out[code] = {
            'walletTarget': 0 if pd.isna(r[col['Wallet Target']]) else float(r[col['Wallet Target']]),
            'walletSales':  0 if pd.isna(r[col['Sales']])         else float(r[col['Sales']]),
        }
    return out

MANAGER_FIELDS = ('supervisor', 'areaManager', 'regionalManager', 'accountManager', 'channelManager')

def build_manager_reference(data_dir, exclude_month=None):
    """Builds a storeCode -> {supervisor, areaManager, ...} lookup table from
    every other month's already-converted JSON in data_dir, most-recent-wins.
    This is the VLOOKUP-by-store-code idea: some months' source workbook has
    a manager column that's genuinely blank at the source (e.g. March-2026's
    Supervisor column was blank for all 264 stores) — that can't be recovered
    from that month's own file, but the store-to-supervisor/area-manager/etc.
    assignment barely changes month to month, so backfilling from whichever
    other month last had a real value for that store is a reasonable stand-in
    until the source workbook itself gets corrected.
    """
    ref = {}
    for f in sorted(glob.glob(os.path.join(data_dir, '*.json'))):
        base = os.path.splitext(os.path.basename(f))[0]
        if not re.match(r'^\d{4}-\d{2}$', base) or base == exclude_month:
            continue
        try:
            with open(f, encoding='utf-8') as fh:
                month_rows = json.load(fh)
        except Exception:
            continue
        for row in month_rows:
            code = row.get('storeCode')
            if not code:
                continue
            entry = ref.setdefault(code, {})
            for field in MANAGER_FIELDS:
                val = row.get(field)
                if val:  # non-empty; later (sorted-ascending) months overwrite earlier ones
                    entry[field] = val
    return ref


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
        try:
            manager_ref = build_manager_reference(data_dir, exclude_month=f'{yyyy}-{mm}')
            rows, filled_count = build_month(pair['mobile'], pair['wallet'], month_str, manager_ref)
            with open(out_path, 'w', encoding='utf-8') as fh:
                json.dump(sanitize_rows(rows), fh, ensure_ascii=False, allow_nan=False)
            extra = f' — {filled_count} manager field(s) backfilled from other months' if filled_count else ''
            print(f'  ✅ {month_str} → {out_name} ({len(rows)} stores){extra}')
            processed.append(f'{yyyy}-{mm}')
        except Exception as e:
            # One month's workbook having an unexpected/broken layout shouldn't
            # stop every other month from being processed and committed.
            # Print the FULL traceback, not just str(e) — some exceptions
            # (e.g. KeyError(None)) stringify to something unhelpful like
            # "None", which hides the real cause and file/line it happened at.
            import traceback
            print(f'  ❌ {month_str}: failed to convert — skipped, other months continue.')
            print(f'     Exception type: {type(e).__name__}')
            traceback.print_exc()

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


def build_month(mobile_path, wepay_path, month_str, manager_reference=None):
    mm, yyyy = month_str.split('-')
    month2 = MONTH_ABBR[int(mm)]

    db = parse_database_sheet(mobile_path)
    tariffs = parse_tariffs_sheet(mobile_path)
    wallet = parse_wallet_sheet(wepay_path)
    idle = parse_idle_sheet(mobile_path)
    manager_reference = manager_reference or {}

    rows = []
    filled_count = 0
    for code, base in db.items():
        t = tariffs.get(code, {})
        w = wallet.get(code, {'walletTarget': 0, 'walletSales': 0})
        idl = idle.get(code, {})
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

        # VLOOKUP-by-store-code fallback: if this month's source workbook left
        # a manager field blank (e.g. March-2026's Supervisor column), backfill
        # it from whatever other month last had a real value for that same
        # store code, rather than shipping an empty field to the dashboard.
        ref_entry = manager_reference.get(code)
        if ref_entry:
            for field in MANAGER_FIELDS:
                if not row.get(field) and ref_entry.get(field):
                    row[field] = ref_entry[field]
                    filled_count += 1

        # Activation figures (Mobile/Fixed/FBB only) — omitted entirely for
        # months whose IDLE sheet didn't have a usable split, so the frontend
        # cleanly falls back to Sales for those months.
        if idl:
            row['mobileSubsAct'] = idl['mobileSubsAct']
            row['fixedSubsAct']  = idl['fixedSubsAct']
            row['fbbSubsAct']    = idl['fbbSubsAct']
        for k, v in t.items():
            if k.startswith('kix') or k.startswith('taz'):
                row[k] = v
        rows.append(row)
    return rows, filled_count

if __name__ == '__main__':
    if len(sys.argv) >= 2 and sys.argv[1] == '--scan':
        raw_dir, data_dir = sys.argv[2:4]
        scan_and_convert(raw_dir, data_dir)
    else:
        mobile_path, wepay_path, month_str, out_path = sys.argv[1:5]
        manager_ref = build_manager_reference(os.path.dirname(out_path) or '.')
        rows, filled_count = build_month(mobile_path, wepay_path, month_str, manager_ref)
        with open(out_path, 'w', encoding='utf-8') as f:
            json.dump(sanitize_rows(rows), f, ensure_ascii=False, allow_nan=False)
        print(f'Wrote {len(rows)} store rows to {out_path}')
