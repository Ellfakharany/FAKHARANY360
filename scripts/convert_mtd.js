/**
 * convert_mtd.js — converts daily MTD Dashboard files (Mobile + WE Pay)
 * dropped into raw_mtd/ into the same JSON shape the browser's MTD tab
 * already knows how to render, so the dashboard can auto-load the latest
 * day for the whole team without anyone touching the manual "Upload MTD" /
 * "Upload Wallet" buttons.
 *
 * IMPORTANT — keep in sync with index.html:
 * The parsing functions below (parseMTDWorkbook / parseMTDFlat /
 * parseMTDMultiLevel / parseIDLESheet / parsePkgMixSheet / parseDailySheet /
 * parseWalletSheet / stripXlsbMetadata) are a Node port of the identically
 * named functions in index.html. If the MTD/WE Pay workbook layout changes
 * and you fix the browser parser, mirror the same fix here — and vice versa.
 *
 * Usage (used by the GitHub Action):
 *   node scripts/convert_mtd.js --scan raw_mtd data/mtd
 *
 * File naming convention expected inside raw_mtd/:
 *   Mobile file  -> must contain "mobile"                + a date, e.g.
 *                   "Mobile_MTD_10-06-2026.xlsb", "Mobile Dashboard 2026-06-10.xlsb"
 *   WE Pay file  -> must contain "wallet"/"we_pay"/"wepay"/"we pay" + a date, e.g.
 *                   "We_Pay_MTD_10-06-2026.xlsb", "Wallet 2026-06-10.xlsb"
 * Accepted date formats: YYYY-MM-DD, DD-MM-YYYY, DD_MM_YYYY, DD.MM.YYYY
 * (separators -, _, ., or space all work).
 */
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const JSZip = require('jszip');

// ══════════════════════════════════════════════════════════════
//  Shared config — copied verbatim from index.html
// ══════════════════════════════════════════════════════════════
const MTD_COL_MAP = {
  'branch name as bss':'store','storecodebss':'storeCode','storecodebs':'storeCode',
  'partner':'partner','classification':'classification','region':'region',
  'partner account manager':'accountManager','partner channel manager':'channelManager',
  'area manager':'areaManager','area manger':'areaManager',
  'supervisor':'supervisor','sinor supervisor':'seniorSupervisor',
  'regional manager':'regionalManager',
  'mobile target':'mobileTarget','mobile subscriptions':'mobileSubs',
  'mobile %':'mobilePct','mobile proj':'mobileProj','mobile proj %':'mobileProjPct',
  'fbb target':'fbbTarget','fbb subscriptions':'fbbSubs',
  'fbb %':'fbbPct','fbb proj':'fbbProj','fbb proj %':'fbbProjPct',
  'fwa target':'fwaTarget','fwa subscriptions':'fwaSubs',
  'fwa %':'fwaPct','fwa proj':'fwaProj','fwa proj %':'fwaProjPct',
  'fixed target':'fixedTarget','fixed subscriptions':'fixedSubs',
  'fixed %':'fixedPct','fixed proj':'fixedProj','fixed proj %':'fixedProjPct',
  'we gold target':'goldTarget','we gold subscriptions':'goldSubs',
  'we gold %':'goldPct','we gold proj':'goldProj','we gold proj %':'goldProjPct',
  'we mix target':'weMixTarget','we mix subscriptions':'weMixSubs',
  'we mix %':'weMixPct','we mix proj':'weMixProj','we mix proj %':'weMixProjPct',
  'control kix target':'kixTarget','control kix subscriptions':'kixSubs','control kix proj':'kixProj',
  'control tazbeet target':'tazbeetTarget','control tazbeet subscriptions':'tazbeetSubs',
};
const MTD_TEXT_KEYS = ['store','storeCode','partner','classification','region',
  'accountManager','channelManager','areaManager','supervisor','seniorSupervisor','regionalManager'];

const MTD_PKG_CATS = [
  { key:'kix',     test:(c)=>/kix/i.test(c) },
  { key:'tazbeet', test:(c)=>/tazbeet/i.test(c) },
  { key:'mifi',    test:(c)=>/mifi/i.test(c) },
  { key:'weClub',  test:(c)=>/we\s*club/i.test(c) },
  { key:'gold',    test:(c)=>/gold/i.test(c) },
  { key:'weMix',   test:(c)=>/we\s*mix/i.test(c) },
  { key:'payg',    test:(c)=>/^12\s*pt$/i.test(c) },
  { key:'other',   test:()=>false },
];
const MTD_PKG_DIM_COLS = ['Channel Manager','Regional Manager','Sinor supervisor','SuperVisor','Area Manger',
  'Region','Branch Name as BSS','Classification','Tariff Plan','Partner','Partner Account Manager','Partner Manager'];

// ══════════════════════════════════════════════════════════════
//  xlsb metadata.bin strip (fixes "Unexpected record 0x3b")
// ══════════════════════════════════════════════════════════════
async function stripXlsbMetadata(buffer) {
  try {
    const zip = await JSZip.loadAsync(buffer);
    if (!zip.files['xl/metadata.bin']) return buffer;
    zip.file('xl/metadata.bin', new Uint8Array(0));
    return await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  } catch (e) {
    console.warn('  ⚠️ stripXlsbMetadata failed, using original file:', e.message);
    return buffer;
  }
}

// ══════════════════════════════════════════════════════════════
//  MTD Database sheet parsing (Mobile file)
// ══════════════════════════════════════════════════════════════
function parseMTDFlat(raw) {
  return raw.map(function (row) {
    var obj = {};
    Object.keys(row).forEach(function (k) {
      var mapped = MTD_COL_MAP[String(k).trim().toLowerCase()];
      if (!mapped) return;
      var v = row[k];
      if (MTD_TEXT_KEYS.indexOf(mapped) >= 0) {
        obj[mapped] = (v === null || v === undefined || v === '(blank)') ? '' : String(v).trim();
      } else {
        obj[mapped] = (v === null || v === undefined || v === '') ? 0 : (isNaN(Number(v)) ? 0 : Number(v));
      }
    });
    return obj;
  }).filter(function (r) { return r.store && r.store.length > 0; });
}

function parseMTDMultiLevel(ws) {
  var aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
  if (!aoa || aoa.length < 5) throw new Error('بنية الملف غير متوقعة — عدد صفوف قليل جداً');

  var fieldRowIdx = -1;
  for (var i = 0; i < Math.min(15, aoa.length); i++) {
    if (aoa[i] && aoa[i][0] === 'StoreCodeBSS') { fieldRowIdx = i; break; }
  }
  if (fieldRowIdx < 0) throw new Error('مش لاقي صف الهيدر (StoreCodeBSS) في أول 15 صف');

  var PRODUCTS = ['Mobile', 'FWA', 'Fixed', 'FBB', 'WE Gold', 'WE Mix', 'Control Kix', 'Control Tazbeet'];
  function rowHasProductLabel(row) {
    if (!row) return false;
    return row.some(function (c) {
      if (typeof c !== 'string') return false;
      return PRODUCTS.some(function (p) { return c.indexOf(p) === 0; });
    });
  }
  var groupRows = [];
  for (var g = Math.max(0, fieldRowIdx - 4); g < fieldRowIdx; g++) {
    if (rowHasProductLabel(aoa[g])) groupRows.push(g);
  }

  var INFO_NAMES = ['StoreCodeBSS', 'Branch Name as BSS', 'Partner', 'Classification', 'Region',
    'Partner Account Manager', 'Partner Channel Manager', 'Area Manager',
    'SuperVisor', 'Sinor supervisor', 'Regional Manager', 'Channel Manager'];

  function hdr(ci) {
    var fieldVal = aoa[fieldRowIdx][ci];
    var fieldStr = (fieldVal === null || fieldVal === undefined) ? '' : String(fieldVal).trim();
    var labels = [];
    for (var gi = groupRows.length - 1; gi >= 0; gi--) {
      var v = aoa[groupRows[gi]][ci];
      if (v !== null && v !== undefined && String(v).trim() !== '') labels.push(String(v).trim());
    }
    if (labels.length > 0) {
      var nearest = labels[0];
      if (nearest.indexOf(' ') > 0 && !fieldStr) return nearest;
      if (fieldStr) return nearest + ' ' + fieldStr;
      return nearest;
    }
    return fieldStr || ('Col_' + ci);
  }

  var totalCols = aoa[fieldRowIdx].length;
  var keepCols = [];
  for (var c = 0; c < totalCols; c++) keepCols.push(c);
  var headers = keepCols.map(function (ci) { return ci < 12 ? INFO_NAMES[ci] : hdr(ci); });

  var dataStartRow = fieldRowIdx + 1;
  var fakeRaw = [];
  for (var ri = dataStartRow; ri < aoa.length; ri++) {
    var row = aoa[ri];
    if (!row || !row[0]) continue;
    if (String(row[0]).trim().toLowerCase() === 'grand total') continue;
    var obj = {};
    keepCols.forEach(function (ci, idx) { obj[headers[idx]] = row[ci]; });
    fakeRaw.push(obj);
  }
  return parseMTDFlat(fakeRaw);
}

function parseMTDWorkbook(wb) {
  var sheetName = wb.SheetNames.find(function (s) { return s.toLowerCase() === 'database'; }) || wb.SheetNames[0];
  var ws = wb.Sheets[sheetName];
  if (!ws) throw new Error('مش لاقي شيت "Database". الشيتات: ' + wb.SheetNames.join(', '));

  var raw = XLSX.utils.sheet_to_json(ws, { defval: null, raw: true });
  if (raw && raw.length > 0) {
    var firstKey = Object.keys(raw[0])[0] || '';
    if (!/unnamed|^__/i.test(firstKey)) return parseMTDFlat(raw);
  }
  return parseMTDMultiLevel(ws);
}

// ══════════════════════════════════════════════════════════════
//  IDLE sheet (Activation vs Sales split)
// ══════════════════════════════════════════════════════════════
function parseIDLESheet(wb) {
  var sheetName = wb.SheetNames.find(function (s) { return s.toLowerCase() === 'idle'; });
  if (!sheetName) return null;
  var ws = wb.Sheets[sheetName];
  var aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
  if (!aoa || aoa.length < 5) return null;

  var headerRowIdx = -1;
  for (var i = 0; i < Math.min(15, aoa.length); i++) {
    if (aoa[i] && aoa[i][0] === 'StoreCodeBSS') { headerRowIdx = i; break; }
  }
  if (headerRowIdx < 0) return null;
  var storeColIdx = 0;
  var fieldRow = aoa[headerRowIdx] || [];
  var totalCols = fieldRow.length;
  var colLabel = [];

  for (var r = Math.max(0, headerRowIdx - 3); r < headerRowIdx; r++) {
    var rr = aoa[r] || [];
    for (var ci = 0; ci < totalCols; ci++) {
      var v = rr[ci];
      if (typeof v === 'string' && /All (Sales|Activation)/.test(v)) colLabel[ci] = v.trim();
    }
  }
  for (var ci2 = 0; ci2 < totalCols; ci2++) {
    if (colLabel[ci2]) continue;
    var fv = fieldRow[ci2];
    var fStr = (typeof fv === 'string') ? fv.trim() : '';
    if (fStr !== 'All Sales' && fStr !== 'All Activation') continue;
    var grp = null;
    for (var r2 = headerRowIdx - 1; r2 >= Math.max(0, headerRowIdx - 3) && !grp; r2--) {
      var gv = (aoa[r2] || [])[ci2];
      if (typeof gv === 'string' && gv.trim() !== '') grp = gv.trim();
    }
    if (grp) colLabel[ci2] = grp + ' ' + fStr;
  }

  function findProductCols(productLabel) {
    var salesCol = -1, actCol = -1;
    for (var ci = 0; ci < totalCols; ci++) {
      if (colLabel[ci] === productLabel + ' All Sales') salesCol = ci;
      if (colLabel[ci] === productLabel + ' All Activation') actCol = ci;
    }
    return { salesCol: salesCol, actCol: actCol };
  }

  var mobileCols = findProductCols('Mobile');
  var fixedCols = findProductCols('Fixed');
  var fbbCols = findProductCols('FBB');
  var goldCols = findProductCols('WE Gold');
  var wemixCols = findProductCols('WE Mix');
  if (mobileCols.salesCol < 0 && fixedCols.salesCol < 0 && fbbCols.salesCol < 0 && goldCols.salesCol < 0) return null;

  var result = {};
  for (var ri = headerRowIdx + 1; ri < aoa.length; ri++) {
    var row = aoa[ri];
    if (!row || !row[storeColIdx]) continue;
    var code = String(row[storeColIdx]).trim().toUpperCase();
    if (!code || code === 'GRAND TOTAL') continue;
    result[code] = {
      mobileSales: Number(row[mobileCols.salesCol]) || 0,
      mobileAct: Number(row[mobileCols.actCol]) || 0,
      fixedSales: Number(row[fixedCols.salesCol]) || 0,
      fixedAct: Number(row[fixedCols.actCol]) || 0,
      fbbSales: Number(row[fbbCols.salesCol]) || 0,
      fbbAct: Number(row[fbbCols.actCol]) || 0,
      goldSales: goldCols.salesCol >= 0 ? (Number(row[goldCols.salesCol]) || 0) : null,
      goldAct: goldCols.actCol >= 0 ? (Number(row[goldCols.actCol]) || 0) : null,
      wemixSales: wemixCols.salesCol >= 0 ? (Number(row[wemixCols.salesCol]) || 0) : null,
      wemixAct: wemixCols.actCol >= 0 ? (Number(row[wemixCols.actCol]) || 0) : null,
    };
  }
  return result;
}

// Mirrors the merge blocks inside index.html's loadMTDFile(): builds the
// Sales dataset and rebuilds the Activation (default) dataset from IDLE.
function mergeIdleIntoRows(parsedRows, idleData) {
  var salesRows = parsedRows.map(function (r) {
    var code = String(r.storeCode || '').trim().toUpperCase();
    var idle = idleData[code];
    var sr = Object.assign({}, r);
    if (idle) {
      sr.mobileSubs = idle.mobileSales;
      sr.mobilePct = r.mobileTarget > 0 ? idle.mobileSales / r.mobileTarget : 0;
      sr.fixedSubs = idle.fixedSales;
      sr.fixedPct = r.fixedTarget > 0 ? idle.fixedSales / r.fixedTarget : 0;
      sr.fbbSubs = idle.fbbSales;
      sr.fbbPct = r.fbbTarget > 0 ? idle.fbbSales / r.fbbTarget : 0;
      if (idle.goldSales !== null) { sr.goldSubs = idle.goldSales; sr.goldPct = r.goldTarget > 0 ? idle.goldSales / r.goldTarget : 0; }
      if (idle.wemixSales !== null) { sr.weMixSubs = idle.wemixSales; sr.weMixPct = r.weMixTarget > 0 ? idle.wemixSales / r.weMixTarget : 0; }
    }
    return sr;
  });
  var actRows = parsedRows.map(function (r) {
    var code = String(r.storeCode || '').trim().toUpperCase();
    var idle = idleData[code];
    if (!idle) return r;
    var sr = Object.assign({}, r);
    sr.mobileSubs = idle.mobileAct;
    sr.mobilePct = r.mobileTarget > 0 ? idle.mobileAct / r.mobileTarget : 0;
    sr.fixedSubs = idle.fixedAct;
    sr.fixedPct = r.fixedTarget > 0 ? idle.fixedAct / r.fixedTarget : 0;
    sr.fbbSubs = idle.fbbAct;
    sr.fbbPct = r.fbbTarget > 0 ? idle.fbbAct / r.fbbTarget : 0;
    if (idle.goldAct !== null) { sr.goldSubs = idle.goldAct; sr.goldPct = r.goldTarget > 0 ? idle.goldAct / r.goldTarget : 0; }
    if (idle.wemixAct !== null) { sr.weMixSubs = idle.wemixAct; sr.weMixPct = r.weMixTarget > 0 ? idle.wemixAct / r.weMixTarget : 0; }
    return sr;
  });
  return { salesRows: salesRows, actRows: actRows };
}

// ══════════════════════════════════════════════════════════════
//  Tariffs Per Stoers (Package Mix) + Stoers Daily (Daily Trend)
// ══════════════════════════════════════════════════════════════
function parsePkgMixSheet(wb) {
  var sheetName = wb.SheetNames.find(function (s) { return /tariffs per sto/i.test(s); });
  if (!sheetName) return null;
  var ws = wb.Sheets[sheetName];
  var aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
  if (!aoa || aoa.length < 3) return null;

  var headerRowIdx = -1;
  for (var i = 0; i < 10; i++) { if (aoa[i] && aoa[i].indexOf('StoreCodeBSS') >= 0) { headerRowIdx = i; break; } }
  if (headerRowIdx < 0) return null;
  var fieldRow = aoa[headerRowIdx];
  var storeColIdx = fieldRow.indexOf('StoreCodeBSS');
  var totalColIdx = fieldRow.indexOf('Grand Total');
  if (storeColIdx < 0) return null;

  function classify(name) {
    for (var ci = 0; ci < MTD_PKG_CATS.length - 1; ci++) { if (MTD_PKG_CATS[ci].test(name)) return MTD_PKG_CATS[ci].key; }
    return 'other';
  }

  var cols = [];
  for (var ci2 = 0; ci2 < fieldRow.length; ci2++) {
    if (ci2 === storeColIdx || ci2 === totalColIdx) continue;
    var name = fieldRow[ci2];
    if (!name || typeof name !== 'string') continue;
    if (MTD_PKG_DIM_COLS.indexOf(name) >= 0) continue;
    cols.push({ ci: ci2, label: name, cat: classify(name) });
  }

  var rows = [];
  for (var ri = headerRowIdx + 1; ri < aoa.length; ri++) {
    var row = aoa[ri];
    if (!row || !row[storeColIdx]) continue;
    var code = String(row[storeColIdx]).trim().toUpperCase();
    if (!code || code === 'GRAND TOTAL') continue;
    var vals = {};
    for (var ci3 = 0; ci3 < cols.length; ci3++) { var v = row[cols[ci3].ci]; if (v) vals[cols[ci3].ci] = Number(v) || 0; }
    rows.push({ storeCode: code, vals: vals });
  }
  return { cols: cols, rows: rows };
}

function parseDailySheet(wb) {
  var sheetName = wb.SheetNames.find(function (s) { return /sto[er]+s daily/i.test(s) || /sales daily/i.test(s); });
  if (!sheetName) return null;
  var ws = wb.Sheets[sheetName];
  var aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
  if (!aoa || aoa.length < 3) return null;

  var headerRowIdx = -1;
  for (var i = 0; i < 10; i++) { if (aoa[i] && aoa[i].indexOf('StoreCodeBSS') >= 0) { headerRowIdx = i; break; } }
  if (headerRowIdx < 0) return null;
  var fieldRow = aoa[headerRowIdx];
  var storeColIdx = fieldRow.indexOf('StoreCodeBSS');
  if (storeColIdx < 0) return null;

  var dayCols = [];
  for (var ci = 0; ci < fieldRow.length; ci++) {
    var v = fieldRow[ci];
    if (typeof v === 'number' && v >= 1 && v <= 31) dayCols.push({ ci: ci, day: v });
  }
  dayCols.sort(function (a, b) { return a.day - b.day; });
  if (dayCols.length === 0) return null;

  var rows = [];
  for (var ri = headerRowIdx + 1; ri < aoa.length; ri++) {
    var row = aoa[ri];
    if (!row || !row[storeColIdx]) continue;
    var code = String(row[storeColIdx]).trim().toUpperCase();
    if (!code || code === 'GRAND TOTAL') continue;
    var days = {};
    for (var di = 0; di < dayCols.length; di++) { var dv = row[dayCols[di].ci]; if (dv) days[dayCols[di].day] = (days[dayCols[di].day] || 0) + (Number(dv) || 0); }
    rows.push({ storeCode: code, days: days });
  }
  return { dayCols: dayCols.map(function (d) { return d.day; }), rows: rows };
}

// ══════════════════════════════════════════════════════════════
//  WE Pay ("Sales VS Target") — mirrors loadWalletFile() in index.html
// ══════════════════════════════════════════════════════════════
function parseWalletSheet(wb) {
  var sheetName = wb.SheetNames.find(function (s) { return /sales\s*vs\s*target/i.test(s); });
  if (!sheetName) {
    sheetName = wb.SheetNames.find(function (s) {
      var test = XLSX.utils.sheet_to_json(wb.Sheets[s], { defval: null, raw: true, range: 0 });
      if (!test || !test.length) return false;
      var keys = Object.keys(test[0]).map(function (k) { return k.toLowerCase(); });
      return keys.some(function (k) { return k.indexOf('storecode') >= 0 || k.indexOf('wallet target') >= 0; });
    });
  }
  if (!sheetName) sheetName = wb.SheetNames[0];

  var ws = wb.Sheets[sheetName];
  var raw = XLSX.utils.sheet_to_json(ws, { defval: null, raw: true });
  if (!raw || raw.length === 0) throw new Error('الشيت "' + sheetName + '" فاضي');

  var WCOL = {
    'storecodebss': 'storeCode', 'storecodebs': 'storeCode',
    'branchnameasbss': 'store',
    'wallettarget': 'walletTarget', 'target': 'walletTarget',
    'sales': 'walletSales', 'walletsales': 'walletSales',
    'ach%': 'walletPct', '%': 'walletPct', 'wallet%': 'walletPct',
    'proj': 'walletProj', 'walletproj': 'walletProj',
    'proj%': 'walletProjPct', 'walletproj%': 'walletProjPct',
  };
  function normKey(k) { return String(k).trim().toLowerCase().replace(/\s+/g, ''); }

  var WALLET_DATA = {};
  raw.forEach(function (row) {
    var obj = {};
    Object.keys(row).forEach(function (k) {
      var mapped = WCOL[normKey(k)];
      if (!mapped) return;
      var v = row[k];
      if (mapped === 'store' || mapped === 'storeCode') {
        obj[mapped] = (v === null || v === undefined) ? '' : String(v).trim();
      } else {
        obj[mapped] = (v === null || v === undefined || v === '') ? 0 : (isNaN(Number(v)) ? 0 : Number(v));
      }
    });
    var key = String(obj.storeCode || '').trim().toUpperCase();
    if (key) WALLET_DATA[key] = obj;
  });

  var count = Object.keys(WALLET_DATA).length;
  if (count === 0) throw new Error('مش لاقي بيانات Wallet في شيت "' + sheetName + '"');
  return WALLET_DATA;
}

// ══════════════════════════════════════════════════════════════
//  raw_mtd/ filename → date + kind detection
// ══════════════════════════════════════════════════════════════
const AR_MONTHS = ['يناير', 'فبراير', 'مارس', 'أبريل', 'مايو', 'يونيو', 'يوليو', 'أغسطس', 'سبتمبر', 'أكتوبر', 'نوفمبر', 'ديسمبر'];

function detectDate(filename) {
  var name = String(filename).replace(/\.[^.]+$/, '');
  var m = name.match(/(20\d{2})[-_. ](\d{1,2})[-_. ](\d{1,2})/);
  if (m) {
    var y = +m[1], mo = +m[2], d = +m[3];
    if (mo >= 1 && mo <= 12 && d >= 1 && d <= 31) return iso(y, mo, d);
  }
  m = name.match(/(\d{1,2})[-_. ](\d{1,2})[-_. ](20\d{2})/);
  if (m) {
    var d2 = +m[1], mo2 = +m[2], y2 = +m[3];
    if (mo2 >= 1 && mo2 <= 12 && d2 >= 1 && d2 <= 31) return iso(y2, mo2, d2);
  }
  return null;
  function iso(y, mo, d) {
    var isoStr = y + '-' + String(mo).padStart(2, '0') + '-' + String(d).padStart(2, '0');
    var label = d + ' ' + AR_MONTHS[mo - 1] + ' ' + y;
    return { iso: isoStr, label: label };
  }
}

function detectKind(filename) {
  var low = filename.toLowerCase();
  if (low.indexOf('we_pay') >= 0 || low.indexOf('wepay') >= 0 || low.indexOf('we pay') >= 0 || low.indexOf('wallet') >= 0) return 'wallet';
  if (low.indexOf('mobile') >= 0) return 'mobile';
  return null;
}

function safeCall(fn, label) {
  try { return fn(); } catch (e) { console.warn('  ⚠️ ' + (label || 'parse') + ' failed:', e.message); return null; }
}

function sanitize(obj) {
  // Replace NaN/Infinity so JSON.stringify never silently emits invalid JSON.
  return JSON.parse(JSON.stringify(obj, function (k, v) {
    if (typeof v === 'number' && !isFinite(v)) return 0;
    return v;
  }));
}

// ══════════════════════════════════════════════════════════════
//  Main scan/convert
// ══════════════════════════════════════════════════════════════
async function scanAndConvert(rawDir, dataDir) {
  fs.mkdirSync(dataDir, { recursive: true });
  const files = fs.readdirSync(rawDir).filter(function (f) { return /\.xlsb$/i.test(f); });

  const manifestPath = path.join(dataDir, 'manifest.json');
  let manifest = { mobile: null, wallet: null };
  if (fs.existsSync(manifestPath)) {
    try { manifest = JSON.parse(fs.readFileSync(manifestPath, 'utf8')); } catch (e) { /* start fresh */ }
  }

  if (files.length === 0) {
    console.log('  ℹ️ No new .xlsb files in ' + rawDir);
    return;
  }

  // If several files of the same kind are pushed at once, keep only the one
  // with the latest date (a team member may re-upload a corrected file).
  const byKind = { mobile: null, wallet: null };
  for (const file of files) {
    const kind = detectKind(file);
    const date = detectDate(file);
    if (!kind || !date) { console.warn('  ⚠️  Skipping (could not detect type/date): ' + file); continue; }
    if (!byKind[kind] || date.iso > byKind[kind].date.iso) byKind[kind] = { file: file, date: date };
  }

  for (const kind of ['mobile', 'wallet']) {
    const pick = byKind[kind];
    if (!pick) continue;
    const fullPath = path.join(rawDir, pick.file);
    console.log('  → Processing ' + pick.file + ' as ' + kind + ' for ' + pick.date.iso);
    const buf = fs.readFileSync(fullPath);
    const cleaned = await stripXlsbMetadata(buf);
    const wb = XLSX.read(cleaned, { type: 'buffer', cellDates: false, raw: true });

    if (kind === 'mobile') {
      const parsed = parseMTDWorkbook(wb);
      if (!parsed || parsed.length === 0) { console.warn('  ⚠️  ' + pick.file + ': لم يتم العثور على بيانات صالحة — skipped'); continue; }
      const idle = safeCall(function () { return parseIDLESheet(wb); }, 'IDLE');
      let mtdData = parsed, mtdSalesData = [];
      if (idle) {
        const merged = mergeIdleIntoRows(parsed, idle);
        mtdData = merged.actRows;
        mtdSalesData = merged.salesRows;
      }
      const pkgMix = safeCall(function () { return parsePkgMixSheet(wb); }, 'PkgMix');
      const daily = safeCall(function () { return parseDailySheet(wb); }, 'Daily');
      const out = sanitize({ date: pick.date.iso, dateLabel: pick.date.label, mtdData: mtdData, mtdSalesData: mtdSalesData, pkgMix: pkgMix, daily: daily });
      const outName = 'mobile-' + pick.date.iso + '.json';
      fs.writeFileSync(path.join(dataDir, outName), JSON.stringify(out));
      manifest.mobile = { date: pick.date.iso, dateLabel: pick.date.label, file: outName, stores: mtdData.length };
      console.log('  ✅ Mobile ' + pick.date.iso + ' → ' + outName + ' (' + mtdData.length + ' stores)');
    } else {
      const walletData = safeCall(function () { return parseWalletSheet(wb); }, 'Wallet');
      if (!walletData) { console.warn('  ⚠️  ' + pick.file + ': WE Pay parse failed — skipped'); continue; }
      const walletDaily = safeCall(function () { return parseDailySheet(wb); }, 'WalletDaily');
      const out = sanitize({ date: pick.date.iso, dateLabel: pick.date.label, walletData: walletData, walletDaily: walletDaily });
      const outName = 'wallet-' + pick.date.iso + '.json';
      fs.writeFileSync(path.join(dataDir, outName), JSON.stringify(out));
      manifest.wallet = { date: pick.date.iso, dateLabel: pick.date.label, file: outName, stores: Object.keys(walletData).length };
      console.log('  ✅ Wallet ' + pick.date.iso + ' → ' + outName + ' (' + Object.keys(walletData).length + ' stores)');
    }
  }

  fs.writeFileSync(manifestPath, JSON.stringify(manifest, null, 2));
  console.log('  📄 data/mtd/manifest.json updated:', JSON.stringify(manifest));
}

const args = process.argv.slice(2);
if (args[0] === '--scan' && args[1] && args[2]) {
  scanAndConvert(args[1], args[2]).catch(function (e) { console.error(e); process.exit(1); });
} else {
  console.error('Usage: node convert_mtd.js --scan <raw_mtd_dir> <data/mtd_dir>');
  process.exit(1);
}
