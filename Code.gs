/**
 * Nasi Cokot Daisafa — Apps Script backend
 *
 * Setup:
 *   1. Convert NASI COKOT DAISAFA.xlsx to a Google Sheet (old tabs stay as archive).
 *   2. Extensions > Apps Script, paste Code.gs, index.html, appsscript.json.
 *   3. Run ensureSchema() once from the editor (safe, never deletes).
 *   4. Deploy > New deployment > Web app (execute as you, access only you).
 *   5. In the app: Pengaturan > Impor data lama.
 *
 * API actions:
 *   bootstrap, dashboard, orderBoard, listOrders, getOrder,
 *   createOrder, updateOrder, cancelOrder,
 *   listCustomers, saveCustomer, mergeCustomers,
 *   listMenu, saveMenu, listExpenses, saveExpense,
 *   report, getSettings, saveSettings, importLegacy
 *
 * Safety: ensureSchema() only creates missing sheets / writes headers on new ones.
 * There is intentionally no destructive setupSheets().
 */

var TZ = 'Asia/Jakarta';
var APP_TITLE = 'Nasi Cokot Daisafa';
var DAY_NAMES = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
var MONTH_SHORT = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
var MONTH_LONG = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];

var SCHEMA = {
  Customers: ['customer_id', 'name', 'name_key', 'whatsapp', 'customer_key', 'aliases', 'merged_into', 'notes', 'active', 'created_at', 'updated_at'],
  Menu: ['menu_id', 'name', 'short_name', 'price', 'cost', 'active', 'sort_order'],
  Orders: ['invoice_no', 'delivery_date', 'customer_id', 'customer_name', 'menu1_id', 'menu2_id', 'qty', 'unit_price', 'unit_cost', 'line_total', 'margin', 'payment_status', 'payment_method', 'order_status', 'notes', 'paid_at', 'created_at', 'updated_at'],
  Expenses: ['expense_id', 'expense_date', 'category', 'description', 'amount', 'payment_method', 'notes', 'created_at', 'updated_at'],
  Settings: ['setting_key', 'setting_value', 'description', 'updated_at'],
  Import_Log: ['logged_at', 'source', 'source_row', 'level', 'issue', 'raw']
};

var MENU_SEED = [
  ['M01', 'Kulit Krengsengan', 'Kulit', 7000, 6000, true, 1],
  ['M02', 'Mercon', 'Mercon', 7000, 6000, true, 2],
  ['M03', 'Rica Kemangi', 'Rica', 7000, 6000, true, 3],
  ['M04', 'Bumbu Rendang', 'Rendang', 7000, 6000, true, 4],
  ['M05', 'Bumbu Bebek', 'Bebek', 7000, 6000, true, 5]
];

var SETTINGS_SEED = [
  ['business_name', 'Nasi Cokot Daisafa', 'Nama usaha'],
  ['owner_whatsapp', '', 'Nomor WhatsApp pemilik'],
  ['qris_url', '', 'URL gambar QRIS'],
  ['invoice_footer', 'Terima kasih!', 'Teks bawah nota'],
  ['default_price', '7000', 'Harga jual per porsi'],
  ['default_cost', '6000', 'Harga modal per porsi'],
  ['invoice_prefix', 'NCD', 'Awalan nomor nota'],
  ['timezone', 'Asia/Jakarta', 'Zona waktu'],
  ['currency', 'IDR', 'Mata uang'],
  ['legacy_import_done', 'false', 'Status impor data lama']
];

/* ------------------------------------------------------------------ */
/* Entry points                                                        */
/* ------------------------------------------------------------------ */

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle(APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

function api(req) {
  try {
    req = req || {};
    var action = String(req.action || '');
    var p = req.payload || {};
    var data;
    switch (action) {
      case 'bootstrap': data = bootstrap_(p); break;
      case 'dashboard': data = dashboard_(p); break;
      case 'orderBoard': data = orderBoard_(p); break;
      case 'listOrders': data = listOrders_(p); break;
      case 'getOrder': data = getOrder_(p); break;
      case 'createOrder': data = createOrder_(p); break;
      case 'updateOrder': data = updateOrder_(p); break;
      case 'cancelOrder': data = cancelOrder_(p); break;
      case 'listCustomers': data = listCustomers_(p); break;
      case 'saveCustomer': data = saveCustomer_(p); break;
      case 'mergeCustomers': data = mergeCustomers_(p); break;
      case 'listMenu': data = listMenu_(p); break;
      case 'saveMenu': data = saveMenu_(p); break;
      case 'listExpenses': data = listExpenses_(p); break;
      case 'saveExpense': data = saveExpense_(p); break;
      case 'report': data = report_(p); break;
      case 'getSettings': data = { settings: getSettings_() }; break;
      case 'saveSettings': data = saveSettings_(p); break;
      case 'importLegacy': data = importLegacy_(p); break;
      default: return fail_('VALIDATION', 'Aksi tidak dikenal: ' + action);
    }
    return { ok: true, data: data };
  } catch (err) {
    var code = err && err.code ? err.code : 'INTERNAL';
    var msg = err && err.message ? err.message : 'Terjadi kesalahan.';
    return { ok: false, code: code, message: msg };
  }
}

function ok_(data) { return { ok: true, data: data }; }
function fail_(code, message) { return { ok: false, code: code, message: message }; }
function appError_(code, message) {
  var e = new Error(message);
  e.code = code;
  return e;
}

/* ------------------------------------------------------------------ */
/* Schema                                                              */
/* ------------------------------------------------------------------ */

function ensureSchema() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var created = [];
  var verified = [];

  Object.keys(SCHEMA).forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.getRange(1, 1, 1, SCHEMA[name].length).setValues([SCHEMA[name]]);
      sh.setFrozenRows(1);
      created.push(name);
    } else {
      assertHeaders_(sh, name);
      verified.push(name);
    }
    styleSheet_(sh, name);
  });

  seedMenu_();
  seedSettings_();

  var summary = { created: created, verified: verified };
  Logger.log('ensureSchema selesai: ' + JSON.stringify(summary));
  return summary;
}

function styleSheet_(sh, name) {
  var headers = SCHEMA[name];
  var rows = Math.max(sh.getMaxRows() - 1, 1);
  sh.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#1A3263')
    .setFontColor('#FFFFFF')
    .setVerticalAlignment('middle');
  sh.setFrozenRows(1);

  if (name === 'Orders') {
    sh.getRange(2, 1, rows, headers.length).setNumberFormat('@');
    sh.getRange(2, 7, rows, 1).setNumberFormat('0');
    sh.getRange(2, 8, rows, 4).setNumberFormat('#,##0');
  } else if (name === 'Menu') {
    sh.getRange(2, 1, rows, headers.length).setNumberFormat('@');
    sh.getRange(2, 4, rows, 2).setNumberFormat('#,##0');
    sh.getRange(2, 7, rows, 1).setNumberFormat('0');
  } else if (name === 'Expenses') {
    sh.getRange(2, 1, rows, headers.length).setNumberFormat('@');
    sh.getRange(2, 5, rows, 1).setNumberFormat('#,##0');
  } else if (name === 'Import_Log') {
    sh.getRange(2, 1, rows, headers.length).setNumberFormat('@');
    sh.getRange(2, 3, rows, 1).setNumberFormat('0');
  } else {
    sh.getRange(2, 1, rows, headers.length).setNumberFormat('@');
  }

  var widths = {
    Customers: [100, 160, 140, 130, 170, 160, 100, 180, 70, 150, 150],
    Menu: [80, 180, 110, 90, 90, 70, 90],
    Orders: [150, 110, 100, 150, 80, 80, 60, 90, 90, 100, 90, 110, 110, 100, 220, 150, 150, 150],
    Expenses: [90, 110, 110, 220, 100, 110, 180, 150, 150],
    Settings: [170, 220, 220, 150],
    Import_Log: [150, 140, 90, 70, 320, 320]
  };
  (widths[name] || []).forEach(function (w, i) {
    sh.setColumnWidth(i + 1, w);
  });
}

function assertHeaders_(sh, name) {
  var expected = SCHEMA[name];
  var actual = sh.getRange(1, 1, 1, expected.length).getValues()[0];
  for (var i = 0; i < expected.length; i++) {
    if (String(actual[i] || '').trim() !== expected[i]) {
      throw appError_('SCHEMA',
        'Header sheet ' + name + ' tidak cocok. Kolom ' + (i + 1) +
        ' seharusnya "' + expected[i] + '" tetapi "' + String(actual[i] || '') +
        '". Jangan ubah urutan header; jalankan ensureSchema() setelah memperbaiki.');
    }
  }
}

function seedMenu_() {
  var sh = sheet_('Menu');
  var last = lastDataRow_(sh, 1);
  if (last >= 2) return;
  sh.getRange(2, 1, MENU_SEED.length, SCHEMA.Menu.length).setValues(MENU_SEED);
}

function seedSettings_() {
  var existing = {};
  readTable_('Settings').forEach(function (r) { existing[r.setting_key] = true; });
  var rows = [];
  var now = nowIso_();
  SETTINGS_SEED.forEach(function (s) {
    if (!existing[s[0]]) rows.push([s[0], s[1], s[2], now]);
  });
  if (rows.length) appendRows_('Settings', rows);
}

/* ------------------------------------------------------------------ */
/* Sheet helpers                                                       */
/* ------------------------------------------------------------------ */

function sheet_(name) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) {
    throw appError_('SCHEMA', 'Sheet ' + name + ' belum ada. Jalankan ensureSchema() dari editor Apps Script terlebih dahulu.');
  }
  return sh;
}

function lastDataRow_(sh, keyCol) {
  var last = sh.getLastRow();
  if (last < 2) return 1;
  var vals = sh.getRange(2, keyCol || 1, last - 1, 1).getValues();
  for (var i = vals.length - 1; i >= 0; i--) {
    if (String(vals[i][0] || '').trim() !== '') return i + 2;
  }
  return 1;
}

function readTable_(name) {
  var sh = sheet_(name);
  var headers = SCHEMA[name];
  var last = lastDataRow_(sh, 1);
  if (last < 2) return [];
  var values = sh.getRange(2, 1, last - 1, headers.length).getValues();
  var out = [];
  values.forEach(function (row) {
    var any = false;
    for (var i = 0; i < row.length; i++) {
      if (String(row[i] || '').trim() !== '') { any = true; break; }
    }
    if (!any) return;
    var obj = {};
    headers.forEach(function (h, i) { obj[h] = row[i]; });
    out.push(obj);
  });
  return out;
}

function objectToRow_(name, obj) {
  return SCHEMA[name].map(function (h) {
    var v = obj[h];
    return v === undefined || v === null ? '' : v;
  });
}

function appendRows_(name, rows) {
  if (!rows || !rows.length) return;
  var sh = sheet_(name);
  var start = lastDataRow_(sh, 1) + 1;
  sh.getRange(start, 1, rows.length, SCHEMA[name].length).setValues(rows);
}

function writeObject_(name, rowIndex, obj) {
  if (rowIndex < 2) throw appError_('INTERNAL', 'Baris data tidak valid.');
  sheet_(name).getRange(rowIndex, 1, 1, SCHEMA[name].length)
    .setValues([objectToRow_(name, obj)]);
}

function findRowIndex_(name, key, value) {
  var sh = sheet_(name);
  var last = lastDataRow_(sh, 1);
  if (last < 2) return -1;
  var col = SCHEMA[name].indexOf(key) + 1;
  if (col < 1) throw appError_('INTERNAL', 'Kolom ' + key + ' tidak ada di ' + name);
  var vals = sh.getRange(2, col, last - 1, 1).getValues();
  var target = String(value);
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]) === target) return i + 2;
  }
  return -1;
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

/* ------------------------------------------------------------------ */
/* Time / value helpers                                                */
/* ------------------------------------------------------------------ */

function nowIso_() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
}

function todayIso_() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
}

function addDaysIso_(iso, days) {
  var d = isoToDate_(iso);
  d.setUTCDate(d.getUTCDate() + days);
  return Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd');
}

function isoToDate_(iso) {
  return new Date(String(iso).slice(0, 10) + 'T00:00:00Z');
}

function parseDate_(v) {
  if (v === null || v === undefined || v === '') return '';
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  }
  if (typeof v === 'number' && isFinite(v)) {
    var serialDate = new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000);
    return Utilities.formatDate(serialDate, 'UTC', 'yyyy-MM-dd');
  }
  var s = String(v).trim();
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[1] + '-' + m[2] + '-' + m[3];
  m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m) return m[3] + '-' + pad_(parseInt(m[2], 10), 2) + '-' + pad_(parseInt(m[1], 10), 2);
  m = s.match(/^(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})$/);
  if (m) {
    var idx = monthIndex_(m[2]);
    if (idx >= 0) return m[3] + '-' + pad_(idx + 1, 2) + '-' + pad_(parseInt(m[1], 10), 2);
  }
  var d = new Date(s);
  if (!isNaN(d.getTime())) return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
  return '';
}

function monthIndex_(name) {
  var n = String(name || '').toLowerCase();
  var full = ['januari', 'februari', 'maret', 'april', 'mei', 'juni', 'juli', 'agustus', 'september', 'oktober', 'november', 'desember'];
  var short = ['jan', 'feb', 'mar', 'apr', 'mei', 'jun', 'jul', 'agu', 'sep', 'okt', 'nov', 'des'];
  var i = full.indexOf(n);
  if (i >= 0) return i;
  i = short.indexOf(n.slice(0, 3));
  return i;
}

function parseMoney_(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return Math.round(v);
  var s = String(v).replace(/[^0-9\-]/g, '');
  if (!s || s === '-') return 0;
  return parseInt(s, 10) || 0;
}

function toBool_(v) {
  return v === true || String(v).toLowerCase() === 'true' || String(v) === '1';
}

function pad_(n, len) {
  var s = String(n);
  while (s.length < len) s = '0' + s;
  return s;
}

function normalizePhone_(v) {
  var s = String(v || '').replace(/[^0-9+]/g, '').replace(/\+/g, '');
  if (!s) return '';
  if (s.indexOf('6208') === 0) s = '62' + s.slice(4);
  if (s.indexOf('62') === 0) {
    // already international
  } else if (s.charAt(0) === '0') {
    s = '62' + s.slice(1);
  } else if (s.charAt(0) === '8') {
    s = '62' + s;
  }
  s = s.replace(/[^0-9]/g, '');
  if (s.length < 10 || s.length > 15) return '';
  return s;
}

function normalizeNameKey_(v) {
  return String(v || '')
    .normalize('NFC')
    .replace(/[\u2060\uFEFF\u200B-\u200D\uFE0F]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function normalizeFlavor_(v) {
  var s = String(v || '')
    .normalize('NFC')
    .replace(/[\u2060\uFEFF\u200B-\u200D\uFE0F]/g, '')
    .toLowerCase()
    .replace(/^\s*\d+\s*[\.\)]\s*/, '')
    .replace(/[^a-z0-9\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  var map = {
    'kulit krengsengan': 'M01', 'kulit': 'M01', 'krengsengan': 'M01',
    'mercon': 'M02',
    'rica': 'M03', 'rica kemangi': 'M03', 'kemangi': 'M03',
    'bumbu rendang': 'M04', 'rendang': 'M04',
    'bumbu bebek': 'M05', 'bebek': 'M05'
  };
  return map[s] || null;
}

function normalizeMethod_(v) {
  var s = String(v || '').trim().toLowerCase();
  if (!s) return '';
  if (s === 'cash' || s === 'tunai') return 'Tunai';
  if (s === 'qris') return 'QRIS';
  if (s === 'bsi' || s === 'transfer' || s === 'bank') return 'Transfer';
  return '';
}

function methodFromStatus_(v) {
  var m = String(v || '').match(/paid\s*-\s*(bsi|qris|cash)/i);
  return m ? normalizeMethod_(m[1]) : '';
}

function sortMenuIds_(id1, id2, menuMap) {
  if (!id2 || id1 === id2) return [id1, ''];
  var a = menuMap[id1];
  var b = menuMap[id2];
  if (!a || !b) return [id1, id2];
  return Number(a.sort_order) <= Number(b.sort_order) ? [id1, id2] : [id2, id1];
}

function menuLabel_(id1, id2, menuMap) {
  var a = menuMap[id1];
  var b = id2 ? menuMap[id2] : null;
  if (!a) return id1 || '';
  if (!b || id1 === id2) return a.short_name || a.name;
  return Number(a.sort_order) <= Number(b.sort_order)
    ? (a.short_name || a.name) + '-' + (b.short_name || b.name)
    : (b.short_name || b.name) + '-' + (a.short_name || a.name);
}

function dayLabel_(iso) {
  if (!iso) return '';
  var d = isoToDate_(iso);
  return DAY_NAMES[d.getUTCDay()] + ', ' + d.getUTCDate() + ' ' + MONTH_SHORT[d.getUTCMonth()] + ' ' + d.getUTCFullYear();
}

function monthLabel_(key) {
  var m = String(key || '').match(/^(\d{4})-(\d{2})$/);
  if (!m) return String(key || '');
  var idx = parseInt(m[2], 10) - 1;
  return MONTH_LONG[idx] + ' ' + m[1];
}

/* ------------------------------------------------------------------ */
/* Settings / lookups                                                  */
/* ------------------------------------------------------------------ */

function getSettings_() {
  var out = {};
  readTable_('Settings').forEach(function (r) {
    out[r.setting_key] = r.setting_value;
  });
  return out;
}

function getSetting_(key, fallback) {
  var v = getSettings_()[key];
  return v === undefined || v === null || v === '' ? (fallback === undefined ? '' : fallback) : v;
}

function setSettingLocked_(key, value) {
  var idx = findRowIndex_('Settings', 'setting_key', key);
  var now = nowIso_();
  if (idx > 0) {
    sheet_('Settings').getRange(idx, 2).setValue(value);
    sheet_('Settings').getRange(idx, 4).setValue(now);
  } else {
    appendRows_('Settings', [[key, value, '', now]]);
  }
}

function getMenuMap_() {
  var map = {};
  readTable_('Menu').forEach(function (m) {
    m.price = Number(m.price) || 0;
    m.cost = Number(m.cost) || 0;
    m.sort_order = Number(m.sort_order) || 0;
    m.active = toBool_(m.active);
    map[m.menu_id] = m;
  });
  return map;
}

function getCustomerMap_() {
  var map = {};
  readTable_('Customers').forEach(function (c) { map[c.customer_id] = c; });
  return map;
}

function nextCustomerIdLocked_(customers) {
  var max = 0;
  (customers || []).forEach(function (c) {
    var m = String(c.customer_id || '').match(/^CUS-(\d+)$/);
    if (m) max = Math.max(max, parseInt(m[1], 10));
  });
  return 'CUS-' + pad_(max + 1, 4);
}

function nextExpenseIdLocked_(expenses) {
  var max = 0;
  (expenses || []).forEach(function (e) {
    var m = String(e.expense_id || '').match(/^EXP-(\d+)$/);
    if (m) max = Math.max(max, parseInt(m[1], 10));
  });
  return 'EXP-' + pad_(max + 1, 4);
}

function buildSummary_(date, orders, menuMap) {
  var groups = {};
  orders.forEach(function (o) {
    if (o.order_status === 'Batal') return;
    var pair = sortMenuIds_(o.menu1_id, o.menu2_id, menuMap);
    var key = pair[0] + '+' + (pair[1] || '');
    if (!groups[key]) {
      groups[key] = {
        type: pair[1] ? 'mix' : 'single',
        label: menuLabel_(pair[0], pair[1], menuMap),
        qty: 0
      };
    }
    groups[key].qty += Number(o.qty) || 0;
  });

  var mix = [];
  var single = [];
  Object.keys(groups).forEach(function (k) {
    (groups[k].type === 'mix' ? mix : single).push(groups[k]);
  });
  function byQty(a, b) {
    if (b.qty !== a.qty) return b.qty - a.qty;
    return a.label.localeCompare(b.label);
  }
  mix.sort(byQty);
  single.sort(byQty);

  var total = 0;
  mix.concat(single).forEach(function (g) { total += g.qty; });

  var lines = [dayLabel_(date)];
  if (mix.length) {
    lines.push('--- Menu Mix ---');
    mix.forEach(function (g) { lines.push(g.qty + ' - ' + g.label); });
    lines.push('');
  }
  if (single.length) {
    lines.push('--- Menu 1 Rasa ---');
    single.forEach(function (g) { lines.push(g.qty + ' - ' + g.label); });
    lines.push('');
  }
  lines.push('Total: ' + total);
  return lines.join('\n');
}



/* ------------------------------------------------------------------ */
/* Queries                                                             */
/* ------------------------------------------------------------------ */

function bootstrap_() {
  var settings = getSettings_();
  var menu = readTable_('Menu')
    .map(function (m) {
      return {
        menu_id: m.menu_id,
        name: m.name,
        short_name: m.short_name,
        price: Number(m.price) || 0,
        cost: Number(m.cost) || 0,
        active: toBool_(m.active),
        sort_order: Number(m.sort_order) || 0
      };
    })
    .filter(function (m) { return m.active; })
    .sort(function (a, b) { return a.sort_order - b.sort_order; });
  var today = todayIso_();
  return {
    today: today,
    tomorrow: addDaysIso_(today, 1),
    settings: settings,
    menu: menu
  };
}

function dashboard_() {
  var settings = getSettings_();
  var today = todayIso_();
  var monthKey = today.slice(0, 7);
  var orders = readTable_('Orders').filter(function (o) { return o.order_status !== 'Batal'; });
  var todayOrders = orders.filter(function (o) { return o.delivery_date === today; });
  var monthOrders = orders.filter(function (o) { return String(o.delivery_date).slice(0, 7) === monthKey; });
  var customerMap = getCustomerMap_();

  function agg(list) {
    var out = { portions: 0, omzet: 0, margin: 0, orders: list.length };
    list.forEach(function (o) {
      out.portions += Number(o.qty) || 0;
      out.omzet += Number(o.line_total) || 0;
      out.margin += Number(o.margin) || 0;
    });
    return out;
  }

  var unpaid = orders.filter(function (o) { return o.payment_status === 'Belum Lunas'; });
  var unpaidAmount = 0;
  unpaid.forEach(function (o) { unpaidAmount += Number(o.line_total) || 0; });
  var unpaidList = unpaid.map(function (o) {
    var c = customerMap[o.customer_id] || {};
    return {
      invoice_no: o.invoice_no,
      customer_name: o.customer_name,
      whatsapp: c.whatsapp || '',
      amount: Number(o.line_total) || 0,
      delivery_date: o.delivery_date
    };
  }).sort(function (a, b) { return String(a.delivery_date).localeCompare(String(b.delivery_date)); });

  var cash = { tunai: 0, qris: 0, transfer: 0, total: 0 };
  monthOrders.filter(function (o) { return o.payment_status === 'Lunas'; }).forEach(function (o) {
    var amt = Number(o.line_total) || 0;
    cash.total += amt;
    if (o.payment_method === 'Tunai') cash.tunai += amt;
    else if (o.payment_method === 'QRIS') cash.qris += amt;
    else if (o.payment_method === 'Transfer') cash.transfer += amt;
  });

  var expensesMonth = 0;
  readTable_('Expenses').forEach(function (e) {
    if (String(e.expense_date).slice(0, 7) === monthKey) expensesMonth += Number(e.amount) || 0;
  });

  var todayAgg = agg(todayOrders);
  todayAgg.date = today;
  var monthAgg = agg(monthOrders);
  monthAgg.key = monthKey;
  monthAgg.label = monthLabel_(monthKey);

  return {
    today: todayAgg,
    month: monthAgg,
    receivables: { count: unpaid.length, amount: unpaidAmount },
    cash: cash,
    expenses_month: expensesMonth,
    net_cash_month: cash.total - expensesMonth,
    unpaid: unpaidList,
    qris_url: settings.qris_url || ''
  };
}

function orderBoard_(p) {
  var date = parseDate_(p.date) || todayIso_();
  var orders = readTable_('Orders').filter(function (o) {
    return o.delivery_date === date && o.order_status !== 'Batal';
  });
  var menuMap = getMenuMap_();
  var customerMap = getCustomerMap_();
  var groups = {};
  var total = 0;

  orders.forEach(function (o) {
    var pair = sortMenuIds_(o.menu1_id, o.menu2_id, menuMap);
    var key = pair[0] + '+' + (pair[1] || '');
    if (!groups[key]) {
      groups[key] = {
        type: pair[1] ? 'mix' : 'single',
        label: menuLabel_(pair[0], pair[1], menuMap),
        qty: 0
      };
    }
    groups[key].qty += Number(o.qty) || 0;
    total += Number(o.qty) || 0;
  });

  var mix = [];
  var single = [];
  Object.keys(groups).forEach(function (k) {
    (groups[k].type === 'mix' ? mix : single).push(groups[k]);
  });
  function byQty(a, b) {
    if (b.qty !== a.qty) return b.qty - a.qty;
    return a.label.localeCompare(b.label);
  }
  mix.sort(byQty);
  single.sort(byQty);

  return {
    date: date,
    date_label: dayLabel_(date),
    summary_text: buildSummary_(date, orders, menuMap),
    groups: mix.concat(single),
    total: total,
    orders: orders.map(function (o) { return serializeOrder_(o, customerMap, menuMap); })
      .sort(function (a, b) { return String(a.customer_name).localeCompare(String(b.customer_name)); })
  };
}

function listOrders_(p) {
  var from = parseDate_(p.from) || '';
  var to = parseDate_(p.to) || '';
  var status = String(p.status || 'all');
  var today = todayIso_();
  var tomorrow = addDaysIso_(today, 1);

  var orders = readTable_('Orders').filter(function (o) {
    if (o.order_status === 'Batal' && status !== 'all') return false;
    if (status === 'unpaid' && o.payment_status !== 'Belum Lunas') return false;
    if (status === 'today' && o.delivery_date !== today) return false;
    if (status === 'tomorrow' && o.delivery_date !== tomorrow) return false;
    if (from && String(o.delivery_date) < from) return false;
    if (to && String(o.delivery_date) > to) return false;
    return true;
  });

  var customerMap = getCustomerMap_();
  var menuMap = getMenuMap_();
  orders.sort(function (a, b) {
    var d = String(b.delivery_date).localeCompare(String(a.delivery_date));
    if (d !== 0) return d;
    return String(b.invoice_no).localeCompare(String(a.invoice_no));
  });

  return { orders: orders.map(function (o) { return serializeOrder_(o, customerMap, menuMap); }) };
}

function getOrder_(p) {
  var inv = String(p.invoice_no || '').trim();
  if (!inv) throw appError_('VALIDATION', 'Nomor nota wajib diisi.');
  var order = null;
  readTable_('Orders').some(function (o) {
    if (String(o.invoice_no) === inv) { order = o; return true; }
    return false;
  });
  if (!order) throw appError_('NOT_FOUND', 'Nota ' + inv + ' tidak ditemukan.');

  var customerMap = getCustomerMap_();
  var menuMap = getMenuMap_();
  return {
    order: serializeOrder_(order, customerMap, menuMap),
    customer: customerMap[order.customer_id] ? serializeCustomer_(customerMap[order.customer_id]) : null,
    menu1: menuMap[order.menu1_id] || null,
    menu2: order.menu2_id ? (menuMap[order.menu2_id] || null) : null,
    settings: getSettings_()
  };
}

/* ------------------------------------------------------------------ */
/* Orders                                                              */
/* ------------------------------------------------------------------ */

function createOrder_(p) {
  var date = parseDate_(p.delivery_date);
  if (!date) throw appError_('VALIDATION', 'Tanggal antar tidak valid.');
  var qty = parseInt(p.qty, 10);
  if (!qty || qty < 1) throw appError_('VALIDATION', 'Jumlah porsi minimal 1.');

  var menuMap = getMenuMap_();
  var m1 = menuMap[p.menu1_id];
  if (!m1) throw appError_('VALIDATION', 'Menu utama tidak valid.');
  var m2 = p.menu2_id ? menuMap[p.menu2_id] : null;
  if (p.menu2_id && !m2) throw appError_('VALIDATION', 'Menu campuran tidak valid.');
  var pair = sortMenuIds_(m1.menu_id, m2 ? m2.menu_id : '', menuMap);

  var status = p.payment_status === 'Lunas' ? 'Lunas' : 'Belum Lunas';
  var method = normalizeMethod_(p.payment_method);
  if (status === 'Lunas' && !method) throw appError_('VALIDATION', 'Metode bayar wajib diisi untuk pesanan lunas.');
  if (status !== 'Lunas') method = '';

  var unitPrice = Number(m1.price) || 0;
  var unitCost = Number(m1.cost) || 0;

  return withLock_(function () {
    var customer = resolveCustomerLocked_(p);
    var invoice = nextInvoiceNoLocked_(date);
    var now = nowIso_();
    var obj = {
      invoice_no: invoice,
      delivery_date: date,
      customer_id: customer.customer_id,
      customer_name: customer.name,
      menu1_id: pair[0],
      menu2_id: pair[1],
      qty: qty,
      unit_price: unitPrice,
      unit_cost: unitCost,
      line_total: qty * unitPrice,
      margin: qty * (unitPrice - unitCost),
      payment_status: status,
      payment_method: method,
      order_status: 'Baru',
      notes: String(p.notes || ''),
      paid_at: status === 'Lunas' ? now : '',
      created_at: now,
      updated_at: now
    };
    appendRows_('Orders', [objectToRow_('Orders', obj)]);
    var cmap = {};
    cmap[customer.customer_id] = customer;
    return { order: serializeOrder_(obj, cmap, menuMap) };
  });
}

function updateOrder_(p) {
  var inv = String(p.invoice_no || '').trim();
  if (!inv) throw appError_('VALIDATION', 'Nomor nota wajib diisi.');

  return withLock_(function () {
    var rowIndex = findRowIndex_('Orders', 'invoice_no', inv);
    if (rowIndex < 0) throw appError_('NOT_FOUND', 'Nota ' + inv + ' tidak ditemukan.');
    var existing = null;
    readTable_('Orders').some(function (o) {
      if (String(o.invoice_no) === inv) { existing = o; return true; }
      return false;
    });
    if (!existing) throw appError_('NOT_FOUND', 'Nota ' + inv + ' tidak ditemukan.');

    var menuMap = getMenuMap_();
    var menu1Id = p.menu1_id || existing.menu1_id;
    var menu2Id = p.menu2_id === '' ? '' : (p.menu2_id || existing.menu2_id);
    var m1 = menuMap[menu1Id];
    if (!m1) throw appError_('VALIDATION', 'Menu utama tidak valid.');
    var m2 = menu2Id ? menuMap[menu2Id] : null;
    if (menu2Id && !m2) throw appError_('VALIDATION', 'Menu campuran tidak valid.');
    var pair = sortMenuIds_(m1.menu_id, m2 ? m2.menu_id : '', menuMap);

    var date = p.delivery_date ? parseDate_(p.delivery_date) : existing.delivery_date;
    if (!date) throw appError_('VALIDATION', 'Tanggal antar tidak valid.');
    var qty = p.qty !== undefined ? parseInt(p.qty, 10) : parseInt(existing.qty, 10);
    if (!qty || qty < 1) throw appError_('VALIDATION', 'Jumlah porsi minimal 1.');

    var status = p.payment_status === 'Lunas' ? 'Lunas' : (p.payment_status === 'Belum Lunas' ? 'Belum Lunas' : existing.payment_status);
    var method = normalizeMethod_(p.payment_method !== undefined ? p.payment_method : existing.payment_method);
    if (status === 'Lunas' && !method) throw appError_('VALIDATION', 'Metode bayar wajib diisi untuk pesanan lunas.');
    if (status !== 'Lunas') method = '';

    var customer = null;
    if (p.customer_id && p.customer_id !== existing.customer_id) {
      customer = getCustomerMap_()[p.customer_id];
      if (!customer) throw appError_('NOT_FOUND', 'Pelanggan tidak ditemukan.');
    }

    var unitPrice = Number(m1.price) || 0;
    var unitCost = Number(m1.cost) || 0;
    var now = nowIso_();
    var obj = {
      invoice_no: inv,
      delivery_date: date,
      customer_id: customer ? customer.customer_id : existing.customer_id,
      customer_name: customer ? customer.name : existing.customer_name,
      menu1_id: pair[0],
      menu2_id: pair[1],
      qty: qty,
      unit_price: unitPrice,
      unit_cost: unitCost,
      line_total: qty * unitPrice,
      margin: qty * (unitPrice - unitCost),
      payment_status: status,
      payment_method: method,
      order_status: p.order_status === 'Batal' ? 'Batal' : (p.order_status === 'Selesai' ? 'Selesai' : (p.order_status === 'Baru' ? 'Baru' : existing.order_status)),
      notes: p.notes !== undefined ? String(p.notes) : existing.notes,
      paid_at: status === 'Lunas' ? (existing.paid_at || now) : '',
      created_at: existing.created_at,
      updated_at: now
    };
    writeObject_('Orders', rowIndex, obj);
    var cmap = getCustomerMap_();
    return { order: serializeOrder_(obj, cmap, menuMap) };
  });
}

function cancelOrder_(p) {
  var inv = String(p.invoice_no || '').trim();
  if (!inv) throw appError_('VALIDATION', 'Nomor nota wajib diisi.');
  return withLock_(function () {
    var rowIndex = findRowIndex_('Orders', 'invoice_no', inv);
    if (rowIndex < 0) throw appError_('NOT_FOUND', 'Nota ' + inv + ' tidak ditemukan.');
    var existing = null;
    readTable_('Orders').some(function (o) {
      if (String(o.invoice_no) === inv) { existing = o; return true; }
      return false;
    });
    if (!existing) throw appError_('NOT_FOUND', 'Nota ' + inv + ' tidak ditemukan.');
    existing.order_status = 'Batal';
    existing.updated_at = nowIso_();
    writeObject_('Orders', rowIndex, existing);
    var menuMap = getMenuMap_();
    return { order: serializeOrder_(existing, getCustomerMap_(), menuMap) };
  });
}

function nextInvoiceNoLocked_(date) {
  var settings = getSettings_();
  var prefix = settings.invoice_prefix || 'NCD';
  var dateKey = String(date).replace(/-/g, '').slice(2);
  var seq = parseInt(settings['invoice_seq:' + dateKey] || '0', 10) || 0;

  readTable_('Orders').forEach(function (o) {
    var m = String(o.invoice_no || '').match(/^(.+)-(\d{6})-(\d+)$/);
    if (m && m[2] === dateKey) {
      seq = Math.max(seq, parseInt(m[3], 10) || 0);
    }
  });

  seq += 1;
  setSettingLocked_('invoice_seq:' + dateKey, String(seq));
  return prefix + '-' + dateKey + '-' + pad_(seq, 3);
}

function resolveCustomerLocked_(p) {
  var customers = readTable_('Customers');
  var customerId = String(p.customer_id || '').trim();
  if (customerId) {
    var byId = null;
    customers.some(function (c) {
      if (c.customer_id === customerId) { byId = c; return true; }
      return false;
    });
    if (!byId) throw appError_('NOT_FOUND', 'Pelanggan tidak ditemukan.');
    return byId;
  }

  var name = String(p.customer_name || '').trim();
  var phone = normalizePhone_(p.whatsapp);
  var nameKey = normalizeNameKey_(name);
  var found = null;

  if (phone) {
    customers.some(function (c) {
      if (c.whatsapp === phone && !c.merged_into) { found = c; return true; }
      return false;
    });
    if (found) return found;
  }

  if (nameKey) {
    customers.some(function (c) {
      var aliases = String(c.aliases || '').split('|');
      if (!c.merged_into && (c.name_key === nameKey || aliases.indexOf(nameKey) >= 0)) {
        found = c;
        return true;
      }
      return false;
    });
    if (found) {
      if (phone && !found.whatsapp) {
        var conflict = false;
        customers.forEach(function (c) {
          if (c.customer_id !== found.customer_id && c.whatsapp === phone) conflict = true;
        });
        if (!conflict) {
          var idx = findRowIndex_('Customers', 'customer_id', found.customer_id);
          found.whatsapp = phone;
          found.customer_key = 'wa:' + phone;
          found.updated_at = nowIso_();
          writeObject_('Customers', idx, found);
        }
      }
      return found;
    }
  }

  if (!name) name = 'Tanpa Nama';
  if (!nameKey) nameKey = normalizeNameKey_(name) || 'tanpa nama';
  var now = nowIso_();
  var cust = {
    customer_id: nextCustomerIdLocked_(customers),
    name: name,
    name_key: nameKey,
    whatsapp: phone,
    customer_key: phone ? ('wa:' + phone) : ('name:' + nameKey),
    aliases: '',
    merged_into: '',
    notes: '',
    active: true,
    created_at: now,
    updated_at: now
  };
  appendRows_('Customers', [objectToRow_('Customers', cust)]);
  return cust;
}

/* ------------------------------------------------------------------ */
/* Customers                                                           */
/* ------------------------------------------------------------------ */

function listCustomers_(p) {
  var q = String(p.q || '').trim().toLowerCase();
  var customers = readTable_('Customers');
  var orders = readTable_('Orders').filter(function (o) { return o.order_status !== 'Batal'; });

  var stats = {};
  orders.forEach(function (o) {
    if (!stats[o.customer_id]) {
      stats[o.customer_id] = { count: 0, last: '', total: 0 };
    }
    stats[o.customer_id].count += 1;
    stats[o.customer_id].total += Number(o.line_total) || 0;
    if (String(o.delivery_date) > String(stats[o.customer_id].last)) {
      stats[o.customer_id].last = o.delivery_date;
    }
  });

  var out = customers.map(function (c) {
    var s = stats[c.customer_id] || { count: 0, last: '', total: 0 };
    return {
      customer_id: c.customer_id,
      name: c.name,
      name_key: c.name_key,
      whatsapp: c.whatsapp,
      customer_key: c.customer_key,
      aliases: c.aliases,
      merged_into: c.merged_into,
      notes: c.notes,
      active: toBool_(c.active),
      order_count: s.count,
      last_order_date: s.last,
      total_spent: s.total
    };
  }).filter(function (c) {
    if (!q) return true;
    return String(c.name).toLowerCase().indexOf(q) >= 0 ||
      String(c.whatsapp).indexOf(q) >= 0;
  });

  out.sort(function (a, b) {
    if (a.active !== b.active) return a.active ? -1 : 1;
    return String(a.name).localeCompare(String(b.name));
  });
  return { customers: out };
}

function saveCustomer_(p) {
  var name = String(p.name || '').trim();
  if (!name) throw appError_('VALIDATION', 'Nama pelanggan wajib diisi.');
  var phone = normalizePhone_(p.whatsapp);
  var nameKey = normalizeNameKey_(name);

  return withLock_(function () {
    var customers = readTable_('Customers');
    var id = String(p.customer_id || '').trim();
    var now = nowIso_();

    if (id) {
      var rowIndex = findRowIndex_('Customers', 'customer_id', id);
      if (rowIndex < 0) throw appError_('NOT_FOUND', 'Pelanggan tidak ditemukan.');
      var existing = null;
      customers.some(function (c) {
        if (c.customer_id === id) { existing = c; return true; }
        return false;
      });
      if (phone) {
        customers.forEach(function (c) {
          if (c.customer_id !== id && c.whatsapp === phone && !c.merged_into) {
            throw appError_('VALIDATION', 'Nomor WhatsApp sudah dipakai pelanggan lain.');
          }
        });
      }
      existing.name = name;
      existing.name_key = nameKey;
      existing.whatsapp = phone;
      existing.customer_key = phone ? ('wa:' + phone) : ('name:' + nameKey);
      existing.notes = p.notes !== undefined ? String(p.notes) : existing.notes;
      existing.active = p.active !== undefined ? toBool_(p.active) : toBool_(existing.active);
      existing.updated_at = now;
      writeObject_('Customers', rowIndex, existing);
      return { customer: serializeCustomer_(existing), existing: true };
    }

    var found = null;
    customers.some(function (c) {
      if (c.merged_into) return false;
      if (phone && c.whatsapp === phone) { found = c; return true; }
      if (!phone && c.name_key === nameKey) { found = c; return true; }
      return false;
    });

    if (found) {
      if (phone && !found.whatsapp) {
        var conflict = false;
        customers.forEach(function (c) {
          if (c.customer_id !== found.customer_id && c.whatsapp === phone) conflict = true;
        });
        if (!conflict) {
          found.whatsapp = phone;
          found.customer_key = 'wa:' + phone;
        }
      }
      var aliases = String(found.aliases || '').split('|').filter(Boolean);
      if (nameKey && aliases.indexOf(nameKey) < 0 && found.name_key !== nameKey) aliases.push(nameKey);
      found.aliases = aliases.join('|');
      found.notes = p.notes !== undefined ? String(p.notes) : found.notes;
      found.updated_at = now;
      writeObject_('Customers', findRowIndex_('Customers', 'customer_id', found.customer_id), found);
      return { customer: serializeCustomer_(found), existing: true };
    }

    var cust = {
      customer_id: nextCustomerIdLocked_(customers),
      name: name,
      name_key: nameKey,
      whatsapp: phone,
      customer_key: phone ? ('wa:' + phone) : ('name:' + nameKey),
      aliases: '',
      merged_into: '',
      notes: String(p.notes || ''),
      active: true,
      created_at: now,
      updated_at: now
    };
    appendRows_('Customers', [objectToRow_('Customers', cust)]);
    return { customer: serializeCustomer_(cust), existing: false };
  });
}

function mergeCustomers_(p) {
  var sourceId = String(p.source_id || '').trim();
  var targetId = String(p.target_id || '').trim();
  if (!sourceId || !targetId) throw appError_('VALIDATION', 'Pelanggan sumber dan tujuan wajib diisi.');
  if (sourceId === targetId) throw appError_('VALIDATION', 'Pilih dua pelanggan yang berbeda.');

  return withLock_(function () {
    var customers = readTable_('Customers');
    var source = null;
    var target = null;
    customers.forEach(function (c) {
      if (c.customer_id === sourceId) source = c;
      if (c.customer_id === targetId) target = c;
    });
    if (!source || !target) throw appError_('NOT_FOUND', 'Pelanggan tidak ditemukan.');

    var orders = readTable_('Orders');
    var now = nowIso_();
    orders.forEach(function (o) {
      if (o.customer_id === sourceId) {
        o.customer_id = targetId;
        o.customer_name = target.name;
        o.updated_at = now;
        writeObject_('Orders', findRowIndex_('Orders', 'invoice_no', o.invoice_no), o);
      }
    });

    var aliases = String(target.aliases || '').split('|').filter(Boolean);
    String(source.name_key + '|' + (source.aliases || '')).split('|').forEach(function (a) {
      a = String(a || '').trim();
      if (a && aliases.indexOf(a) < 0) aliases.push(a);
    });
    target.aliases = aliases.join('|');
    target.updated_at = now;
    writeObject_('Customers', findRowIndex_('Customers', 'customer_id', targetId), target);

    source.merged_into = targetId;
    source.active = false;
    source.updated_at = now;
    writeObject_('Customers', findRowIndex_('Customers', 'customer_id', sourceId), source);

    return { customer: serializeCustomer_(target) };
  });
}

/* ------------------------------------------------------------------ */
/* Menu                                                                */
/* ------------------------------------------------------------------ */

function listMenu_() {
  var menu = readTable_('Menu').map(function (m) {
    return {
      menu_id: m.menu_id,
      name: m.name,
      short_name: m.short_name,
      price: Number(m.price) || 0,
      cost: Number(m.cost) || 0,
      active: toBool_(m.active),
      sort_order: Number(m.sort_order) || 0
    };
  }).sort(function (a, b) { return a.sort_order - b.sort_order; });
  return { menu: menu };
}

function saveMenu_(p) {
  var id = String(p.menu_id || '').trim();
  var name = String(p.name || '').trim();
  var shortName = String(p.short_name || '').trim();
  var price = parseInt(p.price, 10);
  var cost = parseInt(p.cost, 10);
  var sort = parseInt(p.sort_order, 10);
  if (!id) throw appError_('VALIDATION', 'ID menu wajib diisi.');
  if (!name || !shortName) throw appError_('VALIDATION', 'Nama dan nama pendek wajib diisi.');
  if (!price || price < 1) throw appError_('VALIDATION', 'Harga jual tidak valid.');
  if (!(cost >= 0)) throw appError_('VALIDATION', 'Harga modal tidak valid.');
  if (!(sort >= 0)) sort = 0;

  return withLock_(function () {
    var obj = {
      menu_id: id,
      name: name,
      short_name: shortName,
      price: price,
      cost: cost,
      active: p.active !== undefined ? toBool_(p.active) : true,
      sort_order: sort
    };
    var rowIndex = findRowIndex_('Menu', 'menu_id', id);
    if (rowIndex > 0) {
      writeObject_('Menu', rowIndex, obj);
    } else {
      appendRows_('Menu', [objectToRow_('Menu', obj)]);
    }
    return { menu: obj };
  });
}

/* ------------------------------------------------------------------ */
/* Expenses                                                            */
/* ------------------------------------------------------------------ */

function listExpenses_(p) {
  var from = parseDate_(p.from) || '';
  var to = parseDate_(p.to) || '';
  var expenses = readTable_('Expenses').filter(function (e) {
    if (from && String(e.expense_date) < from) return false;
    if (to && String(e.expense_date) > to) return false;
    return true;
  }).map(function (e) {
    return {
      expense_id: e.expense_id,
      expense_date: e.expense_date,
      category: e.category,
      description: e.description,
      amount: Number(e.amount) || 0,
      payment_method: e.payment_method || '',
      notes: e.notes || '',
      created_at: e.created_at,
      updated_at: e.updated_at
    };
  }).sort(function (a, b) {
    return String(b.expense_date).localeCompare(String(a.expense_date));
  });
  return { expenses: expenses };
}

function saveExpense_(p) {
  var date = parseDate_(p.expense_date);
  var category = String(p.category || '').trim();
  var description = String(p.description || '').trim();
  var amount = parseMoney_(p.amount);
  if (!date) throw appError_('VALIDATION', 'Tanggal pengeluaran tidak valid.');
  if (['Beli Stok', 'Operasional', 'Lainnya'].indexOf(category) < 0) {
    throw appError_('VALIDATION', 'Kategori pengeluaran tidak valid.');
  }
  if (!description) throw appError_('VALIDATION', 'Keterangan pengeluaran wajib diisi.');
  if (amount <= 0) throw appError_('VALIDATION', 'Nominal pengeluaran harus lebih dari 0.');

  return withLock_(function () {
    var expenses = readTable_('Expenses');
    var id = String(p.expense_id || '').trim();
    var now = nowIso_();
    var obj;
    if (id) {
      var rowIndex = findRowIndex_('Expenses', 'expense_id', id);
      if (rowIndex < 0) throw appError_('NOT_FOUND', 'Pengeluaran tidak ditemukan.');
      var existing = null;
      expenses.some(function (e) {
        if (e.expense_id === id) { existing = e; return true; }
        return false;
      });
      obj = {
        expense_id: id,
        expense_date: date,
        category: category,
        description: description,
        amount: amount,
        payment_method: normalizeMethod_(p.payment_method),
        notes: p.notes !== undefined ? String(p.notes) : existing.notes,
        created_at: existing.created_at,
        updated_at: now
      };
      writeObject_('Expenses', rowIndex, obj);
    } else {
      obj = {
        expense_id: nextExpenseIdLocked_(expenses),
        expense_date: date,
        category: category,
        description: description,
        amount: amount,
        payment_method: normalizeMethod_(p.payment_method),
        notes: String(p.notes || ''),
        created_at: now,
        updated_at: now
      };
      appendRows_('Expenses', [objectToRow_('Expenses', obj)]);
    }
    return { expense: obj };
  });
}

/* ------------------------------------------------------------------ */
/* Reports / settings                                                  */
/* ------------------------------------------------------------------ */

function report_(p) {
  var period = p.period === 'monthly' ? 'monthly' : 'daily';
  var from = parseDate_(p.from) || '';
  var to = parseDate_(p.to) || '';
  var month = String(p.month || '').trim();
  if (!from && !to && !month) month = todayIso_().slice(0, 7);

  function keyFor(date) {
    return period === 'monthly' ? String(date).slice(0, 7) : String(date);
  }
  function inRange(date) {
    if (!date) return false;
    if (month) return String(date).slice(0, 7) === month;
    if (from && String(date) < from) return false;
    if (to && String(date) > to) return false;
    return true;
  }

  var rowsMap = {};
  function ensureRow(key) {
    if (!rowsMap[key]) {
      rowsMap[key] = { key: key, portions: 0, omzet: 0, margin: 0, expenses: 0, paid: 0 };
    }
    return rowsMap[key];
  }

  readTable_('Orders').forEach(function (o) {
    if (o.order_status === 'Batal' || !inRange(o.delivery_date)) return;
    var r = ensureRow(keyFor(o.delivery_date));
    r.portions += Number(o.qty) || 0;
    r.omzet += Number(o.line_total) || 0;
    r.margin += Number(o.margin) || 0;
    if (o.payment_status === 'Lunas') r.paid += Number(o.line_total) || 0;
  });

  readTable_('Expenses').forEach(function (e) {
    if (!inRange(e.expense_date)) return;
    var r = ensureRow(keyFor(e.expense_date));
    r.expenses += Number(e.amount) || 0;
  });

  var rows = Object.keys(rowsMap).sort().map(function (k) {
    var r = rowsMap[k];
    r.label = period === 'monthly' ? monthLabel_(k) : dayLabel_(k);
    r.net_cash = r.paid - r.expenses;
    return r;
  });

  var totals = { portions: 0, omzet: 0, margin: 0, expenses: 0, paid: 0, net_cash: 0 };
  rows.forEach(function (r) {
    totals.portions += r.portions;
    totals.omzet += r.omzet;
    totals.margin += r.margin;
    totals.expenses += r.expenses;
    totals.paid += r.paid;
    totals.net_cash += r.net_cash;
  });

  return { period: period, rows: rows, totals: totals };
}

function saveSettings_(p) {
  return withLock_(function () {
    Object.keys(p || {}).forEach(function (key) {
      if (key === 'action' || key === 'payload') return;
      var v = p[key];
      if (v === undefined || v === null) return;
      setSettingLocked_(key, String(v));
    });
    return { settings: getSettings_() };
  });
}

/* ------------------------------------------------------------------ */
/* Serialization                                                       */
/* ------------------------------------------------------------------ */

function serializeOrder_(o, customerMap, menuMap) {
  var c = customerMap && customerMap[o.customer_id];
  return {
    invoice_no: o.invoice_no,
    delivery_date: o.delivery_date,
    customer_id: o.customer_id,
    customer_name: o.customer_name,
    customer_whatsapp: c ? (c.whatsapp || '') : '',
    menu1_id: o.menu1_id,
    menu1_name: menuMap[o.menu1_id] ? menuMap[o.menu1_id].name : o.menu1_id,
    menu2_id: o.menu2_id || '',
    menu2_name: o.menu2_id && menuMap[o.menu2_id] ? menuMap[o.menu2_id].name : '',
    menu_label: menuLabel_(o.menu1_id, o.menu2_id, menuMap),
    qty: Number(o.qty) || 0,
    unit_price: Number(o.unit_price) || 0,
    unit_cost: Number(o.unit_cost) || 0,
    line_total: Number(o.line_total) || 0,
    margin: Number(o.margin) || 0,
    payment_status: o.payment_status,
    payment_method: o.payment_method || '',
    order_status: o.order_status,
    notes: o.notes || '',
    paid_at: o.paid_at || '',
    created_at: o.created_at || '',
    updated_at: o.updated_at || ''
  };
}

function serializeCustomer_(c) {
  return {
    customer_id: c.customer_id,
    name: c.name,
    name_key: c.name_key,
    whatsapp: c.whatsapp || '',
    customer_key: c.customer_key || '',
    aliases: c.aliases || '',
    merged_into: c.merged_into || '',
    notes: c.notes || '',
    active: toBool_(c.active),
    created_at: c.created_at || '',
    updated_at: c.updated_at || ''
  };
}



/* ------------------------------------------------------------------ */
/* Legacy import                                                       */
/* ------------------------------------------------------------------ */

function importLegacy_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tb = ss.getSheetByName('Transaksi Baru');
  var lp = ss.getSheetByName('List Pesanan');
  var pe = ss.getSheetByName('Pengeluaran');

  if (!tb && !lp && !pe) {
    throw appError_('VALIDATION', 'Sheet sumber impor (Transaksi Baru / List Pesanan / Pengeluaran) tidak ditemukan. Konversi NASI COKOT DAISAFA.xlsx ke Google Sheets dulu.');
  }

  if (getSetting_('legacy_import_done') === 'true') {
    return {
      already: true,
      counts: { customers: 0, orders: 0, expenses: 0, duplicates: 0, skipped: 0 },
      warnings: []
    };
  }

  var existingOrders = readTable_('Orders');
  var hasImport = false;
  existingOrders.forEach(function (o) {
    if (String(o.notes || '').indexOf('Impor:') === 0) hasImport = true;
  });
  if (hasImport) {
    throw appError_('VALIDATION', 'Impor sebelumnya terdeteksi pada sheet Orders. Kosongkan data hasil impor atau set legacy_import_done=true dulu.');
  }

  return withLock_(function () {
    var today = todayIso_();
    var menuMap = getMenuMap_();

    var ordersOut = [];
    var expensesOut = [];
    var customersOut = [];
    var logs = [];
    var dedupe = {};
    var dedupeExpense = {};
    var counts = { customers: 0, orders: 0, expenses: 0, duplicates: 0, skipped: 0 };

    function logIssue_(source, row, level, issue, raw) {
      logs.push({
        logged_at: nowIso_(),
        source: source,
        source_row: row,
        level: level,
        issue: issue,
        raw: raw ? JSON.stringify(raw).slice(0, 500) : ''
      });
    }

    function trackCustomer(c, key) {
      if (!key) return;
      var seen = customersOut.map(function (r) { return String(r[1]); });
      for (var i = 0; i < seen.length; i++) {
        if (seen[i] === key) return;
      }
      customersOut.push(objectToRow_('Customers', c));
    }

    var customersAll = readTable_('Customers');
    var customerKeys = {};
    customersAll.forEach(function (c) { customerKeys[c.customer_key] = true; });

    /* ---------------- Transaksi Baru ---------------- */
    if (tb) {
      var tbVals = tb.getDataRange().getValues();
      for (var i = 1; i < tbVals.length; i++) {
        var row = tbVals[i];
        var isEmpty = true;
        for (var q = 0; q < row.length; q++) {
          if (String(row[q] || '').trim() !== '') { isEmpty = false; break; }
        }
        if (isEmpty) continue;

        var date = parseDate_(row[2]);
        if (!date) { logIssue_('Transaksi Baru', i + 1, 'SKIP', 'Tanggal tidak valid', row); counts.skipped += 1; continue; }
        var menu1 = normalizeFlavor_(row[4]);
        var menu2 = normalizeFlavor_(row[5]);
        if (!menu1) { logIssue_('Transaksi Baru', i + 1, 'SKIP', 'Rasa utama tidak dikenal', row); counts.skipped += 1; continue; }

        var qty = parseInt(row[6], 10);
        if (!qty || qty < 1) {
          var hj = parseMoney_(row[7]);
          qty = (hj > 0 && hj % 7000 === 0) ? (hj / 7000) : 1;
          logIssue_('Transaksi Baru', i + 1, 'WARN', 'Qty dihitung dari Harga Jual', row);
        }

        var name = String(row[3] || '').trim();
        if (!name) { name = 'Tanpa Nama'; logIssue_('Transaksi Baru', i + 1, 'WARN', 'Nama kosong', row); }
        var customer = resolveCustomerLocked_({ customer_name: name });
        trackCustomer(customer, customer.customer_key);

        var status = String(row[11] || '').trim().toLowerCase() === 'paid' ? 'Lunas' : 'Belum Lunas';
        var method = normalizeMethod_(row[12]);
        if (status === 'Lunas' && !method) logIssue_('Transaksi Baru', i + 1, 'WARN', 'Lunas tanpa metode', row);

        var pair = sortMenuIds_(menu1, menu2, menuMap);
        var lineTotal = qty * 7000;
        var key = [date, customer.customer_key, pair[0], pair[1] || '', qty, status, method].join('|');
        if (dedupe[key]) {
          counts.duplicates += 1;
          logIssue_('Transaksi Baru', i + 1, 'DUP', 'Duplikat dengan order sebelumnya', row);
          continue;
        }
        dedupe[key] = true;

        ordersOut.push({
          invoice_no: '',
          delivery_date: date,
          customer_id: customer.customer_id,
          customer_name: customer.name,
          menu1_id: pair[0],
          menu2_id: pair[1],
          qty: qty,
          unit_price: 7000,
          unit_cost: 6000,
          line_total: lineTotal,
          margin: qty * 1000,
          payment_status: status,
          payment_method: method,
          order_status: date < today ? 'Selesai' : 'Baru',
          notes: 'Impor: Transaksi Baru #' + (i + 1),
          paid_at: status === 'Lunas' ? (date + ' 12:00:00') : '',
          created_at: nowIso_(),
          updated_at: nowIso_(),
          source: 'Transaksi Baru',
          source_row: i + 1
        });
        counts.orders += 1;
      }
    }

    /* ---------------- List Pesanan ---------------- */
    if (lp) {
      var lpVals = lp.getDataRange().getValues();
      var lpRows = [];
      var lastDate = '';
      for (var i2 = 5; i2 < lpVals.length; i2++) {
        var row2 = lpVals[i2];
        var isEmpty2 = true;
        for (var q2 = 0; q2 < row2.length; q2++) {
          if (String(row2[q2] || '').trim() !== '') { isEmpty2 = false; break; }
        }
        if (isEmpty2) continue;

        var d2 = parseDate_(row2[6]);
        if (d2) lastDate = d2;
        var name2 = String(row2[1] || '').trim();
        var rasa1 = String(row2[2] || '');
        var rasa2 = String(row2[3] || '');
        var combined = String(row2[4] || '');
        if (!name2 && !rasa1 && !combined) continue;

        var menu1l = normalizeFlavor_(rasa1);
        var menu2l = rasa2 ? normalizeFlavor_(rasa2) : '';
        if (!menu1l && combined) {
          var parts = combined.split('-').map(function (s) { return normalizeFlavor_(s); }).filter(Boolean);
          if (parts.length === 2) { menu1l = parts[0]; menu2l = parts[1]; }
          else if (parts.length === 1) { menu1l = parts[0]; menu2l = ''; }
        }
        if (!menu1l) { logIssue_('List Pesanan', i2 + 1, 'SKIP', 'Rasa tidak dikenal', row2); counts.skipped += 1; continue; }

        var date2 = lastDate;
        if (!date2) { logIssue_('List Pesanan', i2 + 1, 'SKIP', 'Tanggal belum diisi', row2); counts.skipped += 1; continue; }

        lpRows.push({
          source_row: i2 + 1,
          date: date2,
          name: name2,
          menu1: menu1l,
          menu2: menu2l,
          payment: parseMoney_(row2[8]),
          status: String(row2[9] || '').trim() === '' ? 'Belum Lunas' : (String(row2[9] || '').trim().toLowerCase() === 'paid' ? 'Lunas' : 'Belum Lunas'),
          method: methodFromStatus_(row2[9])
        });
      }

      var groupsByDate = {};
      var groupOrder = [];
      lpRows.forEach(function (r) {
        if (!groupsByDate[r.date]) { groupsByDate[r.date] = []; groupOrder.push(r.date); }
        groupsByDate[r.date].push(r);
      });
      Object.keys(groupsByDate).sort().forEach(function (date) {
        var group = groupsByDate[date];
        var paidAmount = 0;
        group.forEach(function (r) { paidAmount += Number(r.payment) || 0; });
        var paidQty = Math.round(paidAmount / 7000);
        var methodVal = '';
        group.forEach(function (r) { if (!methodVal && r.method) methodVal = r.method; });
        var paidRemaining = Math.min(paidQty, group.length);
        group.forEach(function (r) {
          var status = 'Belum Lunas';
          if (paidRemaining > 0) { status = 'Lunas'; paidRemaining -= 1; }
          var customer = resolveCustomerLocked_({ customer_name: r.name });
          trackCustomer(customer, customer.customer_key);
          var pair = sortMenuIds_(r.menu1, r.menu2, menuMap);
          var key = [date, customer.customer_key, pair[0], pair[1] || '', 1, status, methodVal].join('|');
          if (dedupe[key]) {
            counts.duplicates += 1;
            logIssue_('List Pesanan', r.source_row, 'DUP', 'Duplikat dengan order sebelumnya', r);
            return;
          }
          dedupe[key] = true;
          ordersOut.push({
            invoice_no: '',
            delivery_date: date,
            customer_id: customer.customer_id,
            customer_name: customer.name,
            menu1_id: pair[0],
            menu2_id: pair[1],
            qty: 1,
            unit_price: 7000,
            unit_cost: 6000,
            line_total: 7000,
            margin: 1000,
            payment_status: status,
            payment_method: methodVal,
            order_status: date < today ? 'Selesai' : 'Baru',
            notes: 'Impor: List Pesanan #' + r.source_row,
            paid_at: status === 'Lunas' ? (date + ' 12:00:00') : '',
            created_at: nowIso_(),
            updated_at: nowIso_(),
            source: 'List Pesanan',
            source_row: r.source_row
          });
          counts.orders += 1;
        });
        if (paidQty > group.length) {
          logIssue_('List Pesanan', group[0].source_row, 'WARN', 'Total pembayaran melebihi jumlah baris', { paidQty: paidQty, rows: group.length });
        }
      });
    }

    /* ---------------- Pengeluaran ---------------- */
    if (pe) {
      var peVals = pe.getDataRange().getValues();
      var dedupeExpense = {};
      for (var i3 = 1; i3 < peVals.length; i3++) {
        var row3 = peVals[i3];
        var isEmpty3 = true;
        for (var q3 = 0; q3 < row3.length; q3++) {
          if (String(row3[q3] || '').trim() !== '') { isEmpty3 = false; break; }
        }
        if (isEmpty3) continue;
        var date3 = parseDate_(row3[2]);
        var desc3 = String(row3[3] || '').trim();
        var amount3 = parseMoney_(row3[4]);
        if (!date3) { logIssue_('Pengeluaran', i3 + 1, 'SKIP', 'Tanggal tidak valid', row3); counts.skipped += 1; continue; }
        if (!desc3 && !amount3) continue;
        var ekey = [date3, desc3, amount3].join('|');
        if (dedupeExpense[ekey]) continue;
        dedupeExpense[ekey] = true;
        expensesOut.push({
          expense_id: '',
          expense_date: date3,
          category: 'Beli Stok',
          description: desc3,
          amount: amount3,
          payment_method: '',
          notes: 'Impor: Pengeluaran #' + (i3 + 1),
          created_at: nowIso_(),
          updated_at: nowIso_()
        });
        counts.expenses += 1;
      }
    }

    /* ---------------- Assign invoice numbers ---------------- */
    var seqMap = {};
    var prefix = getSetting_('invoice_prefix') || 'NCD';
    readTable_('Orders').forEach(function (o) {
      var m = String(o.invoice_no || '').match(/^(.+)-(\d{6})-(\d+)$/);
      if (m) seqMap[m[2]] = Math.max(seqMap[m[2]] || 0, parseInt(m[3], 10) || 0);
    });
    ordersOut.sort(function (a, b) {
      var d = String(a.delivery_date).localeCompare(String(b.delivery_date));
      if (d !== 0) return d;
      return String(a.source).localeCompare(String(b.source));
    });
    ordersOut.forEach(function (o) {
      var dk = String(o.delivery_date).replace(/-/g, '').slice(2);
      seqMap[dk] = (seqMap[dk] || 0) + 1;
      o.invoice_no = prefix + '-' + dk + '-' + pad_(seqMap[dk], 3);
    });
    Object.keys(seqMap).forEach(function (dk) {
      setSettingLocked_('invoice_seq:' + dk, String(seqMap[dk]));
    });

    /* ---------------- Write ---------------- */
    if (customersOut.length) {
      var customerObjs = customersOut.map(function (r) { return { customer_id: r[0], name: r[1], name_key: r[2], whatsapp: r[3], customer_key: r[4], aliases: r[5], merged_into: r[6], notes: r[7], active: r[8], created_at: r[9], updated_at: r[10] }; });
      appendRows_('Customers', customerObjs.map(function (c) { return objectToRow_('Customers', c); }));
    }
    if (ordersOut.length) {
      appendRows_('Orders', ordersOut.map(function (o) { return objectToRow_('Orders', o); }));
    }
    if (expensesOut.length) {
      var expenseObjs = expensesOut.map(function (e) { return e; });
      var maxExp = 0;
      readTable_('Expenses').forEach(function (e) {
        var m = String(e.expense_id || '').match(/^EXP-(\d+)$/);
        if (m) maxExp = Math.max(maxExp, parseInt(m[1], 10) || 0);
      });
      var nextExp = maxExp;
      expenseObjs.forEach(function (e) {
        nextExp += 1;
        e.expense_id = 'EXP-' + pad_(nextExp, 4);
      });
      appendRows_('Expenses', expenseObjs.map(function (e) { return objectToRow_('Expenses', e); }));
    }
    if (logs.length) {
      appendRows_('Import_Log', logs.map(function (l) {
        return [l.logged_at, l.source, l.source_row, l.level, l.issue, l.raw];
      }));
    }

    setSettingLocked_('legacy_import_done', 'true');

    var warnLogs = logs.filter(function (l) { return l.level === 'WARN' || l.level === 'SKIP'; });
    return {
      already: false,
      counts: {
        customers: counts.customers,
        orders: counts.orders,
        expenses: counts.expenses,
        duplicates: counts.duplicates,
        skipped: counts.skipped
      },
      warnings: warnLogs.map(function (l) { return { source: l.source, source_row: l.source_row, issue: l.issue }; })
    };
  });
}

function importLegacy() {
  return importLegacy_();
}