/**
 * =====================================================================
 *  SRU Secure Access Kit v1.4 — Google Apps Script Web App template
 *  Access: Sulaiman Al Rajhi University Google accounts (@sr.edu.sa) only
 *  ICTD — ITOC
 * ---------------------------------------------------------------------
 *  Files:  Code.gs · Index.html · Logos.html · Denied.html
 *  Deploy: Deploy ▸ New deployment ▸ Web app
 *          Execute as: Me  ·  Who has access: Anyone within Sulaiman Al Rajhi University
 *  Access log: every sign-in → sheet "AccessLog" (see LOG_* settings)
 *  Access list: only active emails in sheet "AllowedUsers" (ACCESS_MODE = 'LIST')
 * =====================================================================
 */

/* ▼▼▼ CONFIG — change per platform ▼▼▼ */
const APP_NAME       = 'جرد أصول البنية التحتية';   // Arabic title shown in the page
const APP_NAME_EN    = 'Infrastructure Inventory';  // English name
const APP_VERSION    = '3.6';
const ALLOWED_DOMAIN = 'sr.edu.sa';
const FAVICON_URL    = 'https://ictsru.github.io/ITOC/assets/itoc-logo.png';

/* Access log */
const LOG_ENABLED        = true;
const LOG_SPREADSHEET_ID = '';            // '' = the spreadsheet this script is bound to; or paste a Sheet ID
const SHEET_LOG          = 'AccessLog';
const LOG_TIMEZONE       = 'Asia/Riyadh';
const LOG_DEDUP_MINUTES  = 30;            // same user within N min = one entry (0 = log every page open)
const LOG_RETENTION_DAYS = 180;           // menu: delete entries older than this
/* Access list (who may open the platform) */
const ACCESS_MODE        = 'LIST';         // 'LIST' = only active emails in the AllowedUsers tab · 'DOMAIN' = any @sr.edu.sa account
const ACCESS_SPREADSHEET_ID = ''           ;  // where the AllowedUsers tab lives
const SHEET_ACCESS       = 'AllowedUsers';
const OWNER_EMAILS       = ['m.elmahdy@sr.edu.sa']; // always allowed, so the owner can never be locked out

/* Default access list — added to AllowedUsers by setupAccessList() (existing rows are kept) */
const DEFAULT_ALLOWED = [
  // [email, name, role, notes]
  ['nabounar@sr.edu.sa',    'Nasser Abounar', 'admin', 'CIO'],
  ['m.elmahdy@sr.edu.sa',   'Mohamed ElMahdy', 'admin', 'Owner — IT Operations Manager'],
  ['n.nawaz@sr.edu.sa',     '', 'user', ''],
  ['m.omar@sr.edu.sa',      '', 'user', ''],
  ['m.saleh@sr.edu.sa',     '', 'user', ''],
  ['a.alrowedan@sr.edu.sa', '', 'user', ''],
  ['n.alrashidi@sr.edu.sa', '', 'user', ''],
  ['a.alawaji@sr.edu.sa',   '', 'user', ''],
  ['a.alaqeel@sr.edu.sa',   '', 'user', ''],
  ['a.alolayan@sr.edu.sa',  '', 'user', ''],
  ['s.albusayli@sr.edu.sa', '', 'user', ''],
  ['i.almutairi@sr.edu.sa', '', 'user', ''],
  ['a.elnadi@sr.edu.sa',    '', 'user', '']
];
const ACCESS_CACHE_SECONDS = 120;          // list changes take effect within 2 minutes
/* ▲▲▲ CONFIG ▲▲▲ */

const LOG_HEADERS = ['timestamp', 'date', 'time', 'day', 'email', 'result', 'version'];
const DAYS_AR_    = ['الأحد', 'الاثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'الجمعة', 'السبت']; // index = ISO day % 7

/* ----------------------------- Web app ------------------------------ */

function doGet() {
  const email = currentEmail_();
  const status = accessStatus_(email);
  if (status !== 'ok') {
    logAccess_(email || '(unknown)', status === 'list' ? 'not-listed' : 'denied');
    const d = HtmlService.createTemplateFromFile('Denied');
    d.appName = APP_NAME; d.appNameEn = APP_NAME_EN;
    d.userEmail = email || '';
    d.reason = status;
    return d.evaluate().setTitle(APP_NAME + ' | غير مصرح')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  logAccess_(email, 'granted');

  const t = HtmlService.createTemplateFromFile('Index');
  t.appName   = APP_NAME;
  t.appNameEn = APP_NAME_EN;
  t.version   = APP_VERSION;
  t.userEmail = email;
  const out = t.evaluate().setTitle(APP_NAME_EN + ' | ' + APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  if (FAVICON_URL) out.setFaviconUrl(FAVICON_URL);
  return out;
}

/* ------------------- Server calls from the page --------------------- */
/*  EVERY function called via google.script.run must start with
 *  requireSruUser_() — the page check alone is not enough.            */

/** Returns every device + supplier list + per-type config. Replaces the old public /exec GET endpoint. */
function inv_list() {
  requireSruUser_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const result = {};
  TYPES.forEach(function (t) {
    const sheet = ss.getSheetByName(t);
    result[t] = sheet ? readSheet(sheet) : [];
  });
  return { success: true, data: result, suppliers: readSuppliers(ss), config: readConfig(ss) };
}

/** add / update / delete one device record. Replaces the old public /exec POST endpoint. */
function inv_write(action, type, record, idForDelete) {
  requireSruUser_();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (TYPES.indexOf(type) === -1) {
      return { success: false, error: 'Unknown device type: ' + type };
    }
    const sheet = getOrCreateSheet(ss, type);

    if (action === 'add' || action === 'update') {
      const r = record;

      // Defensive server-side validation, mirroring the page's own check.
      if (!r || !String(r.tag || '').trim() || !String(r.location || '').trim() || !type) {
        return { success: false, error: 'رقم الأصل والموقع ونوع الجهاز مطلوبة.' };
      }

      // Resolve the photo separately from the row write: if the Drive upload
      // fails for any reason, the device's actual data must still be saved.
      let photoUrl = r.photoUrl || '';
      if (r.photoBase64) {
        try {
          const uploaded = savePhotoIfProvided(r.photoBase64, r.photoMime, r.photoName);
          if (uploaded) photoUrl = uploaded;
        } catch (photoErr) {
          photoUrl = r.photoUrl || '';
        }
      }

      const existingRowIndex = (action === 'update') ? findRowById(sheet, r.id) : -1;
      const savedRowIndex = writeRecord(sheet, existingRowIndex, {
        ID: r.id,
        Tag: r.tag,
        Brand: r.brand,
        Model: r.model,
        Serial: r.serial,
        Supplier: r.supplier,
        Specs: JSON.stringify(r.specs || {}),
        Location: r.location,
        User: r.user,
        Status: r.status,
        Date: r.date,
        Notes: r.notes,
        Warranty: r.warranty,
        WarrantyExpiry: r.warrantyExpiry,
        UpdatedAt: new Date().toISOString(),
        AddedBy: r.addedBy || '',   // staff attribution picked from the page's own list — unrelated to the Google login
        PhotoUrl: photoUrl
      });
      upsertSupplier(ss, r.supplier, r.supplierContactName, r.supplierEmail, r.supplierPhone);
      SpreadsheetApp.flush();

      // Read the row straight back and confirm the ID actually landed where expected.
      const verifyId = sheet.getRange(savedRowIndex, headerMap(sheet)['ID']).getValue();
      if (String(verifyId) !== String(r.id)) {
        return { success: false, error: 'تعذر التحقق من حفظ السجل في الشيت (Row ' + savedRowIndex + ').' };
      }

    } else if (action === 'delete') {
      const rowIndex = findRowById(sheet, idForDelete);
      if (rowIndex > -1) sheet.deleteRow(rowIndex);
      SpreadsheetApp.flush();

    } else {
      return { success: false, error: 'Unknown action: ' + action };
    }

    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/** Scans every device-type sheet for a duplicate ID and gives each later occurrence a fresh, unique one. */
function inv_repairDuplicateIds() {
  requireSruUser_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fixes = repairDuplicateIds(ss);
  return { success: true, count: fixes.length, fixed: fixes };
}

/* ------------------------------ Access ------------------------------ */

function currentEmail_() {
  return String(Session.getActiveUser().getEmail() || '').toLowerCase();
}
function isAllowed_(email) {
  return !!email && email.endsWith('@' + ALLOWED_DOMAIN);
}
function requireSruUser_() {
  const email = currentEmail_();
  const status = accessStatus_(email);
  if (status === 'domain') throw new Error('غير مصرح — الدخول لمنسوبي الجامعة فقط');
  if (status === 'list') throw new Error('غير مصرح — حسابك غير مضاف إلى قائمة المصرّح لهم');
  return email;
}

/* ---------------------------- Access list --------------------------- */

const ACCESS_HEADERS = ['email', 'name', 'role', 'active', 'notes'];

/** 'ok' | 'domain' (not @sr.edu.sa) | 'list' (not on the active list) */
function accessStatus_(email) {
  if (!isAllowed_(email)) return 'domain';
  if (ACCESS_MODE !== 'LIST') return 'ok';
  if (OWNER_EMAILS.map(e => e.toLowerCase()).indexOf(email) !== -1) return 'ok';
  return allowedEmails_().indexOf(email) !== -1 ? 'ok' : 'list';
}

/** Active emails from the AllowedUsers tab (cached briefly). */
function allowedEmails_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('acl_v1');
  if (hit) return JSON.parse(hit);
  let list = [];
  try {
    const sh = accessSheet_(false);
    if (sh && sh.getLastRow() > 1) {
      list = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues()
        .filter(r => r[0] && r[3] !== false && String(r[3]).toUpperCase() !== 'FALSE')
        .map(r => String(r[0]).trim().toLowerCase());
    }
  } catch (e) { console.error('allowedEmails_ failed: ' + e); }
  cache.put('acl_v1', JSON.stringify(list), ACCESS_CACHE_SECONDS);
  return list;
}

function accessSheet_(create) {
  const ss = ACCESS_SPREADSHEET_ID ? SpreadsheetApp.openById(ACCESS_SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_ACCESS);
  if (!sh && create) {
    sh = ss.insertSheet(SHEET_ACCESS);
    sh.getRange(1, 1, 1, ACCESS_HEADERS.length).setValues([ACCESS_HEADERS])
      .setFontWeight('bold').setFontColor('#ffffff').setBackground('#501e8c');
    sh.setFrozenRows(1);
    sh.setRightToLeft(true);
    [260, 200, 100, 70, 260].forEach((w, i) => sh.setColumnWidth(i + 1, w));
    sh.getRange(2, 4, 300, 1).insertCheckboxes();
    sh.getRange('A1').setNote('البريد الجامعي كاملًا — name@sr.edu.sa');
    sh.getRange('C1').setNote('admin / user (للتوثيق — يمكن استخدامه لاحقًا لصلاحيات أدق)');
    sh.getRange('D1').setNote('ألغِ العلامة لإيقاف الوصول دون حذف الصف');
  }
  return sh;
}

/** Run once from the editor: creates the AllowedUsers tab and adds the owner(s). */
function setupAccessList() {
  const sh = accessSheet_(true);
  const n = Math.max(sh.getLastRow() - 1, 0);            // includes empty rows that only hold checkboxes
  const data = n ? sh.getRange(2, 1, n, 5).getValues() : [];

  // Keep every existing entry (and its active state), compacted to the top — fixes rows that
  // earlier versions appended below the pre-made checkbox rows.
  const kept = data.filter(r => String(r[0]).trim())
    .map(r => [String(r[0]).trim(), r[1], r[2], r[3] === true || String(r[3]).toUpperCase() === 'TRUE', r[4]]);
  const have = kept.map(r => r[0].toLowerCase());

  const seed = OWNER_EMAILS.map(e => [e, 'Owner', 'admin', 'owner'])
    .concat(typeof DEFAULT_ALLOWED !== 'undefined' ? DEFAULT_ALLOWED : []);
  let added = 0;
  seed.forEach(r => {
    const e = String(r[0]).trim().toLowerCase();
    if (e && have.indexOf(e) === -1) {
      kept.push([e, r[1] || '', r[2] || 'user', true, r[3] || 'added by setupAccessList']);
      have.push(e); added++;
    }
  });

  if (n) sh.getRange(2, 1, n, 5).clearContent();
  const rows = Math.max(n, kept.length, 50);
  sh.getRange(2, 4, rows, 1).insertCheckboxes();          // checkboxes first (this sets them to unchecked)
  if (kept.length) sh.getRange(2, 1, kept.length, 5).setValues(kept);  // then the real values

  CacheService.getScriptCache().remove('acl_v1');
  console.log('AllowedUsers ready: ' + kept.length + ' entries (' + added + ' added).');
  try { SpreadsheetApp.getActive().toast('قائمة المصرّح لهم جاهزة: ' + kept.length + ' بريدًا (أُضيف ' + added + ').', SHEET_ACCESS, 8); } catch (e) {}
}

/* ---------------------------- Access log ---------------------------- */

/** Appends one row to AccessLog. Never blocks the page if logging fails. */
function logAccess_(email, result) {
  if (!LOG_ENABLED) return;
  try {
    if (LOG_DEDUP_MINUTES > 0 && result === 'granted') {
      const cache = CacheService.getScriptCache();
      const key = 'log_' + email;
      if (cache.get(key)) return;
      cache.put(key, '1', LOG_DEDUP_MINUTES * 60);
    }
    const now = new Date();
    const f = p => Utilities.formatDate(now, LOG_TIMEZONE, p);
    getLogSheet_().appendRow([now, f('yyyy-MM-dd'), f('HH:mm:ss'), DAYS_AR_[Number(f('u')) % 7],
      email, result === 'granted' ? 'دخول' : (result === 'not-listed' ? 'غير مدرج' : 'مرفوض'), APP_VERSION]);
  } catch (e) {
    console.error('logAccess_ failed: ' + e);
  }
}

function getLogSheet_() {
  const ss = LOG_SPREADSHEET_ID ? SpreadsheetApp.openById(LOG_SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('No spreadsheet: bind the script to a Sheet or set LOG_SPREADSHEET_ID');
  let sh = ss.getSheetByName(SHEET_LOG);
  if (!sh) {
    sh = ss.insertSheet(SHEET_LOG);
    sh.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS])
      .setFontWeight('bold').setFontColor('#ffffff').setBackground('#3a1464');
    sh.setFrozenRows(1);
    sh.setRightToLeft(true);
    sh.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm:ss');
    [170, 100, 80, 80, 260, 80, 70].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  }
  return sh;
}

/** Menu: quick summary of the last 30 days. */
function showLogSummary() {
  const sh = getLogSheet_();
  const ui = SpreadsheetApp.getUi();
  const last = sh.getLastRow();
  if (last < 2) { ui.alert('سجل الدخول', 'لا توجد سجلات بعد.', ui.ButtonSet.OK); return; }
  const since = new Date(Date.now() - 30 * 864e5);
  const rows = sh.getRange(2, 1, last - 1, LOG_HEADERS.length).getValues()
    .filter(r => r[0] instanceof Date && r[0] >= since && r[5] === 'دخول');
  const users = {};
  rows.forEach(r => { users[r[4]] = (users[r[4]] || 0) + 1; });
  const top = Object.keys(users).sort((a, b) => users[b] - users[a]).slice(0, 10)
    .map((u, i) => (i + 1) + '. ' + u + ' — ' + users[u]).join('\n');
  ui.alert('سجل الدخول — آخر 30 يومًا — ' + APP_NAME,
    'عدد الزيارات: ' + rows.length + '\nعدد المستخدمين: ' + Object.keys(users).length + '\n\nالأكثر دخولًا:\n' + (top || '—'),
    ui.ButtonSet.OK);
}

/** Menu: delete entries older than LOG_RETENTION_DAYS. */
function purgeOldLogs() {
  const sh = getLogSheet_();
  const last = sh.getLastRow();
  let removed = 0;
  if (last > 1) {
    const cutoff = new Date(Date.now() - LOG_RETENTION_DAYS * 864e5);
    const stamps = sh.getRange(2, 1, last - 1, 1).getValues();
    for (let i = stamps.length - 1; i >= 0; i--) {
      if (stamps[i][0] instanceof Date && stamps[i][0] < cutoff) { sh.deleteRow(i + 2); removed++; }
    }
  }
  SpreadsheetApp.getActive().toast('حُذف ' + removed + ' سجل أقدم من ' + LOG_RETENTION_DAYS + ' يومًا.', APP_NAME_EN, 6);
}

/** Sheet menu (bound scripts). Add platform items here. */
function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu(APP_NAME_EN)
      .addItem('إعداد قائمة المصرّح لهم (AllowedUsers)', 'setupAccessList')
      .addSeparator()
      .addItem('ملخص سجل الدخول (30 يومًا)', 'showLogSummary')
      .addItem('حذف سجلات الدخول القديمة', 'purgeOldLogs')
      .addSeparator()
      .addItem('فحص المعرّفات المكررة الآن', 'repairDuplicateIdsFromMenu')
      .addSeparator()
      .addItem('عرض رابط المنصة', 'showWebAppUrl')
      .addToUi();
  } catch (e) { /* standalone script: no sheet UI */ }
}

function showWebAppUrl() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('رابط ' + APP_NAME, ScriptApp.getService().getUrl() || 'لم يُنشر بعد: Deploy ▸ New deployment ▸ Web app', ui.ButtonSet.OK);
}

/* ----------------------------- Helpers ------------------------------ */

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}
function safeJson_(obj) {
  return JSON.stringify(obj).replace(/</g, '\\u003c').replace(/\u2028/g, '\\u2028').replace(/\u2029/g, '\\u2029');
}

/* =====================================================================
 *  جرد أصول البنية التحتية — منطق البيانات (Sheet CRUD)
 *  مُدمَج من sru_infrastructure_inventory Code.gs v3.5 داخل حزمة الوصول الآمن
 *  كل القراءة/الكتابة بالاسم (headerMap) لا بموضع العمود - أمان من إعادة الترتيب
 * =====================================================================
 */

const TYPES = [
  'PC','Laptop','Screen','Printer','DesktopPrinter','PolycomVC','SmartProjector','Projector',
  'LHD','MRS','CoreSwitch','DistributionSwitch','Routers','LoadBalancers','Switch','UPS','Scanner',
  'Accessories','ScreenVertical','MicrosoftService','NetworkReceivers','Firewall','LargeScreens',
  'HallScreenHuawei','HallScreenBenq','CiscoPhone','WirelessCiscoPhone','Other'
];

// AddedBy and PhotoUrl were added after the original launch - kept at the very
// end on purpose so older sheets/rows never shift when they're backfilled in.
const HEADERS = [
  'ID','Tag','Brand','Model','Serial','Supplier','Specs','Location','User','Status',
  'Date','Notes','Warranty','WarrantyExpiry','UpdatedAt','AddedBy','PhotoUrl'
];
const SUPPLIER_HEADERS = ['SupplierName','ContactName','Email','Phone','UpdatedAt'];
const CONFIG_HEADERS = ['typeKey','typeLabel','extra1Label','extra2Label'];
const PHOTOS_FOLDER_NAME = 'SRU Infrastructure Inventory - Photos';

// ---- sheet + header helpers (name-based, position-safe) ----

function getOrCreateSheet(ss, type) {
  let sheet = ss.getSheetByName(type);
  if (!sheet) {
    sheet = ss.insertSheet(type);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  } else {
    ensureHeaders(sheet);
  }
  return sheet;
}

// Adds any header from HEADERS that's missing from an existing sheet's first
// row, appending it as a NEW column at the end - never reorders or removes
// an existing header, so already-saved data is never shifted.
function ensureHeaders(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const existing = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const missing = HEADERS.filter(function (h) { return existing.indexOf(h) === -1; });
  if (missing.length) {
    sheet.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
  }
}

function headerMap(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const row = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const map = {};
  row.forEach(function (h, i) { if (h) map[String(h)] = i + 1; }); // 1-based column index
  return map;
}

// Writes a {HeaderName: value} object into one row, by header name, in a
// SINGLE getRange/setValues call. Returns the row index actually written to.
//
// Deliberately avoids the previous "append an empty/holey row, then set each
// cell one at a time" pattern: a sparse array (`new Array(n)`) handed to
// Apps Script's Sheets bridge, and 17 separate setValue() calls per record,
// are both unnecessary risk for zero benefit - one full, non-sparse array
// written in one call is faster and behaves predictably every time.
function writeRecord(sheet, existingRowIndex, dataByHeader) {
  ensureHeaders(sheet); // make sure every header we're about to write actually exists
  const hMap = headerMap(sheet);
  const numCols = sheet.getLastColumn();

  const rowIndex = (existingRowIndex > -1) ? existingRowIndex : (sheet.getLastRow() + 1);

  // Start from the row's current values on update (so any column this call
  // doesn't touch is preserved), or a real, fully-populated blank row on add.
  let rowValues = (existingRowIndex > -1)
    ? sheet.getRange(rowIndex, 1, 1, numCols).getValues()[0]
    : new Array(numCols).fill('');

  Object.keys(dataByHeader).forEach(function (key) {
    const col = hMap[key];
    if (col) rowValues[col - 1] = dataByHeader[key];
  });

  sheet.getRange(rowIndex, 1, 1, numCols).setValues([rowValues]);
  return rowIndex;
}

function readSheet(sheet) {
  const hMap = headerMap(sheet);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const col = function (name) { return hMap[name] ? hMap[name] - 1 : -1; };
  const idCol = col('ID');
  return values.slice(1)
    .filter(function (r) { return idCol > -1 && r[idCol]; })
    .map(function (r) {
      const get = function (name) { const i = col(name); return i > -1 ? r[i] : ''; };
      const dateVal = get('Date');
      const warrExpVal = get('WarrantyExpiry');
      return {
        id: get('ID'),
        tag: get('Tag'),
        brand: get('Brand'),
        model: get('Model'),
        serial: get('Serial'),
        supplier: get('Supplier'),
        specs: safeParse(get('Specs')),
        location: get('Location'),
        user: get('User'),
        status: get('Status'),
        date: dateVal instanceof Date ? Utilities.formatDate(dateVal, Session.getScriptTimeZone(), 'yyyy-MM-dd') : dateVal,
        notes: get('Notes'),
        warranty: get('Warranty'),
        warrantyExpiry: warrExpVal instanceof Date ? Utilities.formatDate(warrExpVal, Session.getScriptTimeZone(), 'yyyy-MM-dd') : warrExpVal,
        addedBy: get('AddedBy'),
        photoUrl: get('PhotoUrl')
      };
    });
}

// ---- suppliers ----

function getOrCreateSuppliersSheet(ss) {
  let sheet = ss.getSheetByName('Suppliers');
  if (!sheet) {
    sheet = ss.insertSheet('Suppliers');
    sheet.appendRow(SUPPLIER_HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function readSuppliers(ss) {
  const sheet = getOrCreateSuppliersSheet(ss);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  return values.slice(1)
    .filter(function (r) { return r[0]; })
    .map(function (r) {
      return { name: r[0], contactName: r[1], email: r[2], phone: r[3] };
    });
}

// Records a supplier's contact details once; if the supplier already exists,
// only overwrites fields that were actually provided (non-empty), so entering
// the same supplier name on another device without contact info won't erase
// what's already on file - it just reuses it.
function upsertSupplier(ss, name, contactName, email, phone) {
  if (!name) return;
  const sheet = getOrCreateSuppliersSheet(ss);
  const values = sheet.getDataRange().getValues();
  let rowIndex = -1;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim().toLowerCase() === String(name).trim().toLowerCase()) {
      rowIndex = i + 1;
      break;
    }
  }
  if (rowIndex > -1) {
    const existing = sheet.getRange(rowIndex, 1, 1, 4).getValues()[0];
    sheet.getRange(rowIndex, 1, 1, 5).setValues([[
      name,
      contactName || existing[1],
      email || existing[2],
      phone || existing[3],
      new Date().toISOString()
    ]]);
  } else {
    sheet.appendRow([name, contactName || '', email || '', phone || '', new Date().toISOString()]);
  }
}

// ---- optional per-type config (custom labels) - safe to leave empty ----

function readConfig(ss) {
  const sheet = ss.getSheetByName('Config');
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  return values.slice(1)
    .filter(function (r) { return r[0]; })
    .map(function (r) {
      return { typeKey: r[0], typeLabel: r[1], extra1Label: r[2], extra2Label: r[3] };
    });
}

// ---- device photo upload ----

function getOrCreatePhotosFolder() {
  const it = DriveApp.getFoldersByName(PHOTOS_FOLDER_NAME);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(PHOTOS_FOLDER_NAME);
}

// Saves a base64-encoded image to Drive and returns a URL that renders
// reliably inline in an <img> tag (drive.google.com/uc?export=view does NOT
// render reliably when hotlinked - always use the /thumbnail endpoint).
function savePhotoIfProvided(base64, mime, name) {
  if (!base64) return '';
  const folder = getOrCreatePhotosFolder();
  const bytes = Utilities.base64Decode(base64);
  const blob = Utilities.newBlob(bytes, mime || 'image/jpeg', name || ('photo_' + Date.now() + '.jpg'));
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1000';
}

// ---- misc ----

function findRowById(sheet, id) {
  const hMap = headerMap(sheet);
  const idCol = hMap['ID'];
  if (!idCol) return -1;
  const values = sheet.getRange(1, idCol, sheet.getLastRow(), 1).getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) return i + 1;
  }
  return -1;
}

function safeParse(s) {
  try { return JSON.parse(s); } catch (e) { return {}; }
}




function repairDuplicateIds(ss) {
  const sheetsByType = {};
  const occurrencesById = {}; // id -> [{type, row, tag}, ...]
  TYPES.forEach(function (t) {
    const sheet = ss.getSheetByName(t);
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    sheetsByType[t] = sheet;
    const hMap = headerMap(sheet);
    const idCol = hMap['ID'];
    if (!idCol) return;
    const tagCol = hMap['Tag'];
    const idVals = sheet.getRange(2, idCol, lastRow - 1, 1).getValues();
    const tagVals = tagCol ? sheet.getRange(2, tagCol, lastRow - 1, 1).getValues() : null;
    for (let i = 0; i < idVals.length; i++) {
      const idVal = String(idVals[i][0] || '').trim();
      if (!idVal) continue;
      const rowIndex = i + 2;
      const tagVal = tagVals ? String(tagVals[i][0] || '') : '';
      if (!occurrencesById[idVal]) occurrencesById[idVal] = [];
      occurrencesById[idVal].push({ type: t, row: rowIndex, tag: tagVal });
    }
  });

  const idsInUse = {};
  Object.keys(occurrencesById).forEach(function (id) { idsInUse[id] = true; });
  function makeUniqueId() {
    let id;
    do { id = 'd' + Date.now() + Math.random().toString(16).slice(2, 8); } while (idsInUse[id]);
    idsInUse[id] = true;
    return id;
  }

  const fixes = [];
  Object.keys(occurrencesById).forEach(function (id) {
    const occurrences = occurrencesById[id];
    if (occurrences.length < 2) return;
    for (let i = 1; i < occurrences.length; i++) {
      const occ = occurrences[i];
      const sheet = sheetsByType[occ.type];
      const idCol = headerMap(sheet)['ID'];
      const newId = makeUniqueId();
      sheet.getRange(occ.row, idCol).setValue(newId);
      fixes.push({ type: occ.type, row: occ.row, tag: occ.tag, oldId: id, newId: newId });
    }
  });

  if (fixes.length) SpreadsheetApp.flush();
  return fixes;
}

function repairDuplicateIdsFromMenu() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fixes = repairDuplicateIds(ss);
  SpreadsheetApp.getActive().toast(
    fixes.length ? ('تم إصلاح ' + fixes.length + ' معرّف مكرر.') : 'لا توجد معرّفات مكررة.',
    APP_NAME_EN, 6
  );
}
