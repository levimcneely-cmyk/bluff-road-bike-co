/**
 * Bluff Road Bike Co — /admin backend (shop operations)
 *
 * Lives in the SAME Apps Script project as waiver-backend.gs (the one bound to the
 * private "BRBC Waiver" Google Sheet). Add this as a second file:
 *   Extensions → Apps Script → ＋ (Add a file) → Script → name it "admin-backend" → paste.
 *
 * What it does
 *   - The "Waivers" tab stays the source of truth: one row = one bike.
 *     Waiver signatures land there automatically; bikes added in /admin land there too.
 *   - "Monthly" tab holds the per-month Profit/Loss you used to type into the Stats tab.
 *   - doGet() (in waiver-backend.gs) serves the public about-page charts: brand + style
 *     counts only — never names, emails or phones.
 *   - Every /admin call is checked against a password (salted hash in Script Properties).
 *
 * Setup: run setupAdmin() once → reload the Sheet → BRBC Admin → Set password…
 * See apps-script/README.md.
 */

var ADM_MAX_FAILS = 10, ADM_LOCK_MIN = 15, ADM_SESSION_H = 12, ADM_REMEMBER_D = 30;
var ADM_MONTHLY_TAB = 'Monthly';
var ADM_MONTHLY_HEAD = ['Month', 'Profit/Loss', 'Notes'];
var ADM_CHARGED = 'Charged';
var ADM_STATS_CACHE = 'brbcPublicStats';

var ADM_STYLES = ['Road', 'Gravel', 'MTB', 'Hybrid', 'TT / Triathlon', 'Beach Cruiser', 'Kids', 'Other'];

// Brand families — matches how the about page has been grouping them.
var ADM_BRAND_ALIASES = {
  'gt': 'GT/Mongoose/Schwinn', 'mongoose': 'GT/Mongoose/Schwinn', 'schwinn': 'GT/Mongoose/Schwinn',
  'roadmaster': 'GT/Mongoose/Schwinn', 'gt/mongoose/schwinn': 'GT/Mongoose/Schwinn',
  'fuji': 'Fuji/Diamondback/Nishiki', 'diamondback': 'Fuji/Diamondback/Nishiki', 'nishiki': 'Fuji/Diamondback/Nishiki',
  'nikishi': 'Fuji/Diamondback/Nishiki', 'fuji/diamondback/nikishi': 'Fuji/Diamondback/Nishiki',
  'fuji/diamondback/nishiki': 'Fuji/Diamondback/Nishiki',
  'other': 'Other/Unbranded', 'unbranded': 'Other/Unbranded', 'other/unbranded': 'Other/Unbranded', 'none': 'Other/Unbranded',
  'cervélo': 'Cervelo', 'cervelo': 'Cervelo'
};

/* ───────── menu + setup ───────── */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('BRBC Admin')
    .addItem('Set password…', 'admSetPasswordPrompt')
    .addItem('Sign out all devices', 'admSignOutEverywhere')
    .addItem('Run setup (safe to re-run)', 'setupAdmin')
    .addToUi();
}

/** Run once from the editor (safe to re-run). */
function setupAdmin() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SHEET_TAB);
  if (!sh) throw new Error('No "' + SHEET_TAB + '" tab found.');
  var col = ensureHeaders_(sh);          // waiver columns (Ref ID, Source, …)
  admEnsureCol_(sh, col, ADM_CHARGED);   // optional $ per bike
  admBackfillRefs_(sh);

  var m = ss.getSheetByName(ADM_MONTHLY_TAB);
  if (!m) {
    m = ss.insertSheet(ADM_MONTHLY_TAB);
    m.getRange(1, 1, 1, 3).setValues([ADM_MONTHLY_HEAD]).setFontWeight('bold');
    m.setFrozenRows(1);
    m.getRange('A:A').setNumberFormat('@');
    m.getRange('B:B').setNumberFormat('$#,##0.00;-$#,##0.00');
    admSeedMonthly_(ss, m);
  }
  var p = PropertiesService.getScriptProperties();
  if (!p.getProperty('ADM_SECRET')) p.setProperty('ADM_SECRET', Utilities.getUuid() + Utilities.getUuid());
  admBustStats_();
  return 'Setup done. Next: reload the Sheet → BRBC Admin → Set password…';
}

function admEnsureCol_(sh, col, name) {
  if (col[name]) return col[name];
  var next = sh.getLastColumn() + 1;
  if (sh.getMaxColumns() < next) sh.insertColumnsAfter(sh.getMaxColumns(), 1);
  sh.getRange(1, next).setValue(name).setFontWeight('bold');
  col[name] = next;
  return next;
}

/** Give every existing bike row a Ref ID so /admin can edit it safely. */
function admBackfillRefs_(sh) {
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var col = admCols_(sh);
  var vals = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  var refCol = col[H.ref], n = 0;
  vals.forEach(function (r, i) {
    if (!admRowHasData_(r, col) || String(r[refCol - 1]).trim()) return;
    var ts = r[col[H.ts] - 1] instanceof Date ? r[col[H.ts] - 1] : new Date();
    sh.getRange(i + 2, refCol).setValue(admNewRef_(ts));
    n++;
  });
  return n;
}

/** Copy the Profit/Loss numbers you typed in the old Stats tab (Month/Year · Bikes · Profit/Loss). */
function admSeedMonthly_(ss, m) {
  var st = ss.getSheetByName('Stats');
  if (!st) return;
  var vals = st.getDataRange().getValues(), rows = [];
  for (var r = 0; r < vals.length; r++) {
    for (var c = 0; c < vals[r].length; c++) {
      if (String(vals[r][c]).trim() !== 'Month/Year') continue;
      var plCol = -1;
      for (var k = c; k < vals[r].length; k++) if (/profit/i.test(String(vals[r][k]))) plCol = k;
      for (var rr = r + 1; rr < vals.length; rr++) {
        var d = vals[rr][c];
        if (!(d instanceof Date)) break;
        var pl = plCol >= 0 ? vals[rr][plCol] : '';
        if (pl === '' || pl === null) continue;
        rows.push([Utilities.formatDate(d, TZ, 'yyyy-MM'), Number(pl), 'From old Stats tab']);
      }
      r = vals.length; break;
    }
  }
  if (rows.length) m.getRange(2, 1, rows.length, 3).setValues(rows);
}

function admSetPasswordPrompt() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('Set /admin password', 'At least 10 characters. A short phrase works well.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var pw = r.getResponseText();
  if (!pw || pw.length < 10) { ui.alert('Password must be at least 10 characters. Nothing changed.'); return; }
  admSetPassword_(pw);
  ui.alert('Password set. Any other signed-in devices were signed out.');
}

function admSignOutEverywhere() {
  var p = PropertiesService.getScriptProperties();
  p.setProperty('ADM_PW_VERSION', String(Number(p.getProperty('ADM_PW_VERSION') || 0) + 1));
}

function admSetPassword_(pw) {
  var p = PropertiesService.getScriptProperties();
  if (!p.getProperty('ADM_SECRET')) p.setProperty('ADM_SECRET', Utilities.getUuid() + Utilities.getUuid());
  var salt = Utilities.getUuid();
  p.setProperty('ADM_PW_SALT', salt);
  p.setProperty('ADM_PW_HASH', admHash_(pw, salt));
  p.setProperty('ADM_PW_VERSION', String(Number(p.getProperty('ADM_PW_VERSION') || 0) + 1));
  p.deleteProperty('ADM_FAILS'); p.deleteProperty('ADM_LOCK_UNTIL');
}

/* ───────── crypto ───────── */

function admHex_(bytes) { return bytes.map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join(''); }
function admHash_(pw, salt) {
  var h = salt + '|' + pw;
  for (var i = 0; i < 2000; i++) h = admHex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + '|' + salt, Utilities.Charset.UTF_8));
  return h;
}
function admSign_(payload) {
  return admHex_(Utilities.computeHmacSha256Signature(payload, PropertiesService.getScriptProperties().getProperty('ADM_SECRET')));
}
function admEq_(a, b) {
  a = String(a); b = String(b);
  if (a.length !== b.length) return false;
  var d = 0; for (var i = 0; i < a.length; i++) d |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return d === 0;
}
function admMakeToken_(remember) {
  var p = PropertiesService.getScriptProperties();
  var exp = Date.now() + (remember ? ADM_REMEMBER_D * 864e5 : ADM_SESSION_H * 36e5);
  var payload = 'b1.' + exp + '.' + (p.getProperty('ADM_PW_VERSION') || '0');
  return { token: payload + '.' + admSign_(payload), exp: exp };
}
function admCheckToken_(token) {
  var parts = String(token || '').split('.');
  if (parts.length !== 4) return false;
  var payload = parts.slice(0, 3).join('.');
  if (!admEq_(admSign_(payload), parts[3])) return false;
  if (Number(parts[1]) < Date.now()) return false;
  return parts[2] === (PropertiesService.getScriptProperties().getProperty('ADM_PW_VERSION') || '0');
}

/* ───────── router (called from doPost in waiver-backend.gs) ───────── */

function adminRoute_(req) {
  try {
    var a = req.action;
    if (a === 'login') return admLogin_(req);
    if (!admCheckToken_(req.token)) return { ok: false, auth: false, error: 'Signed out — please log in again.' };
    return admWithLock_(function () {
      if (a === 'bootstrap') return admBootstrap_();
      if (a === 'saveBike') { var b = admSaveBike_(req.bike || {}); admBustStats_(); return { ok: true, bike: b }; }
      if (a === 'deleteBike') { admDeleteBike_(req.ref); admBustStats_(); return { ok: true }; }
      if (a === 'saveMonth') return { ok: true, month: admSaveMonth_(req.month || {}) };
      if (a === 'changePassword') return admChangePassword_(req);
      throw new Error('Unknown action');
    });
  } catch (err) {
    return { ok: false, error: String(err && err.message || err) };
  }
}

function admWithLock_(fn) {
  var l = LockService.getScriptLock();
  l.waitLock(20000);
  try { return fn(); } finally { l.releaseLock(); }
}

function admLogin_(req) {
  var p = PropertiesService.getScriptProperties();
  if (!p.getProperty('ADM_PW_HASH')) return { ok: false, error: 'No password set yet. In the Sheet: BRBC Admin → Set password…' };
  var until = Number(p.getProperty('ADM_LOCK_UNTIL') || 0);
  if (until > Date.now()) return { ok: false, error: 'Too many wrong tries. Locked for ' + Math.ceil((until - Date.now()) / 60000) + ' more min.' };
  if (!admEq_(admHash_(String(req.password || ''), p.getProperty('ADM_PW_SALT')), p.getProperty('ADM_PW_HASH'))) {
    var fails = Number(p.getProperty('ADM_FAILS') || 0) + 1;
    if (fails >= ADM_MAX_FAILS) { p.setProperty('ADM_LOCK_UNTIL', String(Date.now() + ADM_LOCK_MIN * 60000)); fails = 0; }
    p.setProperty('ADM_FAILS', String(fails));
    Utilities.sleep(600);
    return { ok: false, error: 'Wrong password.' };
  }
  p.deleteProperty('ADM_FAILS');
  var t = admMakeToken_(!!req.remember);
  return { ok: true, token: t.token, exp: t.exp };
}

function admChangePassword_(req) {
  var p = PropertiesService.getScriptProperties();
  if (!admEq_(admHash_(String(req.oldPassword || ''), p.getProperty('ADM_PW_SALT')), p.getProperty('ADM_PW_HASH'))) return { ok: false, error: 'Current password is wrong.' };
  if (String(req.newPassword || '').length < 10) return { ok: false, error: 'New password must be at least 10 characters.' };
  admSetPassword_(String(req.newPassword));
  var t = admMakeToken_(!!req.remember);
  return { ok: true, token: t.token, exp: t.exp };
}

/* ───────── data ───────── */

function admSheet_() { return SpreadsheetApp.getActive().getSheetByName(SHEET_TAB); }

function admCols_(sh) {
  var head = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(String);
  var map = {};
  head.forEach(function (h, i) { if (h && !map[h]) map[h] = i + 1; });
  var hit = head.findIndex(function (x) { return x.indexOf('CUSTOMER ACKNOWLEDGMENT') === 0; });
  if (!map[H.ack] && hit >= 0) map[H.ack] = hit + 1;
  return map;
}

function admRowHasData_(r, col) {
  var get = function (k) { return col[k] ? String(r[col[k] - 1]).trim() : ''; };
  return !!(get(H.ts) || get(H.brand) || get(H.sigName) || get(H.style));
}

function admNewRef_(d) {
  return 'BRBC-' + Utilities.formatDate(d || new Date(), TZ, 'yyMMdd') + '-' +
    Utilities.getUuid().replace(/-/g, '').slice(0, 5).toUpperCase();
}

function admYmd_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  var s = String(v || '').trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return m[3] + '-' + ('0' + m[1]).slice(-2) + '-' + ('0' + m[2]).slice(-2);
  return '';
}

function normBrand_(b) {
  var s = String(b || '').trim().replace(/\s+/g, ' ');
  if (!s) return '';
  var hit = ADM_BRAND_ALIASES[s.toLowerCase()];
  if (hit) return hit;
  // "kona" → "Kona" (leave anything with capitals alone: BMC, GT, etc.)
  return s === s.toLowerCase() ? s.replace(/(^|[\s\/-])([a-z])/g, function (m, a, b) { return a + b.toUpperCase(); }) : s;
}

function normStyle_(s) {
  var t = String(s || '').trim().replace(/\s+/g, ' ');
  if (!t) return '';
  var low = t.toLowerCase();
  for (var i = 0; i < ADM_STYLES.length; i++) if (ADM_STYLES[i].toLowerCase() === low) return ADM_STYLES[i];
  if (/mountain|mtb/.test(low)) return 'MTB';
  if (/cruiser/.test(low)) return 'Beach Cruiser';
  if (/tri|^tt/.test(low)) return 'TT / Triathlon';
  if (/kid|child|youth/.test(low)) return 'Kids';
  return t;
}

function admReadBikes_() {
  var sh = admSheet_(), last = sh.getLastRow();
  if (last < 2) return [];
  var col = admCols_(sh);
  var vals = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  var get = function (r, k) { return col[k] ? r[col[k] - 1] : ''; };
  var out = [];
  vals.forEach(function (r, i) {
    if (!admRowHasData_(r, col)) return;
    var ref = String(get(r, H.ref)).trim();
    if (!ref) { // a row added by hand in the Sheet — give it an ID
      ref = admNewRef_(get(r, H.ts) instanceof Date ? get(r, H.ts) : new Date());
      sh.getRange(i + 2, col[H.ref]).setValue(ref);
    }
    var comp = get(r, H.completed);
    var charged = get(r, ADM_CHARGED);
    var src = String(get(r, H.source) || '');
    out.push({
      ref: ref,
      logged: admYmd_(get(r, H.ts)),
      date: admYmd_(get(r, H.date)) || admYmd_(get(r, H.ts)),
      customer: String(get(r, H.sigName) || '').trim(),
      phone: String(get(r, H.phone) || '').trim(),
      email: String(get(r, H.email) || '').trim(),
      brand: normBrand_(get(r, H.brand)),
      style: normStyle_(get(r, H.style)),
      model: String(get(r, H.makeModel) || '').trim(),
      notes: String(get(r, H.notes) || '').trim(),
      completed: comp === true || String(comp).toUpperCase() === 'TRUE',
      charged: charged === '' || charged === null ? '' : Number(charged),
      signed: src.indexOf('waiver') >= 0 || String(get(r, H.ack)).trim() === 'I Agree',
      pdf: String(get(r, H.pdf) || ''),
      source: src
    });
  });
  return out;
}

function admFindRow_(sh, col, ref) {
  var last = sh.getLastRow();
  if (last < 2 || !ref) return -1;
  var ids = sh.getRange(2, col[H.ref], last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]).trim() === String(ref)) return i + 2;
  return -1;
}

function admSaveBike_(b) {
  var brand = normBrand_(b.brand), style = normStyle_(b.style);
  if (!brand) throw new Error('Pick a brand.');
  if (!style) throw new Error('Pick a style.');
  var date = admYmd_(b.date) || Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  var cut = function (v, n) { return String(v == null ? '' : v).trim().slice(0, n); };
  var safe = function (v) { return /^[=+\-@]/.test(v) ? "'" + v : v; }; // never let typed text become a formula
  var charged = b.charged === '' || b.charged == null || isNaN(Number(b.charged)) ? '' : Math.round(Number(b.charged) * 100) / 100;

  var sh = admSheet_();
  var col = ensureHeaders_(sh);
  admEnsureCol_(sh, col, ADM_CHARGED);
  var row = admFindRow_(sh, col, b.ref);
  var fields = {};
  fields[H.date] = Utilities.parseDate(date, TZ, 'yyyy-MM-dd');
  fields[H.brand] = brand;
  fields[H.style] = style;
  fields[H.makeModel] = safe(cut(b.model, 200));
  fields[H.notes] = safe(cut(b.notes, 2000));
  fields[H.completed] = !!b.completed;
  fields[ADM_CHARGED] = charged;

  if (row < 0) {
    var arr = new Array(sh.getLastColumn()).fill('');
    var ref = admNewRef_(new Date());
    arr[col[H.ts] - 1] = new Date();
    arr[col[H.sigName] - 1] = safe(cut(b.customer, 120));
    arr[col[H.phone] - 1] = safe(cut(b.phone, 40));
    arr[col[H.email] - 1] = cut(b.email, 160);
    arr[col[H.ref] - 1] = ref;
    arr[col[H.source] - 1] = 'admin';
    Object.keys(fields).forEach(function (k) { arr[col[k] - 1] = fields[k]; });
    sh.appendRow(arr);
    row = sh.getLastRow();
    b.ref = ref;
  } else {
    // Signed waivers keep the customer's own name/phone/email exactly as signed.
    var signed = String(sh.getRange(row, col[H.source]).getValue()).indexOf('waiver') >= 0 ||
                 String(sh.getRange(row, col[H.ack]).getValue()).trim() === 'I Agree';
    if (!signed) {
      fields[H.sigName] = safe(cut(b.customer, 120));
      fields[H.phone] = safe(cut(b.phone, 40));
      fields[H.email] = cut(b.email, 160);
    }
    Object.keys(fields).forEach(function (k) { sh.getRange(row, col[k]).setValue(fields[k]); });
  }
  SpreadsheetApp.flush();
  var saved = admReadBikes_().filter(function (x) { return x.ref === b.ref; })[0];
  return saved;
}

function admDeleteBike_(ref) {
  var sh = admSheet_();
  var col = admCols_(sh);
  var row = admFindRow_(sh, col, ref);
  if (row < 0) throw new Error('That bike was not found (maybe already deleted).');
  sh.deleteRow(row);
}

function admReadMonthly_() {
  var m = SpreadsheetApp.getActive().getSheetByName(ADM_MONTHLY_TAB);
  if (!m || m.getLastRow() < 2) return [];
  return m.getRange(2, 1, m.getLastRow() - 1, 3).getValues()
    .filter(function (r) { return String(r[0]).trim(); })
    .map(function (r) {
      var k = r[0] instanceof Date ? Utilities.formatDate(r[0], TZ, 'yyyy-MM') : String(r[0]).trim().slice(0, 7);
      return { month: k, pl: r[1] === '' ? '' : Number(r[1]), notes: String(r[2] || '') };
    });
}

function admSaveMonth_(mo) {
  var key = String(mo.month || '');
  if (!/^\d{4}-\d{2}$/.test(key)) throw new Error('Bad month.');
  var pl = mo.pl === '' || mo.pl == null || isNaN(Number(mo.pl)) ? '' : Math.round(Number(mo.pl) * 100) / 100;
  var notes = String(mo.notes || '').slice(0, 500);
  if (/^[=+\-@]/.test(notes)) notes = "'" + notes;
  var m = SpreadsheetApp.getActive().getSheetByName(ADM_MONTHLY_TAB);
  if (!m) { setupAdmin(); m = SpreadsheetApp.getActive().getSheetByName(ADM_MONTHLY_TAB); }
  var last = m.getLastRow(), row = -1;
  if (last >= 2) {
    var keys = m.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < keys.length; i++) {
      var k = keys[i][0] instanceof Date ? Utilities.formatDate(keys[i][0], TZ, 'yyyy-MM') : String(keys[i][0]).trim().slice(0, 7);
      if (k === key) { row = i + 2; break; }
    }
  }
  if (row < 0) row = last + 1;
  m.getRange(row, 1, 1, 3).setValues([[key, pl, notes]]);
  return { month: key, pl: pl, notes: String(mo.notes || '').slice(0, 500) };
}

function admBootstrap_() {
  return {
    ok: true,
    bikes: admReadBikes_(),
    monthly: admReadMonthly_(),
    styles: ADM_STYLES,
    sheetUrl: SpreadsheetApp.getActive().getUrl(),
    today: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd')
  };
}

/* ───────── public stats for the about page (counts only, no customer info) ───────── */

function publicStats_() {
  var c = CacheService.getScriptCache();
  var hit = c.get(ADM_STATS_CACHE);
  if (hit) return JSON.parse(hit);
  var brand = {}, style = {}, total = 0;
  admReadBikes_().forEach(function (b) {
    if (!b.brand || !b.style) return;
    brand[b.brand] = (brand[b.brand] || 0) + 1;
    style[b.style] = (style[b.style] || 0) + 1;
    total++;
  });
  var out = { ok: true, service: 'BRBC', brand: brand, style: style, total: total };
  c.put(ADM_STATS_CACHE, JSON.stringify(out), 600);
  return out;
}

function admBustStats_() { try { CacheService.getScriptCache().remove(ADM_STATS_CACHE); } catch (e) {} }
