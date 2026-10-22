/**
 * EYECON OPTICS — Google Apps Script backend.
 * Every app section is stored in its own sheet of the spreadsheet this script is bound to.
 * Columns: A id | B updatedAt (ms) | C summary (for reading only) | D… record JSON (split into chunks).
 * Users live in the "_users" sheet; deletions are logged in "_deleted" so other devices sync them.
 */
var CHUNK = 45000;
var TOKEN_TTL = 12 * 3600 * 1000;
var RESERVED = { _users: 1, _deleted: 1, _log: 1 };
var VIEW_PREFIX = 'عرض - ';

function doGet() {
  try { HtmlService.createHtmlOutputFromFile('index'); } catch (e) {
    return ContentService.createTextOutput('EYECON server OK — ' + (ss_() ? ss_().getName() : '') + ' — ' + new Date().toISOString());
  }
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('EYECON OPTICS')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/** JSON API used by the GitHub Pages site: POST {fn, args} as text/plain. */
function doPost(e) {
  var out;
  try {
    var req = JSON.parse(e.postData.contents);
    var name = 'api_' + String(req.fn || '');
    var f = globalThis[name];
    if (!/^api_[A-Za-z]+$/.test(name) || typeof f !== 'function') throw new Error('bad_fn');
    out = { ok: f.apply(null, Array.isArray(req.args) ? req.args : []) };
  } catch (err) { out = { err: String(err && err.message || err) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- helpers ---------- */
// Optional: if this script was created from script.google.com (not from the Sheet's Extensions menu),
// paste the Sheet id here (the long part of the Sheet URL between /d/ and /edit).
var SHEET_ID = '';
function ss_() {
  var s = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  if (!s) throw new Error('no_sheet');
  return s;
}
function props_() { return PropertiesService.getScriptProperties(); }
function sheet_(name, create) {
  var sh = ss_().getSheetByName(name);
  if (!sh && create) {
    sh = ss_().insertSheet(name);
    if (name === '_users') sh.appendRow(['username', 'name', 'role', 'active', 'salt', 'hash', 'created', 'branches']);
    else if (name === '_deleted') sh.appendRow(['collection', 'id', 'time', 'by', 'deleted record (JSON)']);
    else if (name === '_log') sh.appendRow(['الوقت', 'المستخدم', 'العملية', 'القسم', 'رقم السجل', 'ملخص', 'تفاصيل']);
    else sh.appendRow(['id', 'updatedAt', 'summary', 'data']);
    sh.setFrozenRows(1);
  }
  return sh;
}
function validCol_(c) {
  if (typeof c !== 'string' || !/^[A-Za-z][A-Za-z0-9_]{0,40}$/.test(c) || RESERVED[c]) throw new Error('bad_collection');
  return c;
}
function validId_(id) {
  if (typeof id !== 'string' || !/^[A-Za-z0-9_\-.]{1,80}$/.test(id)) throw new Error('bad_id');
  return id;
}
function hash_(salt, pass) {
  var h = salt + '|' + pass;
  for (var i = 0; i < 300; i++) {
    var b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h, Utilities.Charset.UTF_8);
    h = Utilities.base64Encode(b);
  }
  return h;
}
function uid_() { return Utilities.getUuid().replace(/-/g, ''); }
function summary_(c, d) {
  var keys = ['code', 'no', 'name', 'brand', 'customer', 'supplier', 'title', 'date', 'phone', 'note', 'total', 'amount', 'qty', 'status'];
  var out = [];
  for (var i = 0; i < keys.length && out.length < 6; i++) {
    var v = d[keys[i]];
    if (v !== undefined && v !== null && v !== '' && typeof v !== 'object') out.push(String(v).slice(0, 60));
  }
  return out.join(' | ');
}
function rowFor_(c, id, d, t) {
  var json = JSON.stringify(d), sm = summary_(c, d);
  var row = [id, t, /^[=+\-@]/.test(sm) ? ' ' + sm : sm];
  for (var i = 0; i < json.length; i += CHUNK) row.push('~' + json.slice(i, i + CHUNK));
  return row;
}
function readRow_(r) {
  var json = '';
  for (var i = 3; i < r.length; i++) { if (r[i] === '' || r[i] === null) break; var x = String(r[i]); json += x.charAt(0) === '~' ? x.slice(1) : x; }
  try { return JSON.parse(json || '{}'); } catch (e) { return {}; }
}
function findRow_(sh, id) {
  if (sh.getLastRow() < 2) return 0;
  var f = sh.getRange(2, 1, sh.getLastRow() - 1, 1).createTextFinder(id).matchEntireCell(true).findNext();
  return f ? f.getRow() : 0;
}
function writeRow_(sh, rowNo, row) {
  var width = Math.max(sh.getLastColumn(), row.length);
  var full = row.slice();
  while (full.length < width) full.push('');
  if (sh.getMaxColumns() < width) sh.insertColumnsAfter(sh.getMaxColumns(), width - sh.getMaxColumns());
  var rng = sh.getRange(rowNo, 1, 1, width);
  rng.setNumberFormat('@');
  rng.setValues([full.map(function (x) { return typeof x === 'number' ? String(x) : x; })]);
}
function touch_(c, t) { props_().setProperty('m_' + c, String(t)); props_().setProperty('ver', String(t)); }
function now_() { return Date.now(); }

/* ---------- activity log (sheet _log) ---------- */
var AR_COL = { patients: 'الزبائن', exams: 'الفحوصات', sales: 'المبيعات', items: 'المخزون', purchases: 'المشتريات', expenses: 'المصاريف', vouchers: 'سندات الصرف', partners: 'الشركاء', withdrawals: 'سحوبات الشركاء', fixedAssets: 'الأصول', exchanges: 'صرف العملة', campaigns: 'حملات واتساب', branches: 'الفروع', guides: 'النماذج التثقيفية', settings: 'الإعدادات', rates: 'أسعار الصرف', counters: 'العدادات' };
function cut_(v, n) { v = String(v == null ? '' : v); return v.length > n ? v.slice(0, n) + '…' : v; }
function safe_(v) { return typeof v === 'string' && /^[=+@]/.test(v) ? "'" + v : v; }
function log_(user, action, c, id, sum, details) {
  try {
    sheet_('_log', true).appendRow([new Date(), String(user || ''), action, AR_COL[c] || c || '', id || '', safe_(cut_(sum, 300)), safe_(cut_(details, 2000))]);
  } catch (e) {}
}
function changed_(a, b) {
  var keys = {}, out = [];
  for (var k in a) keys[k] = 1; for (var k2 in b) keys[k2] = 1;
  for (var x in keys) { if (x === 'updatedAt' || x === 'createdAt') continue; if (JSON.stringify(a[x]) !== JSON.stringify(b[x])) out.push(x + ': ' + cut_(JSON.stringify(a[x] === undefined ? '' : a[x]), 60) + ' → ' + cut_(JSON.stringify(b[x] === undefined ? '' : b[x]), 60)); }
  return out.join(' | ');
}

/* ---------- auth ---------- */
function session_(token) {
  if (!token || typeof token !== 'string') throw new Error('auth');
  var raw = props_().getProperty('tok_' + token);
  if (!raw) throw new Error('auth');
  var s = JSON.parse(raw);
  if (s.exp < now_()) { props_().deleteProperty('tok_' + token); throw new Error('auth'); }
  return s;
}
var LEGACY = { admin: 'owner', staff: 'user', viewer: 'user' };
function role_(r) { r = String(r || ''); return LEGACY[r] || (['owner', 'manager', 'user'].indexOf(r) >= 0 ? r : 'user'); }
function writer_(token) { return session_(token); }
function admin_(token) { var s = session_(token); if (role_(s.role) !== 'owner' && role_(s.role) !== 'manager') throw new Error('perm'); s.role = role_(s.role); return s; }
function owner_(token) { var s = session_(token); if (role_(s.role) !== 'owner') throw new Error('perm'); s.role = 'owner'; return s; }
function users_() {
  var sh = sheet_('_users', true), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, 8).getValues().map(function (r, i) {
    return { row: i + 2, username: String(r[0]), name: String(r[1]), role: role_(r[2]), active: r[3] === true || String(r[3]).toUpperCase() === 'TRUE', salt: String(r[4]), hash: String(r[5]), branches: brList_(r[7]) };
  });
}
/* ---------- branches: every branch is isolated; a user with a branches list only receives and changes that branch's records ---------- */
var SCOPED = { items: 1, patients: 1, exams: 1, sales: 1, purchases: 1, expenses: 1, fixedAssets: 1, vouchers: 1, exchanges: 1, campaigns: 1 };
function brList_(v) {
  if (Array.isArray(v)) v = v.join(',');
  var a = String(v == null ? '' : v).split(',').map(function (x) { return x.trim(); }).filter(function (x) { return /^[A-Za-z0-9_\-]{1,60}$/.test(x); });
  return a.length ? a : null;
}
function vis_(s) {
  if (!s.br || role_(s.role) === 'owner') return null;
  var m = {}; s.br.forEach(function (b) { m[b] = 1; }); return m;
}
function setOwner_(id) {
  var m = /^(?:b|logo|brands)_(.+)$/.exec(id); if (m) return m[1];
  return (id === 'main' || id === 'logo' || id === 'brands') ? 'main' : null;
}
function docOk_(al, c, id, d, pv) {
  if (!al) return true; d = d || {};
  if (SCOPED[c]) return !!al[d.branchId || 'main'];
  if (c === 'branches') return !!al[id];
  if (c === 'settings') { var o = setOwner_(id); return !o || !!al[o]; }
  if (c === 'partners') return !!al[d.scope && d.scope !== 'group' ? d.scope : 'main'];
  if (c === 'withdrawals') return !!(pv && pv[d.partnerId]);
  return true;
}
function partnersVis_(al) {
  if (!al) return null; var m = {};
  readCol_('partners', 0).forEach(function (x) { if (docOk_(al, 'partners', x[0], x[1])) m[x[0]] = 1; });
  return m;
}
function cleanTokens_() {
  var all = props_().getProperties(), t = now_();
  for (var k in all) if (k.indexOf('tok_') === 0) { try { if (JSON.parse(all[k]).exp < t) props_().deleteProperty(k); } catch (e) { props_().deleteProperty(k); } }
}
function issue_(u) {
  cleanTokens_();
  var tok = uid_() + uid_();
  var br = role_(u.role) === 'owner' ? null : (u.branches || null);
  props_().setProperty('tok_' + tok, JSON.stringify({ u: u.username, role: role_(u.role), name: u.name, br: br, exp: now_() + TOKEN_TTL }));
  return { token: tok, user: { username: u.username, name: u.name, role: u.role, branches: br } };
}

function api_status() { return JSON.stringify({ hasUsers: users_().length > 0 }); }

function api_setup(username, password, name) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    if (users_().length) throw new Error('exists');
    username = String(username || '').trim().toLowerCase();
    if (!/^[a-z0-9_.\-]{3,30}$/.test(username) || String(password || '').length < 6) throw new Error('weak');
    var salt = uid_();
    sheet_('_users', true).appendRow([username, String(name || username), 'owner', true, salt, hash_(salt, password), new Date()]);
    log_(username, 'إنشاء حساب المالك', '', '', String(name || username), '');
    return JSON.stringify(issue_({ username: username, name: String(name || username), role: 'owner' }));
  } finally { lock.releaseLock(); }
}

function api_login(username, password) {
  username = String(username || '').trim().toLowerCase();
  var u = users_().filter(function (x) { return x.username === username; })[0];
  Utilities.sleep(300);
  if (!u || !u.active || hash_(u.salt, String(password || '')) !== u.hash) { log_(username, 'محاولة دخول فاشلة', '', '', '', ''); throw new Error('login'); }
  log_(u.username, 'تسجيل دخول', '', '', u.name, '');
  return JSON.stringify(issue_(u));
}

function api_logout(token) { try { log_(session_(token).u, 'تسجيل خروج', '', '', '', ''); } catch (e) {} props_().deleteProperty('tok_' + token); return '1'; }

function api_me(token) { var s = session_(token); return JSON.stringify({ username: s.u, name: s.name, role: role_(s.role), branches: vis_(s) ? s.br : null }); }

function api_users(token) {
  admin_(token);
  return JSON.stringify(users_().map(function (u) { return { username: u.username, name: u.name, role: u.role, active: u.active, branches: u.role === 'owner' ? null : u.branches }; }));
}

function api_saveUser(token, o) {
  var me = admin_(token);
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var username = String(o.username || '').trim().toLowerCase();
    if (!/^[a-z0-9_.\-]{3,30}$/.test(username)) throw new Error('weak');
    var role = role_(o.role);
    var list = users_(), u = list.filter(function (x) { return x.username === username; })[0];
    var sh = sheet_('_users', true);
    if (sh.getRange(1, 8).getValue() === '') sh.getRange(1, 8).setValue('branches');
    var br = role === 'owner' ? null : brList_(o.branches), mine = vis_(me);
    if (mine) { if (!br) throw new Error('perm'); br.forEach(function (b) { if (!mine[b]) throw new Error('perm'); }); if (u && u.branches && u.branches.some(function (b) { return !mine[b]; })) throw new Error('perm'); }
    var brTxt = br ? br.join(',') : '';
    if (me.role !== 'owner' && (role !== 'user' || (u && u.role !== 'user'))) throw new Error('perm');
    if (u && username === me.u && (role !== me.role || o.active === false)) throw new Error('self');
    if (u && u.role === 'owner' && (role !== 'owner' || o.active === false) &&
        list.filter(function (x) { return x.role === 'owner' && x.active && x.username !== username; }).length === 0) throw new Error('last_owner');
    if (!u) {
      if (String(o.password || '').length < 6) throw new Error('weak');
      var salt = uid_();
      sh.appendRow([username, String(o.name || username), role, o.active !== false, salt, hash_(salt, o.password), new Date(), brTxt]);
    } else {
      sh.getRange(u.row, 2, 1, 3).setValues([[String(o.name || u.name), role, o.active !== false]]);
      var brChanged = (u.branches ? u.branches.join(',') : '') !== brTxt;
      sh.getRange(u.row, 8).setValue(brTxt);
      if (o.password) {
        if (String(o.password).length < 6) throw new Error('weak');
        var s2 = uid_(); sh.getRange(u.row, 5, 1, 2).setValues([[s2, hash_(s2, o.password)]]);
      }
      if (o.password || o.active === false || role !== u.role || brChanged) {
        var all = props_().getProperties();
        for (var k in all) if (k.indexOf('tok_') === 0) { try { if (JSON.parse(all[k]).u === username) props_().deleteProperty(k); } catch (e) {} }
      }
    }
    log_(me.u, u ? 'تعديل مستخدم' : 'إضافة مستخدم', '', username, String(o.name || ''), 'الصلاحية: ' + role + ' | الفروع: ' + (brTxt || 'الكل') + ' | فعّال: ' + (o.active !== false) + (o.password ? ' | كلمة مرور جديدة' : ''));
    return '1';
  } finally { lock.releaseLock(); }
}

function api_changePassword(token, oldPass, newPass) {
  var s = session_(token);
  var u = users_().filter(function (x) { return x.username === s.u; })[0];
  if (!u || hash_(u.salt, String(oldPass || '')) !== u.hash) throw new Error('login');
  if (String(newPass || '').length < 6) throw new Error('weak');
  var salt = uid_();
  sheet_('_users', true).getRange(u.row, 5, 1, 2).setValues([[salt, hash_(salt, newPass)]]);
  log_(s.u, 'تغيير كلمة المرور', '', '', '', '');
  return '1';
}

/* ---------- data ---------- */
function readCol_(c, since) {
  var sh = sheet_(c, false), out = [];
  if (!sh || sh.getLastRow() < 2) return out;
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  for (var i = 0; i < vals.length; i++) {
    var r = vals[i]; if (!r[0]) continue;
    var t = Number(r[1]) || 0;
    if (since && t <= since) continue;
    out.push([String(r[0]), readRow_(r), t]);
  }
  return out;
}

function api_load(token) {
  var s = session_(token), al = vis_(s), pv = partnersVis_(al);
  var res = { t: now_(), cols: {} };
  ss_().getSheets().forEach(function (sh) {
    var n = sh.getName(); if (RESERVED[n] || !/^[A-Za-z][A-Za-z0-9_]{0,40}$/.test(n)) return;
    res.cols[n] = readCol_(n, 0).filter(function (x) { return docOk_(al, n, x[0], x[1], pv); }).map(function (x) { return [x[0], x[1]]; });
  });
  return JSON.stringify(res);
}

function api_poll(token, since) {
  var s = session_(token), al = vis_(s), pv = null;
  since = Number(since) || 0;
  var t = now_(), all = props_().getProperties(), res = { t: t, cols: {}, del: [] };
  if (Number(all.ver || 0) <= since) return JSON.stringify(res);
  for (var k in all) {
    if (k.indexOf('m_') !== 0 || Number(all[k]) <= since) continue;
    var c = k.slice(2);
    if (c === '_deleted') {
      var dsh = sheet_('_deleted', false);
      if (dsh && dsh.getLastRow() > 1) dsh.getRange(2, 1, dsh.getLastRow() - 1, 3).getValues().forEach(function (r) { if (Number(r[2]) > since) res.del.push([String(r[0]), String(r[1])]); });
    } else {
      if (al && c === 'withdrawals' && !pv) pv = partnersVis_(al);
      res.cols[c] = readCol_(c, since).filter(function (x) { return docOk_(al, c, x[0], x[1], pv); }).map(function (x) { return [x[0], x[1]]; });
    }
  }
  return JSON.stringify(res);
}

function api_set(token, c, id, json, merge) {
  var me = writer_(token); validCol_(c); validId_(id);
  var d = JSON.parse(json);
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = sheet_(c, true), rowNo = findRow_(sh, id), t = now_(), prev = null;
    if (rowNo) prev = readRow_(sh.getRange(rowNo, 1, 1, sh.getLastColumn()).getValues()[0]);
    if (merge && prev) {
      var old = JSON.parse(JSON.stringify(prev));
      for (var k in d) old[k] = d[k];
      d = old;
    }
    var al = vis_(me);
    if (al) { var pv = c === 'withdrawals' ? partnersVis_(al) : null; if (!docOk_(al, c, id, d, pv) || (prev && !docOk_(al, c, id, prev, pv))) throw new Error('perm'); }
    log_(me.u, prev ? 'تعديل' : 'إضافة', c, id, summary_(c, d), prev ? changed_(prev, d) : '');
    if (rowNo) sh.getRange(rowNo, 4, 1, Math.max(1, sh.getLastColumn() - 3)).clearContent();
    writeRow_(sh, rowNo || sh.getLastRow() + 1, rowFor_(c, id, d, t));
    touch_(c, t);
    return JSON.stringify({ t: t, data: d });
  } finally { lock.releaseLock(); }
}

function api_del(token, c, id) {
  var me = writer_(token); validCol_(c); validId_(id);
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = sheet_(c, false), rowNo = sh ? findRow_(sh, id) : 0, t = now_(), prev = {};
    if (rowNo) {
      prev = readRow_(sh.getRange(rowNo, 1, 1, sh.getLastColumn()).getValues()[0]);
      var al = vis_(me); if (al && !docOk_(al, c, id, prev, c === 'withdrawals' ? partnersVis_(al) : null)) throw new Error('perm');
      sh.deleteRow(rowNo);
    }
    var pj = JSON.stringify(prev);
    sheet_('_deleted', true).appendRow([c, id, t, me.u, '~' + (pj.length > 49000 ? pj.slice(0, 49000) : pj)]);
    log_(me.u, 'حذف', c, id, summary_(c, prev), 'السجل المحذوف محفوظ في ورقة _deleted');
    touch_('_deleted', t);
    return JSON.stringify({ t: t });
  } finally { lock.releaseLock(); }
}

function api_incr(token, path, min) {
  writer_(token);
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var key = 'ctr_' + String(path).replace(/[^A-Za-z0-9_\/]/g, '');
    var n = Math.max(Number(props_().getProperty(key) || 0) + 1, Number(min) || 0);
    props_().setProperty(key, String(n));
    var cid = String(path).split('/').pop().replace(/[^A-Za-z0-9_]/g, '') || 'n', csh = sheet_('counters', true), cr = findRow_(csh, cid), ct = now_();
    writeRow_(csh, cr || csh.getLastRow() + 1, rowFor_('counters', cid, { n: n }, ct)); touch_('counters', ct);
    return String(n);
  } finally { lock.releaseLock(); }
}

/** Import an app backup (JSON from «تحميل نسخة احتياطية»). Admin only; existing ids are overwritten. */
function api_import(token, json) {
  owner_(token);
  var data = JSON.parse(json), n = 0;
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var t = now_();
    var put = function (c, id, d) {
      validCol_(c); validId_(id);
      var sh = sheet_(c, true), rowNo = findRow_(sh, id);
      if (rowNo) sh.getRange(rowNo, 4, 1, Math.max(1, sh.getLastColumn() - 3)).clearContent();
      writeRow_(sh, rowNo || sh.getLastRow() + 1, rowFor_(c, id, d, t)); n++;
    };
    if (data.settings) put('settings', 'main', data.settings);
    if (data.logoDoc) put('settings', 'logo', data.logoDoc); else if (data.logo) put('settings', 'logo', { data: data.logo });
    if (data.brandLogos) put('settings', 'brands', { logos: data.brandLogos });
    var touched = { settings: 1 };
    for (var c in data) {
      if (!Array.isArray(data[c]) || RESERVED[c]) continue;
      data[c].forEach(function (r) { if (r && r.id) { var d = {}; for (var k in r) if (k !== 'id') d[k] = r[k]; put(c, String(r.id), d); } });
      touched[c] = 1;
    }
    for (var k2 in touched) touch_(k2, t);
    log_(session_(token).u, 'استيراد نسخة احتياطية', '', '', n + ' سجل', Object.keys(touched).join(', '));
    return String(n);
  } finally { lock.releaseLock(); }
}


/* ---------- hourly readable archive inside this Sheet ---------- */
/** Run once from the Apps Script editor (choose setupHourly → Run) to refresh the readable sheets every hour. */
function setupHourly() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'hourlySnapshot') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('hourlySnapshot').timeBased().everyHours(1).create();
  return snapshot_(true);
}
function hourlySnapshot() { snapshot_(false); }

function flat_(o, pre, out) {
  out = out || {};
  for (var k in o) {
    if (k === 'updatedAt' || k === 'createdAt') continue;
    var v = o[k], key = pre ? pre + '.' + k : k;
    if (v === null || v === undefined) out[key] = '';
    else if (Array.isArray(v)) out[key] = v.length && typeof v[0] !== 'object' ? v.join(', ') : (v.length ? cut_(JSON.stringify(v), 45000) : '');
    else if (typeof v === 'object') flat_(v, key, out);
    else if (typeof v === 'string' && v.indexOf('data:') === 0) out[key] = '[صورة]';
    else out[key] = typeof v === 'string' ? safe_(cut_(v, 45000)) : v;
  }
  return out;
}
function writeView_(name, rows) {
  var sh = ss_().getSheetByName(name) || ss_().insertSheet(name);
  sh.clear();
  if (!rows.length) { sh.getRange(1, 1).setValue('—'); return; }
  var cols = [], seen = {};
  rows.forEach(function (r) { for (var k in r) if (!seen[k] && cols.length < 120) { seen[k] = 1; cols.push(k); } });
  var data = [cols].concat(rows.map(function (r) { return cols.map(function (c) { return r[c] === undefined ? '' : r[c]; }); }));
  sh.getRange(1, 1, data.length, cols.length).setValues(data);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, cols.length).setFontWeight('bold').setBackground('#E1F1FB');
}
function snapshot_(force) {
  var ver = props_().getProperty('ver') || '0';
  if (!force && props_().getProperty('snap_ver') === ver) return 'unchanged';
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var all = {};
    ss_().getSheets().forEach(function (sh) { var n = sh.getName(); if (RESERVED[n] || !/^[A-Za-z][A-Za-z0-9_]{0,40}$/.test(n)) return; all[n] = readCol_(n, 0); });
    var pname = {}; (all.patients || []).forEach(function (x) { pname[x[0]] = (x[1].code ? x[1].code + ' · ' : '') + (x[1].name || ''); });
    var stamp = function (x) { var o = { id: x[0] }; var f = flat_(x[1]); for (var k in f) o[k] = f[k]; o['آخر تعديل'] = x[2] ? new Date(x[2]) : ''; return o; };
    for (var c in all) {
      var rows = all[c].map(stamp);
      if (c === 'exams' || c === 'sales') rows.forEach(function (r, i) { r['الزبون'] = pname[all[c][i][1].patientId] || all[c][i][1].customer || ''; });
      writeView_(VIEW_PREFIX + (AR_COL[c] || c), rows);
    }
    var lines = [], pays = [];
    (all.sales || []).forEach(function (x) {
      var s = x[1], cust = pname[s.patientId] || s.customer || '';
      (s.lines || []).forEach(function (l) { lines.push({ 'رقم الفاتورة': s.no || '', 'التاريخ': s.date || '', 'الزبون': cust, 'المادة': l.name || '', 'الدرجة': l.pow || '', 'الكمية': l.qty, 'السعر': l.price, 'المجموع': (Number(l.qty) || 0) * (Number(l.price) || 0), 'العملة': s.currency || '' }); });
      (s.payments || []).forEach(function (p) { pays.push({ 'رقم الفاتورة': s.no || '', 'الزبون': cust, 'التاريخ': p.date || '', 'المبلغ': p.amount, 'العملة': p.cur || s.currency || '', 'الصندوق': p.method || '' }); });
    });
    writeView_(VIEW_PREFIX + 'مواد الفواتير', lines);
    writeView_(VIEW_PREFIX + 'الدفعات', pays);
    props_().setProperty('snap_ver', ver);
    props_().setProperty('snap_last', String(now_()));
    return 'ok';
  } finally { lock.releaseLock(); }
}
function api_archiveStatus(token) {
  admin_(token);
  var on = false; try { on = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'hourlySnapshot'; }); } catch (e) {}
  return JSON.stringify({ on: on, last: Number(props_().getProperty('snap_last') || 0), url: ss_().getUrl() });
}
function api_archiveNow(token) { var s = admin_(token); var r = snapshot_(true); log_(s.u, 'تحديث الأرشيف المقروء', '', '', '', ''); return r; }
