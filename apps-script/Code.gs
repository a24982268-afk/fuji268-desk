/**
 * 蹦火仔 接團台 後端（Google Apps Script）
 *
 * 用法：在 Google 試算表的「擴充功能 → Apps Script」整段貼上，詳細步驟見 README.md。
 * 密碼不寫在程式裡：用試算表上方選單「接團台 → 設定／更改登入密碼」設定，只存雜湊值。
 */

var TZ = 'Asia/Taipei';
var TOKEN_DAYS = 30;          // 登入後幾天內不用再輸入密碼
var MAX_FAILS = 10;           // 連續輸錯幾次就暫停登入
var LOCK_SECONDS = 15 * 60;   // 暫停登入多久

var SHEETS = {
  bookings: { title: '團單', cols: [
    ['id', '編號'], ['status', '狀態'], ['date', '日期'], ['time', '集合時間'], ['type', '類型'],
    ['org', '學校／單位'], ['name', '聯絡人'], ['phone', '手機'], ['adult', '大人'], ['kid', '小孩'],
    ['tourId', '遊程編號'], ['slot', '抵達時段'], ['quote', '報價'], ['deposit', '訂金'], ['depositPaid', '已收訂金'],
    ['guides', '導覽員'], ['bus', '遊覽車／停車'], ['meal', '用餐'], ['tourNote', '客製行程內容'], ['note', '備註'],
    ['createdAt', '建立時間']
  ]},
  guides: { title: '導覽員', cols: [['id', '編號'], ['name', '姓名'], ['note', '備註']] },
  tours: { title: '遊程', cols: [['id', '編號'], ['name', '遊程名稱'], ['meet', '集合地點'], ['content', '遊程內容']] }
};
var SETTINGS_TITLE = '設定';
var SETTINGS = [['perGuide', '每位導覽員最多帶幾人']];
var NUM = { adult: 1, kid: 1, quote: 1, deposit: 1 };

/* ---------- 試算表選單與初始化 ---------- */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('接團台')
    .addItem('設定／更改登入密碼', 'setPasswordMenu')
    .addItem('讓所有裝置重新登入', 'logoutAllMenu')
    .addToUi();
}

function 初始化() {
  var p = PropertiesService.getScriptProperties();
  p.setProperty('SHEET_ID', SpreadsheetApp.getActiveSpreadsheet().getId());
  if (!p.getProperty('SALT')) p.setProperty('SALT', Utilities.getUuid());
  if (!p.getProperty('SECRET')) p.setProperty('SECRET', Utilities.getUuid() + Utilities.getUuid());
  Object.keys(SHEETS).forEach(sheetOf);
  settingsSheet();
  Logger.log('初始化完成。請回到試算表，重新整理後用「接團台」選單設定密碼。');
}

function setPasswordMenu() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('設定登入密碼', '請輸入新的共用密碼（至少 8 個字）。\n更改密碼後，所有已登入的手機都要用新密碼重新登入。', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var pw = r.getResponseText().trim();
  if (pw.length < 8) { ui.alert('密碼太短，至少要 8 個字。沒有變更。'); return; }
  初始化();
  PropertiesService.getScriptProperties().setProperty('PW_HASH', hash(pw));
  ui.alert('密碼已設定。');
}

function logoutAllMenu() {
  PropertiesService.getScriptProperties().setProperty('SECRET', Utilities.getUuid() + Utilities.getUuid());
  SpreadsheetApp.getUi().alert('已讓所有裝置登出，下次打開接團台要重新輸入密碼。');
}

/* ---------- 網頁 API ---------- */

function doGet() {
  return ContentService.createTextOutput('蹦火仔接團台後端運作中');
}

function doPost(e) {
  var out;
  try {
    out = handle(JSON.parse(e.postData.contents));
  } catch (err) {
    out = { ok: false, code: 'error', message: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function handle(req) {
  if (req.action === 'login') return login(String(req.password || ''));
  if (!checkToken(req.token)) return { ok: false, code: 'auth' };
  if (req.action === 'load') return { ok: true, data: loadAll() };

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var id = null, coll = req.collection;
    if (req.action === 'settings') {
      writeSettings(req.data || {});
    } else {
      if (!SHEETS[coll]) return { ok: false, code: 'error', message: '未知的資料類別' };
      if (req.action === 'add') {
        id = newId();
        var d = req.data || {};
        if (coll === 'bookings') d.createdAt = Date.now();
        putRow(coll, id, d, false);
      } else if (req.action === 'update') {
        if (findRow(sheetOf(coll), req.id) < 0) return { ok: false, code: 'not_found' };
        putRow(coll, String(req.id), req.data || {}, true);
      } else if (req.action === 'delete') {
        var sh = sheetOf(coll), r = findRow(sh, req.id);
        if (r >= 0) sh.deleteRow(r + 1);
      } else {
        return { ok: false, code: 'error', message: '未知的動作' };
      }
    }
    SpreadsheetApp.flush();
    return { ok: true, id: id, data: loadAll() };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 驗證 ---------- */

function hash(pw) {
  var salt = PropertiesService.getScriptProperties().getProperty('SALT');
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + pw, Utilities.Charset.UTF_8));
}

function sign(exp) {
  var p = PropertiesService.getScriptProperties();
  var key = p.getProperty('SECRET') + '|' + p.getProperty('PW_HASH');
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(String(exp), key, Utilities.Charset.UTF_8));
}

function login(pw) {
  var p = PropertiesService.getScriptProperties();
  var stored = p.getProperty('PW_HASH');
  if (!stored) return { ok: false, code: 'no_password' };
  var cache = CacheService.getScriptCache();
  var fails = +(cache.get('fails') || 0);
  if (fails >= MAX_FAILS) return { ok: false, code: 'locked' };
  if (!pw || hash(pw) !== stored) {
    Utilities.sleep(800);
    cache.put('fails', String(fails + 1), LOCK_SECONDS);
    return { ok: false, code: 'wrong_password' };
  }
  cache.remove('fails');
  var exp = Date.now() + TOKEN_DAYS * 864e5;
  return { ok: true, token: exp + '.' + sign(exp) };
}

function checkToken(t) {
  if (typeof t !== 'string' || !PropertiesService.getScriptProperties().getProperty('PW_HASH')) return false;
  var i = t.indexOf('.');
  if (i < 1) return false;
  var exp = t.slice(0, i);
  if (!/^\d+$/.test(exp) || +exp < Date.now()) return false;
  return t.slice(i + 1) === sign(exp);
}

/* ---------- 試算表讀寫 ---------- */

function book() {
  var id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}

function sheetOf(coll) {
  var def = SHEETS[coll], ss = book(), sh = ss.getSheetByName(def.title);
  if (!sh) {
    sh = ss.insertSheet(def.title);
    sh.getRange(1, 1, sh.getMaxRows(), def.cols.length).setNumberFormat('@');
    sh.getRange(1, 1, 1, def.cols.length).setValues([def.cols.map(function (c) { return c[1]; })]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function settingsSheet() {
  var ss = book(), sh = ss.getSheetByName(SETTINGS_TITLE);
  if (!sh) {
    sh = ss.insertSheet(SETTINGS_TITLE);
    sh.getRange(1, 1, 1, 3).setValues([['項目代碼', '值', '說明']]).setFontWeight('bold');
    sh.getRange(2, 1, SETTINGS.length, 3).setValues(SETTINGS.map(function (s) { return [s[0], '', s[1]]; }));
    sh.setFrozenRows(1);
  }
  return sh;
}

// 依標題列找欄位，店家調換欄位順序也不會壞
function columnsOf(sh, coll) {
  var header = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getDisplayValues()[0];
  var byLabel = {};
  SHEETS[coll].cols.forEach(function (c) { byLabel[c[1]] = c[0]; byLabel[c[0]] = c[0]; });
  return header.map(function (h) { return byLabel[String(h).trim()] || null; });
}

function findRow(sh, id) {
  if (!id || sh.getLastRow() < 2) return -1;
  var keys = columnsOf(sh, coll_(sh)), ci = keys.indexOf('id');
  var ids = sh.getRange(2, ci + 1, sh.getLastRow() - 1, 1).getDisplayValues();
  for (var i = 0; i < ids.length; i++) if (ids[i][0] === String(id)) return i + 1;
  return -1;
}

function coll_(sh) {
  var name = sh.getName();
  for (var k in SHEETS) if (SHEETS[k].title === name) return k;
  throw new Error('找不到工作表 ' + name);
}

function toCell(k, v) {
  if (k === 'guides') return (Array.isArray(v) ? v : []).map(String).join('、');
  if (k === 'depositPaid') return v ? '是' : '';
  if (k === 'createdAt') return v ? Utilities.formatDate(new Date(+v), TZ, 'yyyy-MM-dd HH:mm:ss') : '';
  if (NUM[k]) return (v === '' || v == null) ? '' : String(+v || 0);
  var s = String(v == null ? '' : v).slice(0, 5000);
  return s.charAt(0) === '=' ? ' ' + s : s;   // 避免被當成公式
}

function fromCell(k, v) {
  if (k === 'guides') return v ? v.split('、').map(function (s) { return s.trim(); }).filter(String) : [];
  if (k === 'depositPaid') return v === '是' || v === 'TRUE' || v === 'true';
  if (k === 'createdAt') {
    var t = v ? new Date(v.replace(' ', 'T') + '+08:00').getTime() : 0;
    return isNaN(t) ? 0 : t;
  }
  if (NUM[k]) return +v || 0;
  return v;
}

function readRows(coll) {
  var sh = sheetOf(coll);
  if (sh.getLastRow() < 2) return [];
  var keys = columnsOf(sh, coll);
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, keys.length).getDisplayValues();
  var out = [];
  vals.forEach(function (row) {
    var o = {};
    keys.forEach(function (k, i) { if (k) o[k] = fromCell(k, row[i]); });
    if (o.id) out.push(o);
  });
  return out;
}

function putRow(coll, id, data, merge) {
  var sh = sheetOf(coll), keys = columnsOf(sh, coll), r = findRow(sh, id);
  var allowed = {};
  SHEETS[coll].cols.forEach(function (c) { allowed[c[0]] = 1; });
  var cur = {};
  if (merge && r >= 0) {
    var old = sh.getRange(r + 1, 1, 1, keys.length).getDisplayValues()[0];
    keys.forEach(function (k, i) { if (k) cur[k] = fromCell(k, old[i]); });
  }
  Object.keys(data).forEach(function (k) { if (allowed[k] && k !== 'id') cur[k] = data[k]; });
  cur.id = id;
  var row = keys.map(function (k) { return k ? toCell(k, cur[k]) : ''; });
  if (r < 0) {
    var at = sh.getLastRow() + 1;
    var range = sh.getRange(at, 1, 1, keys.length);
    range.setNumberFormat('@').setValues([row]);
  } else {
    var target = sh.getRange(r + 1, 1, 1, keys.length);
    // 保留沒對應到欄位的儲存格（店家自己加的欄）
    var existing = target.getValues()[0];
    target.setValues([row.map(function (v, i) { return keys[i] ? v : existing[i]; })]);
  }
}

function readSettings() {
  var vals = settingsSheet().getDataRange().getDisplayValues(), o = { perGuide: 0 };
  vals.slice(1).forEach(function (r) { if (r[0] === 'perGuide') o.perGuide = +r[1] || 0; });
  return o;
}

function writeSettings(d) {
  var sh = settingsSheet(), vals = sh.getDataRange().getDisplayValues();
  SETTINGS.forEach(function (s) {
    if (!(s[0] in d)) return;
    var v = String(+d[s[0]] || 0), found = false;
    for (var i = 1; i < vals.length; i++) if (vals[i][0] === s[0]) { sh.getRange(i + 1, 2).setValue(v); found = true; }
    if (!found) sh.appendRow([s[0], v, s[1]]);
  });
}

function loadAll() {
  return { bookings: readRows('bookings'), guides: readRows('guides'), tours: readRows('tours'), settings: readSettings() };
}

function newId() {
  return Utilities.getUuid().replace(/-/g, '').slice(0, 16);
}
