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
    ['createdAt', '建立時間'], ['passengers', '旅客名單'], ['source', '來源'], ['lineUser', 'LINE 使用者ID'], ['boat', '船'],
    ['code', '團號'], ['inbox', '客人訊息'], ['unread', '有新訊息'], ['confirm', '客人確認'], ['confirmAt', '確認時間'],
    ['paxAt', '名單填寫時間'], ['paxReminded', '已提醒填名單']
  ]},
  guides: { title: '導覽員', cols: [['id', '編號'], ['name', '姓名'], ['note', '備註']] },
  tours: { title: '遊程', cols: [['id', '編號'], ['name', '遊程名稱'], ['meet', '集合地點'], ['content', '遊程內容'], ['boat', '搭船']] },
  boats: { title: '船隻', cols: [['id', '編號'], ['name', '船名'], ['capacity', '載客上限']],
    seed: [['金波羅', 34], ['航海家16號', 45], ['公主268號', 45], ['航海家18號', 42]] },
  // 一天一艘船一列；載客上限留空＝用「船隻」的預設值
  sailings: { title: '船班', cols: [['id', '編號'], ['date', '日期'], ['boat', '船名'], ['capacity', '當日載客上限'], ['stop', '停航'], ['note', '備註']] }
};
var SETTINGS_TITLE = '設定';
var SETTINGS = [['perGuide', '每位導覽員最多帶幾人']];
var NUM = { adult: 1, kid: 1, quote: 1, deposit: 1, capacity: 1 };

/* ---------- 試算表選單與初始化 ---------- */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('接團台')
    .addItem('設定／更改登入密碼', 'setPasswordMenu')
    .addItem('讓所有裝置重新登入', 'logoutAllMenu')
    .addSeparator()
    .addItem('產生指定日期的船班名冊', 'manifestMenu')
    .addItem('開啟名冊資料夾', 'openFolderMenu')
    .addSeparator()
    .addItem('設定 LINE 機器人', 'lineSetupMenu')
    .addItem('LINE 機器人測試名單', 'lineTestersMenu')
    .addItem('LINE 機器人正式上線／改回測試', 'lineGoLiveMenu')
    .addToUi();
}

function 初始化() {
  var p = PropertiesService.getScriptProperties();
  p.setProperty('SHEET_ID', SpreadsheetApp.getActiveSpreadsheet().getId());
  if (!p.getProperty('SALT')) p.setProperty('SALT', Utilities.getUuid());
  if (!p.getProperty('SECRET')) p.setProperty('SECRET', Utilities.getUuid() + Utilities.getUuid());
  Object.keys(SHEETS).forEach(sheetOf);
  settingsSheet();
  replySheet();
  lineLogSheet();
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

function doGet(e) {
  if (e && e.parameter && e.parameter.pax) {
    return ContentService.createTextOutput(JSON.stringify(paxGet_(String(e.parameter.pax), String(e.parameter.k || ''))))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.avail != null) {
    return ContentService.createTextOutput(JSON.stringify({ ok: true, days: availability_() })).setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput('蹦火仔接團台後端運作中');
}

/* ---------- 預約頁用：未來每天的名額狀態（不含船名與人數） ---------- */
var SEAT = ['已成單', '導覽員已排定', '已通知導覽員', '已完成'];
var FEW_SEATS = 10;   // 剩不到幾位顯示「名額有限」

function tourNeedsBoat_(t) {
  if (!t) return false;
  if (t.boat === '是') return true;
  if (t.boat === '否') return false;
  return /海上|海陸|船/.test(t.name || '');
}

function availability_() {
  var cache = CacheService.getScriptCache(), hit = cache.get('avail');
  if (hit) return JSON.parse(hit);
  var today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  var boatCap = {}, tours = {}, days = {};
  readRows('boats').forEach(function (b) { boatCap[b.name] = +b.capacity || 0; });
  readRows('tours').forEach(function (t) { tours[t.id] = t; });
  readRows('sailings').forEach(function (s) {
    if (s.date < today) return;
    var d = days[s.date] = days[s.date] || { cap: 0, open: 0, stop: 0, used: 0 };
    if (s.stop) { d.stop++; return; }
    d.open++; d.cap += +s.capacity || boatCap[s.boat] || 0;
  });
  readRows('bookings').forEach(function (b) {
    if (days[b.date] && SEAT.indexOf(b.status) >= 0 && tourNeedsBoat_(tours[b.tourId])) days[b.date].used += (+b.adult || 0) + (+b.kid || 0);
  });
  var out = {};
  Object.keys(days).forEach(function (k) {
    var d = days[k], left = d.cap - d.used;
    out[k] = !d.open ? (d.stop ? 'stop' : 'none') : left <= 0 ? 'full' : left < FEW_SEATS ? 'few' : 'open';
  });
  cache.put('avail', JSON.stringify(out), 300);
  return out;
}

function doPost(e) {
  if (e && e.parameter && e.parameter.line != null) return lineWebhook(e);
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
  if (req.action === 'paxSubmit') return paxSubmit_(String(req.id || ''), String(req.k || ''), req.passengers);
  if (!checkToken(req.token)) return { ok: false, code: 'auth' };
  if (req.action === 'load') return { ok: true, data: loadAll() };
  if (req.action === 'manifest') {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(String(req.date))) return { ok: false, code: 'error', message: '日期格式不對' };
    return { ok: true, manifest: makeManifest(req.date) };
  }
  if (req.action === 'linePush') return linePush_(String(req.id || ''), String(req.text || ''));
  if (req.action === 'lineConfirm') return lineConfirm_(String(req.id || ''));

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
        if (coll === 'bookings') { d.createdAt = Date.now(); d.code = nextCode_(); }
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
    CacheService.getScriptCache().remove('avail');
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
    (def.seed || []).forEach(function (r) { putRow(coll, newId(), { name: r[0], capacity: r[1] }, false); });
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
  var keys = header.map(function (h) { return byLabel[String(h).trim()] || null; });
  // 程式新增的欄位（例如「旅客名單」）自動補在最右邊
  SHEETS[coll].cols.forEach(function (c) {
    if (keys.indexOf(c[0]) >= 0) return;
    var at = keys.length + 1;
    sh.getRange(1, at, sh.getMaxRows(), 1).setNumberFormat('@');
    sh.getRange(1, at).setValue(c[1]).setFontWeight('bold');
    keys.push(c[0]);
  });
  return keys;
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
  if (k === 'passengers') {
    // 一行一位：姓名｜證號｜出生年月日｜電話
    var clean = function (x) { return String(x == null ? '' : x).replace(/[｜\r\n]/g, ' ').trim().slice(0, 100); };
    var s2 = (Array.isArray(v) ? v : []).slice(0, 500).map(function (p) {
      p = p || {};
      return [clean(p.name), clean(p.id), clean(p.birth), clean(p.phone)].join('｜');
    }).join('\n');
    return s2.charAt(0) === '=' ? ' ' + s2 : s2;
  }
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
  if (k === 'passengers') {
    return v ? v.split('\n').map(function (line) {
      var f = line.split('｜');
      return { name: (f[0] || '').trim(), id: (f[1] || '').trim(), birth: (f[2] || '').trim(), phone: (f[3] || '').trim() };
    }).filter(function (p) { return p.name || p.id; }) : [];
  }
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
  var bookings = readRows('bookings');
  bookings.forEach(function (b) { b.paxUrl = paxUrl_(b.id); });
  return { bookings: bookings, guides: readRows('guides'), tours: readRows('tours'), boats: readRows('boats'), sailings: readRows('sailings'), settings: readSettings() };
}

function newId() {
  return Utilities.getUuid().replace(/-/g, '').slice(0, 16);
}

/* ---------- 船班名冊（出發前自動產生，放在雲端硬碟） ---------- */

var ACTIVE = ['已成單', '導覽員已排定', '已通知導覽員'];
var MANIFEST_FOLDER = '蹦火仔船班名冊';
var MANIFEST_HOUR = 7;   // 每天幾點自動產生

// 執行一次即可：每天早上自動產生 3 天後的名冊，並更新 1、2 天後的名冊
function 設定每日排程() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'dailyManifest') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dailyManifest').timeBased().everyDays(1).atHour(MANIFEST_HOUR).inTimezone(TZ).create();
  manifestFolder();
  dailyManifest();
  Logger.log('已設定：每天 ' + MANIFEST_HOUR + ' 點自動產生船班名冊，放在雲端硬碟「' + MANIFEST_FOLDER + '」資料夾。');
}

function dailyManifest() {
  [3, 2, 1].forEach(function (d) {
    makeManifest(Utilities.formatDate(new Date(Date.now() + d * 864e5), TZ, 'yyyy-MM-dd'));
  });
  try { paxReminders_(); } catch (err) { console.error('名單提醒失敗', err); }
}

function manifestFolder() {
  var p = PropertiesService.getScriptProperties(), id = p.getProperty('FOLDER_ID');
  if (id) {
    try { var f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) {}
  }
  var it = DriveApp.getFoldersByName(MANIFEST_FOLDER);
  var folder = it.hasNext() ? it.next() : DriveApp.createFolder(MANIFEST_FOLDER);
  p.setProperty('FOLDER_ID', folder.getId());
  return folder;
}

function safe_(v) {
  var s = String(v == null ? '' : v);
  return s.charAt(0) === '=' ? ' ' + s : s;
}

// 產生某一天的名冊（試算表＋PDF），同名的舊檔會移到垃圾桶。沒有團就不產生。
function makeManifest(dateStr) {
  var data = loadAll(), tours = {}, boats = {};
  data.tours.forEach(function (t) { tours[t.id] = t; });
  data.boats.forEach(function (t) { boats[t.id] = t.name; });
  var boatOf = function (b) { return boats[b.boat] || ''; };
  // 同一艘船的團排在一起，沒排船的放最後
  var list = data.bookings.filter(function (b) { return b.date === dateStr && ACTIVE.indexOf(b.status) >= 0; })
    .sort(function (a, b) {
      return (boatOf(a) ? 0 : 1) - (boatOf(b) ? 0 : 1) || boatOf(a).localeCompare(boatOf(b)) ||
        (a.time || '').localeCompare(b.time || '') || (a.createdAt || 0) - (b.createdAt || 0);
    });

  var folder = manifestFolder(), title = '蹦火仔船班名冊 ' + dateStr;
  [title, title + '.pdf'].forEach(function (n) {
    var it = folder.getFilesByName(n);
    while (it.hasNext()) it.next().setTrashed(true);
  });
  if (!list.length) return { date: dateStr, groups: 0, count: 0, listed: 0 };

  var rows = [], kinds = [], total = 0, listed = 0, seq = 0;
  list.forEach(function (b) {
    var n = (+b.adult || 0) + (+b.kid || 0), pax = b.passengers || [], t = tours[b.tourId];
    total += n; listed += pax.length;
    var who = b.org ? b.org + '（' + (b.name || '未填聯絡人') + '）' : (b.name || '未填聯絡人');
    var line = (boatOf(b) ? '【' + boatOf(b) + '】' : '') + (b.time ? b.time + ' 集合　' : '') + who + (t ? '　' + t.name : '') + '　共 ' + n + ' 人' + (b.phone ? '　聯絡電話 ' + b.phone : '');
    rows.push([safe_(line), '', '', '', '']); kinds.push('group');
    pax.forEach(function (p) {
      seq++;
      rows.push([String(seq), safe_(p.name), safe_(p.id), safe_(p.birth), safe_(p.phone)]); kinds.push('pax');
    });
    if (pax.length < n) { rows.push(['', '（尚缺 ' + (n - pax.length) + ' 位旅客資料）', '', '', '']); kinds.push('lack'); }
  });

  var d = dateStr.split('-');
  var wd = '日一二三四五六'.charAt(new Date(+d[0], +d[1] - 1, +d[2]).getDay());
  var head = [
    ['蹦火仔 船班旅客名冊　' + d[0] + '/' + (+d[1]) + '/' + (+d[2]) + '（' + wd + '）', '', '', '', ''],
    ['共 ' + list.length + ' 團、' + total + ' 人，名單已填 ' + listed + ' 人　｜　產生時間 ' + Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'), '', '', '', ''],
    ['序號', '姓名', '身分證字號／護照', '出生年月日', '聯絡電話']
  ];
  var all = head.concat(rows), H = head.length;

  var ss = SpreadsheetApp.create(title);
  DriveApp.getFileById(ss.getId()).moveTo(folder);
  var sh = ss.getSheets()[0];
  sh.setName('名冊');
  if (sh.getMaxColumns() > 5) sh.deleteColumns(6, sh.getMaxColumns() - 5);
  if (sh.getMaxRows() < all.length) sh.insertRowsAfter(sh.getMaxRows(), all.length - sh.getMaxRows());
  if (sh.getMaxRows() > all.length) sh.deleteRows(all.length + 1, sh.getMaxRows() - all.length);
  sh.getRange(1, 1, all.length, 5).setNumberFormat('@').setValues(all).setFontSize(11).setVerticalAlignment('middle');
  [56, 120, 170, 120, 130].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.getRange(1, 1, 1, 5).merge().setFontSize(15).setFontWeight('bold');
  sh.getRange(2, 1, 1, 5).merge().setFontColor('#555555');
  sh.getRange(3, 1, 1, 5).setFontWeight('bold').setBackground('#eeeeee');
  sh.getRange(3, 1, all.length - 2, 5).setBorder(true, true, true, true, true, true, '#999999', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(H + 1, 1, rows.length, 1).setHorizontalAlignment('center');
  kinds.forEach(function (k, i) {
    var r = H + 1 + i;
    if (k === 'group') sh.getRange(r, 1, 1, 5).merge().setFontWeight('bold').setBackground('#fde7dc').setHorizontalAlignment('left');
    if (k === 'lack') sh.getRange(r, 2, 1, 4).merge().setFontColor('#b42318');
  });
  sh.setFrozenRows(3);
  SpreadsheetApp.flush();

  var pdfUrl = '', blob = null, pdfError = [];
  try {
    // 排版較好的匯出方式（A4 直式、縮放到頁寬、每頁重複標題列）
    var url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=pdf&size=A4&portrait=true&fitw=true' +
      '&gridlines=false&printtitle=false&sheetnames=false&pagenum=CENTER&fzr=true' +
      '&top_margin=0.5&bottom_margin=0.5&left_margin=0.5&right_margin=0.5';
    var res = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
    if (res.getResponseCode() === 200) blob = res.getBlob();
    else pdfError.push('匯出回應 ' + res.getResponseCode());
  } catch (e) {
    pdfError.push('匯出：' + e.message);
  }
  try {
    if (!blob) blob = DriveApp.getFileById(ss.getId()).getAs(MimeType.PDF);
    pdfUrl = folder.createFile(blob.setName(title + '.pdf')).getUrl();
  } catch (e) {
    pdfError.push('轉檔：' + e.message);
  }
  if (pdfError.length) Logger.log('PDF 問題：' + pdfError.join('；'));
  return { date: dateStr, groups: list.length, count: total, listed: listed, sheetUrl: ss.getUrl(), pdfUrl: pdfUrl,
    pdfError: pdfUrl ? '' : pdfError.join('；') };
}

// 從編輯器執行，下方「執行紀錄」會顯示結果與 PDF 的錯誤原因
function 測試產生PDF() {
  var d = Utilities.formatDate(new Date(Date.now() + 3 * 864e5), TZ, 'yyyy-MM-dd');
  var list = loadAll().bookings.filter(function (b) { return ACTIVE.indexOf(b.status) >= 0 && b.date >= Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); });
  if (list.length) d = list.map(function (b) { return b.date; }).sort()[0];
  var res = makeManifest(d);
  Logger.log(d + '：' + (res.groups ? (res.pdfUrl ? 'PDF 產生成功 ' + res.pdfUrl : 'PDF 失敗，原因：' + res.pdfError) : '這天沒有已成單的團'));
}

function manifestMenu() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('產生船班名冊', '請輸入日期，例如 2026-10-20 或 10/20', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var t = r.getResponseText().trim(), m, d = null;
  if ((m = t.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})$/))) d = [m[1], m[2], m[3]];
  else if ((m = t.match(/^(\d{1,2})[-\/.](\d{1,2})$/))) d = [Utilities.formatDate(new Date(), TZ, 'yyyy'), m[1], m[2]];
  if (!d) { ui.alert('看不懂這個日期：' + t); return; }
  var res = makeManifest(d[0] + '-' + ('0' + d[1]).slice(-2) + '-' + ('0' + d[2]).slice(-2));
  if (!res.groups) { ui.alert(res.date + ' 沒有已成單的團，沒有產生名冊。'); return; }
  showLinks_('名冊已產生', res.date + '：' + res.groups + ' 團、' + res.count + ' 人（名單已填 ' + res.listed + ' 人）',
    [['開啟 PDF（可直接列印）', res.pdfUrl], ['開啟試算表版', res.sheetUrl]]);
}

function openFolderMenu() {
  showLinks_('船班名冊資料夾', '每天早上 ' + MANIFEST_HOUR + ' 點自動產生 3 天後的名冊。', [['開啟資料夾', manifestFolder().getUrl()]]);
}

function showLinks_(title, text, links) {
  var esc = function (s) {
    return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; });
  };
  var html = '<div style="font-family:sans-serif;font-size:14px"><p>' + esc(text) + '</p>' +
    links.filter(function (l) { return l[1]; }).map(function (l) {
      return '<p><a href="' + esc(l[1]) + '" target="_blank">' + esc(l[0]) + '</a></p>';
    }).join('') + '</div>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(360).setHeight(180), title);
}

/* ---------- LINE 機器人：自動回覆、收預約詢問建立團單 ---------- */
// 客人傳「【蹦火仔 預約詢問】」→ 建立「待確認」團單並回覆已收到。
// 其他訊息比對「自動回覆」工作表的關鍵字，對到才回；沒對到就不回，留給真人。
// 所有收到的訊息都記在「LINE 訊息」工作表。

var REPLY_TITLE = '自動回覆';
var LINE_LOG_TITLE = 'LINE 訊息';
var INQUIRY_MARK = '【蹦火仔 預約詢問】';
var FOLLOW_KEY = '（加好友）';
var BOOKING_PAGE = 'https://a24982268-afk.github.io/fuji268/';

var DEFAULT_REPLIES = [
  [FOLLOW_KEY, '您好，歡迎來到蹦火仔！\n想預約遊程，請到預約頁填寫後按「用 LINE 傳送」，我們收到會盡快回覆：\n' + BOOKING_PAGE + '\n\n也可以直接傳「價格」「集合地點」「季節」等關鍵字查詢。'],
  ['價格,價錢,費用,多少錢,收費,票價', '2026 遊程價格（全票／7–12 歲／6 歲以下）：\n・陸上漁村導覽 $650／$500／$350\n・海上蹦火導覽 $1,800／$1,400／$900\n・海陸雙響蹦 $2,300／$1,800／$1,150\n公司客製行程請直接告訴我們需求。'],
  ['集合,地點,在哪,地址,怎麼去', '在煮飯間集合。陸上與海陸遊程 14:20 集合，海上遊程 16:20 集合。'],
  ['多久,時數,幾個小時', '陸上約 2.5 小時，海上約 4.5 小時，海陸雙響蹦約 6.5 小時。'],
  ['季節,幾月,期間,開放,四季', '陸上漁村導覽全年皆有。海上蹦火導覽與海陸雙響蹦為季節限定，2026 年期間是 5/9 至 9/27。'],
  ['包團', '平日皆可包團，請到預約頁填寫日期與人數後傳給我們：\n' + BOOKING_PAGE],
  ['天氣,下雨,颱風,海況', '為了安全，本公司保有依情況調整行程的權利與義務。實際安排會在確認時與您說明。'],
  ['預約,報名,訂位,怎麼訂,想訂', '請到預約頁填寫，最後按「用 LINE 傳送」，我們收到會盡快回覆確認：\n' + BOOKING_PAGE],
  ['電話,客服', '富吉漁創有限公司 電話：0955-058-522']
];

function replySheet() {
  var ss = book(), sh = ss.getSheetByName(REPLY_TITLE);
  if (!sh) {
    sh = ss.insertSheet(REPLY_TITLE);
    sh.getRange(1, 1, 1, 2).setValues([['關鍵字（用逗號分開，對到任一個就回）', '回覆內容（留空＝不回）']]).setFontWeight('bold');
    sh.getRange(2, 1, DEFAULT_REPLIES.length, 2).setNumberFormat('@').setValues(DEFAULT_REPLIES).setWrap(true);
    sh.setColumnWidth(1, 220); sh.setColumnWidth(2, 420);
    sh.setFrozenRows(1);
  }
  return sh;
}

function lineLogSheet() {
  var ss = book(), sh = ss.getSheetByName(LINE_LOG_TITLE);
  if (!sh) {
    sh = ss.insertSheet(LINE_LOG_TITLE);
    sh.getRange(1, 1, sh.getMaxRows(), 5).setNumberFormat('@');
    sh.getRange(1, 1, 1, 5).setValues([['時間', 'LINE 名稱', 'LINE 使用者ID', '訊息', '機器人處理']]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function lineSetupMenu() {
  var ui = SpreadsheetApp.getUi(), p = PropertiesService.getScriptProperties();
  var r = ui.prompt('設定 LINE 機器人', '貼上 LINE Developers 的 Channel access token（長期）。\n留空直接按確定＝只重新顯示 Webhook 網址。', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  初始化();
  var tok = r.getResponseText().trim();
  if (tok) p.setProperty('LINE_TOKEN', tok);
  if (!p.getProperty('LINE_TOKEN')) { ui.alert('還沒有 Channel access token，機器人無法回覆。'); return; }
  if (!p.getProperty('LINE_KEY')) p.setProperty('LINE_KEY', Utilities.getUuid().replace(/-/g, ''));
  var url = ScriptApp.getService().getUrl();
  var key = '?line=' + p.getProperty('LINE_KEY');
  ui.alert('LINE 機器人已設定',
    '請把下面的網址貼到 LINE Developers →「Messaging API」→「Webhook URL」，並開啟「Use webhook」：\n\n' +
    (url ? url + key : '（網頁應用程式網址）' + key + '\n\n找不到網址時：「部署」→「管理部署作業」複製網頁應用程式網址，後面接上 ' + key) +
    '\n\n這串網址等同機器人的鑰匙，不要公開。', ui.ButtonSet.OK);
}

// 還沒「正式上線」前，機器人只處理測試名單上的 LINE 帳號；其他人的訊息只記錄、不回覆不建單（照舊由真人回）
function lineLive_() { return PropertiesService.getScriptProperties().getProperty('LINE_LIVE') === '1'; }

function lineTesters_() {
  return String(PropertiesService.getScriptProperties().getProperty('LINE_TESTERS') || '')
    .split(/[\s,，、]+/).filter(function (x) { return /^U[0-9a-f]{32}$/.test(x); });
}

function lineTestersMenu() {
  var ui = SpreadsheetApp.getUi(), p = PropertiesService.getScriptProperties(), cur = lineTesters_();
  var r = ui.prompt('LINE 機器人測試名單',
    '目前：' + (lineLive_() ? '已正式上線，對所有人作用' : '測試中，只對名單上 ' + cur.length + ' 個帳號作用') + '\n\n' +
    '貼上要測試的 LINE 使用者ID（U 開頭，多個用逗號分開）。先用自己的手機傳一則訊息給官方帳號，' +
    '再到「LINE 訊息」工作表 C 欄複製自己的ID。\n\n設定名單會讓機器人回到測試模式。', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var ids = r.getResponseText().split(/[\s,，、]+/).filter(function (x) { return /^U[0-9a-f]{32}$/.test(x); });
  if (r.getResponseText().trim() && !ids.length) { ui.alert('看不懂這些ID，沒有變更。ID 是 U 開頭、後面 32 個英數字。'); return; }
  p.setProperty('LINE_TESTERS', ids.join(','));
  p.deleteProperty('LINE_LIVE');
  ui.alert('測試模式：機器人只對名單上 ' + ids.length + ' 個帳號作用，其他遊客不受影響。');
}

function lineGoLiveMenu() {
  var ui = SpreadsheetApp.getUi(), p = PropertiesService.getScriptProperties();
  if (lineLive_()) {
    if (ui.alert('LINE 機器人', '目前已正式上線。要改回測試模式（只對測試名單作用）嗎？', ui.ButtonSet.YES_NO) === ui.Button.YES) {
      p.deleteProperty('LINE_LIVE'); ui.alert('已改回測試模式。');
    }
    return;
  }
  if (ui.alert('LINE 機器人正式上線', '上線後，所有遊客傳給官方帳號的訊息都會由機器人處理（自動回覆、建立團單）。\n' +
    '請先到 LINE 官方帳號後台關閉「自動回應訊息」，避免重複回覆。確定上線嗎？', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  p.setProperty('LINE_LIVE', '1');
  ui.alert('機器人已正式上線。');
}

function lineWebhook(e) {
  var p = PropertiesService.getScriptProperties();
  var key = p.getProperty('LINE_KEY');
  // Apps Script 讀不到 LINE 的簽章標頭，改用網址上的密鑰確認是 LINE 打來的
  if (key && e.parameter.line === key) {
    var events = [];
    try { events = JSON.parse(e.postData.contents).events || []; } catch (err) {}
    events.forEach(function (ev) {
      try { lineEvent_(ev); } catch (err) { console.error('LINE 事件處理失敗', err); }
    });
  }
  return ContentService.createTextOutput('ok');
}

function lineEvent_(ev) {
  var who = (ev.source && ev.source.userId) || '';
  if (!lineLive_() && lineTesters_().indexOf(who) < 0) {
    if (ev.type === 'message' && ev.message && ev.message.type === 'text') {
      var t = String(ev.message.text || '').slice(0, 5000);
      lineLogSheet().appendRow([Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'), '', who,
        t.charAt(0) === '=' ? ' ' + t : t, '測試模式：不處理']);
    }
    return;
  }
  if (ev.type === 'follow') {
    var w = matchReply_(FOLLOW_KEY, true);
    if (w) lineReply_(ev.replyToken, w.text);
    return;
  }
  if (ev.type === 'postback') { linePostback_(ev); return; }
  if (ev.type !== 'message' || !ev.message || ev.message.type !== 'text') return;
  // LINE 重送同一則訊息時不要重複建團單
  var cache = CacheService.getScriptCache(), seen = 'ev_' + (ev.webhookEventId || ev.message.id);
  if (cache.get(seen)) return;
  cache.put(seen, '1', 21600);

  var text = String(ev.message.text || ''), uid = (ev.source && ev.source.userId) || '';
  var name = lineName_(uid), handled = '';
  if (text.indexOf(INQUIRY_MARK) >= 0) {
    var b = parseInquiry_(text);
    b.status = '待確認'; b.depositPaid = false; b.guides = []; b.createdAt = Date.now();
    b.source = 'LINE'; b.lineUser = uid;
    if (name) b.note = 'LINE 名稱：' + name + (b.note ? '\n' + b.note : '');
    var id = newId(), lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try { b.code = nextCode_(); putRow('bookings', id, b, false); SpreadsheetApp.flush(); } finally { lock.releaseLock(); }
    var n = (+b.adult || 0) + (+b.kid || 0);
    lineReply_(ev.replyToken, (b.name || name || '您') + '您好，這裡是蹦火仔，已收到您' +
      (b.date ? ' ' + Utilities.formatDate(new Date(b.date + 'T00:00:00+08:00'), TZ, 'M/d') + ' ' : '') +
      '的預約詢問' + (n ? '（共 ' + n + ' 位）' : '') + '，團號 ' + b.code + '。我們確認導覽員後會盡快回覆您，之後有任何修改直接在這裡告訴我們即可，謝謝！');
    handled = '建立待確認團單 ' + b.code;
  } else {
    var m = matchReply_(text, false);
    if (m) { lineReply_(ev.replyToken, m.text); handled = '自動回覆：' + m.key; }
    // 客人已有團單：訊息掛到團單上，沒被自動回覆的標「有新訊息」
    var mine = activeBookingOf_(uid);
    if (mine) {
      addInbox_(mine.id, text, !m);
      handled = (handled ? handled + '；' : '') + '記到團單 ' + (mine.code || mine.name || '');
    }
  }
  var safeText = text.slice(0, 5000);
  lineLogSheet().appendRow([Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'), name, uid,
    safeText.charAt(0) === '=' ? ' ' + safeText : safeText, handled]);
}

// 回傳第一列對到的 {key, text}；exact=true 時關鍵字要完全相同（給「（加好友）」用）
function matchReply_(text, exact) {
  var rows = replySheet().getDataRange().getDisplayValues().slice(1);
  for (var i = 0; i < rows.length; i++) {
    var reply = String(rows[i][1] || '').trim();
    if (!reply) continue;
    var keys = String(rows[i][0] || '').split(/[,，、]/).map(function (k) { return k.trim(); }).filter(String);
    for (var j = 0; j < keys.length; j++) {
      if (exact ? text === keys[j] : (keys[j] !== FOLLOW_KEY && text.indexOf(keys[j]) >= 0)) return { key: keys[j], text: reply };
    }
  }
  return null;
}

function lineReply_(token, text) {
  var tok = PropertiesService.getScriptProperties().getProperty('LINE_TOKEN');
  if (!tok || !token) return;
  var res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + tok },
    payload: JSON.stringify({ replyToken: token, messages: [{ type: 'text', text: String(text).slice(0, 5000) }] })
  });
  if (res.getResponseCode() !== 200) console.error('LINE 回覆失敗', res.getResponseCode(), res.getContentText());
}

function lineName_(uid) {
  if (!uid) return '';
  var cache = CacheService.getScriptCache(), k = 'name_' + uid, hit = cache.get(k);
  if (hit != null) return hit;
  var tok = PropertiesService.getScriptProperties().getProperty('LINE_TOKEN'), name = '';
  if (tok) {
    var res = UrlFetchApp.fetch('https://api.line.me/v2/bot/profile/' + encodeURIComponent(uid),
      { headers: { Authorization: 'Bearer ' + tok }, muteHttpExceptions: true });
    if (res.getResponseCode() === 200) name = JSON.parse(res.getContentText()).displayName || '';
  }
  cache.put(k, name, 21600);
  return name;
}

// 跟接團台網頁「貼上預約內容」用同一套規則
function parseInquiry_(txt) {
  var o = {}, notes = [];
  txt.split(/\r?\n/).forEach(function (line) {
    var i = line.search(/[：:]/);
    if (i < 0) return;
    var k = line.slice(0, i).trim(), v = line.slice(i + 1).trim();
    if (k === '類型' && /散客|學校|公司團體/.test(v)) o.type = v.match(/散客|學校|公司團體/)[0];
    else if (k === '學校' || k === '單位') o.org = v;
    else if (k === '聯絡人' && v.indexOf('（未填）') < 0) o.name = v;
    else if (k === '手機' && v.indexOf('（未填）') < 0) o.phone = v;
    else if (k === '首選日期') { var m = v.match(/(\d{4})\/(\d{2})\/(\d{2})/); if (m) o.date = m[1] + '-' + m[2] + '-' + m[3]; }
    else if (k === '抵達時段') o.slot = v;
    else if (k === '人數') {
      var a = v.match(/(?:全票|大人)\s*(\d+)/), b = v.match(/7[–-]12\s*歲\s*(\d+)/), c = v.match(/6\s*歲以下\s*(\d+)/), kk = v.match(/小孩\s*(\d+)/);
      if (a) o.adult = +a[1];
      if (b || c) o.kid = (b ? +b[1] : 0) + (c ? +c[1] : 0); else if (kk) o.kid = +kk[1];
      if (b || c) notes.push('人數明細：' + v.replace(/，共.*$/, ''));
    }
    else if (k === '遊程') {
      readRows('tours').forEach(function (t) { if (t.name === v) o.tourId = t.id; });
      if (!o.tourId) notes.push('遊程：' + v);
    }
    else if (k === '預估費用') { var q = v.replace(/[^\d]/g, ''); if (q) o.quote = +q; }
    else if (k === '交通') o.bus = v;
    else if (k === '用餐') o.meal = v;
    else if (['備案日期', 'LINE', '年級／班級', '帶隊老師', '單據', '備註'].indexOf(k) >= 0) notes.push(k + '：' + v);
  });
  if (notes.length) o.note = notes.join('\n');
  return o;
}

// 接團台按「用 LINE 傳給客人」：用推播傳到這張團單記錄的 LINE 帳號（會用掉官方帳號的每月推播額度）
function linePush_(id, text) {
  var tok = PropertiesService.getScriptProperties().getProperty('LINE_TOKEN');
  if (!tok) return { ok: false, code: 'no_line', message: '還沒設定 LINE 機器人' };
  text = text.trim().slice(0, 5000);
  if (!text) return { ok: false, code: 'error', message: '訊息是空的' };
  var b = null;
  readRows('bookings').forEach(function (r) { if (r.id === id) b = r; });
  if (!b) return { ok: false, code: 'not_found' };
  if (!b.lineUser) return { ok: false, code: 'error', message: '這張團單不是從 LINE 來的' };
  var r = pushTo_(b.lineUser, [{ type: 'text', text: text }]);
  if (!r.ok) return r;
  lineLogSheet().appendRow([Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'), '（接團台傳送）', b.lineUser,
    text.charAt(0) === '=' ? ' ' + text : text, '傳給 ' + (b.name || b.org || '')]);
  return { ok: true };
}

function pushTo_(uid, messages) {
  var tok = PropertiesService.getScriptProperties().getProperty('LINE_TOKEN');
  if (!tok) return { ok: false, code: 'no_line', message: '還沒設定 LINE 機器人' };
  var res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + tok },
    payload: JSON.stringify({ to: uid, messages: messages })
  });
  var c = res.getResponseCode();
  if (c === 200) return { ok: true };
  console.error('LINE 推播失敗', c, res.getContentText());
  var why = c === 429 ? '本月推播額度已用完，請改用複製後到 LINE 貼上' : c === 400 ? '客人可能已封鎖官方帳號' : 'LINE 回應 ' + c;
  return { ok: false, code: 'line', message: why };
}

/* ---------- 團號與客人訊息 ---------- */

// 團號＝建立日 MMdd－當天第幾團，例如 1018-03（呼叫前要先拿到鎖）
function nextCode_() {
  var p = PropertiesService.getScriptProperties(), day = Utilities.formatDate(new Date(), TZ, 'MMdd');
  var cur = String(p.getProperty('CODE_SEQ') || '').split(':');
  var n = cur[0] === day ? (+cur[1] || 0) + 1 : 1;
  p.setProperty('CODE_SEQ', day + ':' + n);
  return day + '-' + (n < 10 ? '0' + n : n);
}

function getBooking_(id) {
  var b = null;
  readRows('bookings').forEach(function (r) { if (r.id === id) b = r; });
  return b;
}

// 這個 LINE 帳號目前進行中的團單：還沒出發的取最近一團，否則取最新建立的
function activeBookingOf_(uid) {
  if (!uid) return null;
  var today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  var list = readRows('bookings').filter(function (b) {
    return b.lineUser === uid && b.status !== '取消' && b.status !== '已完成' && (!b.date || b.date >= today);
  });
  list.sort(function (a, b) { return (a.date || '9999').localeCompare(b.date || '9999') || b.createdAt - a.createdAt; });
  return list[0] || null;
}

function addInbox_(id, text, unread) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var b = getBooking_(id);
    if (!b) return;
    var line = Utilities.formatDate(new Date(), TZ, 'M/d HH:mm') + '　' + String(text).replace(/\r?\n/g, ' ').slice(0, 500);
    var all = (b.inbox ? b.inbox + '\n' : '') + line;
    var d = { inbox: all.length > 4000 ? all.slice(all.length - 4000).replace(/^[^\n]*\n/, '') : all };
    if (unread) d.unread = '是';
    putRow('bookings', id, d, true);
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
}

function setBooking_(id, d) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { putRow('bookings', id, d, true); SpreadsheetApp.flush(); } finally { lock.releaseLock(); }
}

/* ---------- 成單確認卡（LINE） ---------- */

function mdw_(d) {
  if (!d) return '日期未定';
  var p = d.split('-'), w = '日一二三四五六'.charAt(new Date(+p[0], +p[1] - 1, +p[2]).getDay());
  return (+p[1]) + '/' + (+p[2]) + '（' + w + '）';
}

function lineConfirm_(id) {
  var b = getBooking_(id);
  if (!b) return { ok: false, code: 'not_found' };
  if (!b.lineUser) return { ok: false, code: 'error', message: '這張團單不是從 LINE 來的' };
  var t = null;
  readRows('tours').forEach(function (x) { if (x.id === b.tourId) t = x; });
  var n = (+b.adult || 0) + (+b.kid || 0);
  var rows = [
    ['日期', mdw_(b.date)], ['集合時間', b.time || '待定'], ['集合地點', (t && t.meet) || '待定'],
    ['遊程', (t && t.name) || '待定'], ['人數', n + ' 位（大人 ' + (+b.adult || 0) + '／小孩 ' + (+b.kid || 0) + '）']
  ];
  if (b.quote) rows.push(['費用', '$' + (+b.quote).toLocaleString('en-US')]);
  if (b.deposit) rows.push(['訂金', '$' + (+b.deposit).toLocaleString('en-US') + (b.depositPaid ? '（已收）' : '（未收）')]);
  var body = [
    { type: 'text', text: '行程確認', weight: 'bold', size: 'lg' },
    { type: 'text', text: (b.name || b.org || '') + (b.code ? '　團號 ' + b.code : ''), size: 'sm', color: '#888888', wrap: true }
  ].concat(rows.map(function (r) {
    return { type: 'box', layout: 'baseline', spacing: 'sm', contents: [
      { type: 'text', text: r[0], size: 'sm', color: '#888888', flex: 2 },
      { type: 'text', text: String(r[1]), size: 'sm', wrap: true, flex: 5 }] };
  })).concat([{ type: 'text', text: '內容正確請按「確認無誤」；有要修改的地方請按「需要修改」。', size: 'xs', color: '#888888', wrap: true, margin: 'md' }]);
  var msg = { type: 'flex', altText: '蹦火仔 行程確認：' + mdw_(b.date) + '，共 ' + n + ' 位', contents: {
    type: 'bubble',
    body: { type: 'box', layout: 'vertical', spacing: 'sm', contents: body },
    footer: { type: 'box', layout: 'vertical', spacing: 'sm', contents: [
      { type: 'button', style: 'primary', color: '#c2410c', action: { type: 'postback', label: '確認無誤', data: 'ok:' + b.id, displayText: '確認無誤' } },
      { type: 'button', style: 'secondary', action: { type: 'postback', label: '需要修改', data: 'chg:' + b.id, displayText: '需要修改' } },
      { type: 'button', style: 'link', action: { type: 'uri', label: '填寫旅客名單', uri: paxUrl_(b.id) } }
    ] }
  } };
  var r = pushTo_(b.lineUser, [msg]);
  if (!r.ok) return r;
  setBooking_(b.id, { confirm: '已傳送', confirmAt: '' });
  lineLogSheet().appendRow([Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'), '（接團台傳送）', b.lineUser,
    '【行程確認卡】' + rows.map(function (x) { return x[0] + '：' + x[1]; }).join('；'), '傳給 ' + (b.name || b.org || '')]);
  return { ok: true, data: loadAll() };
}

function linePostback_(ev) {
  var data = String((ev.postback && ev.postback.data) || ''), uid = (ev.source && ev.source.userId) || '';
  var m = data.match(/^(ok|chg):(\w+)$/);
  if (!m) return;
  var b = getBooking_(m[2]);
  if (!b || b.lineUser !== uid) { lineReply_(ev.replyToken, '找不到這筆預約，請直接傳訊息給我們，謝謝！'); return; }
  var stamp = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
  if (m[1] === 'ok') {
    setBooking_(b.id, { confirm: '已確認', confirmAt: stamp });
    addInbox_(b.id, '（按了「確認無誤」）', false);
    var lack = (+b.adult || 0) + (+b.kid || 0) - (b.passengers || []).length;
    lineReply_(ev.replyToken, '已收到您的確認，謝謝！' +
      (lack > 0 ? '\n\n還需要 ' + lack + ' 位旅客的資料（印船班名冊用），請由此填寫：\n' + paxUrl_(b.id) : '\n我們出發前會再提醒您集合資訊。'));
  } else {
    setBooking_(b.id, { confirm: '要修改', confirmAt: stamp });
    addInbox_(b.id, '（按了「需要修改」）', true);
    lineReply_(ev.replyToken, '好的，請直接在這裡告訴我們要修改的地方（日期、人數、集合時間…），我們會盡快處理並重新確認。');
  }
  lineLogSheet().appendRow([stamp, lineName_(uid), uid, m[1] === 'ok' ? '（確認無誤）' : '（需要修改）', '團單 ' + (b.code || '')]);
}

/* ---------- 客人填寫旅客名單 ---------- */

var PAX_PAGE = 'https://a24982268-afk.github.io/fuji268/pax.html';

function paxKey_(id) {
  var p = PropertiesService.getScriptProperties(), k = p.getProperty('PAX_KEY');
  if (!k) { k = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('PAX_KEY', k); }
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(String(id), k, Utilities.Charset.UTF_8)).slice(0, 22);
}

function paxUrl_(id) { return PAX_PAGE + '?b=' + encodeURIComponent(id) + '&k=' + paxKey_(id); }

// 連結有效：金鑰正確、團單沒取消、出發日隔天以前
function paxCheck_(id, k) {
  if (!id || !k || k !== paxKey_(id)) return { err: 'bad' };
  var b = getBooking_(id);
  if (!b || b.status === '取消') return { err: 'bad' };
  var today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  if (b.date && b.date < Utilities.formatDate(new Date(Date.now() - 864e5), TZ, 'yyyy-MM-dd')) return { err: 'expired' };
  return { b: b };
}

function paxGet_(id, k) {
  var c = paxCheck_(id, k);
  if (c.err) return { ok: false, code: c.err };
  var b = c.b, t = null;
  readRows('tours').forEach(function (x) { if (x.id === b.tourId) t = x; });
  return { ok: true, booking: { date: b.date, dateText: mdw_(b.date), tour: t ? t.name : '', count: (+b.adult || 0) + (+b.kid || 0),
    who: b.org || b.name || '', code: b.code || '', passengers: b.passengers || [], paxAt: b.paxAt || '' } };
}

function paxSubmit_(id, k, list) {
  var c = paxCheck_(id, k);
  if (c.err) return { ok: false, code: c.err };
  if (!Array.isArray(list) || !list.length) return { ok: false, code: 'error', message: '名單是空的' };
  var clean = function (x) { return String(x == null ? '' : x).replace(/[｜\r\n]/g, ' ').trim().slice(0, 60); };
  var pax = list.slice(0, 200).map(function (p) {
    p = p || {};
    return { name: clean(p.name), id: clean(p.id).toUpperCase(), birth: clean(p.birth), phone: clean(p.phone) };
  }).filter(function (p) { return p.name || p.id; });
  if (!pax.length) return { ok: false, code: 'error', message: '請至少填一位的姓名' };
  var stamp = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
  setBooking_(id, { passengers: pax, paxAt: stamp });
  addInbox_(id, '（客人在網頁填寫旅客名單，共 ' + pax.length + ' 位）', false);
  return { ok: true, count: pax.length };
}

// 每天早上跑：3 天後出發、名單還不齊、從 LINE 來的團，提醒一次
function paxReminders_() {
  var target = Utilities.formatDate(new Date(Date.now() + 3 * 864e5), TZ, 'yyyy-MM-dd');
  readRows('bookings').forEach(function (b) {
    if (b.date !== target || ACTIVE.indexOf(b.status) < 0 || !b.lineUser || b.paxReminded) return;
    var lack = (+b.adult || 0) + (+b.kid || 0) - (b.passengers || []).length;
    if (lack <= 0) return;
    var r = pushTo_(b.lineUser, [{ type: 'text', text: (b.name || '您') + '您好，這裡是蹦火仔。' + mdw_(b.date) +
      ' 的行程還需要 ' + lack + ' 位旅客的資料（印船班名冊用），麻煩您由此填寫，謝謝！\n' + paxUrl_(b.id) }]);
    if (r.ok) {
      setBooking_(b.id, { paxReminded: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm') });
      addInbox_(b.id, '（系統已提醒客人填寫旅客名單）', false);
    }
  });
}
