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
    ['createdAt', '建立時間'], ['passengers', '旅客名單']
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
    .addSeparator()
    .addItem('產生指定日期的船班名冊', 'manifestMenu')
    .addItem('開啟名冊資料夾', 'openFolderMenu')
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
  if (req.action === 'manifest') {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(String(req.date))) return { ok: false, code: 'error', message: '日期格式不對' };
    return { ok: true, manifest: makeManifest(req.date) };
  }

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
  return { bookings: readRows('bookings'), guides: readRows('guides'), tours: readRows('tours'), settings: readSettings() };
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
  var data = loadAll(), tours = {};
  data.tours.forEach(function (t) { tours[t.id] = t; });
  var list = data.bookings.filter(function (b) { return b.date === dateStr && ACTIVE.indexOf(b.status) >= 0; })
    .sort(function (a, b) { return (a.time || '').localeCompare(b.time || '') || (a.createdAt || 0) - (b.createdAt || 0); });

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
    var line = (b.time ? b.time + ' 集合　' : '') + who + (t ? '　' + t.name : '') + '　共 ' + n + ' 人' + (b.phone ? '　聯絡電話 ' + b.phone : '');
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
