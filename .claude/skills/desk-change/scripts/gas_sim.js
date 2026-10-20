#!/usr/bin/env node
// 在 Node 裡模擬 Google Apps Script 執行 Code.gs（試算表、屬性、快取、UrlFetch 都是假的）。
// 用法：node gas_sim.js apps-script/Code.gs        → 語法檢查＋冒煙測試
//      const sim = require('./gas_sim.js').load('apps-script/Code.gs'); sim.run("loadAll()")
'use strict';
const fs = require('fs');
const vm = require('vm');
const crypto = require('crypto');

function load(codePath, opts) {
  opts = opts || {};
  const src = fs.readFileSync(codePath, 'utf8');
  const sheets = {};
  const props = Object.assign({ SHEET_ID: 'sim' }, opts.props || {});
  const cache = {};
  const fetched = [];

  function mkSheet(name) {
    const data = [];
    const lastRow = () => { let n = data.length; while (n > 0 && !(data[n - 1] || []).some(x => String(x == null ? '' : x) !== '')) n--; return n; };
    const sh = {
      name, data,
      getName: () => name, getMaxRows: () => 1000, getMaxColumns: () => 26,
      getLastRow: lastRow,
      getLastColumn: () => Math.max(0, ...data.map(r => r.length)),
      getRange(r, c, nr = 1, nc = 1) {
        return {
          setNumberFormat() { return this; }, setFontWeight() { return this; }, setWrap() { return this; },
          setFontSize() { return this; }, setVerticalAlignment() { return this; }, setHorizontalAlignment() { return this; },
          setBackground() { return this; }, setFontColor() { return this; }, setBorder() { return this; }, merge() { return this; },
          setValues(v) { v.forEach((row, i) => { data[r - 1 + i] = data[r - 1 + i] || []; row.forEach((x, j) => { data[r - 1 + i][c - 1 + j] = x; }); }); return this; },
          setValue(x) { data[r - 1] = data[r - 1] || []; data[r - 1][c - 1] = x; return this; },
          getDisplayValues() { const o = []; for (let i = 0; i < nr; i++) { const row = []; for (let j = 0; j < nc; j++) row.push(String(((data[r - 1 + i] || [])[c - 1 + j]) ?? '')); o.push(row); } return o; },
          getValues() { return this.getDisplayValues(); },
          clearContent() { for (let i = 0; i < nr; i++) for (let j = 0; j < nc; j++) if (data[r - 1 + i]) data[r - 1 + i][c - 1 + j] = ''; return this; }
        };
      },
      getDataRange() { return this.getRange(1, 1, Math.max(lastRow(), 1), Math.max(this.getLastColumn(), 1)); },
      appendRow(row) { data.splice(lastRow(), 0, row); },
      deleteRow(n) { data.splice(n - 1, 1); },
      setFrozenRows() {}, setColumnWidth() {}, setName() {}, deleteColumns() {}, deleteRows() {}, insertRowsAfter() {}
    };
    return sh;
  }
  const ss = { getSheetByName: n => sheets[n], insertSheet: n => (sheets[n] = mkSheet(n)), getId: () => 'sim', getSheets: () => Object.values(sheets) };
  const fmt = (d, tz, f) => {
    const t = new Date(d.getTime() + 8 * 3600e3), P = n => String(n).padStart(2, '0');
    return f.replace('yyyy', t.getUTCFullYear()).replace('MM', P(t.getUTCMonth() + 1)).replace('dd', P(t.getUTCDate()))
      .replace('HH', P(t.getUTCHours())).replace('mm', P(t.getUTCMinutes())).replace('ss', P(t.getUTCSeconds()))
      .replace(/\bM\b/, t.getUTCMonth() + 1).replace(/\bd\b/, t.getUTCDate());
  };
  const ctx = {
    console, JSON, Date, Math, String, Array, Object, Number, Error,
    Logger: { log: (...a) => { if (opts.verbose) console.log('[Logger]', ...a); } },
    SpreadsheetApp: { getActiveSpreadsheet: () => ss, openById: () => ss, flush() {}, getUi() { throw new Error('getUi 只能在試算表選單裡用'); } },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => props[k] ?? null, setProperty: (k, v) => { props[k] = String(v); }, deleteProperty: k => { delete props[k]; } }) },
    CacheService: { getScriptCache: () => ({ get: k => cache[k] ?? null, put: (k, v) => { cache[k] = v; }, remove: k => { delete cache[k]; } }) },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    Utilities: {
      getUuid: () => crypto.randomUUID(),
      formatDate: fmt,
      computeDigest: (alg, s) => crypto.createHash('sha256').update(s).digest(),
      computeHmacSha256Signature: (v, k) => crypto.createHmac('sha256', k).update(String(v)).digest(),
      base64Encode: b => Buffer.from(b).toString('base64'),
      base64EncodeWebSafe: b => Buffer.from(b).toString('base64url') + '=',
      sleep() {}, DigestAlgorithm: { SHA_256: 1 }, Charset: { UTF_8: 1 }
    },
    ContentService: { createTextOutput: t => ({ t, setMimeType() { return this; } }), MimeType: { JSON: 1 } },
    UrlFetchApp: {
      fetch: (u, o) => {
        fetched.push({ url: u, payload: o && o.payload ? JSON.parse(o.payload) : null });
        const code = opts.fetchCode || 200;
        return { getResponseCode: () => code, getContentText: () => JSON.stringify({ displayName: '測試客人' }), getBlob: () => ({}) };
      }
    },
    ScriptApp: { getService: () => ({ getUrl: () => null }), getOAuthToken: () => 'tok', newTrigger() { return { timeBased() { return this; }, everyDays() { return this; }, atHour() { return this; }, inTimezone() { return this; }, create() {} }; }, getProjectTriggers: () => [] },
    DriveApp: { getFolderById() { throw new Error('DriveApp 沒有模擬'); }, getFoldersByName() { throw new Error('DriveApp 沒有模擬'); } },
    HtmlService: { createHtmlOutput: () => ({ setWidth() { return this; }, setHeight() { return this; } }) }
  };
  vm.createContext(ctx);
  vm.runInContext(src, ctx, { filename: codePath });
  // 測試時不檢查登入
  if (!opts.keepAuth) vm.runInContext('checkToken = function () { return true; };', ctx);
  return {
    ctx, sheets, props, cache, fetched,
    run: js => vm.runInContext(js, ctx),
    post: body => JSON.parse(ctx.doPost({ parameter: {}, postData: { contents: JSON.stringify(Object.assign({ token: 't' }, body)) } }).t),
    get: params => JSON.parse(ctx.doGet({ parameter: params }).t),
    line: (events, key) => ctx.doPost({ parameter: { line: key || props.LINE_KEY }, postData: { contents: JSON.stringify({ destination: 'U', events }) } })
  };
}

function smoke(codePath) {
  const ok = (c, m) => { if (!c) { console.error('✗ ' + m); process.exitCode = 1; } else console.log('✓ ' + m); };
  const sim = load(codePath, { props: { LINE_KEY: 'k', LINE_TOKEN: 'tok' } });
  sim.run('初始化()');
  ok(sim.run("readRows('boats').length") > 0, '初始化建立工作表與預設船隻');
  const r = sim.post({ action: 'batch', ops: [
    { action: 'add', collection: 'tours', data: { name: '海上蹦火導覽', meet: '煮飯間' } },
    { action: 'add', collection: 'bookings', data: { status: '已成單', date: '2099-01-02', name: '甲', adult: 3 } },
    { action: 'add', collection: 'bookings', data: { status: '已成單', date: '2099-01-02', name: '乙', adult: 2 } },
    { action: 'update', collection: 'bookings', id: 'nope', data: {} }
  ] });
  ok(r.ok && r.results.length === 4 && r.results[3].code === 'not_found' && r.data.features && r.data.features.batch, 'batch 寫入與逐筆結果');
  const ids = r.results.slice(1, 3).map(x => x.id);
  const tid = sim.run(`tripOp_({op:'merge',date:'2099-01-02',ids:${JSON.stringify(ids)}})`);
  sim.run(`tripOp_({op:'guides',tripId:'${tid}',guides:['林']})`);
  const bs = sim.run("readRows('bookings')");
  ok(bs.every(b => b.trip === tid && /導覽員已排定/.test(b.marks)), '行程合併＋排導覽員自動勾進度');
  sim.line([{ type: 'message', replyToken: 'r', source: { userId: 'U' + 'a'.repeat(32) }, webhookEventId: 'e1', message: { type: 'text', id: 'm1', text: '價格' } }]);
  ok(sim.fetched.length === 0 && /測試模式/.test(sim.sheets['LINE 訊息'].data.slice(-1)[0][4]), 'LINE 測試模式：名單外不回覆');
  ok(sim.get({ avail: '1' }).ok, '預約頁名額查詢');
  ok(Array.isArray(sim.run('loadAll().templates')), '訊息範本載入');
}

module.exports = { load, smoke };
if (require.main === module) {
  const p = process.argv[2] || 'apps-script/Code.gs';
  smoke(p);
}
