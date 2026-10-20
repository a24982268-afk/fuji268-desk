#!/usr/bin/env node
// 用假後端打開接團台 index.html，印出頁面錯誤並截圖。需要 playwright（見 SKILL.md）。
// 用法：node desk_mock.js --data data.json --view 導覽員安排 --shot out.png [--width 390] [--latency 1500] [--html index.html]
//      const m = await require('./desk_mock.js').open({ data, width: 390 }); await m.page.click(...); await m.close();
'use strict';
const fs = require('fs');
const path = require('path');

function pw() {
  try { return require('playwright'); } catch (e) {
    throw new Error('找不到 playwright：在暫存目錄 npm install playwright@1，執行時設定 NODE_PATH=<目錄>/node_modules');
  }
}
function chromePath() {
  const base = '/opt/pw-browsers';
  if (!fs.existsSync(base)) return undefined;
  const d = fs.readdirSync(base).filter(x => /^chromium-\d+$/.test(x)).sort().pop();
  const p = d && path.join(base, d, 'chrome-linux', 'chrome');
  return p && fs.existsSync(p) ? p : undefined;
}

// 假後端：load、batch、add/update/delete、trip（簡化版），其他動作回 ok
function makeBackend(data) {
  let nid = 0;
  function op(q) {
    const C = data[q.collection];
    if (q.action === 'add') { const id = 'n' + (++nid); C.push(Object.assign({ id }, q.data)); return { ok: true, id }; }
    if (q.action === 'update') { const x = (C || []).find(x => x.id === q.id); if (!x) return { ok: false, code: 'not_found' }; Object.assign(x, q.data); return { ok: true }; }
    if (q.action === 'delete') { data[q.collection] = C.filter(x => x.id !== q.id); return { ok: true }; }
    if (q.action === 'trip') {
      const B = data.bookings; data.trips = data.trips || [];
      if (q.op === 'merge') { const id = q.tripId || ('n' + (++nid)); if (!q.tripId) data.trips.push({ id, date: q.date, name: '', boat: '', guides: [] }); B.forEach(b => { if (q.ids.includes(b.id)) b.trip = id; }); data.trips = data.trips.filter(t => B.some(b => b.trip === t.id)); return { ok: true, id }; }
      if (q.op === 'leave') B.forEach(b => { if (q.ids.includes(b.id)) b.trip = ''; });
      if (q.op === 'dissolve') { B.forEach(b => { if (b.trip === q.tripId) b.trip = ''; }); data.trips = data.trips.filter(t => t.id !== q.tripId); }
      if (q.op === 'guides') { const t = data.trips.find(t => t.id === q.tripId); if (t) t.guides = q.guides; }
      return { ok: true, id: q.tripId };
    }
    return { ok: true };
  }
  return op;
}

async function open(o) {
  const data = o.data;
  data.bookings = data.bookings || []; data.guides = data.guides || []; data.tours = data.tours || [];
  data.boats = data.boats || []; data.settings = data.settings || {};
  const browser = await pw().chromium.launch({ executablePath: chromePath() });
  const page = await browser.newPage({ viewport: { width: o.width || 390, height: o.height || 1400 }, acceptDownloads: true });
  const errors = [], requests = [];
  page.on('pageerror', e => errors.push(e.message));
  const op = makeBackend(data);
  await page.route(/script\.google\.com/, async r => {
    const q = JSON.parse(r.request().postData() || '{}');
    requests.push(q.action + (q.ops ? '(' + q.ops.length + ')' : ''));
    let out = { ok: true, data };
    if (q.action === 'batch') out.results = q.ops.map(op);
    else if (q.action && q.action !== 'load') Object.assign(out, op(q));
    if (o.latency && q.action !== 'load') await new Promise(z => setTimeout(z, o.latency));
    await r.fulfill({ contentType: 'application/json', body: JSON.stringify(out) });
  });
  await page.addInitScript(() => localStorage.setItem('fuji268-desk-token', 'test'));
  const html = path.resolve(o.html || 'index.html');
  await page.goto('file://' + html);
  await page.waitForTimeout(700);
  if (o.view) { await page.click('#nav >> text=' + o.view); await page.waitForTimeout(300); }
  return { page, data, errors, requests, close: () => browser.close() };
}

module.exports = { open, makeBackend };

if (require.main === module) {
  const a = process.argv.slice(2), arg = k => { const i = a.indexOf('--' + k); return i >= 0 ? a[i + 1] : undefined; };
  const data = arg('data') ? JSON.parse(fs.readFileSync(arg('data'), 'utf8')) : { features: { batch: 1 }, sailings: [], trips: [], bookings: [] };
  (async () => {
    const m = await open({ data, view: arg('view'), width: +arg('width') || 390, latency: +arg('latency') || 0, html: arg('html') });
    if (arg('shot')) await m.page.screenshot({ path: arg('shot'), fullPage: true });
    console.log('sync:', await m.page.textContent('#sync'));
    console.log('requests:', m.requests.join(', '));
    console.log(m.errors.length ? 'pageerror:\n' + m.errors.join('\n') : 'pageerror: 無');
    await m.close();
    if (m.errors.length) process.exitCode = 1;
  })();
}
