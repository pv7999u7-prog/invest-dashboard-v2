// 0050 長期投資：Apps Script 這一端（抓行情、讀 0050 實際投入、存每日建議、送推播通知）
// 每個台股開盤日 8:25～8:35（每 10 分鐘的 liveJob 順便檢查）自動跑一次；App 也可以要求「重新分析」。
// Claude 的雲端排程每天 8:00 會把新聞解讀寫進「0050 建議」分頁同一天那一列的 E 欄，這裡推播時一起帶上。
const ETF_CFG_SHEET = '0050 設定', ETF_LOG_SHEET = '0050 建議';
const ETF_PUSH_SUBJECT = 'https://pv7999u7-prog.github.io/invest-dashboard-v2/';
const ETF_LEVEL = { core: '核心投入', buy: '加碼', wait: '等待間隔', full: '預算用完', yearend: '年底投入', pause: '暫停加碼', hold: '照常', closed: '休市' };

function etfPrim() {
  const toB = a => a.map(x => ((x & 255) > 127 ? (x & 255) - 256 : x & 255));
  const fromB = a => Array.from(a, x => x & 255);
  return {
    toB,
    hmac: (k, d) => fromB(Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_256, toB(d), toB(k))),
    sha256: d => fromB(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, toB(d))),
    random: n => {
      let out = [];
      while (out.length < n) out = out.concat(fromB(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid() + Utilities.getUuid() + Date.now() + Math.random())));
      return out.slice(0, n);
    }
  };
}
const etfToday = () => Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');

/* ---------- 資料 ---------- */
// Yahoo 日線：[{d:'YYYY-MM-DD'（交易所當地日期）, c:還原權息收盤, raw:收盤}]
function etfChart(sym, range) {
  let last = '';
  for (const host of ['query1', 'query2']) {
    const res = UrlFetchApp.fetch(`https://${host}.finance.yahoo.com/v8/finance/chart/${encodeURIComponent(sym)}?range=${range}&interval=1d&events=div%2Csplit`,
      { muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' } });
    if (res.getResponseCode() !== 200) { last = sym + ' HTTP ' + res.getResponseCode(); continue; }
    const r = JSON.parse(res.getContentText()).chart.result[0], q = r.indicators.quote[0];
    const adj = ((r.indicators.adjclose || [])[0] || {}).adjclose || q.close, off = (r.meta && r.meta.gmtoffset) || 0;
    const out = [];
    (r.timestamp || []).forEach((t, i) => {
      if (q.close[i] == null || adj[i] == null || q.close[i] <= 0) return;
      const d = Utilities.formatDate(new Date((t + off) * 1000), 'UTC', 'yyyy-MM-dd');
      if (out.length && out[out.length - 1].d === d) out.pop();
      out.push({ d, c: adj[i], raw: q.close[i] });
    });
    if (out.length) return out;
    last = sym + ' 沒有資料';
  }
  throw new Error('抓不到行情：' + last);
}
// 證交所休市日（平日但不開盤）；一年抓一次存起來
function etfHolidays(year) {
  const props = PropertiesService.getScriptProperties(), key = 'etf-hol-' + year, cached = props.getProperty(key);
  if (cached) return JSON.parse(cached);
  try {
    const res = UrlFetchApp.fetch('https://openapi.twse.com.tw/v1/holidaySchedule/holidaySchedule', { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return [];
    const out = [];
    JSON.parse(res.getContentText()).forEach(x => {
      const s = String(x.Date || '').replace(/\D/g, ''), name = String(x.Name || '') + String(x.Description || '');
      if (/開始交易|最後交易/.test(name)) return;   // 這兩種日子有開盤
      let d = '';
      if (s.length === 7) d = `${Number(s.slice(0, 3)) + 1911}-${s.slice(3, 5)}-${s.slice(5)}`;
      else if (s.length === 8) d = `${s.slice(0, 4)}-${s.slice(4, 6)}-${s.slice(6)}`;
      if (d) out.push(d);
    });
    if (out.some(d => d.slice(0, 4) === String(year))) props.setProperty(key, JSON.stringify(out));
    return out;
  } catch (err) { return []; }
}
// 0050 的單筆買進：台股分頁交易表（股號 0050、有買入股數）
function etfLumps() {
  const v = sh('台股').getDataRange().getDisplayValues(), h = v.findIndex(r => String(r[0]).trim() === '日期' && /股號/.test(r[3]));
  if (h < 0) return [];
  const out = [];
  for (let i = h + 1; i < v.length; i++) {
    const r = v[i], m = String(r[0]).match(/^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})/);
    if (!m || String(r[3]).trim().toUpperCase() !== '0050' || !(num(r[5]) > 0)) continue;
    const amount = Math.abs(num(r[7]) || 0) || Math.round(num(r[4]) * num(r[5]));
    out.push({ d: `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`, amount });
  }
  return out;
}
function etfCfgSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(ETF_CFG_SHEET);
}
function etfCfg() {
  const s = etfCfgSheet(), out = {};
  if (s) s.getDataRange().getDisplayValues().slice(1).forEach(r => { if (String(r[0]).trim()) out[String(r[0]).trim()] = String(r[1]).trim(); });
  return out;
}
function etfCompute() {
  const today = etfToday(), year = Number(today.slice(0, 4)), cfg = etfCfg(), budgets = {};
  Object.keys(cfg).forEach(k => { const m = k.match(/^預算\s*(\d{4})$/); if (m && num(cfg[k]) > 0) budgets[m[1]] = num(cfg[k]); });
  const adv = etfAdvise({
    today, holidays: etfHolidays(year).concat(etfHolidays(year + 1)),
    p0050: etfChart('0050.TW', '2y'), vix: etfChart('^VIX', '6mo'), sox: etfChart('^SOX', '6mo'), tnx: etfChart('^TNX', '1y'),
    ndc: cfg['景氣燈號分數'] ? { score: num(cfg['景氣燈號分數']), month: cfg['景氣燈號月份'] || '' } : null,
    lumps: etfLumps(), budgets, start: num(cfg['開始年份']) || year
  });
  adv.at = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm');
  return adv;
}

/* ---------- 每日紀錄 ---------- */
function etfLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(ETF_LOG_SHEET);
  if (!s) { s = ss.insertSheet(ETF_LOG_SHEET); s.getRange('A:G').setNumberFormat('@'); s.getRange(1, 1, 1, 7).setValues([['日期', '燈號', '建議', '建議金額', 'Claude 解讀', '解讀時間', '詳細資料（App 用）']]); }
  return s;
}
// 同一天只有一列：找得到就更新 B～D、G（E、F 是 Claude 寫的，不動），找不到就加在最後
function etfSave(adv) {
  const s = etfLogSheet(), last = s.getLastRow();
  const dates = last > 1 ? s.getRange(2, 1, last - 1, 1).getDisplayValues().map(r => r[0]) : [];
  const i = dates.lastIndexOf(adv.date), row = i >= 0 ? i + 2 : last + 1;
  s.getRange(row, 1, 1, 4).setNumberFormat('@').setValues([[adv.date, ETF_LEVEL[adv.level] || adv.level, adv.title, adv.amount ? String(adv.amount) : '']]);
  s.getRange(row, 7).setNumberFormat('@').setValue(JSON.stringify(adv));
  return row;
}
function etfRows(n) {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ETF_LOG_SHEET);
  if (!s || s.getLastRow() < 2) return [];
  const from = Math.max(2, s.getLastRow() - n + 1);
  return s.getRange(from, 1, s.getLastRow() - from + 1, 7).getDisplayValues().filter(r => r[0]).map(r => {
    let adv = null; try { adv = r[6] ? JSON.parse(r[6]) : null; } catch (e) {}
    return { date: r[0], level: r[1], title: r[2], amount: num(r[3]) || 0, news: r[4], newsAt: r[5], adv };
  });
}

/* ---------- 推播 ---------- */
function etfVapid() {
  const props = PropertiesService.getScriptProperties();
  let v = props.getProperty('ETF_VAPID');
  if (!v) { v = JSON.stringify(wpVapidKeys(etfPrim())); props.setProperty('ETF_VAPID', v); }
  return JSON.parse(v);
}
const etfSubs = () => JSON.parse(PropertiesService.getScriptProperties().getProperty('ETF_SUBS') || '[]');
const etfSetSubs = a => PropertiesService.getScriptProperties().setProperty('ETF_SUBS', JSON.stringify(a));
function etfPushAll(msg) {
  const subs = etfSubs();
  if (!subs.length) return '還沒有裝置開啟通知';
  const vapid = etfVapid(), prim = etfPrim(), payload = JSON.stringify(msg);
  const reqs = subs.map(s => { const r = wpRequest(s, payload, vapid, prim, ETF_PUSH_SUBJECT); r.payload = prim.toB(r.payload); return r; });
  const res = UrlFetchApp.fetchAll(reqs), keep = [], codes = [];
  res.forEach((r, i) => { const c = r.getResponseCode(); codes.push(c); if (c !== 404 && c !== 410) keep.push(subs[i]); });
  if (keep.length !== subs.length) etfSetSubs(keep);
  PropertiesService.getScriptProperties().setProperty('etf-push', Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm') + ' ' + codes.join(','));
  return `已送出 ${subs.length} 台裝置（${codes.join('、')}）`;
}
function etfMessage(adv, news) {
  const head = adv.level === 'buy' || adv.level === 'core' || adv.level === 'yearend' ? '🟢' : adv.level === 'pause' ? '🟠' : '⚪';
  const first = String(news || '').split(/(?<=[。！？])/)[0].trim();
  return { title: `${head} 0050 今日建議：${ETF_LEVEL[adv.level] || ''}`, body: adv.title + (first ? '\n' + first : ''), url: './#etfpop', tag: 'etf-' + adv.date };
}

// liveJob 每 10 分鐘叫一次：開盤日 8:25 之後還沒做就做
function etfJob(force) {
  const props = PropertiesService.getScriptProperties(), today = etfToday();
  const hm = Utilities.formatDate(new Date(), 'Asia/Taipei', 'HH:mm');
  if (!force && (hm < '08:25' || hm > '10:30' || props.getProperty('etf-done') === today)) return;
  try {
    const adv = etfCompute();
    props.setProperty('etf-done', today);
    if (!adv.trading) return;
    const row = etfSave(adv), news = etfLogSheet().getRange(row, 5).getDisplayValue();
    etfPushAll(etfMessage(adv, news));
    props.deleteProperty('etf-error');
  } catch (err) {
    props.setProperty('etf-error', etfToday() + ' ' + String(err.message || err));
  }
}

/* ---------- App 的請求 ---------- */
function etfState() {
  const props = PropertiesService.getScriptProperties();
  return { cfg: etfCfg(), rows: etfRows(60), vapid: etfVapid().publicKey, subs: etfSubs().map(s => ({ device: s.device || '', at: s.at || '', id: s.endpoint.slice(-12) })),
    done: props.getProperty('etf-done') || '', error: props.getProperty('etf-error') || '', push: props.getProperty('etf-push') || '' };
}
const ETF_ACTIONS = {
  cfg(p) {
    const key = String(p.key || ''), val = String(p.value == null ? '' : p.value).trim();
    if (/^預算 \d{4}$/.test(key)) { if (!(num(val) >= 100000 && num(val) <= 3000000)) throw new Error('預算請填 10 萬～300 萬'); }
    else if (key === '景氣燈號分數') { if (!(num(val) >= 9 && num(val) <= 45)) throw new Error('景氣燈號分數是 9～45'); }
    else if (key === '開始年份') { if (!(num(val) >= 2020 && num(val) <= 2100)) throw new Error('年份不對'); }
    else if (key !== '景氣燈號月份') throw new Error('不能改這個項目：' + key);
    const s = etfCfgSheet(); if (!s) throw new Error(`找不到「${ETF_CFG_SHEET}」分頁`);
    const v = s.getDataRange().getDisplayValues(), i = v.findIndex(r => String(r[0]).trim() === key);
    const row = i >= 0 ? i + 1 : v.length + 1;
    s.getRange(row, 1, 1, 2).setNumberFormat('@').setValues([[key, val]]);
    return `已更新${key}`;
  },
  sub(p) {
    const x = p.sub || {}, keys = x.keys || {};
    if (!/^https:\/\//.test(x.endpoint || '') || !keys.p256dh || !keys.auth) throw new Error('通知訂閱資料不完整');
    const subs = etfSubs().filter(s => s.endpoint !== x.endpoint);
    subs.push({ endpoint: x.endpoint, p256dh: keys.p256dh, auth: keys.auth, device: String(p.device || '').slice(0, 40), at: Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm') });
    etfSetSubs(subs.slice(-5));
    return '已開啟這台裝置的通知';
  },
  unsub(p) { etfSetSubs(etfSubs().filter(s => s.endpoint !== p.endpoint)); return '已關閉這台裝置的通知'; },
  run() { const adv = etfCompute(); if (adv.trading) etfSave(adv); return adv.trading ? '已重新分析' : '今天休市，沒有新的建議'; },
  test() { return etfPushAll({ title: '0050 通知測試', body: '收到這則就代表每天 8:30 的通知會送到這台裝置。', url: './#etfpop', tag: 'etf-test' }); }
};
function etfRequest(body) {
  try {
    let message = '';
    if (body.etfAction) {
      const fn = ETF_ACTIONS[body.etfAction];
      if (!fn) throw new Error('不認得的動作：' + body.etfAction);
      const lock = LockService.getScriptLock(); lock.waitLock(20000);
      try { message = fn(body.payload || {}); SpreadsheetApp.flush(); } finally { lock.releaseLock(); }
    }
    return json({ ok: true, message, etf: etfState(), ms: msOut() });
  } catch (err) {
    return json({ ok: false, error: String(err.message || err) });
  }
}
