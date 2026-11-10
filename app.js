// 投資總覽：打開時從 Google 試算表（透過 Apps Script）抓最新數據，畫成總覽與各項目頁面。
// 所有財務數字都只存在使用者自己的裝置（localStorage），程式碼裡沒有任何資料。
'use strict';

const LS = { cfg: 'inv2-cfg-v1', cache: 'inv2-cache-v1', hide: 'inv2-hide-v1' };   // 舊版 App 在同一個網域，key 要不同
const AUTO_REFRESH_MIN = 5;   // App 從背景回到前景時，超過幾分鐘沒更新就重抓
const APP_VERSION = '3.3.2';   // 顯示在設定頁，用來確認 iPhone 上是不是新版
const LATEST_SCRIPT = 29;   // 最新 Code.gs 的 SCRIPT_VERSION
// 目前部署的 Code.gs 是第幾版：舊版沒有版本號，用回傳的內容推斷
const scriptVersion = () => { const r = state.raw; if (!r) return null; if (r._scriptVersion) return r._scriptVersion; return r.bb ? 3 : r._sheetUrl ? 2 : 1; };

// 配置類別的顏色跟著類別走（依試算表 Total Overall 的順序），不跟排名走
const CAT_COLORS = { '虛擬貨幣': '--s2', '放貸': '--s7', '台股': '--s3', "Kids' 台股": '--s5', "Kids' 美股": '--s4', '美股': '--s1' };
const CAT_NAMES = { '台股': '我的台股', "Kids' 台股": '小孩台股', "Kids' 美股": '小孩美股' };

/* ================================================================
   小工具
   ================================================================ */
const $ = s => document.querySelector(s);
const css = n => getComputedStyle(document.documentElement).getPropertyValue(n).trim();
const str = v => (v == null ? '' : String(v).trim());
function num(v) {
  if (typeof v === 'number') return isFinite(v) ? v : null;
  const s = str(v).replace(/[,\s]|NT\$|US\$|\$/g, '');
  if (!s) return null;
  const n = parseFloat(s);
  if (!isFinite(n)) return null;
  return s.endsWith('%') ? n / 100 : n;
}
const pad = n => String(n).padStart(2, '0');
// 試算表的日期可能以「2026/10/30」文字、或 ISO 時間字串（2026-10-29T16:00:00.000Z）傳過來，統一成 yyyy/mm/dd
function toDateStr(v) {
  const s = str(v);
  let m = s.match(/^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})$/);
  if (m) return `${m[1]}/${pad(m[2])}/${pad(m[3])}`;
  if (/^\d{4}-\d{2}-\d{2}T/.test(s)) {
    const d = new Date(s);
    if (!isNaN(d)) return `${d.getFullYear()}/${pad(d.getMonth() + 1)}/${pad(d.getDate())}`;
  }
  return '';
}
const isDate = v => !!toDateStr(v);
const normDate = toDateStr;
const dateText = v => toDateStr(v) || str(v);
const toTime = d => { const [y, m, dd] = d.split('/').map(Number); return new Date(y, m - 1, dd).getTime(); };
const esc = s => str(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

let hidden = false;
try { hidden = localStorage.getItem(LS.hide) === '1'; } catch (e) {}
const MASK = '<span class="masked">••••</span>';
const int = n => Math.round(n).toLocaleString('en-US');
const dec = (n, d) => n.toLocaleString('en-US', { minimumFractionDigits: d, maximumFractionDigits: d });
const ntd = n => (n == null ? '—' : hidden ? MASK : 'NT$' + int(n));
const usd = (n, d = 0) => (n == null ? '—' : hidden ? MASK : 'US$' + dec(n, d));
const wan = n => (n == null ? '—' : hidden ? MASK : dec(n / 10000, 1) + ' 萬');
const signedNtd = n => (n == null ? '—' : hidden ? MASK : (n >= 0 ? '+' : '−') + 'NT$' + int(Math.abs(n)));
const signedUsd = n => (n == null ? '—' : hidden ? MASK : (n >= 0 ? '+' : '−') + 'US$' + int(Math.abs(n)));
const signedWan = n => (n == null ? '—' : hidden ? MASK : (n >= 0 ? '+' : '−') + dec(Math.abs(n) / 10000, 1) + ' 萬');
const pct = (f, d = 1) => (f == null ? '—' : (f > 0 ? '+' : f < 0 ? '−' : '') + dec(Math.abs(f * 100), d) + '%');
const pctPlain = (f, d = 1) => (f == null ? '—' : dec(f * 100, d) + '%');
const cls = f => (f == null ? '' : f >= 0 ? 'up' : 'down');
const qty = n => (n == null ? '—' : hidden ? MASK : (Number.isInteger(n) ? int(n) : dec(n, n < 1 ? 4 : 2)));

/* ================================================================
   解析試算表（Apps Script 回傳每張工作表的原始格子）
   ================================================================ */
function findRow(rows, col, text, from = 0) {
  for (let r = from; r < rows.length; r++) if (str(rows[r][col]) === text) return r;
  return -1;
}
const cell = (rows, r, c) => (rows[r] ? rows[r][c] : '');

function parseTotal(rows) {
  const out = {
    fx: num(cell(rows, 1, 1)), cost: num(cell(rows, 1, 4)), valueUSD: num(cell(rows, 1, 5)),
    value: num(cell(rows, 1, 6)), earn: num(cell(rows, 1, 7)), irr: num(cell(rows, 1, 8)),
    items: [], history: [], yearEnd: []
  };
  let cat = '';
  for (let r = 3; r < rows.length; r++) {
    const row = rows[r];
    if (!row || row.slice(0, 9).every(v => str(v) === '')) break;
    cat = str(row[0]) || cat;
    out.items.push({ cat, sub: str(row[1]), where: str(row[2]), costUSD: num(row[3]), costNTD: num(row[4]), valueUSD: num(row[5]), value: num(row[6]) });
  }
  const alloc = {};
  out.items.forEach(it => { alloc[it.cat] = (alloc[it.cat] || 0) + (it.value || 0); });
  out.alloc = Object.keys(alloc).map(k => ({ cat: k, name: CAT_NAMES[k] || k, value: alloc[k], color: CAT_COLORS[k] || '--s8',
    parts: out.items.filter(i => i.cat === k) }));

  const h = findRow(rows, 0, 'Date');
  if (h >= 0) {
    for (let r = h + 1; r < rows.length; r++) {
      const row = rows[r];
      if (!isDate(row[0])) continue;
      out.history.push({ date: normDate(row[0]), value: num(row[1]), cost: num(row[2]), flow: num(row[3]), earn: num(row[4]), fx: num(row[6]), note: str(row[9]) });
    }
  }
  // 年底總結（K 欄起：日期、匯率、Total）
  for (let r = 0; r < rows.length; r++) {
    if (str(cell(rows, r, 10)) === '日期' && str(cell(rows, r, 12)).startsWith('Total')) {
      for (let k = r + 1; k < rows.length; k++) {
        const d = cell(rows, k, 10), v = num(cell(rows, k, 12));
        if (isDate(d) && v) out.yearEnd.push({ date: normDate(d), value: v });
      }
      break;
    }
  }
  return out;
}

function parseBrokerSummary(rows) {
  const hdr = findRow(rows, 0, '券商');
  const list = [];
  if (hdr < 0) return { brokers: list, total: null };
  let total = null;
  for (let r = hdr + 1; r < Math.min(rows.length, hdr + 8); r++) {
    const name = str(rows[r][0]);
    if (!name) continue;
    const o = { name, cost: num(rows[r][1]), cash: num(rows[r][2]), value: num(rows[r][4]), earn: num(rows[r][5]), pct: num(rows[r][6]) };
    if (name === 'Total') { total = o; break; }
    list.push(o);
  }
  return { brokers: list, total };
}
function irrBelow(rows, label) {
  for (let r = 0; r < Math.min(rows.length, 30); r++) {
    if (str(rows[r][0]).replace(/\s+/g, ' ') === label) return num(cell(rows, r + 1, 0));
  }
  return null;
}

function parseTw(rows) {
  const out = parseBrokerSummary(rows);
  out.irr = irrBelow(rows, '台股 Total IRR');
  out.holdings = [];
  const h = findRow(rows, 8, '股號');
  if (h >= 0) {
    for (let r = h + 1; r < rows.length && str(rows[r][8]); r++) {
      const x = rows[r];
      out.holdings.push({ code: str(x[8]), name: str(x[9]), qty: num(x[10]), pct: num(x[11]), pl: num(x[12]), price: num(x[13]), avg: num(x[14]), value: num(x[15]) });
    }
  }
  out.tx = [];
  const t = findRow(rows, 0, '日期');
  if (t >= 0) {
    // 交易表到第一個空白列為止；下面那列「今天」的總結（A 是 TODAY()、C 是台股總值）不是交易
    for (let r = t + 1; r < rows.length && str(rows[r][0]) !== ''; r++) {
      const x = rows[r];
      if (!isDate(x[0])) continue;
      out.tx.push({ row: r + 1, date: normDate(x[0]), name: str(x[1]), flow: num(x[2]), code: str(x[3]), price: num(x[4]), buy: num(x[5]), sell: num(x[6]), note: str(x[9]) });
    }
  }
  return out;
}

function parseUs(rows) {
  const out = parseBrokerSummary(rows);
  out.irr = irrBelow(rows, '美股 Total IRR');
  out.brokerIrr = { CUBE: irrBelow(rows, 'CUBE IRR'), Firstrade: irrBelow(rows, 'Firstrade IRR') };
  out.holdings = [];
  const txHdr = findRow(rows, 0, '日期');
  const end = txHdr >= 0 ? txHdr : rows.length;
  for (let r = 0; r < end; r++) {
    if (str(rows[r][8]) !== '代號') continue;
    const broker = str(cell(rows, r - 1, 8)) || '其他';
    for (let k = r + 1; k < end && str(rows[k][8]); k++) {
      if (str(cell(rows, k + 1, 8)) === '代號') break;
      const x = rows[k];
      out.holdings.push({ broker, code: str(x[8]), qty: num(x[9]), pct: num(x[10]), pl: num(x[11]), price: num(x[12]), avg: num(x[13]), value: num(x[14]), cost: num(x[15]) });
    }
  }
  out.tx = [];
  if (txHdr >= 0) {
    for (let r = txHdr + 1; r < rows.length; r++) {
      const x = rows[r];
      if (!isDate(x[0]) || !str(x[3])) continue;
      out.tx.push({ row: r + 1, date: normDate(x[0]), where: str(x[1]), flow: num(x[2]), name: str(x[3]), price: num(x[4]), buy: num(x[5]), sell: num(x[6]), note: str(x[9]) });
    }
  }
  return out;
}

function parseCrypto(rows) {
  const out = { totalUSD: num(cell(rows, 1, 4)), total: num(cell(rows, 1, 5)), coins: [], realized: [], memo: [] };
  // 時間備忘：幣分頁 T、U 欄（例如「預期下次減半時間｜2028年4月」），只有一欄有字也列
  for (let r = 0; r < Math.min(rows.length, 20); r++) {
    const k = str(cell(rows, r, 19)), v = str(cell(rows, r, 20));
    if (k || v) out.memo.push({ k, v });
  }
  for (let r = 3; r < rows.length; r++) {
    const x = rows[r], q = num(x[2]);
    if (!str(x[0]) || !q) continue;
    out.coins.push({ name: str(x[0]), where: str(x[1]), qty: q, price: num(x[3]), usd: num(x[4]), value: num(x[5]) });
  }
  for (let r = 1; r < rows.length; r++) {
    const d = str(cell(rows, r, 11));
    if (!d) continue;
    out.realized.push({ date: fmtCompactDate(d), coin: str(cell(rows, r, 12)), qty: num(cell(rows, r, 13)) });
  }
  return out;
}
function fmtCompactDate(s) {
  if (toDateStr(s)) return toDateStr(s);
  s = str(s);
  if (/^\d{8}$/.test(s)) return `${s.slice(0, 4)}/${s.slice(4, 6)}/${s.slice(6)}`;
  if (/^\d{6}$/.test(s)) return `${s.slice(0, 4)}/${s.slice(4, 6)}`;
  return s.replace(/(\d{4})(\d{2})~(\d{4})(\d{2})/, '$1/$2 – $3/$4');
}

// 可以修改、刪除的明細；row 是試算表的列號（Apps Script 從 A1 開始讀，所以陣列位置 +1）
function ledgerRows(rows, col, header, map) {
  const out = [];
  const h = rows.findIndex(x => str(x && x[col]) === header);
  if (h < 0) return out;
  for (let r = h + 1; r < rows.length && str(cell(rows, r, col)) !== ''; r++) { const x = map(rows[r], r + 1); if (x) out.push(x); }
  return out;
}
function parseLedger(raw) {
  const trade = (x, row) => ({ row, date: normDate(x[0]), b: str(x[1]), flow: num(x[2]), code: str(x[3]), price: num(x[4]), buy: num(x[5]), sell: num(x[6]),
    amt: num(x[7]) || num(x[8]), note: str(x[9]), key: num(x[2]) || num(x[5]) || num(x[6]) || 0 });
  return {
    flow: ledgerRows(raw.total || [], 0, 'Date', (x, row) => (num(x[3]) ? { row, date: normDate(x[0]), flow: num(x[3]), fx: num(x[6]), note: str(x[9]), key: num(x[3]) } : null)),
    tw: ledgerRows(raw.tw || [], 0, '日期', trade),
    us: ledgerRows(raw.us || [], 0, '日期', trade),
    crypto: ledgerRows(raw.crypto || [], 24, '日期', (x, row) => ({ row, date: normDate(x[24]), type: str(x[25]), coin: str(x[26]), where: str(x[27]),
      price: num(x[28]), qty: num(x[29]), usdt: num(x[30]), note: str(x[31]), key: num(x[30]) || 0 }))
  };
}

/* ---------- IRR（App 自己算，不用試算表的公式） ---------- */
// XIRR：flows [{t: 毫秒, v: 金額}]，投入是負數、拿回（或現值）是正數；用二分法找年化報酬率
function xirr(flows) {
  const f = flows.filter(x => x.v && isFinite(x.t)).sort((a, b) => a.t - b.t);
  if (f.length < 2 || !f.some(x => x.v > 0) || !f.some(x => x.v < 0)) return null;
  const t0 = f[0].t, npv = r => f.reduce((s2, x) => s2 + x.v / Math.pow(1 + r, (x.t - t0) / 864e5 / 365), 0);
  let lo = -0.9999, hi = 100, flo = npv(lo), fhi = npv(hi);
  if (!(flo * fhi < 0)) return null;
  for (let i = 0; i < 300; i++) {
    const mid = (lo + hi) / 2, fm = npv(mid);
    if (Math.abs(fm) < 1e-7 || hi - lo < 1e-10) return mid;
    if (flo * fm < 0) { hi = mid; } else { lo = mid; flo = fm; }
  }
  return (lo + hi) / 2;
}
// 每一年：前一年 12/31 的總資產當成初始投入，加上當年的入金、出金，年底（今年用目前）總資產結算
// 累計：從第一筆入金到現在；台股、美股各帳戶：該帳戶的出入金＋目前現值
function computeIrr(d) {
  const t = d.total;
  if (!t) return;
  const flowsOf = list => list.filter(x => x.flow).map(x => ({ t: toTime(x.date), v: x.flow }));
  const now = Date.now(), dayMs = 864e5;
  const toFlows = (d.ledger && d.ledger.flow) || t.history.filter(h => h.flow);
  const flows = flowsOf(toFlows);
  // 有總資產數字的時間點：歷史表的 B 欄、年底總結表
  const pts = t.history.filter(h => h.value && h.date).map(h => ({ t: toTime(h.date), v: h.value }))
    .concat((t.yearEnd || []).map(y => ({ t: toTime(y.date), v: y.value }))).sort((a, b) => a.t - b.t);
  const near = (time, maxDays) => {   // 最接近這個時間（往前找）的總資產
    const p = pts.filter(x => x.t <= time + dayMs / 2).pop();
    return p && time - p.t <= maxDays * dayMs ? p : null;
  };
  const irr = { all: xirr(flows.concat([{ t: now, v: t.value }])), years: [] };
  const y0 = flows.length ? new Date(flows[0].t).getFullYear() : new Date().getFullYear(), yNow = new Date().getFullYear();
  for (let y = y0; y <= yNow; y++) {
    const s = near(toTime(`${y - 1}/12/31`), 45);
    const e = y === yNow ? { t: now, v: t.value } : near(toTime(`${y}/12/31`), 45);
    if (!s || !e || e.t <= s.t) continue;
    const fl = flows.filter(x => x.t > s.t && x.t <= e.t);
    const r = xirr([{ t: s.t, v: -s.v }].concat(fl, [{ t: e.t, v: e.v }]));
    if (r == null) continue;
    const yrs = (e.t - s.t) / dayMs / 365;
    irr.years.push({ year: y, start: s, end: e, net: -fl.reduce((a, x) => a + x.v, 0), irr: r, ret: Math.pow(1 + r, yrs) - 1,
      exactStart: new Date(s.t).getMonth() === 11 && new Date(s.t).getDate() === 31, partial: y === yNow });
  }
  t.irrCalc = irr;
  if (irr.all != null) t.irr = irr.all;
  // 帳戶：台股（我的）、美股（全部、各券商）
  const L = d.ledger || {};
  if (d.tw && d.tw.total && L.tw) { const r = xirr(flowsOf(L.tw).concat([{ t: now, v: d.tw.total.value }])); if (r != null) d.tw.irr = r; }
  if (d.us && L.us) {
    const all = xirr(flowsOf(L.us).concat([{ t: now, v: d.us.total ? d.us.total.value : 0 }]));
    if (all != null) d.us.irr = all;
    (d.us.brokers || []).forEach(b => {
      const r = xirr(flowsOf(L.us.filter(x => x.b === b.name)).concat([{ t: now, v: b.value }]));
      if (r != null) d.us.brokerIrr[b.name] = r;
    });
  }
}

function parsePlans(rows) {
  // 「股名」表頭底下的定期定額計畫
  const plans = [];
  for (let r = 0; r < rows.length; r++) {
    if (str(rows[r][0]) !== '股名') continue;
    const currency = /USD/i.test(str(rows[r][1])) ? 'USD' : 'NTD';   // 表頭「USD金額/月」或「金額/月」
    for (let k = r + 1; k < rows.length && str(rows[k][0]) && str(rows[k][0]) !== 'Total'; k++) {
      const x = rows[k];
      plans.push({ currency, name: str(x[0]), monthly: num(x[1]), start: dateText(x[2]), to: dateText(x[3]), months: num(x[4]), total: num(x[5]), shares: num(x[6]), price: num(x[7]), value: num(x[8]), pct: num(x[9]) });
    }
  }
  return plans;
}

function parseMyDca(rows) {
  const out = { rows: [], plans: parsePlans(rows) };
  const h = findRow(rows, 0, '券商');
  if (h >= 0) {
    for (let r = h + 1; r < rows.length && str(rows[r][0]); r++) {
      const x = rows[r];
      out.rows.push({ broker: str(x[0]), target: str(x[1]), invUSD: num(x[2]), invNTD: num(x[3]), valUSD: num(x[4]), valNTD: num(x[5]), earnUSD: num(x[6]), earnNTD: num(x[7]), pct: num(x[8]) });
      if (str(x[0]) === 'Total') break;
    }
  }
  return out;
}

function parseKidsUs(rows) {
  const out = { kids: [], total: null, holdings: {} };
  const h = findRow(rows, 0, 'Name');
  if (h >= 0) {
    for (let r = h + 1; r < rows.length && str(rows[r][0]); r++) {
      const x = rows[r];
      const o = { name: str(x[0]), costNTD: num(x[1]), costUSD: num(x[2]), valUSD: num(x[4]), earnUSD: num(x[5]), pctUSD: num(x[6]), valNTD: num(x[7]), earnNTD: num(x[8]), pctNTD: num(x[9]) };
      if (o.name === 'Total') { out.total = o; break; }
      out.kids.push(o);
    }
  }
  for (let r = 1; r < rows.length; r++) {
    if (str(rows[r][0]) !== '代號') continue;
    const kid = str(rows[r - 1][0]);
    out.holdings[kid] = [];
    for (let k = r + 1; k < rows.length && str(rows[k][0]); k++) {
      const x = rows[k];
      out.holdings[kid].push({ code: str(x[0]), qty: num(x[1]), pct: num(x[2]), price: num(x[4]), avg: num(x[5]), value: num(x[6]), cost: num(x[7]) });
    }
  }
  return out;
}

function parseKidsTw(rows) {
  const out = { kids: [], total: null, plans: parsePlans(rows), lumps: [] };
  const h = findRow(rows, 0, 'Name');
  if (h >= 0) {
    for (let r = h + 1; r < rows.length && str(rows[r][0]); r++) {
      const x = rows[r];
      const o = { name: str(x[0]), broker: str(x[1]), inv: num(x[2]), value: num(x[3]), earn: num(x[4]), pct: num(x[5]) };
      if (o.name === 'Total') { out.total = o; break; }
      out.kids.push(o);
    }
  }
  const d = findRow(rows, 0, 'Date');
  if (d >= 0) {
    for (let r = d + 1; r < rows.length; r++) {
      const x = rows[r];
      if (!isDate(x[0])) continue;
      out.lumps.push({ date: normDate(x[0]), name: str(x[1]), price: num(x[2]), inv: num(x[5]), shares: num(x[6]), cur: num(x[7]), value: num(x[8]) });
    }
  }
  return out;
}

function parsePlan(rows) {
  let irr = num(cell(rows, 1, 2));
  if (irr != null && irr > 1) irr /= 100;
  const out = { irr, input: num(cell(rows, 2, 5)), years: [] };
  const h = findRow(rows, 0, 'No.');
  if (h >= 0) {
    for (let r = h + 1; r < rows.length; r++) {
      const x = rows[r], y = num(x[1]);
      if (!y) continue;
      out.years.push({ year: y, age: num(x[2]), target: num(x[4]), input: num(x[5]), actual: num(x[6]), actualInput: num(x[7]) });
    }
  }
  return out;
}

function parseMilestone(rows) {
  return rows.filter(r => str(r[1])).map(r => ({ date: fmtCompactDate(r[0]), text: str(r[1]) }));
}

// 版面：每列月份標題（1、2、3…），底下每個月兩欄「項目｜金額」；再往下是「年度基本總支出」的「項目｜Total」表
function parseExpense(rows) {
  const isMonthRow = r => { const a = num(cell(rows, r, 0)), b = num(cell(rows, r, 2));
    return Number.isInteger(a) && a >= 1 && a <= 12 && b === a + 1 && !str(cell(rows, r, 1)); };
  const blocks = [];
  for (let r = 0; r < rows.length; r++) if (isMonthRow(r)) blocks.push(r);
  const summaryRow = rows.findIndex(row => /^年度/.test(str(row[0])) || str(row[0]) === '項目');
  const months = [];
  blocks.forEach((b, i) => {
    let end = i + 1 < blocks.length ? blocks[i + 1] : rows.length;
    if (summaryRow > b && summaryRow < end) end = summaryRow;
    for (let m = 0; m < 6; m++) {
      const col = m * 2, mo = num(cell(rows, b, col));
      if (!mo) continue;
      const items = [];
      for (let r = b + 1; r < end; r++) {
        const label = str(cell(rows, r, col)), amt = num(cell(rows, r, col + 1));
        if (label && amt) items.push({ label, amt });
      }
      months.push({ month: mo, total: items.reduce((s, it) => s + it.amt, 0), items });
    }
  });
  months.sort((a, b) => a.month - b.month);
  const out = { months, cats: [], title: '' };
  if (summaryRow >= 0) {
    if (/^年度/.test(str(rows[summaryRow][0]))) out.title = str(rows[summaryRow][0]);
    const hdr = rows.findIndex((row, r) => r >= summaryRow && str(row[0]) === '項目');
    for (let r = (hdr >= 0 ? hdr : summaryRow) + 1; r < rows.length && str(rows[r][0]); r++) {
      const a = num(rows[r][1]);
      if (a) out.cats.push({ label: str(rows[r][0]), amt: a, note: str(rows[r][2]) });
    }
  }
  out.monthsTotal = months.reduce((s, m) => s + m.total, 0);
  out.annual = out.cats.length ? out.cats.reduce((s, c) => s + c.amt, 0) : out.monthsTotal;
  return out;
}

// House 分頁上半部是房產，A 欄寫「房貸」那一列以下是房貸：
// A 房產名稱 B 基準日 C 基準日剩餘本金 D 年利率 E 每月還款 F 每月扣款日 G 到期日 H 初貸日 I 初貸金額 J 還本方式
function parseHouse(rows) {
  const list = [];
  let r = 2;
  for (; r < rows.length; r++) {
    const x = rows[r];
    if (str(x[0]) === '房貸') break;
    if (!str(x[0]) || !num(x[2])) continue;
    const housePing = num(x[4]), estPark = num(x[10]), estUnit = num(x[11]);
    list.push({ name: str(x[0]), date: dateText(x[1]).replace(/^(\d{4}\/\d{2})\/01$/, '$1'), price: num(x[2]), ping: num(x[3]), housePing, parkN: num(x[6]), unit: num(x[8]),
      estUnit, estPark, est: housePing && estUnit ? housePing * estUnit + (estPark || 0) : null });
  }
  for (r++; r < rows.length; r++) {
    const x = rows[r], h = list.find(h => h.name === str(x[0]));
    const base = toDateStr(x[1]), bal = num(x[2]), rate = num(x[3]), pmt = num(x[4]);
    if (!h || !base || !bal || rate == null || !pmt) continue;
    h.loan = loanSchedule({ base, bal, rate, pmt, day: num(x[5]) || Number(base.slice(8)), end: toDateStr(x[6]),
      start: toDateStr(x[7]), orig: num(x[8]), method: str(x[9]) || '本息平均攤還' });
  }
  return list;
}

// 本息平均攤還：從基準日的剩餘本金開始，每月扣款日先付當月利息，其餘還本金
function loanSchedule(l) {
  const [by, bm] = l.base.split('/').map(Number);
  const endT = l.end ? toTime(l.end) : Infinity, now = Date.now();
  const fmt = (y, m, d) => `${y}/${String(m).padStart(2, '0')}/${String(d).padStart(2, '0')}`;
  const sched = [];
  let bal = l.bal;
  for (let i = 1; bal > 0.5 && i < 600; i++) {
    const y = by + Math.floor((bm - 1 + i) / 12), m = (bm - 1 + i) % 12 + 1;
    const d = Math.min(l.day, new Date(y, m, 0).getDate());
    let date = fmt(y, m, d);
    if (toTime(date) > endT) date = l.end;
    const interest = Math.round(bal * l.rate / 12);
    // 到期那一期（或剩不到一期）把剩下的本金一次還清
    const last = toTime(date) >= endT || l.pmt - interest >= bal;
    const principal = last ? bal : l.pmt - interest;
    bal -= principal;
    sched.push({ date, interest, principal, pay: principal + interest, bal });
    if (last) break;
  }
  const paid = sched.filter(p => toTime(p.date) <= now);
  const left = sched.slice(paid.length);
  return { ...l, sched, balance: paid.length ? paid[paid.length - 1].bal : l.bal, next: left[0] || null, paidN: paid.length,
    leftN: left.length, leftInterest: left.reduce((s, p) => s + p.interest, 0), payoff: sched.length ? sched[sched.length - 1].date : null };
}

// BB 分頁：每個年度一列（A 欄是年份），底下沒有年份的列是同年度的結構調薪或晚發的年中分紅
// 欄位：B 職稱（↑ 表示升遷）C 年資 D 本次年資 E 年齡｜F 月薪 G 年前分紅（Year-End B，農曆年前）H 職務加給 I 年終 J 月數 K 總月數
//      M 年中分紅時月薪 N 年中分紅（B）O 月數｜Q 調薪 R 調薪%｜T 年收入 U 加班 W 年收入（含加班）AC 相當月數
function parseCareer(rows) {
  const years = [];
  let cur = null, salary = null;
  rows.forEach((x, i) => {
    const y = num(x[0]);
    const take = (k, v) => { if (v != null) cur[k] = v; };
    if (Number.isInteger(y) && y > 1990 && y < 2100) {
      const title = str(x[1]);
      cur = { year: y, title: title.replace(/^↑\s*/, ''), promoted: title.startsWith('↑'), tenure: num(x[2]), qTenure: num(x[3]), age: num(x[4]),
        salary: num(x[5]), yeTotal: num(x[6]), allowance: num(x[7]), yeBonus: num(x[8]), yeMonths: num(x[9]), yeTotalMonths: num(x[10]),
        bSalary: null, bonus: null, bonusMonths: null, raise: null, raisePct: null,
        annual: num(x[19]), ot: num(x[20]), annualTotal: num(x[22]), months: num(x[28]),
        fiscal: num(x[25]), fiscalTotal: num(x[27]), fiscalMonths: num(x[29]), events: [] };
      years.push(cur);
    } else if (!cur) return;
    else {
      const label = str(x[1]), s = num(x[5]) || num(x[12]);
      if (label) cur.events.push({ label, salary: s, row: i + 1 });   // row：BB 分頁的列號，用來修改、刪除
    }
    // 年度調薪（Q、R 欄）照它出現的那一列排進時間軸，調薪後月薪取同一列的 M 欄
    if (num(x[16])) cur.events.push({ raise: num(x[16]), raisePct: num(x[17]) ?? cur.raisePct, salary: num(x[12]) });
    take('bSalary', num(x[12])); take('bonus', num(x[13])); take('bonusMonths', num(x[14]));
    take('raise', num(x[16])); take('raisePct', num(x[17]));
    [num(x[5]), num(x[12])].forEach(v => { if (v) { salary = v; cur.lastSalary = v; } });
  });
  // 年收入那格（T、W 欄）還沒有公式，但已經填了年中分紅：先用試算表 2024 年的算法預估，畫面標示「預估」
  //   年前分紅 ＋ 年中分紅 ＋ 年初月薪 × 6 ＋ 最新月薪 × 6（＋ 加班費）
  years.forEach(y => {
    if (y.annual || y.annualTotal || !y.bonus) return;
    const first = y.salary || y.bSalary || y.lastSalary, last = y.lastSalary || first;
    if (!first) return;
    y.annual = (y.yeTotal || 0) + y.bonus + first * 6 + last * 6;
    y.annualTotal = y.annual + (y.ot || 0);
    y.estFirst = first; y.estLast = last; y.est = true;
  });
  return { years, salary };
}

function parseAll(raw) {
  const safe = (fn, rows) => { try { return fn(rows || []); } catch (e) { console.error(e); return null; } };
  const d = {
    total: safe(parseTotal, raw.total), tw: safe(parseTw, raw.tw), us: safe(parseUs, raw.us),
    crypto: safe(parseCrypto, raw.crypto), myDca: safe(parseMyDca, raw.myDca), kidsUs: safe(parseKidsUs, raw.kidsUs),
    kidsTw: safe(parseKidsTw, raw.kidsTw), plan: safe(parsePlan, raw.plan), milestone: safe(parseMilestone, raw.milestone),
    expense: safe(parseExpense, raw.expense), house: safe(parseHouse, raw.house),
    career: raw.bb ? safe(parseCareer, raw.bb) : null,
    ledger: safe(parseLedger, raw)
  };
  derive(d);
  safe(() => computeIrr(d));
  d.usAll = safe(() => combineUs(d));
  d.twAll = safe(() => combineTw(d));
  return d;
}

// 衍生指標：高點、回檔、今年實際獲利、年度目標
function derive(d) {
  const t = d.total;
  if (!t) return;
  const today = new Date();
  const todayStr = `${today.getFullYear()}/${pad(today.getMonth() + 1)}/${pad(today.getDate())}`;
  const pts = t.history.filter(p => p.value);
  if (t.value && (!pts.length || pts[pts.length - 1].date !== todayStr) && (!pts.length || toTime(pts[pts.length - 1].date) < toTime(todayStr))) {
    pts.push({ date: todayStr, value: t.value, cost: t.cost, note: '' });
  }
  t.series = pts;
  const peak = pts.reduce((a, b) => (b.value > (a ? a.value : 0) ? b : a), null);
  t.peak = peak;
  t.drawdown = peak ? t.value - peak.value : null;
  const year = today.getFullYear();
  const prevEnd = t.yearEnd.filter(y => +y.date.slice(0, 4) === year - 1).pop();
  const costAt = date => {
    let c = null;
    t.history.forEach(h => { if (h.cost != null && toTime(h.date) <= toTime(date)) c = h.cost; });
    return c;
  };
  let base = null;
  if (prevEnd) base = { value: prevEnd.value, cost: costAt(prevEnd.date) };
  else {
    const last = pts.filter(p => +p.date.slice(0, 4) < year).pop();
    if (last) base = { value: last.value, cost: last.cost };
  }
  if (base && base.cost != null) {
    t.ytd = (t.value - t.cost) - (base.value - base.cost);
    t.ytdInput = t.cost - base.cost;
    t.ytdPct = t.ytd / (base.value + Math.max(0, t.ytdInput) / 2);
  }
  if (d.plan) {
    const target = d.plan.years.find(y => y.year === year);
    if (target) t.target = { year, value: target.target * 10000 };
  }
}

// 美股／台股頁要看到全部部位：券商帳戶＋我的定期定額＋小孩帳戶
function position(o) {
  const cost = o.cost != null ? o.cost : (o.avg != null && o.qty != null ? o.avg * o.qty : null);
  const pl = o.pl != null ? o.pl : (o.value != null && cost != null ? o.value - cost : null);
  return Object.assign({}, o, { cost, pl,
    avg: o.avg != null ? o.avg : (cost != null && o.qty ? cost / o.qty : null),
    pct: o.pct != null ? o.pct : (cost ? pl / cost : null) });
}
function summarize(name, items, extra = {}) {
  const value = items.reduce((s, x) => s + (x.value || 0), 0), cost = items.reduce((s, x) => s + (x.cost || 0), 0);
  return Object.assign({ name, value, cost, earn: value - cost, pct: cost ? (value - cost) / cost : null }, extra);
}

function combineUs(d) {
  const items = [], groups = [];
  if (d.us) {
    d.us.holdings.forEach(x => items.push(position({ group: x.broker, label: x.code, code: x.code, qty: x.qty, price: x.price, avg: x.avg, value: x.value, cost: x.cost, pl: x.pl, pct: x.pct })));
    d.us.brokers.forEach(b => groups.push({ name: b.name, value: b.value, cost: b.cost, earn: b.earn, pct: b.pct, irr: d.us.brokerIrr[b.name] }));
  }
  const dca = d.myDca ? d.myDca.plans.filter(p => p.currency === 'USD' && p.value) : [];
  dca.forEach(p => items.push(position({ group: '定期定額', label: p.name + ' 定期', code: p.name, qty: p.shares, price: p.price, value: p.value, cost: p.total, pct: p.pct,
    note: `每月 US$${dec(p.monthly || 0, 0)}・${p.start} 起` })));
  if (dca.length) groups.push(summarize('定期定額', items.filter(x => x.group === '定期定額')));
  if (d.kidsUs) {
    Object.keys(d.kidsUs.holdings).forEach(kid => d.kidsUs.holdings[kid].forEach(x =>
      items.push(position({ group: '小孩', owner: kid, label: x.code + ' ' + kid, code: x.code, qty: x.qty, price: x.price, avg: x.avg, value: x.value, cost: x.cost, pct: x.pct }))));
    if (items.some(x => x.group === '小孩')) groups.push(summarize('小孩', items.filter(x => x.group === '小孩')));
  }
  return { items, groups, total: summarize('全部', groups) };
}

function combineTw(d) {
  const items = [], groups = [];
  const mine = [];
  if (d.tw) d.tw.holdings.forEach(x => mine.push(position({ group: '我的', label: x.name, name: x.name, code: x.code, qty: x.qty, price: x.price, avg: x.avg, value: x.value, pl: x.pl, pct: x.pct })));
  const dca = d.myDca ? d.myDca.plans.filter(p => p.currency === 'NTD' && p.value) : [];
  dca.forEach(p => mine.push(position({ group: '我的', label: p.name + ' 定期', name: p.name, code: p.name, qty: p.shares, price: p.price, value: p.value, cost: p.total })));
  items.push(...mine);
  if (d.tw && d.tw.total) {
    const dv = dca.reduce((s, p) => s + p.value, 0), dc = dca.reduce((s, p) => s + (p.total || 0), 0);
    const v = d.tw.total.value + dv, c = d.tw.total.cost + dc;
    groups.push({ name: '我的', value: v, cost: c, earn: v - c, pct: c ? (v - c) / c : null, irr: d.tw.irr, cash: d.tw.total.cash });
  }
  // 小孩：定期定額與單筆買進都是「每人」的數字，依股票合併成每個小孩一筆部位
  const k = d.kidsTw;
  if (k && k.kids.length) {
    const byStock = {};
    k.plans.forEach(p => { const s = byStock[p.name] = byStock[p.name] || { qty: 0, value: 0, cost: 0, price: null }; s.qty += p.shares || 0; s.value += p.value || 0; s.cost += p.total || 0; s.price = p.price; });
    k.lumps.forEach(l => { const s = byStock[l.name] = byStock[l.name] || { qty: 0, value: 0, cost: 0, price: null }; s.qty += l.shares || 0; s.value += l.value || 0; s.cost += l.inv || 0; if (s.price == null) s.price = l.cur; });
    k.kids.forEach(kid => Object.keys(byStock).forEach(name => {
      const s = byStock[name];
      items.push(position({ group: '小孩', owner: kid.name, label: name + ' ' + kid.name, name, code: name, qty: s.qty, price: s.price, value: s.value, cost: s.cost }));
    }));
    groups.push(summarize('小孩', items.filter(x => x.group === '小孩')));
  }
  return { items, groups, total: summarize('全部', groups) };
}

/* ================================================================
   資料讀取與快取
   ================================================================ */
let state = { cfg: null, raw: null, updated: '', data: null, status: 'idle', error: '', lastFetch: 0, seq: 0, timing: {} };
const LIVE_STALE_MS = 2 * 60000;   // 即時股價超過這麼久沒更新，讀到資料後再另外要求抓一次

// 送到 Apps Script；順便記下這次花了多久（設定頁會顯示，用來看慢在哪裡）
async function post(body, kind) {
  const t0 = Date.now();
  const res = await fetch(state.cfg.url, { method: 'POST', body: JSON.stringify(Object.assign({ key: state.cfg.key }, body)),
    headers: { 'Content-Type': 'text/plain;charset=utf-8' }, cache: 'no-store', redirect: 'follow' });
  const j = await res.json();
  if (j.ok) state.timing[kind] = { total: Date.now() - t0, ms: j.ms || null };
  return j;
}

function loadLocal() {
  loadLockRecord();
  if (lockEnabled()) return;   // 有鎖：資料是加密的，等解鎖後才讀
  try { state.cfg = JSON.parse(localStorage.getItem(LS.cfg) || 'null'); } catch (e) {}
  try {
    const c = JSON.parse(localStorage.getItem(LS.cache) || 'null');
    if (c && c.raw) { state.raw = c.raw; state.updated = c.updated; state.data = parseAll(c.raw); }
  } catch (e) {}
}

async function refresh() {
  if (lock.locked) return;
  if (!state.cfg || !state.cfg.url) { setStatus('idle'); return; }
  if (state.status === 'busy') return;
  setStatus('busy');
  if (!state.data) render();
  state.seq++;
  try {
    const j = await post({}, '更新');
    if (!j.ok) throw new Error(j.error === 'unauthorized' ? '密碼不對' : (j.error || '讀取失敗'));
    state.raw = j.data; state.updated = j.updated; state.data = parseAll(j.data); state.lastFetch = Date.now();
    await persist();
    setStatus('ok');
    render();
    refreshLivePrices();
  } catch (e) {
    state.error = e.message === 'Failed to fetch' || e.name === 'TypeError' ? '連不上，請檢查網路或網址' : e.message;
    setStatus('err');
    if (!state.data) render();
  }
}

// 資料先顯示出來，股價太舊再另外請 Apps Script 抓即時價（抓價會讓試算表重算，比較久，所以不擋住前面的讀取）
// 這段時間如果又做了更新或寫入，就丟掉這次的結果，免得把比較新的資料蓋掉
async function refreshLivePrices() {
  const l = state.raw && state.raw._live;
  if (!l || scriptVersion() < 27 || Date.now() - (l.at || 0) < LIVE_STALE_MS || state.liveBusy) return;
  const seq = state.seq;
  state.liveBusy = true;
  setStatus('ok');
  try {
    const j = await post({ live: true }, '更新股價');
    if (j.ok && seq === state.seq && !lock.locked) {
      state.raw = j.data; state.updated = j.updated; state.data = parseAll(j.data); state.lastFetch = Date.now();
      await persist();
      state.liveBusy = false;
      setStatus('ok');
      render();
    }
  } catch (e) {}
  state.liveBusy = false;
  if (state.status === 'ok') setStatus('ok');
}

function setStatus(s) {
  state.status = s;
  const el = $('#status'), tx = $('#statusText');
  el.className = 'status ' + s;
  const upd = state.updated ? state.updated.slice(5) : '';
  tx.textContent = s === 'busy' ? '更新中…' : s === 'ok' ? (state.liveBusy ? '更新股價中…' : '已更新 ' + upd) : s === 'err' ? (upd ? '更新失敗・顯示 ' + upd : state.error || '更新失敗') : upd ? '上次 ' + upd : '';
}

/* ================================================================
   圖表（純 SVG，點／拖曳可看數值）
   ================================================================ */
const NS = 'http://www.w3.org/2000/svg';
function svgEl(tag, attrs, parent) {
  const e = document.createElementNS(NS, tag);
  for (const k in attrs) e.setAttribute(k, attrs[k]);
  if (parent) parent.appendChild(e);
  return e;
}
function svgText(parent, x, y, text, attrs = {}) {
  const t = svgEl('text', Object.assign({ x, y, 'font-size': 11, fill: css('--text-muted') }, attrs), parent);
  t.textContent = text;
  return t;
}
// 圖表提示：只有點到圖形（長條、線圖區）才顯示，點一下就一直顯示；點其他地方（包括圖上的文字）才收起來。滑鼠則是移開就收
const openTips = new Set();
document.addEventListener('pointerdown', e => {
  openTips.forEach(tp => {
    if (!tp.host.isConnected) openTips.delete(tp);
    else if (!(e.target.closest && e.target.closest('.hit') && tp.host.contains(e.target))) tp.close();
  });
}, true);
const mouseOnly = fn => ev => { if (ev.pointerType === 'mouse') fn(); };
function makeTip(host) {
  const t = document.createElement('div');
  t.className = 'tip'; host.appendChild(t);
  const tip = {
    host, onClose: null,
    show(html, x, y) {
      t.innerHTML = html; t.style.opacity = 1;
      const w = t.offsetWidth, hw = host.clientWidth;
      let left = x - w / 2; left = Math.max(0, Math.min(hw - w, left));
      t.style.left = left + 'px'; t.style.top = Math.max(-8, y - t.offsetHeight - 12) + 'px';
      openTips.add(tip);
    },
    hide() { t.style.opacity = 0; openTips.delete(tip); },
    close() { if (tip.onClose) tip.onClose(); else tip.hide(); }
  };
  return tip;
}
const tr = (k, v) => `<div class="tr"><span>${k}</span><span>${v}</span></div>`;
// 只在資料端圓角、貼著基線的長條
function barPath(x, y, w, h, dir) {
  const r = Math.max(0, Math.min(4, (dir === 'up' || dir === 'down' ? w : h) / 2, dir === 'up' || dir === 'down' ? h : w));
  if (dir === 'up') return `M${x},${y + h}V${y + r}Q${x},${y} ${x + r},${y}H${x + w - r}Q${x + w},${y} ${x + w},${y + r}V${y + h}Z`;
  if (dir === 'right') return `M${x},${y}H${x + w - r}Q${x + w},${y} ${x + w},${y + r}V${y + h - r}Q${x + w},${y + h} ${x + w - r},${y + h}H${x}Z`;
  return `M${x + w},${y}H${x + r}Q${x},${y} ${x},${y + r}V${y + h - r}Q${x},${y + h} ${x + r},${y + h}H${x + w}Z`; // left
}
function niceStep(max, n) {
  const raw = max / n, p = Math.pow(10, Math.floor(Math.log10(raw)));
  const m = raw / p;
  return (m <= 1 ? 1 : m <= 2 ? 2 : m <= 2.5 ? 2.5 : m <= 5 ? 5 : 10) * p;
}
function pointerX(ev, svg) {
  const r = svg.getBoundingClientRect();
  return (ev.touches ? ev.touches[0].clientX : ev.clientX) - r.left;
}

function trendChart(host, pts) {
  host.innerHTML = '';
  if (pts.length < 2) { host.innerHTML = '<div class="empty">資料不足</div>'; return; }
  const W = host.clientWidth, H = W < 420 ? 220 : 270;
  const m = { t: 14, r: 50, b: 24, l: 6 };
  const svg = svgEl('svg', { viewBox: `0 0 ${W} ${H}`, height: H }, host);
  const x0 = toTime(pts[0].date), x1 = toTime(pts[pts.length - 1].date);
  const X = d => m.l + (toTime(d) - x0) / Math.max(1, x1 - x0) * (W - m.l - m.r);
  const maxV = Math.max(...pts.map(p => Math.max(p.value, p.cost || 0)));
  const step = niceStep(maxV, 5), yMax = Math.ceil(maxV / step) * step;
  const Y = v => m.t + (1 - v / yMax) * (H - m.t - m.b);
  for (let v = 0; v <= yMax + 1; v += step) {
    svgEl('line', { x1: m.l, x2: W - m.r, y1: Y(v), y2: Y(v), stroke: css(v ? '--grid' : '--axis') }, svg);
    if (v) svgText(svg, W - m.r + 6, Y(v) + 4, hidden ? '' : dec(v / 10000, 0) + '萬');
  }
  // x 軸：首尾日期
  svgText(svg, m.l, H - 6, pts[0].date.slice(0, 7));
  svgText(svg, W - m.r, H - 6, pts[pts.length - 1].date.slice(0, 7), { 'text-anchor': 'end' });
  let line = '';
  pts.forEach((p, i) => { line += (i ? 'L' : 'M') + X(p.date).toFixed(1) + ',' + Y(p.value).toFixed(1); });
  svgEl('path', { d: line + `L${X(pts[pts.length - 1].date)},${Y(0)}L${X(pts[0].date)},${Y(0)}Z`, fill: css('--area') }, svg);
  svgEl('path', { d: line, fill: 'none', stroke: css('--gold'), 'stroke-width': 2, 'stroke-linejoin': 'round' }, svg);
  let cp = '';
  pts.forEach((p, i) => { if (p.cost == null) return; const xx = X(p.date).toFixed(1), yy = Y(p.cost).toFixed(1); cp += cp ? `H${xx}V${yy}` : `M${xx},${yy}`; });
  svgEl('path', { d: cp, fill: 'none', stroke: css('--text-muted'), 'stroke-width': 1.6, 'stroke-dasharray': '4 3' }, svg);
  const peak = pts.reduce((a, b) => (b.value > a.value ? b : a));
  svgEl('circle', { cx: X(peak.date), cy: Y(peak.value), r: 4, fill: css('--gold'), stroke: css('--card'), 'stroke-width': 2 }, svg);
  const cross = svgEl('line', { y1: m.t, y2: H - m.b, stroke: css('--axis'), opacity: 0 }, svg);
  const dot = svgEl('circle', { r: 5, fill: css('--gold'), stroke: css('--card'), 'stroke-width': 2, opacity: 0 }, svg);
  const tip = makeTip(host);
  const hit = svgEl('rect', { class: 'hit', x: m.l, y: m.t, width: W - m.l - m.r, height: H - m.t - m.b, fill: 'transparent' }, svg);
  const move = ev => {
    const mx = pointerX(ev, svg);
    let best = pts[0];
    pts.forEach(p => { if (Math.abs(X(p.date) - mx) < Math.abs(X(best.date) - mx)) best = p; });
    const bx = X(best.date);
    cross.setAttribute('x1', bx); cross.setAttribute('x2', bx); cross.setAttribute('opacity', 1);
    dot.setAttribute('cx', bx); dot.setAttribute('cy', Y(best.value)); dot.setAttribute('opacity', 1);
    const earn = best.cost != null ? best.value - best.cost : null;
    tip.show(`<b>${best.date}</b>${tr('總現值', ntd(best.value))}${best.cost != null ? tr('投入成本', ntd(best.cost)) + tr('獲利', signedNtd(earn) + '（' + pct(earn / best.cost) + '）') : ''}${best.note ? tr('備註', esc(best.note)) : ''}`, bx, Y(best.value));
  };
  const leave = () => { tip.hide(); cross.setAttribute('opacity', 0); dot.setAttribute('opacity', 0); };
  hit.addEventListener('pointermove', move);
  hit.addEventListener('pointerdown', move);
  hit.addEventListener('pointerleave', mouseOnly(leave));
  tip.onClose = leave;
}

// 報酬率長條（正負兩側），rows: {label, pct, html}
function divergingChart(host, rows) {
  host.innerHTML = '';
  if (!rows.length) return;
  const W = host.clientWidth, rowH = 30;
  const labelW = Math.min(96, Math.max(...rows.map(r => r.label.length)) * 13 + 14);
  const m = { t: 2, r: 50, b: 2, l: labelW };
  const H = m.t + m.b + rows.length * rowH;
  const svg = svgEl('svg', { viewBox: `0 0 ${W} ${H}`, height: H }, host);
  const lo = Math.min(0, ...rows.map(r => r.pct)), hi = Math.max(0, ...rows.map(r => r.pct));
  const padL = lo < 0 ? 46 : 0;
  const X = v => m.l + padL + (v - lo) / Math.max(0.0001, hi - lo) * (W - m.l - m.r - padL);
  svgEl('line', { x1: X(0), x2: X(0), y1: m.t, y2: H - m.b, stroke: css('--axis') }, svg);
  const tip = makeTip(host);
  rows.forEach((d, i) => {
    const y = m.t + i * rowH, bh = 16, by = y + (rowH - bh) / 2;
    const xa = X(Math.min(0, d.pct)), xb = X(Math.max(0, d.pct)), w = Math.max(2, xb - xa);
    svgEl('path', { d: barPath(d.pct >= 0 ? xa : xb - w, by, w, bh, d.pct >= 0 ? 'right' : 'left'), fill: css(d.pct >= 0 ? '--pos' : '--neg') }, svg);
    svgText(svg, m.l - 8, y + rowH / 2 + 4, d.label, { 'text-anchor': 'end', 'font-size': 12.5, fill: css('--text-primary') });
    svgText(svg, d.pct >= 0 ? xb + 5 : xa - 5, y + rowH / 2 + 4, pct(d.pct), { 'text-anchor': d.pct >= 0 ? 'start' : 'end', 'font-size': 11.5, fill: css('--text-secondary') });
    const bx = d.pct >= 0 ? xa : xb - w, hw = Math.max(w + 12, 32);
    const hit = svgEl('rect', { class: 'hit', x: bx + w / 2 - hw / 2, y: by - 6, width: hw, height: bh + 12, fill: 'transparent' }, svg);
    const show = () => tip.show(`<b>${esc(d.label)}</b>${d.html}`, (xa + xb) / 2, y + 4);
    hit.addEventListener('pointerenter', show);
    hit.addEventListener('pointerdown', show);
    hit.addEventListener('pointerleave', mouseOnly(() => tip.hide()));
  });
}

// 直條圖：series = [{name, color, values[]}]，labels[]，可多組並排
function columnChart(host, labels, series, opts = {}) {
  host.innerHTML = '';
  const W = host.clientWidth, H = opts.height || 220;
  const m = { t: 18, r: 6, b: 24, l: 40 };
  const svg = svgEl('svg', { viewBox: `0 0 ${W} ${H}`, height: H }, host);
  const maxV = Math.max(1, ...series.flatMap(s => s.values.filter(v => v != null)));
  const step = niceStep(maxV, 4), yMax = Math.ceil(maxV / step) * step;
  const Y = v => m.t + (1 - v / yMax) * (H - m.t - m.b);
  const fmtAxis = opts.axis || (v => v);
  for (let v = 0; v <= yMax + 1e-9; v += step) {
    svgEl('line', { x1: m.l, x2: W - m.r, y1: Y(v), y2: Y(v), stroke: css(v ? '--grid' : '--axis') }, svg);
    svgText(svg, m.l - 6, Y(v) + 4, hidden && v ? '' : fmtAxis(v), { 'text-anchor': 'end' });
  }
  const band = (W - m.l - m.r) / labels.length;
  const n = series.length, gap = 2;
  const bw = Math.min(opts.maxBar || 24, (band * 0.7 - gap * (n - 1)) / n);
  const tip = makeTip(host);
  labels.forEach((lab, i) => {
    const cx = m.l + band * i + band / 2;
    const groupW = bw * n + gap * (n - 1);
    series.forEach((s, j) => {
      const v = s.values[i];
      if (v == null || v <= 0) return;
      const x = cx - groupW / 2 + j * (bw + gap);
      svgEl('path', { d: barPath(x, Y(v), bw, Y(0) - Y(v), 'up'), fill: css(s.color) }, svg);
    });
    if (labels.length <= 14 && (i % (opts.labelEvery || 1) === 0)) svgText(svg, cx, H - 6, lab, { 'text-anchor': 'middle' });
    if (opts.topLabel && !hidden) {
      (opts.labelSeries || [0]).forEach(j => {   // 要標數字的系列（預設只有第一個）
        const v = series[j].values[i];
        if (v > 0) svgText(svg, cx - groupW / 2 + j * (bw + gap) + bw / 2, Y(v) - 5, opts.topLabel(v), { 'text-anchor': 'middle', fill: css('--text-secondary') });
      });
    }
    const top = Math.min(Y(0) - 24, Y(Math.max(0, ...series.map(s => s.values[i] || 0))) - 2);
    const hit = svgEl('rect', { class: 'hit', x: cx - groupW / 2 - 6, y: top, width: groupW + 12, height: Y(0) - top, fill: 'transparent' }, svg);
    const show = () => tip.show(opts.tip(i), cx, Y(Math.max(...series.map(s => s.values[i] || 0))));
    hit.addEventListener('pointerenter', show);
    hit.addEventListener('pointerdown', show);
    hit.addEventListener('pointerleave', mouseOnly(() => tip.hide()));
  });
}

// 甜甜圈圖：中間顯示總額，點區塊時換成該類別
function donutChart(host, slices, total) {
  host.innerHTML = '';
  if (!slices.length || !total) return;
  const S = Math.min(host.clientWidth, 260), H = S;
  const cx = host.clientWidth / 2, cy = H / 2, R = S / 2 - 6, r = R * 0.62;
  const svg = svgEl('svg', { viewBox: `0 0 ${host.clientWidth} ${H}`, height: H }, host);
  const arc = (a0, a1, rr) => [cx + rr * Math.sin(a0), cy - rr * Math.cos(a0), cx + rr * Math.sin(a1), cy - rr * Math.cos(a1)];
  const center1 = svgText(svg, cx, cy - 4, '', { 'text-anchor': 'middle', 'font-size': 22, 'font-weight': 700, fill: css('--text-primary') });
  const center2 = svgText(svg, cx, cy + 18, '', { 'text-anchor': 'middle', 'font-size': 12.5, fill: css('--text-secondary') });
  const setCenter = s => {
    center1.textContent = s ? pctPlain(s.value / total) : (hidden ? '••••' : dec(total / 10000, 0) + ' 萬');
    center2.textContent = s ? s.label + (hidden ? '' : '・' + dec(s.value / 10000, 1) + ' 萬') : '總資產';
  };
  let a = 0, active = null;
  const paths = slices.map(s => {
    const sweep = s.value / total * Math.PI * 2;
    const a0 = a, a1 = a + sweep; a = a1;
    const large = sweep > Math.PI ? 1 : 0;
    const [x0, y0, x1, y1] = arc(a0, a1, R), [x2, y2, x3, y3] = arc(a1, a0, r);
    const d = sweep >= Math.PI * 2 - 1e-6
      ? `M${cx},${cy - R}A${R},${R} 0 1 1 ${cx - 0.01},${cy - R}ZM${cx},${cy - r}A${r},${r} 0 1 0 ${cx - 0.01},${cy - r}Z`
      : `M${x0},${y0}A${R},${R} 0 ${large} 1 ${x1},${y1}L${x2},${y2}A${r},${r} 0 ${large} 0 ${x3},${y3}Z`;
    const p = svgEl('path', { d, fill: css(s.color), stroke: css('--card'), 'stroke-width': 2, 'stroke-linejoin': 'round', style: 'cursor:pointer;transition:transform .15s,opacity .15s;transform-origin:' + cx + 'px ' + cy + 'px' }, svg);
    const mid = (a0 + a1) / 2;
    const pick = () => {
      active = active === s ? null : s;
      paths.forEach(q => { q.el.style.opacity = active && q.s !== active ? 0.35 : 1; q.el.style.transform = q.s === active ? `translate(${4 * Math.sin(q.mid)}px, ${-4 * Math.cos(q.mid)}px)` : ''; });
      setCenter(active);
    };
    p.addEventListener('click', pick);
    return { el: p, s, mid };
  });
  setCenter(null);
}

// 橫條：rows {label, value, color, html, text}
function hbarChart(host, rows, right = 64) {
  host.innerHTML = '';
  const W = host.clientWidth, rowH = 30;
  const m = { t: 2, r: right, b: 2, l: 76 };
  const H = m.t + m.b + rows.length * rowH;
  const svg = svgEl('svg', { viewBox: `0 0 ${W} ${H}`, height: H }, host);
  const max = Math.max(1, ...rows.map(r => r.value));
  const tip = makeTip(host);
  rows.forEach((r, i) => {
    const y = m.t + i * rowH, w = Math.max(2, r.value / max * (W - m.l - m.r));
    svgEl('path', { d: barPath(m.l, y + 7, w, 16, 'right'), fill: css(r.color || '--s1') }, svg);
    svgText(svg, m.l - 8, y + rowH / 2 + 4, r.label, { 'text-anchor': 'end', 'font-size': 12.5, fill: css('--text-primary') });
    svgText(svg, m.l + w + 6, y + rowH / 2 + 4, hidden ? '••••' : r.text, { 'font-size': 11.5, fill: css('--text-secondary') });
    const hit = svgEl('rect', { class: 'hit', x: m.l, y: y + 2, width: Math.max(w, 32) + 4, height: rowH - 4, fill: 'transparent' }, svg);
    const show = () => tip.show(`<b>${esc(r.label)}</b>${r.html}`, m.l + w / 2, y + 4);
    hit.addEventListener('pointerenter', show);
    hit.addEventListener('pointerdown', show);
    hit.addEventListener('pointerleave', mouseOnly(() => tip.hide()));
  });
}

/* ================================================================
   頁面
   ================================================================ */
const charts = [];   // 重畫用：[{id, draw}]
function chart(id, draw) { charts.push({ id, draw }); return `<div class="chart" id="${id}"></div>`; }
function drawCharts() { charts.forEach(c => { const el = document.getElementById(c.id); if (el) c.draw(el); }); }

let trendRange = '1Y';
let usBroker = 'all';
let twGroup = 'all';

function pageOverview(d) {
  const t = d.total;
  if (!t) return '<div class="empty">找不到 Total Overall 的資料</div>';
  const earnPct = t.cost ? t.earn / t.cost : null;
  let h = `${head('總覽', [['＋ 成本變更', 'cost']])}
  <div class="card hero">
    <div class="lbl">總資產現值（不含房產）</div>
    <div class="big">${ntd(t.value)}</div>
    <div class="meta">約 ${usd(t.valueUSD)}・匯率 ${t.fx ? dec(t.fx, 2) : '—'}</div>
    <div style="margin-top:10px;display:flex;gap:8px;flex-wrap:wrap">
      <span class="pill ${cls(t.earn)}">累計 ${signedNtd(t.earn)}（${pct(earnPct)}）</span>
      <span class="pill up" style="background:var(--tint-wash);color:var(--tint)">累計 IRR ${pctPlain(t.irr)}</span>
    </div>
  </div>
  <div class="tiles t4">
    <div class="tile"><div class="k">投入成本</div><div class="v">${wan(t.cost)}</div><div class="n">累計淨投入</div></div>
    <div class="tile"><div class="k">距歷史高點</div><div class="v ${cls(t.drawdown)}">${t.drawdown >= 0 ? '新高' : signedWan(t.drawdown)}</div><div class="n">${t.peak ? '高點 ' + wan(t.peak.value) + '・' + t.peak.date.slice(5) : ''}</div></div>
    <div class="tile"><div class="k">今年實際獲利</div><div class="v ${cls(t.ytd)}">${signedWan(t.ytd)}</div><div class="n">${t.ytdInput != null ? '已扣今年投入 ' + wan(t.ytdInput) : ''}</div></div>
    <div class="tile"><div class="k">${t.target ? t.target.year + ' 年底目標' : '年度目標'}</div><div class="v">${t.target ? wan(t.target.value) : '—'}</div><div class="n">${t.target ? (t.value >= t.target.value ? '已達成' : '還差 ' + wan(t.target.value - t.value)) : ''}</div></div>
  </div>`;

  h += todayCard(d);
  h += irrCard(t);
  h += `<div class="card"><h3>資產走勢</h3><p class="cap">總現值（藍）與累計投入成本（虛線）</p>
    <div class="seg" id="rangeSeg">${['6M', '1Y', 'ALL'].map(k => `<button data-r="${k}" class="${trendRange === k ? 'on' : ''}">${{ '6M': '6 個月', '1Y': '1 年', ALL: '全部' }[k]}</button>`).join('')}</div>
    ${chart('c-trend', el => trendChart(el, rangePts(t.series)))}</div>`;

  const alloc = t.alloc.filter(a => a.value > 0);
  const sum = alloc.reduce((s, a) => s + a.value, 0);
  h += `<div class="card"><h3>資產配置</h3><p class="cap">依 Total Overall 的分類，點區塊可看金額</p>
    ${chart('c-alloc', el => donutChart(el, alloc.map(a => ({ label: a.name, value: a.value, color: a.color })), sum))}</div>
  <div class="list">${alloc.map(a => `<div class="row link" data-go="${allocLink(a.cat)}">
      <span class="swatch" style="background:var(${a.color})"></span>
      <div class="main"><div class="t1">${esc(a.name)}</div><div class="t2">${a.parts.map(p => esc(p.sub === 'NTD' || p.sub === 'USD' ? p.where : p.sub)).filter(Boolean).join('・')}</div></div>
      <div class="r"><div class="t1">${wan(a.value)}</div><div class="t2" style="color:var(--text-secondary)">${pctPlain(a.value / sum)}</div></div>
      <span class="chev">›</span></div>`).join('')}</div>`;

  h += cashSection('', '各帳戶閒置現金');
  const flows = (d.ledger ? d.ledger.flow : t.history.filter(x => x.note)).slice(-6).reverse();
  if (flows.length) {
    h += `<div class="sec">最近資金異動<span class="seclink" data-go="more/ledger">全部明細 ›</span></div><div class="list">${flows.map(f => `<div class="row${f.row && ed() ? ' link' : ''}"${f.row && ed() ? ` data-form="ledger" data-arg="flow:${f.row}"` : ''}>
      <div class="main"><div class="t1" style="font-size:15px">${esc(f.note)}</div><div class="t2">${f.date}</div></div>
      <div class="r"><div class="t1 ${f.flow > 0 ? 'down' : ''}" style="font-size:15px">${f.flow ? (f.flow < 0 ? '投入 ' + ntd(-f.flow) : '取出 ' + ntd(f.flow)) : ''}</div></div></div>`).join('')}</div>`;
  }
  return h;
}
// 今日漲跌：Apps Script（第 21 版起）回傳每一檔的漲跌幅，乘上各帳戶持股現值；包含我的、定期定額、小孩帳戶
// 每檔前一個交易日到現在：現值 − 現值÷(1＋漲跌幅)；美股換成台幣；抓不到價格的（例如上櫃股）略過
function todayChange(d) {
  const m = state.raw && state.raw._daily;
  if (!m || !Object.keys(m).length) return null;
  const fx = (d.total && d.total.fx) || 0;
  const sum = (items, key, rate) => {
    let ch = 0, base = 0, n = 0;
    items.forEach(x => { const p = m[key(x)]; if (p == null || !x.value) return; const prev = x.value / (1 + p); ch += x.value - prev; base += prev; n++; });
    return n ? { ch: ch * rate, pct: ch / base } : null;
  };
  return [
    ['美股', d.usAll ? sum(d.usAll.items, x => 'US:' + str(x.code).toUpperCase(), fx) : null],
    ['台股', d.twAll ? sum(d.twAll.items, x => 'TW:' + str(x.code), 1) : null],
    ['加密貨幣', d.crypto ? sum(d.crypto.coins.filter(c => c.name !== 'USDT'), x => 'C:' + str(x.name).toUpperCase(), 1) : null]
  ];
}
// 報酬率（IRR）：每年以前一年 12/31 的總資產為初始金額，加上當年入金、出金，年底（今年用目前）結算
function irrCard(t) {
  const c = t.irrCalc;
  if (!c || !c.years.length) return '';
  const cur = c.years[c.years.length - 1], wanOf = v => (hidden ? MASK : dec(v / 10000, 1) + ' 萬');
  const md = ms => { const x = new Date(ms); return `${x.getFullYear()}/${pad(x.getMonth() + 1)}/${pad(x.getDate())}`; };
  return `<div class="card"><h3>報酬率（IRR）</h3><p class="cap">每年以前一年 12/31 的總資產當初始金額，加上當年的入金、出金計算；今年算到今天</p>
    <div class="today" style="grid-template-columns:repeat(${cur.partial ? 3 : 2},minmax(0,1fr))">
      ${cur.partial ? `<div><div class="k">${cur.year} 年 IRR</div><div class="v ${cls(cur.irr)}">${pct(cur.irr)}</div><div class="n">年化</div></div>
      <div><div class="k">${cur.year} 年至今</div><div class="v ${cls(cur.ret)}">${pct(cur.ret)}</div><div class="n">實際報酬</div></div>` : ''}
      <div><div class="k">累計 IRR</div><div class="v ${cls(c.all)}">${pct(c.all)}</div><div class="n">從第一筆入金</div></div>
    </div>
    <div style="margin-top:12px">${c.years.slice().reverse().map(y => `<div class="kv"><span>${y.year}${y.partial ? '（至今）' : ''}<br><small style="color:var(--text-muted)">${wanOf(y.start.v)}${y.exactStart ? '' : `（${md(y.start.t).slice(5)}）`} → ${wanOf(y.end.v)}・淨投入 ${wanOf(y.net)}</small></span>
      <span class="${cls(y.irr)}">${pct(y.irr)}${y.partial ? '<br><small style="color:var(--text-muted)">至今 ' + pct(y.ret) + '</small>' : ''}</span></div>`).join('')}</div></div>`;
}
function todayCard(d) {
  const t = todayChange(d);
  if (!t) return '';
  return `<div class="card"><h3>今日漲跌</h3><p class="cap">最近一個交易日，依目前持股計算（含定期定額、小孩帳戶）</p>
    <div class="today">${t.map(([name, c]) => `<div><div class="k">${name}</div>
      <div class="v ${c ? cls(c.pct) : ''}">${c ? pct(c.pct) : '—'}</div><div class="n">${c ? signedNtd(c.ch) : '還沒有價格'}</div></div>`).join('')}</div></div>`;
}

function allocLink(cat) {
  return { '美股': 'us', '台股': 'tw', '虛擬貨幣': 'crypto', "Kids' 台股": 'more/kids', "Kids' 美股": 'more/kids', '放貸': 'crypto' }[cat] || 'overview';
}
function rangePts(pts) {
  if (trendRange === 'ALL' || !pts.length) return pts;
  const last = toTime(pts[pts.length - 1].date);
  const days = trendRange === '6M' ? 183 : 365;
  const out = pts.filter(p => toTime(p.date) >= last - days * 864e5);
  return out.length >= 2 ? out : pts;
}

function holdingRows(list, money) {
  return list.map(x => `<div class="row${x.form ? ' link' : ''}"${x.form ? ` data-form="${x.form}" data-arg="${esc(x.arg)}"` : ''}>
    <div class="main"><div class="t1">${esc(x.title)}</div><div class="t2">${x.sub}</div></div>
    <div class="r"><div class="t1">${money(x.value)}</div><div class="t2 ${cls(x.pct)}">${pct(x.pct)}</div></div></div>`).join('');
}

// 持股頁共用：分組按鈕、報酬率圖、明細
const groupLabel = x => (x.merged ? mergedLabel(x) : x.owner ? x.group + '・' + x.owner : x.group);
// 同一檔在不同帳戶（我的、定期定額、小孩）的部位合併成一列：股數、現值、成本加總，均價＝總成本÷總股數
function mergeHoldings(items) {
  const by = {}, order = [];
  items.forEach(x => { const k = str(x.code).toUpperCase() || x.label; if (!by[k]) { by[k] = []; order.push(k); } by[k].push(x); });
  return order.map(k => {
    const ps = by[k];
    if (ps.length === 1) return ps[0];
    const own = ps.find(p => p.group === '我的' && !/ 定期$/.test(p.label)) || ps.find(p => p.group === '我的') || ps[0];
    const sum = f => ps.reduce((t, p) => t + (p[f] || 0), 0);
    const qtyAll = sum('qty'), value = sum('value'), cost = sum('cost');
    return Object.assign({}, own, { label: own.group === '我的' && !/ 定期$/.test(own.label) ? own.label : (own.name || str(own.code)),
      qty: qtyAll, value, cost, pl: value - cost, pct: cost ? (value - cost) / cost : null, avg: qtyAll ? cost / qtyAll : null, merged: ps });
  });
}
function mergedLabel(x) {
  const g = {};
  x.merged.filter(p => p.qty > 0).forEach(p => { const n = p.owner || (/ 定期$/.test(p.label) ? '定期定額' : p.group); g[n] = (g[n] || 0) + (p.qty || 0); });
  return Object.keys(g).map(n => (hidden ? n : `${n} ${dec(g[n], Number.isInteger(g[n]) ? 0 : 2)} 股`)).join('＋');
}
// priceForm：點持股可以打開的表單（台股：更新現價），回傳 [表單, 參數] 或 null
// between：放在「個股報酬率」和「持股明細」中間的內容（持股比重）
// assets：總資產（台幣）；rate：這頁幣別換台幣的匯率（台股 1、美股用美金匯率）
function weightCard(id, items, money, assets, rate = 1) {
  const list = mergeHoldings(items).filter(x => x.value > 0).sort((a, b) => b.value - a.value);
  const tv = list.reduce((s2, x) => s2 + x.value, 0);
  if (!tv) return '';
  const ofAll = assets && rate ? x => x.value * rate / assets : null;
  return `<div class="card"><h3>持股比重</h3><p class="cap">依目前市值，同一檔股票合併計算${ofAll ? '；長條右邊：佔這頁持股／佔總資產' : ''}</p>
    ${chart(id, el => hbarChart(el, list.map(x => ({ label: x.label, value: x.value, text: pctPlain(x.value / tv) + (ofAll ? '／' + pctPlain(ofAll(x)) : ''),
      html: tr('現值', money(x.value)) + tr('佔這頁持股', pctPlain(x.value / tv)) + (ofAll ? tr('佔總資產', pctPlain(ofAll(x))) : '') })), ofAll ? 92 : 64))}</div>`;
}
function holdingsSection(id, all, filter, segId, money, priceFmt, priceForm, between = '') {
  const names = all.groups.map(g => g.name);
  const shown = all.items.filter(x => filter === 'all' || x.group === filter);
  const list = (filter === 'all' ? mergeHoldings(shown) : shown).sort((a, b) => b.pct - a.pct);
  return `<div class="card"><h3>個股報酬率</h3><p class="cap">點長條可看持股與損益</p>
    <div class="seg" id="${segId}"><button data-b="all" class="${filter === 'all' ? 'on' : ''}">全部</button>${names.map(b => `<button data-b="${esc(b)}" class="${filter === b ? 'on' : ''}">${esc(b)}</button>`).join('')}</div>
    ${chart(id, el => divergingChart(el, list.map(x => ({ label: x.label, pct: x.pct,
      html: tr('帳戶', esc(groupLabel(x))) + tr('股數', qty(x.qty)) + tr('均價／現價', hidden ? MASK : priceFmt(x.avg) + '／' + priceFmt(x.price)) + tr('現值', money(x.value)) + tr('損益', money === ntd ? signedNtd(x.pl) : signedUsd(x.pl)) }))))}</div>
  ${between}<div class="sec">持股明細</div><div class="list">${holdingRows(list.slice().sort((a, b) => b.value - a.value).map(x => ({
    title: x.label, value: x.value, pct: x.pct, sub: `${esc(groupLabel(x))}${x.merged ? '' : '・' + qty(x.qty) + ' 股'}・均價 ${hidden ? MASK : priceFmt(x.avg)}${x.manual ? '・現價手動' : ''}`,
    form: priceForm && priceForm(x) ? priceForm(x)[0] : null, arg: priceForm && priceForm(x) ? priceForm(x)[1] : null })), money)}</div>`;
}

function pageUs(d) {
  const u = d.us, all = d.usAll;
  if (!u || !all) return '<div class="empty">找不到美股的資料</div>';
  const t = all.total;
  const fx = d.total && d.total.fx;
  if (!all.groups.some(g => g.name === usBroker)) usBroker = 'all';
  let h = `${head('美股', [['＋ 買賣', 'trade', 'us']])}
  <div class="card hero">
    <div class="lbl">美股總值（含定期定額與小孩）</div>
    <div class="big">${usd(t.value)}</div>
    <div class="meta">投入 ${usd(t.cost)}${fx ? '・約 ' + ntd(t.value * fx) : ''}</div>
    <div style="margin-top:10px;display:flex;gap:8px;flex-wrap:wrap"><span class="pill ${cls(t.earn)}">${signedUsd(t.earn)}（${pct(t.pct)}）</span>
    ${u.irr != null ? `<span class="pill up" style="background:var(--tint-wash);color:var(--tint)">券商帳戶 IRR ${pctPlain(u.irr)}</span>` : ''}</div>
  </div>
  <div class="tiles">${all.groups.map(g => `<div class="tile"><div class="k">${esc(g.name)}</div><div class="v">${usd(g.value)}</div>
    <div class="n"><span class="${cls(g.pct)}">${pct(g.pct)}</span>${g.irr != null ? '・IRR ' + pctPlain(g.irr) : ''}</div></div>`).join('')}</div>`;
  h += cashSection('us');
  // 持股比重放在個股報酬率和持股明細中間；同一檔（例如我的 QQQ 和小孩的 QQQ）合併計算
  h += holdingsSection('c-us', all, usBroker, 'brokerSeg', usd, v => (v == null ? '—' : '$' + dec(v, 2)), null,
    weightCard('c-us-w', all.items.filter(x => usBroker === 'all' || x.group === usBroker), usd, d.total && d.total.value, fx));
  h += txList(u.tx.slice(-12).reverse().map(x => ({
    arg: 'us:' + x.row, date: x.date, title: `${x.buy ? '買進' : x.sell ? '賣出' : ''} ${x.name}`, sub: `${esc(x.where)}・${x.buy || x.sell ? qty(x.buy || x.sell) + ' 股' : ''} @ ${x.price != null ? '$' + dec(x.price, 2) : '—'}`,
    right: x.buy ? `<span>買</span>` : x.sell ? '<span class="up">賣</span>' : '' })));
  return h;
}

// 閒置現金（Apps Script 第 16 版起回傳）：點一下可以改成券商 App 上的數字（例如收到股利）
function cashSection(page, title = '閒置現金') {
  const e = ed();
  if (!e || !e.cash) return '';
  const list = page ? e.cash.filter(a => a.page === page) : e.cash;
  if (!list.length) return '';
  const money = a => (a.cur === 'NTD' ? ntd(a.value) : a.cur === 'USD' ? usd(a.value, 2) : hidden ? MASK : dec(a.value || 0, 2) + ' USDT');
  return `<div class="sec">${title}<span class="seclink" style="cursor:default;color:var(--text-muted)">點一下可以修改</span></div><div class="list">${list.map(a => `<div class="row link" data-form="cash" data-arg="${esc(a.id)}">
    <div class="main"><div class="t1" style="font-size:15px">${esc(a.label)}</div></div><div class="r"><div class="t1" style="font-size:15px">${money(a)}</div></div><span class="chev">›</span></div>`).join('')}</div>`;
}

function txList(items) {
  if (!items.length) return '';
  return `<div class="sec">最近交易<span class="seclink" data-go="more/ledger">全部明細 ›</span></div><div class="list">${items.map(x => `<div class="row${x.arg && ed() ? ' link' : ''}"${x.arg && ed() ? ` data-form="ledger" data-arg="${x.arg}"` : ''}>
    <div class="main"><div class="t1" style="font-size:15px">${esc(x.title)}</div><div class="t2">${x.date}・${x.sub}</div></div>
    <div class="r" style="font-size:14px;font-weight:600">${x.right}</div></div>`).join('')}</div>`;
}

function pageTw(d) {
  const w = d.tw, all = d.twAll;
  if (!w || !all) return '<div class="empty">找不到台股的資料</div>';
  const t = all.total;
  if (!all.groups.some(g => g.name === twGroup)) twGroup = 'all';
  let h = `${head('台股', [['＋ 買賣', 'trade', 'tw']])}
  <div class="card hero">
    <div class="lbl">台股總值（含定期定額與小孩）</div>
    <div class="big">${ntd(t.value)}</div>
    <div class="meta">投入 ${ntd(t.cost)}${w.total ? '・現金 ' + ntd(w.total.cash) : ''}</div>
    <div style="margin-top:10px;display:flex;gap:8px;flex-wrap:wrap"><span class="pill ${cls(t.earn)}">${signedNtd(t.earn)}（${pct(t.pct)}）</span>
    ${w.irr != null ? `<span class="pill up" style="background:var(--tint-wash);color:var(--tint)">我的帳戶 IRR ${pctPlain(w.irr)}</span>` : ''}</div>
  </div>
  <div class="tiles">${all.groups.map(g => `<div class="tile"><div class="k">${esc(g.name)}</div><div class="v">${ntd(g.value)}</div>
    <div class="n"><span class="${cls(g.pct)}">${pct(g.pct)}</span>${g.irr != null ? '・IRR ' + pctPlain(g.irr) : ''}</div></div>`).join('')}</div>`;
  h += cashSection('tw');
  // 我的台股持股：點一下可以手動填現價（上櫃等抓不到價格的）
  const prices = (ed() && ed().tw && ed().tw.prices) || [];
  const priceOf = x => (x.group === '我的' ? prices.find(q => q.code === str(x.code).toUpperCase()) : null);
  all.items.forEach(x => { const p = priceOf(x); x.manual = !!(p && p.manual); });
  // 持股比重放在個股報酬率和持股明細中間；同一檔（例如我的 0050 和小孩的 0050）用代號合併計算
  h += holdingsSection('c-tw', all, twGroup, 'twSeg', ntd, v => (v == null ? '—' : dec(v, 1)),
    x => (priceOf(x) ? ['twPrice', priceOf(x).code] : null), weightCard('c-tw-w', all.items.filter(x => twGroup === 'all' || x.group === twGroup), ntd, d.total && d.total.value));
  if (prices.length) h += '<p class="hint" style="margin:-4px 4px 12px">點自己的持股可以手動填現價（上櫃等抓不到價格的股票）</p>';
  h += txList(w.tx.filter(x => x.buy || x.sell).slice(-12).reverse().map(x => ({
    arg: 'tw:' + x.row, date: x.date, title: x.buy || x.sell ? `${x.buy ? '買進' : '賣出'} ${x.name}` : x.flow ? (x.flow < 0 ? '入金' : '出金') : x.name,
    sub: x.buy || x.sell ? `${qty(x.buy || x.sell)} 股 @ ${x.price != null ? dec(x.price, 1) : '—'}` : [x.flow ? ntd(Math.abs(x.flow)) : '', esc(x.note)].filter(Boolean).join('・'),
    right: x.buy ? '<span>買</span>' : x.sell ? '<span class="up">賣</span>' : '' })));
  return h;
}

function pageCrypto(d) {
  const c = d.crypto;
  if (!c) return '<div class="empty">找不到虛擬貨幣的資料</div>';
  const lend = d.total ? d.total.items.filter(i => i.cat === '放貸') : [];
  const groups = {};
  c.coins.forEach(x => { (groups[x.name] = groups[x.name] || { name: x.name, qty: 0, usd: 0, value: 0, price: x.price, where: [] }); const g = groups[x.name]; g.qty += x.qty; g.usd += x.usd || 0; g.value += x.value || 0; g.where.push(x.where); });
  const gl = Object.values(groups).sort((a, b) => b.value - a.value);
  const colors = { BTC: '--s4', ETH: '--s7', SOL: '--s3', DOGE: '--s2' };
  let h = `${head('加密貨幣', [['＋ 買賣', 'trade', 'crypto']])}
  <div class="card hero">
    <div class="lbl">加密貨幣總值</div>
    <div class="big">${ntd(c.total)}</div>
    <div class="meta">約 ${usd(c.totalUSD)}${d.total && d.total.value ? '・佔總資產 ' + pctPlain(c.total / d.total.value) : ''}</div>
  </div>
  <div class="card"><h3>幣別比重</h3><p class="cap">依目前市值</p>
    <div class="stack">${gl.map(g => `<div style="width:${g.value / c.total * 100}%;background:var(${colors[g.name] || '--s1'})"></div>`).join('')}</div>
    <div class="legend" style="margin:8px 0 0">${gl.map(g => `<span><i style="background:var(${colors[g.name] || '--s1'})"></i>${esc(g.name)} ${pctPlain(g.value / c.total)}${d.total && d.total.value ? '（總 ' + pctPlain(g.value / d.total.value) + '）' : ''}</span>`).join('')}</div>
    ${d.total && d.total.value ? '<p class="cap" style="margin:6px 0 0">括號內是佔總資產的比重</p>' : ''}
  </div>
  <div class="sec">持有</div><div class="list">${gl.map(g => `<div class="row">
    <span class="badge" style="background:var(${colors[g.name] || '--s1'})">${esc(g.name.slice(0, 1))}</span>
    <div class="main"><div class="t1">${esc(g.name)}</div><div class="t2">${qty(g.qty)} 顆・${[...new Set(g.where)].map(esc).join('、')}</div></div>
    <div class="r"><div class="t1">${ntd(g.value)}</div><div class="t2" style="color:var(--text-secondary)">$${g.price != null ? dec(g.price, g.price < 10 ? 3 : 0) : '—'}</div></div></div>`).join('')}</div>`;
  if (c.memo && c.memo.length) {
    h += `<div class="sec">時間備忘</div><div class="card">${c.memo.map(m => m.k
      ? `<div class="kv"><span>${esc(m.k)}</span><span>${esc(m.v)}</span></div>`
      : `<p class="cap" style="margin:8px 0 0">${esc(m.v)}</p>`).join('')}</div>`;
  }
  h += cashSection('crypto');
  if (c.coins.length) {
    h += `<div class="sec">存放位置</div><div class="list">${c.coins.map(x => `<div class="row"><div class="main"><div class="t1" style="font-size:15px">${esc(x.name)}・${esc(x.where)}</div><div class="t2">${qty(x.qty)} 顆</div></div><div class="r"><div class="t1" style="font-size:15px">${ntd(x.value)}</div></div></div>`).join('')}</div>`;
  }
  if (c.realized.length) {
    h += `<div class="sec">已兌現</div><div class="list">${c.realized.map(x => `<div class="row"><div class="main"><div class="t1" style="font-size:15px">賣出 ${esc(x.coin)}</div><div class="t2">${x.date}</div></div><div class="r"><div class="t1" style="font-size:15px">${qty(x.qty)} 顆</div></div></div>`).join('')}</div>`;
  }
  if (lend.length) {
    h += `<div class="sec">放貸</div><div class="list">${lend.map(x => `<div class="row"><div class="main"><div class="t1" style="font-size:15px">${esc(x.sub)}・${esc(x.where)}</div><div class="t2">${x.value ? '' : '目前無部位'}</div></div><div class="r"><div class="t1" style="font-size:15px">${ntd(x.value || 0)}</div></div></div>`).join('')}</div>`;
  }
  return h;
}

// 明細與復原：四種明細切換；每列點進去可以修改、刪除；上面有「復原上一步」
let ledgerKind = 'flow';
const LEDGER_KINDS = [['flow', '資金異動'], ['tw', '台股'], ['us', '美股'], ['crypto', '加密貨幣']];
function ledgerLine(kind, x) {
  const m = v => (hidden ? MASK : v);
  if (kind === 'flow') return { t1: x.note || (x.flow < 0 ? '投入' : '撤出'), t2: x.date, r: `<span class="${x.flow > 0 ? 'down' : ''}">${x.flow < 0 ? '投入' : '撤出'} ${m(ntd(Math.abs(x.flow)))}</span>` };
  if (kind === 'crypto') {
    const trade = x.type === '買進' || x.type === '賣出';
    return { t1: trade ? `${x.type} ${x.coin}` : `加密貨幣 ${x.type}`, t2: [esc(x.date), esc(x.where), trade ? `${qty(x.qty)} 顆 @ $${dec(x.price, 2)}` : '', esc(x.note)].filter(Boolean).join('・'),
      r: `${x.usdt > 0 ? '+' : x.usdt < 0 ? '−' : ''}${m(dec(Math.abs(x.usdt || 0), 2))} USDT` };
  }
  const money = kind === 'tw' ? v => ntd(v) : v => usd(v, 2);
  if (x.buy || x.sell) {
    const name = kind === 'tw' ? x.b : x.code;
    return { t1: `${x.buy ? '買進' : '賣出'} ${name}`, t2: [esc(x.date), kind === 'us' ? esc(x.b) : '', `${qty(x.buy || x.sell)} 股 @ ${dec(x.price, 2)}`, esc(x.note)].filter(Boolean).join('・'),
      r: x.buy ? '<span>買</span>' : '<span class="up">賣</span>' };
  }
  return { t1: `${kind === 'us' ? '美股 ' + x.b : '台股'} ${x.flow < 0 ? '入金' : '出金'}`, t2: [x.date, x.note].filter(Boolean).map(esc).join('・'),
    r: m(money(Math.abs(x.flow || 0))) };
}
// 明細清單：台股、美股、加密貨幣只列買賣；資金異動列 Total Overall 的資金異動，
// 再加上沒有對應資金異動的出入金（例如台股轉到加密貨幣），同一組只列一次
function ledgerList(L, kind) {
  if (kind !== 'flow') return (L[kind] || []).filter(x => !isFlowEntry(kind, x)).slice().reverse().map(x => ({ k: kind, x }));
  const items = (L.flow || []).map(x => ({ k: 'flow', x })), seen = {};
  ['tw', 'us', 'crypto'].forEach(k => (L[k] || []).forEach(x => {
    if (!isFlowEntry(k, x) || seen[k + ':' + x.row]) return;
    const links = linkedOf(k, x);
    if (links.some(l => l.k === 'flow')) return;   // 有對應的資金異動，點那一筆就能改
    links.forEach(l => { seen[l.k + ':' + l.y.row] = true; });
    items.push({ k, x });
  }));
  return items.map((it, i) => Object.assign(it, { i })).sort((a, b) => (a.x.date < b.x.date ? 1 : a.x.date > b.x.date ? -1 : b.i - a.i));
}
function pageLedger(d) {
  const L = d.ledger || {}, e = ed(), u = e && e.undo;
  let h = '<h1 class="title">明細與復原</h1>';
  if (!e) h += `<div class="card"><p class="msg">要把 Apps Script 更新到第 ${LATEST_SCRIPT} 版，才能修改明細和復原。</p></div>`;
  else h += `<div class="card"><h3>復原</h3><p class="cap">${u && u.count ? `上一步：${esc(u.desc)}（${esc(u.time)}）<br>最近 ${u.count} 次從 App 的修改都可以一步步復原` : '目前沒有可以復原的動作'}</p>
    ${u && u.count ? '<button class="btn soft" id="undoBtn" style="margin-top:10px">復原上一步</button>' : ''}</div>`;
  h += `<div class="seg" id="ledgerSeg">${LEDGER_KINDS.map(k => `<button data-k="${k[0]}" class="${k[0] === ledgerKind ? 'on' : ''}">${k[1]}</button>`).join('')}</div>`;
  const list = ledgerList(L, ledgerKind);
  if (!list.length) return h + '<div class="empty">沒有資料</div>';
  h += `<p class="hint" style="margin:0 4px 8px">${e ? '點一筆可以修改或刪除；' : ''}最新的在最上面，共 ${list.length} 筆${ledgerKind === 'flow' ? '；各帳戶的入金、出金都在這裡' : '（只列買賣，入金出金在資金異動）'}</p>`;
  return h + `<div class="list">${list.map(({ k, x }) => { const l = ledgerLine(k, x); return `<div class="row${e ? ' link' : ''}"${e ? ` data-form="ledger" data-arg="${k}:${x.row}"` : ''}>
    <div class="main"><div class="t1" style="font-size:15px">${esc(l.t1)}</div><div class="t2" style="white-space:normal">${l.t2}</div></div>
    <div class="r" style="font-size:14px;font-weight:600">${l.r}</div></div>`; }).join('')}</div>`;
}

const MORE = [
  { id: 'kids', title: '小孩帳戶', sub: '小孩的美股與台股', icon: '★', color: '--s5' },
  { id: 'dca', title: '我的定期定額', sub: '台幣扣款 QQQ／0050', icon: '↻', color: '--s3' },
  { id: 'plan', title: '長期計畫', sub: '目標與實際', icon: '◎', color: '--s1' },
  { id: 'expense', title: '固定支出', sub: '每月與年度支出', icon: '▤', color: '--s2' },
  { id: 'house', title: '房產', sub: '購入價與估計現值', icon: '⌂', color: '--s6' },
  { id: 'milestone', title: '里程碑', sub: '重要資金動作', icon: '⚑', color: '--s7' },
  { id: 'career', title: '職涯與薪資', sub: '年資、職位、薪水與分紅', icon: '▲', color: '--s4' },
  { id: 'ledger', title: '明細與復原', sub: '修改、刪除資金異動與交易；復原上一步', icon: '☰', color: '--s8' },
  { id: 'settings', title: '設定', sub: '資料來源與密碼', icon: '⚙', color: '--axis' }
];
function pageMore() {
  // 試算表網址由 Apps Script 回傳，程式碼裡不寫死
  const url = state.raw && state.raw._sheetUrl;
  const sheetRow = url
    ? `<a class="row link" href="${esc(url)}" target="_blank" rel="noopener" style="text-decoration:none;color:inherit">
        <span class="badge" style="background:#188038">▦</span>
        <div class="main"><div class="t1">原始資料</div><div class="t2">打開 Google 試算表</div></div><span class="chev">↗</span></a>`
    : `<div class="row"><span class="badge" style="background:var(--axis)">▦</span>
        <div class="main"><div class="t1">原始資料</div><div class="t2">請先更新 Apps Script 程式碼並部署新版本</div></div></div>`;
  return `<h1 class="title">更多</h1><div class="list">${sheetRow}</div><div class="list">${MORE.map(m => `<div class="row link" data-go="more/${m.id}">
    <span class="badge" style="background:var(${m.color})">${m.icon}</span>
    <div class="main"><div class="t1">${m.title}</div><div class="t2">${m.sub}</div></div><span class="chev">›</span></div>`).join('')}</div>`;
}

function pageKids(d) {
  const u = d.kidsUs, w = d.kidsTw;
  const names = [...new Set([...(u ? u.kids.map(k => k.name) : []), ...(w ? w.kids.map(k => k.name) : [])])];
  let h = head('小孩帳戶', [['＋ 本月定期定額', 'dca']]);
  const totalVal = (u && u.total ? u.total.valNTD || 0 : 0) + (w && w.total ? w.total.value || 0 : 0);
  const totalInv = (u && u.total ? u.total.costNTD || 0 : 0) + (w && w.total ? w.total.inv || 0 : 0);
  h += `<div class="card hero"><div class="lbl">兩人合計現值</div><div class="big">${ntd(totalVal)}</div>
    <div class="meta">投入 ${ntd(totalInv)}</div><div style="margin-top:10px"><span class="pill ${cls(totalVal - totalInv)}">${signedNtd(totalVal - totalInv)}（${pct(totalInv ? (totalVal - totalInv) / totalInv : null)}）</span></div></div>`;
  // 台股每人的股數：定期定額累計＋單筆買進，依股票加總（例如 0050 1,113＋300＝1,413 股）
  const twShares = {};
  if (w) w.plans.concat(w.lumps).forEach(p => { if (p.name) twShares[p.name] = (twShares[p.name] || 0) + (p.shares || 0); });
  names.forEach(n => {
    const a = u ? u.kids.find(k => k.name === n) : null, b = w ? w.kids.find(k => k.name === n) : null;
    const hold = u && u.holdings[n] ? u.holdings[n] : [];
    h += `<div class="card"><h3>${esc(n)}</h3><p class="cap">美股＋台股定期定額</p>
      ${a ? `<div class="kv"><span>美股現值</span><span>${ntd(a.valNTD)}</span></div>
      <div class="kv"><span>美股報酬（台幣計）</span><span class="${cls(a.pctNTD)}">${pct(a.pctNTD)}</span></div>
      ${hold.map(x => `<div class="kv"><span>${esc(x.code)}</span><span>${qty(x.qty)} 股・均價 ${hidden ? MASK : '$' + dec(x.avg, 2)}</span></div>`).join('')}` : ''}
      ${b ? `${Object.keys(twShares).filter(k => twShares[k] > 0).map(k => `<div class="kv"><span>${esc(k)} 目前股數</span><span>${qty(twShares[k])} 股</span></div>`).join('')}
      <div class="kv"><span>台股現值</span><span>${ntd(b.value)}</span></div>
      <div class="kv"><span>台股報酬</span><span class="${cls(b.pct)}">${pct(b.pct)}</span></div>` : ''}</div>`;
  });
  h += cashSection('more/kids', '美股帳戶閒置現金');
  if (w && w.plans.length) {
    h += `<div class="sec">台股定期定額（每人）</div><div class="list">${w.plans.map(p => `<div class="row">
      <div class="main"><div class="t1">${esc(p.name)}</div><div class="t2" style="white-space:normal">每月 ${ntd(p.monthly)}・${esc(p.start)} 起・${p.months || 0} 個月・累計 ${qty(p.shares)} 股</div></div>
      <div class="r"><div class="t1">${ntd(p.value)}</div><div class="t2 ${cls(p.pct)}">${pct(p.pct)}</div></div></div>`).join('')}</div>`;
  }
  if (w && w.lumps.length) {
    h += `<div class="sec">單筆買進（每人）</div><div class="list">${w.lumps.map(p => `<div class="row">
      <div class="main"><div class="t1">${esc(p.name)}</div><div class="t2">${p.date}・${qty(p.shares)} 股 @ ${p.price != null ? dec(p.price, 1) : '—'}</div></div>
      <div class="r"><div class="t1">${ntd(p.value)}</div><div class="t2" style="color:var(--text-secondary)">投入 ${ntd(p.inv)}</div></div></div>`).join('')}</div>`;
  }
  return h;
}

function pageDca(d) {
  const m = d.myDca;
  if (!m) return '<div class="empty">找不到定期定額的資料</div>';
  const tot = m.rows.find(r => r.broker === 'Total');
  let h = head('我的定期定額', [['＋ 本月定期定額', 'dca']]);
  if (tot) h += `<div class="card hero"><div class="lbl">總現值</div><div class="big">${ntd(tot.valNTD)}</div>
    <div class="meta">投入約 ${ntd(tot.invNTD)}・${usd(tot.invUSD)}</div><div style="margin-top:10px"><span class="pill ${cls(tot.pct)}">${pct(tot.pct)}</span></div></div>`;
  const plans = m.plans.filter(p => p.monthly);
  h += `<div class="sec">扣款計畫</div><div class="list">${(plans.length ? plans : m.plans).map(p => `<div class="row">
    <div class="main"><div class="t1">${esc(p.name)}</div><div class="t2" style="white-space:normal">每月 ${p.monthly ? (p.currency === 'USD' ? usd(p.monthly) : ntd(p.monthly)) : '未扣款'}・${esc(p.start)} 起・${p.months || 0} 個月・累計 ${qty(p.shares)} 股</div></div>
    <div class="r"><div class="t1">${p.value ? (p.currency === 'USD' ? usd(p.value, 0) : ntd(p.value)) : '—'}</div><div class="t2 ${cls(p.pct)}">${p.pct != null ? pct(p.pct) : ''}</div></div></div>`).join('')}</div>`;
  return h;
}

function pagePlan(d) {
  const p = d.plan;
  if (!p) return '<div class="empty">找不到 Plan 的資料</div>';
  const now = d.total ? d.total.value / 10000 : null;
  const year = new Date().getFullYear();
  const ys = p.years.filter(y => y.year <= year + 5);
  const last = p.years[p.years.length - 1];
  // 每年實際出入金淨額（萬）：明細的入金是負數，所以加總後反過來；正數＝淨投入、負數＝淨出金
  const flowRows = (d.ledger && d.ledger.flow) || (d.total ? d.total.history.filter(x => x.flow) : []);
  const netIn = {};
  flowRows.forEach(x => { const t = toTime(x.date); if (x.flow && isFinite(t)) { const yy = new Date(t).getFullYear(); netIn[yy] = (netIn[yy] || 0) - x.flow / 10000; } });
  const flowOf = y => y.year in netIn ? netIn[y.year] : (y.actualInput != null ? y.actualInput : null);
  const signWan = v => (v > 0 ? '+' : v < 0 ? '−' : '') + dec(Math.abs(v), 1) + ' 萬';
  let h = `${head('長期計畫', [['編輯', 'plan']])}
  <div class="tiles">
    <div class="tile"><div class="k">假設年報酬</div><div class="v">${pctPlain(p.irr, 0)}</div><div class="n" style="font-size:14.5px;color:var(--text-secondary)">每年投入 ${p.input || 100} 萬</div></div>
    <div class="tile"><div class="k">${last ? last.year + '（' + last.age + ' 歲）' : '最終'}目標</div><div class="v">${last ? (hidden ? MASK : dec(last.target / 10000, 2) + ' 億') : '—'}</div><div class="n" style="font-size:14.5px;color:var(--text-secondary)">計畫終點</div></div>
  </div>
  <div class="card"><h3>目標 vs 實際</h3><p class="cap">單位：萬元；今年的「實際」是目前現值；出入金是當年入金減出金（淨出金時不畫柱子）</p>
    <div class="legend"><span><i style="background:var(--s1)"></i>實際</span><span><i style="background:var(--axis)"></i>計畫目標</span><span><i style="background:var(--s3)"></i>實際出入金</span></div>
    ${chart('c-plan', el => columnChart(el, ys.map(y => String(y.year)),
      [{ name: '實際', color: '--s1', values: ys.map(y => y.actual || (y.year === year ? now : null)) }, { name: '目標', color: '--axis', values: ys.map(y => y.target) },
       { name: '出入金', color: '--s3', values: ys.map(flowOf) }],
      { axis: v => dec(v, 0), topLabel: v => dec(v, 0), labelSeries: [0, 2], tip: i => { const y = ys[i]; const a = y.actual || (y.year === year ? now : null); const f = flowOf(y);
        return `<b>${y.year}（${y.age} 歲）</b>${tr('計畫目標', hidden ? MASK : dec(y.target, 0) + ' 萬')}${a ? tr(y.actual ? '實際' : '目前', hidden ? MASK : dec(a, 0) + ' 萬') + tr('達成率', pctPlain(a / y.target, 0)) : ''}` +
          (f != null ? tr('實際出入金' + (y.year === year ? '（至今）' : ''), hidden ? MASK : signWan(f)) : '') + (y.input ? tr('計畫投入', hidden ? MASK : dec(y.input, 0) + ' 萬') : ''); } }))}</div>
  <div class="sec">逐年目標</div><div class="list">${p.years.map(y => `<div class="row">
    <div class="main"><div class="t1" style="font-size:15px">${y.year}</div><div class="t2">${y.age} 歲</div></div>
    <div class="r"><div class="t1" style="font-size:15px">${hidden ? MASK : int(y.target) + ' 萬'}</div><div class="t2" style="font-size:15px;color:var(--text-primary)">${[y.actual ? '實際 ' + (hidden ? MASK : int(y.actual) + ' 萬') : '', flowOf(y) != null ? '出入金 ' + (hidden ? MASK : signWan(flowOf(y))) : ''].filter(Boolean).join(' · ')}</div></div></div>`).join('')}</div>`;
  return h;
}

function pageExpense(d) {
  const e = d.expense;
  if (!e) return '<div class="empty">找不到固定支出的資料</div>';
  const months = Array.from({ length: 12 }, (_, i) => e.months.find(m => m.month === i + 1) || { month: i + 1, total: 0, items: [] });
  const curM = new Date().getMonth() + 1;
  const next = [0, 1, 2].map(k => months[(curM - 1 + k) % 12]);
  const diff = Math.abs(e.monthsTotal - e.annual) > 1;
  let h = `${head('固定支出')}
  <div class="card hero"><div class="lbl">${esc(e.title || '年度總支出')}</div><div class="big">${ntd(e.annual)}</div>
    <div class="meta">平均每月 ${ntd(e.annual / 12)}${diff ? '・每月明細合計 ' + ntd(e.monthsTotal) : ''}</div></div>
  <div class="tiles">
    <div class="tile"><div class="k">本月（${curM} 月）</div><div class="v">${wan(months[curM - 1].total)}</div><div class="n">${months[curM - 1].items.length} 個項目</div></div>
    <div class="tile"><div class="k">接下來三個月</div><div class="v">${wan(next.reduce((s, m) => s + m.total, 0))}</div><div class="n">${next.map(m => m.month + '月').join('、')}</div></div>
  </div>
  <div class="card"><h3>每月固定支出</h3><p class="cap">點長條可看當月項目</p>
    ${chart('c-exp', el => columnChart(el, months.map(m => String(m.month)), [{ name: '支出', color: '--s2', values: months.map(m => m.total) }],
      { axis: v => (v ? dec(v / 10000, 0) + '萬' : '0'), maxBar: 20, tip: i => { const m = months[i];
        return `<b>${m.month} 月</b>${m.items.length ? m.items.map(it => tr(esc(it.label), ntd(it.amt))).join('') : tr('項目', '無')}${tr('合計', ntd(m.total))}`; } }))}</div>`;
  if (e.cats.length) {
    h += `<div class="card"><h3>年度分類</h3><p class="cap">${esc(e.title || '試算表的分類小計')}</p>
      ${chart('c-exp-cat', el => hbarChart(el, e.cats.slice().sort((a, b) => b.amt - a.amt).map(c => ({ label: c.label.replace(/\(.*\)/, ''), value: c.amt, color: '--s2',
        text: dec(c.amt / 10000, 1) + ' 萬', html: tr('金額', ntd(c.amt)) + tr('比重', pctPlain(c.amt / e.annual)) + (c.note ? tr('備註', esc(c.note)) : '') }))))}</div>`;
  }
  h += `<div class="sec">每月明細${ed() ? '<em>點月份可以編輯</em>' : ''}</div><div class="list">${months.map(m => `<div class="row${ed() ? ' link' : ''}"${ed() ? ` data-form="expense" data-arg="${m.month}"` : ''}>
    <div class="main"><div class="t1" style="font-size:15px">${m.month} 月${m.month === curM ? '（本月）' : ''}</div>
      <div class="t2" style="white-space:normal">${m.items.length ? m.items.map(it => esc(it.label) + ' ' + (hidden ? '••••' : dec(it.amt / 10000, 1) + '萬')).join('・') : '無'}</div></div>
    <div class="r"><div class="t1" style="font-size:15px">${ntd(m.total)}</div></div></div>`).join('')}</div>`;
  return h;
}

function pageCareer(d) {
  const c = d.career;
  if (!c) return `<h1 class="title">職涯與薪資</h1><div class="card"><p class="msg">還讀不到 BB 分頁：目前連到的 Apps Script 是第 ${scriptVersion() || '?'} 版，需要第 ${LATEST_SCRIPT} 版。請到 GitHub 複製最新的 Code.gs 貼到 Apps Script，再「管理部署作業 → 鉛筆 → 版本選新版本 → 部署」，然後回來往下拉更新。</p></div>`;
  const ys = c.years;
  if (!ys.length) return '<div class="empty">BB 分頁沒有年度資料</div>';
  const last = ys[ys.length - 1];
  const full = ys.filter(y => y.annualTotal || y.annual);
  const lf = full[full.length - 1], pf = full[full.length - 2];
  const inc = y => y.annualTotal || y.annual;
  const growth = lf && pf ? inc(lf) / inc(pf) - 1 : null;
  const first = ys[0];
  const firstSalary = first.salary || first.bSalary;
  const raised = ys.slice().reverse().find(y => y.raise);
  // 年前、年中分紅各取最近一次有發的年度
  const lastYe = ys.slice().reverse().find(y => y.yeTotal);
  const lastMid = ys.slice().reverse().find(y => y.bonus);
  let h = `${head('職涯與薪資')}
  ${ed() ? `<div class="acts wide"><button class="act" data-form="career" data-arg="pay">＋ 分紅</button><button class="act" data-form="career" data-arg="ot">＋ 加班費</button>
    <button class="act" data-form="career" data-arg="raise">＋ 調薪升遷</button></div>` : ''}
  <div class="card hero">
    <div class="lbl">${last.year} 年・${esc(last.title)}</div>
    <div class="big">${hidden ? MASK : 'NT$' + int(c.salary) + '<span style="font-size:17px;font-weight:600;color:var(--text-secondary)"> / 月</span>'}</div>
    <div class="meta">年資 ${last.tenure ?? '—'} 年・這次回任第 ${last.qTenure ?? '—'} 年・${last.age ?? '—'} 歲</div>
    <div style="margin-top:10px;display:flex;gap:8px;flex-wrap:wrap">
      ${firstSalary && c.salary ? `<span class="pill up">回任以來月薪 ${pct(c.salary / firstSalary - 1, 0)}</span>` : ''}
      ${raised ? `<span class="pill up" style="background:var(--tint-wash);color:var(--tint)">${raised.year} 調薪 ${hidden ? '' : '+' + int(raised.raise) + '・'}${pctPlain(raised.raisePct)}</span>` : ''}
    </div>
  </div>
  <div class="tiles t4">
    <div class="tile"><div class="k">${lastYe ? lastYe.year + ' ' : ''}年前分紅</div><div class="v">${lastYe ? wan(lastYe.yeTotal) : '—'}</div><div class="n">${lastYe && lastYe.yeTotalMonths ? dec(lastYe.yeTotalMonths, 1) + ' 個月・農曆年前' : '農曆年前'}</div></div>
    <div class="tile"><div class="k">${lastMid ? lastMid.year + ' ' : ''}年中分紅</div><div class="v">${lastMid ? wan(lastMid.bonus) : '—'}</div><div class="n">${lastMid && lastMid.bonusMonths ? dec(lastMid.bonusMonths, 1) + ' 個月・年中' : '年中'}</div></div>
    <div class="tile"><div class="k">${lf ? lf.year + ' 年收入' + (lf.est ? '（預估）' : '') : '年收入'}</div><div class="v">${lf ? wan(inc(lf)) : '—'}</div><div class="n">${lf && lf.ot ? '含加班 ' + wan(lf.ot) : ''}${lf && lf.est ? (lf.ot ? '・' : '') + '試算表還沒填年收入公式' : ''}</div></div>
    <div class="tile"><div class="k">年收入成長</div><div class="v ${cls(growth)}">${pct(growth)}</div><div class="n">${pf ? '比 ' + pf.year + ' 年' : ''}</div></div>
  </div>`;

  const fin = y => y.fiscalTotal || y.fiscal;
  h += `<div class="card"><h3>年收入</h3><p class="cap">單位：萬元，含加班；點長條看明細</p>
    <div class="legend"><span><i style="background:var(--s1)"></i>年收入：當年 1–2 月年前分紅＋當年年中分紅</span><span><i style="background:var(--s3)"></i>年度年收入：當年年中分紅＋隔年 1–2 月年前分紅</span></div>
    ${chart('c-income', el => columnChart(el, full.map(y => "'" + String(y.year).slice(2)), [{ name: '年收入', color: '--s1', values: full.map(y => inc(y) / 10000) },
      { name: '年度年收入', color: '--s3', values: full.map(y => (fin(y) ? fin(y) / 10000 : null)) }],
      { axis: v => dec(v, 0), topLabel: v => dec(v, 0), maxBar: 18, tip: i => { const y = full[i];
        const fy = fin(y) ? tr('年度年收入', ntd(fin(y)) + (y.fiscalMonths ? `（${dec(y.fiscalMonths, 1)} 個月）` : '')) : tr('年度年收入', '隔年年前分紅還沒發');
        if (y.est) return `<b>${y.year}・預估</b>${tr('年前分紅', ntd(y.yeTotal || 0))}${tr('年中分紅', ntd(y.bonus))}${tr('月薪', (hidden ? MASK : int(y.estFirst) + '×6＋' + int(y.estLast) + '×6'))}${y.ot ? tr('加班', ntd(y.ot)) : ''}${tr('合計', ntd(y.annualTotal))}${fy}`;
        return `<b>${y.year}・${esc(y.title)}</b>${tr('年收入', ntd(y.annual))}${y.ot ? tr('加班', ntd(y.ot)) : ''}${y.annualTotal ? tr('合計', ntd(y.annualTotal)) : ''}${y.yeTotal ? tr('年前分紅', ntd(y.yeTotal)) : ''}${y.bonus ? tr('年中分紅', ntd(y.bonus)) : ''}${y.months ? tr('相當', dec(y.months, 1) + ' 個月') : ''}${fy}`; } }))}</div>`;

  const sal = ys.filter(y => y.salary || y.bSalary);
  h += `<div class="card"><h3>月薪</h3><p class="cap">每年年初的月薪；點長條看當年調薪</p>
    ${chart('c-salary', el => columnChart(el, sal.map(y => "'" + String(y.year).slice(2)), [{ name: '月薪', color: '--s3', values: sal.map(y => (y.salary || y.bSalary) / 1000) }],
      { axis: v => dec(v, 0) + 'k', maxBar: 22, tip: i => { const y = sal[i];
        return `<b>${y.year}・${esc(y.title)}</b>${tr('月薪', ntd(y.salary || y.bSalary))}${y.raise ? tr('調薪', (hidden ? MASK : '+' + int(y.raise)) + '（' + pctPlain(y.raisePct) + '）') : ''}${y.events.filter(e => e.label).map(e => tr(esc(e.label), e.salary ? ntd(e.salary) : '')).join('')}`; } }))}</div>`;

  const bm = ys.filter(y => y.yeTotal || y.bonus);
  h += `<div class="card"><h3>分紅金額</h3><p class="cap">單位：萬元；點長條看各相當幾個月薪</p>
    <div class="legend"><span><i style="background:var(--s4)"></i>年前分紅</span><span><i style="background:var(--s5)"></i>年中分紅</span></div>
    ${chart('c-bonus', el => columnChart(el, bm.map(y => "'" + String(y.year).slice(2)), [
      { name: '年前分紅', color: '--s4', values: bm.map(y => y.yeTotal / 10000) }, { name: '年中分紅', color: '--s5', values: bm.map(y => y.bonus / 10000) }],
      { axis: v => dec(v, 0), maxBar: 14, tip: i => { const y = bm[i];
        return `<b>${y.year}</b>${y.yeTotal ? tr('年前分紅', ntd(y.yeTotal) + '（' + dec(y.yeTotalMonths, 1) + ' 個月）') : ''}${y.allowance ? tr('　職務加給', ntd(y.allowance)) : ''}${y.yeBonus ? tr('　年終', ntd(y.yeBonus)) : ''}${y.bonus ? tr('年中分紅', ntd(y.bonus) + '（' + dec(y.bonusMonths, 1) + ' 個月）') : ''}`; } }))}</div>`;

  // 職位與調薪的時間軸
  const tl = [];
  ys.forEach(y => {
    if (y === ys[0]) tl.push({ d: y.year, t: `回任・${y.title}` + (firstSalary ? `，月薪 ${hidden ? MASK : int(firstSalary)}` : '') });
    else if (y.promoted) tl.push({ d: y.year, t: `升遷為 ${y.title}` });
    y.events.forEach(e => tl.push({ d: y.year, arg: e.label && e.row ? e.row : null, t: (e.raise ? `年度調薪 ${hidden ? '' : '+' + int(e.raise) + '・'}${pctPlain(e.raisePct)}` : esc(e.label))
      + (e.salary ? `，月薪 ${hidden ? MASK : int(e.salary)}` : '') }));
  });
  h += `<div class="sec">職位與調薪${ed() ? '<em>點結構調薪等項目可以修改或刪除</em>' : ''}</div><div class="card"><ul class="tl">${tl.reverse().map(x => `<li${x.arg && ed() ? ` class="link" data-form="careerEvent" data-arg="${x.arg}"` : ''}><div class="d">${x.d}</div><div class="x">${x.t}</div></li>`).join('')}</ul></div>`;

  h += `<div class="sec">逐年明細</div><div class="list">${ys.slice().reverse().map(y => `<div class="row">
    <div class="main"><div class="t1" style="font-size:15px">${y.year}・${esc(y.title)}${y.promoted ? ' <span class="pill up" style="font-size:11px;padding:1px 6px">升遷</span>' : ''}</div>
      <div class="t2" style="white-space:normal">月薪 ${ntd(y.salary || y.bSalary)}${y.yeTotal ? '・年前分紅 ' + wan(y.yeTotal) : ''}${y.bonus ? '・年中分紅 ' + wan(y.bonus) : ''}${y.raise ? '・調薪 ' + pctPlain(y.raisePct) : ''}</div></div>
    <div class="r"><div class="t1" style="font-size:15px">${inc(y) ? wan(inc(y)) : '—'}</div><div class="t2" style="color:var(--text-secondary)">${y.age ? y.age + ' 歲' : ''}</div></div></div>`).join('')}</div>`;
  return h;
}

function pageHouse(d) {
  const list = d.house || [];
  if (!list.length) return '<div class="empty">找不到房產的資料</div>';
  const tb = list.reduce((s, x) => s + x.price, 0), te = list.reduce((s, x) => s + (x.est || 0), 0);
  const loanLeft = list.reduce((s, x) => s + (x.loan ? x.loan.balance : 0), 0);
  let h = `<h1 class="title">房產</h1>
  <div class="card hero"><div class="lbl">估計總值</div><div class="big">${hidden ? MASK : int(te) + ' 萬'}</div>
    <div class="meta">購入合計 ${hidden ? MASK : int(tb) + ' 萬'}${loanLeft ? `・房貸剩 ${wan(loanLeft)}・淨值 ${wan(te * 10000 - loanLeft)}` : ''}</div><div style="margin-top:10px"><span class="pill up">${hidden ? MASK : '+' + int(te - tb) + ' 萬'}（${pct((te - tb) / tb, 0)}）</span></div></div>
  <div class="card"><h3>購入價 vs 估計現值</h3><p class="cap">估值＝房屋坪數 × 估計單價 ＋ 估計車位價</p>
    <div class="legend"><span><i style="background:var(--axis)"></i>購入價</span><span><i style="background:var(--s6)"></i>估計現值</span></div>
    ${chart('c-house', el => columnChart(el, list.map(x => x.name), [{ name: '購入', color: '--axis', values: list.map(x => x.price) }, { name: '估值', color: '--s6', values: list.map(x => x.est) }],
      { axis: v => dec(v, 0), maxBar: 44, height: 200, tip: i => { const x = list[i]; return `<b>${esc(x.name)}</b>${tr('購入', hidden ? MASK : int(x.price) + ' 萬')}${tr('估計', hidden ? MASK : int(x.est) + ' 萬')}`; } }))}</div>`;
  list.forEach(x => {
    h += `<div class="card"><div class="hrow"><h3>${esc(x.name)}</h3>${ed() ? `<button class="act" data-form="house" data-arg="${esc(x.name)}">調整估價</button>` : ''}</div><p class="cap">${esc(x.date)} 購入</p>
      <div class="kv"><span>購入總價</span><span>${hidden ? MASK : int(x.price) + ' 萬'}</span></div>
      <div class="kv"><span>坪數（房屋）</span><span>${x.ping}（${x.housePing}）坪</span></div>
      <div class="kv"><span>車位</span><span>${x.parkN || 0} 個</span></div>
      <div class="kv"><span>購入單價／估計單價</span><span>${x.unit != null ? dec(x.unit, 1) : "—"} → ${x.estUnit != null ? dec(x.estUnit, 1) : "—"} 萬/坪</span></div>
      <div class="kv"><span>估計現值</span><span>${hidden ? MASK : int(x.est) + ' 萬'}</span></div></div>`;
    if (x.loan) h += loanCard(x);
  });
  return h;
}

function loanCard(x) {
  const l = x.loan, n = l.next;
  const y = Math.floor(l.leftN / 12), mo = l.leftN % 12;
  const equity = x.est != null ? x.est * 10000 - l.balance : null;
  const done = l.orig ? (l.orig - l.balance) / l.orig : null;
  // 總期數＝初貸日到到期日的月數
  const ym = d => { const [a, b] = d.split('/').map(Number); return a * 12 + b; };
  const totalN = l.start && l.end ? ym(l.end) - ym(l.start) : null;
  return `<div class="card"><div class="hrow"><h3>${esc(x.name)}・房貸</h3>${ed() ? `<button class="act" data-form="loan" data-arg="${esc(x.name)}">編輯</button>` : ''}</div><p class="cap">年利率 ${dec(l.rate * 100, 2)}%・每月 ${l.day} 號扣 ${ntd(l.pmt)}・${esc(l.method)}</p>
    <div class="lbl" style="margin-top:6px">剩餘房貸</div><div class="big" style="font-size:28px">${ntd(l.balance)}</div>
    <div class="meta">${l.paidN ? `基準日後已扣 ${l.paidN} 期，本金少了 ${ntd(l.bal - l.balance)}` : `基準日 ${esc(l.base)}`}</div>
    ${done != null ? `<div class="kv" style="margin-top:10px"><span>已還本金 ${ntd(l.orig - l.balance)}</span><span>${dec(done * 100, 1)}%</span></div>
    <div class="stack" style="height:10px;margin:6px 0 4px"><div style="width:${done * 100}%;background:var(--gold)"></div><div style="width:${(1 - done) * 100}%;background:var(--card-2)"></div></div>` : ''}
    ${chart('c-loan-' + x.loan.base.replace(/\//g, '') + x.loan.bal, el => loanChart(el, l))}
    ${n ? `<div class="kv"><span>下次扣款</span><span>${esc(n.date)}</span></div>
    <div class="kv"><span>這期利息／本金</span><span>${ntd(n.interest)}／${ntd(n.principal)}</span></div>` : ''}
    <div class="kv"><span>剩餘期數</span><span>${l.leftN} 期（${y} 年${mo ? ' ' + mo + ' 個月' : ''}）${totalN ? '／共 ' + totalN + ' 期' : ''}</span></div>
    ${l.orig ? `<div class="kv"><span>初貸金額</span><span>${ntd(l.orig)}${l.start ? '（' + esc(l.start) + '）' : ''}</span></div>` : ''}
    <div class="kv"><span>剩餘利息（預估）</span><span>${ntd(l.leftInterest)}</span></div>
    <div class="kv"><span>還清日</span><span>${esc(l.end || l.payoff)}</span></div>
    ${equity != null ? `<div class="kv"><span>房屋淨值（估值 − 房貸）</span><span>${wan(equity)}</span></div>` : ''}</div>`;
}

// 房貸餘額走勢：基準日到還清，標出今天
function loanChart(host, l) {
  host.innerHTML = '';
  const pts = [{ date: l.base, bal: l.bal }, ...l.sched];
  const W = host.clientWidth, H = 180, m = { t: 12, r: 50, b: 24, l: 6 };
  const svg = svgEl('svg', { viewBox: `0 0 ${W} ${H}`, height: H }, host);
  const x0 = toTime(pts[0].date), x1 = toTime(pts[pts.length - 1].date);
  const X = t => m.l + (t - x0) / Math.max(1, x1 - x0) * (W - m.l - m.r);
  const step = niceStep(l.bal, 4), yMax = Math.ceil(l.bal / step) * step;
  const Y = v => m.t + (1 - v / yMax) * (H - m.t - m.b);
  for (let v = 0; v <= yMax + 1; v += step) {
    svgEl('line', { x1: m.l, x2: W - m.r, y1: Y(v), y2: Y(v), stroke: css(v ? '--grid' : '--axis') }, svg);
    if (v) svgText(svg, W - m.r + 6, Y(v) + 4, hidden ? '' : dec(v / 10000, 0) + '萬');
  }
  svgText(svg, m.l, H - 6, pts[0].date.slice(0, 4));
  svgText(svg, W - m.r, H - 6, pts[pts.length - 1].date.slice(0, 4), { 'text-anchor': 'end' });
  let line = '';
  pts.forEach((p, i) => { line += (i ? 'L' : 'M') + X(toTime(p.date)).toFixed(1) + ',' + Y(p.bal).toFixed(1); });
  svgEl('path', { d: line + `L${X(x1)},${Y(0)}L${X(x0)},${Y(0)}Z`, fill: css('--area') }, svg);
  svgEl('path', { d: line, fill: 'none', stroke: css('--gold'), 'stroke-width': 2, 'stroke-linejoin': 'round' }, svg);
  const nx = X(Math.min(x1, Math.max(x0, Date.now())));
  svgEl('circle', { cx: nx, cy: Y(l.balance), r: 4.5, fill: css('--gold'), stroke: css('--card'), 'stroke-width': 2 }, svg);
  svgText(svg, nx + 7, Y(l.balance) - 7, '今天', { fill: css('--text-secondary') });
  const cross = svgEl('line', { y1: m.t, y2: H - m.b, stroke: css('--axis'), opacity: 0 }, svg);
  const dot = svgEl('circle', { r: 5, fill: css('--gold'), stroke: css('--card'), 'stroke-width': 2, opacity: 0 }, svg);
  const tip = makeTip(host);
  const hit = svgEl('rect', { class: 'hit', x: m.l, y: m.t, width: W - m.l - m.r, height: H - m.t - m.b, fill: 'transparent' }, svg);
  const move = ev => {
    const mx = pointerX(ev, svg);
    let best = pts[0];
    pts.forEach(p => { if (Math.abs(X(toTime(p.date)) - mx) < Math.abs(X(toTime(best.date)) - mx)) best = p; });
    const bx = X(toTime(best.date));
    cross.setAttribute('x1', bx); cross.setAttribute('x2', bx); cross.setAttribute('opacity', 1);
    dot.setAttribute('cx', bx); dot.setAttribute('cy', Y(best.bal)); dot.setAttribute('opacity', 1);
    tip.show(`<b>${best.date}</b>${tr('剩餘本金', ntd(best.bal))}${best.interest != null ? tr('當期利息', ntd(best.interest)) + tr('當期本金', ntd(best.principal)) : ''}`, bx, Y(best.bal));
  };
  const leave = () => { tip.hide(); cross.setAttribute('opacity', 0); dot.setAttribute('opacity', 0); };
  hit.addEventListener('pointermove', move);
  hit.addEventListener('pointerdown', move);
  hit.addEventListener('pointerleave', mouseOnly(leave));
  tip.onClose = leave;
}

// 里程碑的日期是自己打的文字（20261105、202611、202401~202403、2026/11/05…），取開頭的年月日當排序依據；看不懂的回傳 0
function milestoneKey(s) {
  const m = str(s).match(/^(\d{4})[\/\-.年]?(\d{1,2})?[\/\-.月]?(\d{1,2})?/);
  return m ? Number(m[1]) * 10000 + Number(m[2] || 0) * 100 + Number(m[3] || 0) : 0;
}
// 依日期由舊到新排（後來才補登的早期事件也會排到對的位置）；日期看不懂的跟著試算表裡的上一筆
function sortMilestones(list) {
  let last = 0;
  return list.map((m, i) => ({ m, i, k: (last = milestoneKey(m.date) || last) }))
    .sort((a, b) => a.k - b.k || a.i - b.i).map(x => x.m);
}

function pageMilestone(d) {
  // 有 Apps Script 的編輯資訊時用它（帶列號，點了可以編輯）；否則用一般讀到的資料
  const e = ed(), list = sortMilestones(e ? e.milestones : (d.milestone || []));
  return `${head('里程碑', [['＋ 新增', 'milestone']])}${e ? '<p class="hint" style="margin:-6px 0 10px">點項目可以修改或刪除</p>' : ''}
  <div class="card"><ul class="tl">${list.slice().reverse().map(m => `<li${e ? ` class="link" data-form="milestone" data-arg="${m.row}"` : ''}><div class="d">${esc(m.date)}</div><div class="x">${esc(m.text)}</div></li>`).join('')}</ul></div>`;
}

// 即時股價的狀態（Apps Script 第 23 版起）：上次抓的時間、美股的 Finnhub 金鑰有沒有設定
function liveLine() {
  const l = state.raw && state.raw._live;
  if (!l) return '';
  const t = l.at ? new Date(l.at) : null;
  const when = t ? `${t.getMonth() + 1}/${t.getDate()} ${pad(t.getHours())}:${pad(t.getMinutes())}` : '還沒抓過';
  return `<p class="msg" style="margin:4px 0 0;text-align:center;color:var(--text-muted)">即時股價：${when}・台股證交所、加密貨幣 Kraken・美股 ${l.finnhub ? 'Finnhub' : '未設定 Finnhub 金鑰（用 GOOGLEFINANCE）'}</p>`;
}
// 最近一次更新、寫入、更新股價各花的時間：Google 那邊各階段（Apps Script 回報）＋其餘（網路與 Apps Script 啟動）
function timingLine() {
  const sec = v => dec(v / 1000, 1) + ' 秒';
  return ['更新', '寫入', '更新股價'].filter(k => state.timing[k]).map(k => {
    const t = state.timing[k], m = t.ms;
    const parts = m ? [['抓即時價', m.live], ['寫入', m.write], ['讀分頁', m.ranges], ['讀表單資料', m.edit], ['讀漲跌', m.daily], ['記錄', m.log]]
      .filter(x => x[1] >= 50).map(x => x[0] + ' ' + sec(x[1])) : [];
    if (m) parts.push('網路與啟動 ' + sec(Math.max(0, t.total - m.total)));
    return `<p class="msg" style="margin:4px 0 0;text-align:center;color:var(--text-muted)">上次${k} ${sec(t.total)}${parts.length ? '（' + parts.join('、') + '）' : ''}</p>`;
  }).join('');
}

function pageSettings(first) {
  const c = state.cfg || {};
  return `<h1 class="title">${first ? '開始使用' : '設定'}</h1>
  ${first ? `<div class="card"><h3>第一次使用</h3><p class="cap">App 會透過你試算表裡的 Apps Script 讀取資料。資料只存在這台裝置上。</p>
    <ol class="steps"><li>打開試算表 → 擴充功能 → Apps Script</li><li>貼上 <a href="https://github.com/pv7999u7-prog/invest-dashboard-v2/blob/main/apps-script/Code.gs" target="_blank" rel="noopener" style="color:var(--tint)">apps-script/Code.gs</a> 的內容，把 KEY 改成自己的密碼</li>
    <li>部署 → 新增部署作業 → 網頁應用程式（執行身分：我；存取：所有人）</li><li>把網址和密碼填在下面</li></ol></div>` : ''}
  <div class="card">
    <label class="field"><span>Apps Script 網址</span><input id="cfgUrl" type="url" placeholder="https://script.google.com/macros/s/…/exec" value="${esc(c.url || '')}" autocomplete="off" autocapitalize="off"></label>
    <label class="field"><span>密碼</span><input id="cfgKey" type="password" value="${esc(c.key || '')}" autocomplete="off"></label>
    <button class="btn" id="cfgSave">儲存並更新</button>
    <p class="msg" id="cfgMsg" style="margin:10px 0 0">${state.status === 'err' ? '<span class="msg err">' + esc(state.error) + '</span>' : state.updated ? '上次更新：' + esc(state.updated) : ''}</p>
    ${first ? '' : '<button class="btn ghost" id="cfgClear">清除這台裝置上的資料與設定</button>'}
    <p class="msg" style="margin:12px 0 0;text-align:center;color:var(--text-muted)">App 版本 ${APP_VERSION}${state.raw ? `・Apps Script 第 ${scriptVersion()} 版${scriptVersion() < LATEST_SCRIPT ? `（最新是第 ${LATEST_SCRIPT} 版，請更新）` : '（最新）'}` : ''}</p>
    ${liveLine()}
    ${timingLine()}
  </div>
  ${first ? '' : securitySection(canBio)}`;
}

/* ================================================================
   路由與畫面
   ================================================================ */
function route() { return (location.hash || '#overview').slice(1); }
const PAGES = { overview: pageOverview, us: pageUs, tw: pageTw, crypto: pageCrypto, more: pageMore,
  'more/kids': pageKids, 'more/dca': pageDca, 'more/plan': pagePlan, 'more/expense': pageExpense,
  'more/house': pageHouse, 'more/milestone': pageMilestone, 'more/career': pageCareer, 'more/ledger': pageLedger, 'more/settings': () => pageSettings(false) };

function render() {
  const r = route();
  const tab = r.split('/')[0];
  charts.length = 0;
  ['#tabbar', '#eye', '#refresh'].forEach(s => { $(s).style.display = lock.locked ? 'none' : ''; });
  if (lock.locked) {
    $('#back').classList.remove('show');
    $('#view').innerHTML = pageLock();
    bindLock();
    return;
  }
  document.querySelectorAll('.tabbar button').forEach(b => b.classList.toggle('on', b.dataset.tab === tab));
  $('#back').classList.toggle('show', r.includes('/'));
  let html;
  if (!state.cfg || !state.cfg.url) html = pageSettings(true);
  else if (r === 'more/settings' || r === 'more') html = PAGES[r](state.data);
  else if (!state.data) html = `<div class="empty">${state.status === 'busy' ? '正在讀取試算表…' : state.status === 'err' ? esc(state.error) : '還沒有資料'}</div>`;
  else html = (PAGES[r] || pageOverview)(state.data);
  $('#view').innerHTML = html;
  drawCharts();
  bindPage();
  $('#eye').innerHTML = hidden
    ? '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M3 3l18 18"/><path d="M10.6 6.1A9.8 9.8 0 0 1 12 6c5 0 9 6 9 6a15.6 15.6 0 0 1-2.6 3.1M6.4 7.4C4.2 9 3 12 3 12s4 6 9 6a8.6 8.6 0 0 0 4.1-1"/></svg>'
    : '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M2.5 12S6.5 5.5 12 5.5 21.5 12 21.5 12 17.5 18.5 12 18.5 2.5 12 2.5 12z"/><circle cx="12" cy="12" r="3"/></svg>';
}

function bindPage() {
  document.querySelectorAll('[data-go]').forEach(el => el.addEventListener('click', () => { location.hash = el.dataset.go; }));
  const rs = $('#rangeSeg');
  if (rs) rs.addEventListener('click', e => { const b = e.target.closest('button'); if (!b) return; trendRange = b.dataset.r; render(); });
  const bs = $('#brokerSeg');
  if (bs) bs.addEventListener('click', e => { const b = e.target.closest('button'); if (!b) return; usBroker = b.dataset.b; render(); });
  const ts = $('#twSeg');
  if (ts) ts.addEventListener('click', e => { const b = e.target.closest('button'); if (!b) return; twGroup = b.dataset.b; render(); });
  const save = $('#cfgSave');
  if (save) save.addEventListener('click', async () => {
    const url = $('#cfgUrl').value.trim(), key = $('#cfgKey').value;
    const msg = $('#cfgMsg');
    if (!/^https:\/\/script\.google(usercontent)?\.com\//.test(url)) { msg.innerHTML = '<span class="msg err">網址要是 https://script.google.com/… 開頭</span>'; return; }
    if (!key) { msg.innerHTML = '<span class="msg err">請輸入密碼</span>'; return; }
    const keyChanged = !state.cfg || state.cfg.key !== key;
    state.cfg = { url, key };
    if (keyChanged) await rewrapPassword(key);
    await persist();
    msg.textContent = '讀取中…';
    await refresh();
    if (state.status === 'ok') { location.hash = 'overview'; render(); }
    else msg.innerHTML = '<span class="msg err">' + esc(state.error) + '</span>';
  });
  const clr = $('#cfgClear');
  if (clr) clr.addEventListener('click', () => {
    if (!confirm('要清除這台裝置上的設定和快取資料嗎？試算表本身不受影響。')) return;
    try { [LS.cfg, LS.cache, LOCK_KEY, ENC_KEY].forEach(k => localStorage.removeItem(k)); } catch (e) {}
    lock.rec = null; lock.dk = null; lock.locked = false;
    state = { cfg: null, raw: null, updated: '', data: null, status: 'idle', error: '', lastFetch: 0 };
    setStatus('idle'); location.hash = 'overview'; render();
  });
  bindSecurity();
  bindEdit();
}

/* ================================================================
   啟動
   ================================================================ */
document.querySelectorAll('.tabbar button').forEach(b => b.addEventListener('click', () => { location.hash = b.dataset.tab; window.scrollTo(0, 0); }));
$('#back').addEventListener('click', () => { location.hash = route().split('/')[0]; });
$('#refresh').addEventListener('click', () => refresh());

// 下拉更新：畫面在最頂端時往下拉超過門檻、放開就重新讀取試算表
(() => {
  const TRIGGER = 70, MAX = 120;
  const view = $('#view'), ptr = $('#ptr'), icon = ptr.querySelector('svg');
  const ARROW = '<path d="M12 5v14M6 13l6 6 6-6"/>';
  const SPIN = '<path d="M20 12a8 8 0 1 1-2.3-5.7"/><path d="M20 4v4h-4"/>';
  let startY = null, dist = 0, busy = false;
  const show = d => {
    view.style.transform = d ? `translateY(${d}px)` : '';
    ptr.style.opacity = Math.min(1, d / TRIGGER);
    ptr.style.transform = `translateY(${Math.max(0, d - 44)}px)`;
    if (!busy) icon.style.transform = d >= TRIGGER ? 'rotate(180deg)' : '';
  };
  const reset = () => { view.classList.remove('pulling'); show(0); };
  document.addEventListener('touchstart', e => {
    startY = null;
    if (busy || lock.locked || window.scrollY > 0 || e.touches.length !== 1) return;
    if (e.target.closest('.chart, input, textarea, .seg')) return;   // 圖表上拖曳是看數值，不觸發更新
    startY = e.touches[0].clientY; dist = 0;
  }, { passive: true });
  document.addEventListener('touchmove', e => {
    if (startY == null) return;
    const dy = e.touches[0].clientY - startY;
    if (dy <= 0 || window.scrollY > 0) { startY = null; reset(); return; }
    e.preventDefault();
    view.classList.add('pulling');
    dist = Math.min(MAX, dy * 0.5);
    show(dist);
  }, { passive: false });
  document.addEventListener('touchend', async () => {
    if (startY == null) return;
    startY = null;
    view.classList.remove('pulling');
    if (dist < TRIGGER) { reset(); return; }
    busy = true;
    icon.innerHTML = SPIN; icon.style.transform = ''; ptr.classList.add('spin');
    show(56);
    const t0 = Date.now();
    await refresh();
    await new Promise(r => setTimeout(r, Math.max(0, 600 - (Date.now() - t0))));   // 轉一下再收，避免一閃而過
    ptr.classList.remove('spin'); icon.innerHTML = ARROW;
    busy = false; dist = 0;
    reset();
  });
})();
$('#eye').addEventListener('click', () => { hidden = !hidden; try { localStorage.setItem(LS.hide, hidden ? '1' : '0'); } catch (e) {} render(); });
window.addEventListener('hashchange', () => { render(); window.scrollTo(0, 0); });
let rt;
window.addEventListener('resize', () => { clearTimeout(rt); rt = setTimeout(drawCharts, 150); });
matchMedia('(prefers-color-scheme: dark)').addEventListener('change', () => render());
document.addEventListener('visibilitychange', () => {
  if (document.visibilityState === 'hidden') { lock.hiddenAt = Date.now(); return; }
  if (lockEnabled() && !lock.locked && lock.hiddenAt && Date.now() - lock.hiddenAt > RELOCK_SEC * 1000) {
    relock();
    unlockWithBio();
    return;
  }
  if (Date.now() - state.lastFetch > AUTO_REFRESH_MIN * 60000) refresh();
});

let canBio = false;
bioAvailable().then(v => { canBio = v; if (route() === 'more/settings') render(); });

loadLocal();
setStatus('idle');
render();
if (lock.locked) unlockWithBio();   // 打開就直接跳 Face ID；被系統擋下時，畫面上還有按鈕可以按
else refresh();

// 有新版時自動換上：每次打開／回到前景都檢查一次，新版接手後重新載入一次
if ('serviceWorker' in navigator && location.protocol === 'https:') {
  const hadController = !!navigator.serviceWorker.controller;
  let reloaded = false;
  navigator.serviceWorker.addEventListener('controllerchange', () => {
    if (hadController && !reloaded) { reloaded = true; location.reload(); }
  });
  navigator.serviceWorker.register('sw.js', { updateViaCache: 'none' }).then(reg => {
    reg.update();
    document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible') reg.update(); });
  });
}
