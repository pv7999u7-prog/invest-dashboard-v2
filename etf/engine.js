// 0050 長期投資：每日建議的計算規則（純計算，不含任何個人資料；資料由呼叫的人傳進來）
// 規則來自 2026-11 的回測討論：
//   定期定額每月 5 號 2 萬永不停；每年單筆預算 30～50 萬，核心 60% 在 1 月投入，其餘是預備金；
//   預備金只在四種訊號出現時加碼，每次 1 份＝預備金 ÷ 3，兩次至少隔 10 個交易日；
//   只加碼不賣；預備金可滾到明年但最多一年份預算，超過的部分 12 月最後一個交易日投入。
// 「實際投入」用試算表的紀錄算，所以沒照建議買（或買得比建議多）時，隔天的建議會跟著調整。
'use strict';

const ETF_RULES = {
  coreRatio: 0.6, units: 3, gapDays: 10, monthly: 20000, dcaDay: 5,
  budgetMin: 300000, budgetMax: 500000, budgetDefault: 400000
};

// series: [{d:'YYYY-MM-DD', c:收盤}]，日期由舊到新
function etfSma(a, n, i) { if (i < n - 1) return null; let s = 0; for (let k = i - n + 1; k <= i; k++) s += a[k]; return s / n; }
// 美股的資料：取「日期早於台股今天」的最後幾筆（美股收盤在台灣的晚上）
function etfBefore(series, day) { return (series || []).filter(x => x.d < day); }

/**
 * in = {
 *   today: 'YYYY-MM-DD'（台北日期）,
 *   holidays: ['YYYY-MM-DD'] 台股休市日（平日但不開盤；可省略）,
 *   p0050: [{d, c(還原權息), raw(收盤)}] 至少 260 筆, vix: [{d,c}], sox: [{d,c}], tnx: [{d,c}],
 *   ndc: {score, month} 或 null（國發會景氣對策信號分數）,
 *   lumps: [{d, amount}] 今年與去年的 0050 單筆買進（台幣，正數）,
 *   budgets: {年: 預算}（沒填用 40 萬）, start: 開始照這套規則記帳的年份（之前的年份不算滾存）
 * }
 */
function etfAdvise(inp) {
  const R = ETF_RULES, today = inp.today, year = Number(today.slice(0, 4));
  const P = (inp.p0050 || []).filter(x => x.d <= today);
  if (P.length < 250) throw new Error('0050 價格資料不足');
  const c = P.map(x => x.c), i = c.length - 1, last = P[i];
  const hi250 = Math.max(...c.slice(Math.max(0, i - 249)));
  const dd = c[i] / hi250 - 1, ma240 = etfSma(c, 240, i), bias = ma240 ? c[i] / ma240 - 1 : null;
  const ma240prev = etfSma(c, 240, i - 20), upTrend = ma240 && ma240prev && c[i] > ma240 && ma240 > ma240prev;
  // 價格換算：還原價和實際收盤的比例（除息後兩者不同）
  const ratio = last.raw ? last.raw / last.c : 1;
  const priceAt = d => Math.round(hi250 * (1 + d) * ratio * 100) / 100;

  const V = etfBefore(inp.vix, today), S = etfBefore(inp.sox, today), T = etfBefore(inp.tnx, today);
  const vix = V.length ? V[V.length - 1].c : null, vixMax10 = V.length ? Math.max(...V.slice(-10).map(x => x.c)) : null;
  const soxChg = S.length > 1 ? S[S.length - 1].c / S[S.length - 2].c - 1 : null;
  const tnx = T.length ? T[T.length - 1].c : null;
  let tnx90 = null;
  if (T.length) {
    const t0 = new Date(T[T.length - 1].d + 'T00:00:00Z'); t0.setUTCDate(t0.getUTCDate() - 90);
    const ref = T.filter(x => x.d <= t0.toISOString().slice(0, 10)).pop();
    if (ref) tnx90 = tnx - ref.c;   // ^TNX 是殖利率 %（例如 4.92）
  }
  const rateShock = tnx90 != null && tnx90 > 0.75;
  const ndc = inp.ndc && inp.ndc.score != null ? Number(inp.ndc.score) : null;

  const pct = x => (x < 0 ? '−' : '') + Math.abs(x * 100).toFixed(1) + '%';
  const signals = [
    { id: 'ndc', name: '景氣谷底＋深跌', on: ndc != null && ndc <= 22 && dd <= -0.20,
      need: '景氣燈號分數 ≤ 22 且跌幅 ≤ −20%', now: `景氣分數 ${ndc == null ? '未知' : ndc}、跌幅 ${pct(dd)}` },
    { id: 'vix', name: 'VIX 恐慌退燒', on: vixMax10 != null && vixMax10 >= 30 && vix < vixMax10 * 0.85 && dd <= -0.08,
      need: '近 10 日 VIX 最高 ≥ 30、今天低於最高 × 0.85、跌幅 ≤ −8%', now: `VIX ${vix == null ? '—' : vix.toFixed(1)}（近 10 日最高 ${vixMax10 == null ? '—' : vixMax10.toFixed(1)}）` },
    { id: 'sox', name: '費半夜殺', on: soxChg != null && soxChg < -0.05 && dd <= -0.10,
      need: '前一晚費半跌超過 5% 且跌幅 ≤ −10%', now: `費半前一晚 ${soxChg == null ? '—' : (soxChg >= 0 ? '+' : '') + pct(soxChg)}` },
    { id: 'dd15', name: '回檔 15%', on: dd <= -0.15 && !rateShock,
      need: '跌幅 ≤ −15% 且不是升息衝擊', now: `跌幅 ${pct(dd)}${rateShock ? '（升息衝擊中，暫停）' : ''}` }
  ];
  const fired = signals.filter(s => s.on);

  // ---- 預算與實際投入 ----
  const budgetOf = y => { const b = Number((inp.budgets || {})[y]); return b > 0 ? b : R.budgetDefault; };
  const lumps = (inp.lumps || []).filter(x => x.amount > 0);
  const sumYear = y => lumps.filter(x => x.d.slice(0, 4) === String(y)).reduce((t, x) => t + x.amount, 0);
  const B = budgetOf(year), prevB = budgetOf(year - 1);
  // 去年沒用完的預備金滾進今年，最多一年份預算（去年超過上限的部分，照規則應該在 12 月最後一天投掉）
  const start = Number(inp.start) || year;
  const carry = year - 1 >= start ? Math.max(0, Math.min(prevB, prevB - sumYear(year - 1))) : 0;
  const core = Math.round(B * R.coreRatio), reserve = B - core + carry;
  const used = sumYear(year), coreLeft = Math.max(0, core - used);
  const remaining = Math.max(0, B + carry - used);
  const reserveLeft = Math.max(0, remaining - coreLeft);
  const unit = Math.round(reserve / R.units / 1000) * 1000;

  // 交易日：平日且不是休市日
  const hol = new Set(inp.holidays || []);
  const isTradeDay = d => { const w = new Date(d + 'T00:00:00Z').getUTCDay(); return w > 0 && w < 6 && !hol.has(d); };
  const monthDays = (y, mo) => { const out = []; for (let k = 1; k <= 31; k++) { const d = `${y}-${String(mo).padStart(2, '0')}-${String(k).padStart(2, '0')}`; if (new Date(d + 'T00:00:00Z').getUTCMonth() + 1 === mo && isTradeDay(d)) out.push(d); } return out; };
  const isTrade = isTradeDay(today);
  const lastLump = lumps.map(x => x.d).filter(d => d <= today).sort().pop() || null;
  let gap = null;
  if (lastLump) {
    const hist = P.map(x => x.d).concat(isTrade && P[P.length - 1].d !== today ? [today] : []);
    gap = hist.filter(d => d > lastLump && d <= today).length;   // 上次買進之後經過幾個交易日
  }
  const cooling = gap != null && gap < R.gapDays;
  const ymd = today.slice(5);
  const lastTradeOfYear = monthDays(year, 12).pop();
  const isJan = ymd < '02-01';
  // 定期定額扣款日：每月 5 號，遇假日順延到下一個開盤日
  const dcaDate = monthDays(year, Number(today.slice(5, 7))).find(d => Number(d.slice(8)) >= R.dcaDay) || null;
  const dcaToday = isTrade && dcaDate === today;

  // ---- 決定今天的建議 ----
  const round = x => Math.round(x / 1000) * 1000;
  let level = 'hold', title = '照常定期定額，不用動作', amount = 0;
  const why = [];
  if (!isTrade) {
    level = 'closed'; title = '今天台股休市';
    why.push('休市日不用動作。');
  } else if (isJan && coreLeft > 0) {
    level = 'core'; amount = round(coreLeft);
    title = `一月核心投入：建議單筆買進約 ${(amount / 10000).toFixed(1)} 萬`;
    why.push(`今年單筆預算 ${(B / 10000).toFixed(0)} 萬的 60% 是核心部位（${(core / 10000).toFixed(1)} 萬），1 月直接投入，不等回檔。` + (used ? `今年已投入 ${(used / 10000).toFixed(1)} 萬，還差 ${(coreLeft / 10000).toFixed(1)} 萬。` : ''));
  } else if (fired.length && reserveLeft > 0 && !cooling) {
    level = 'buy'; amount = round(Math.min(unit, reserveLeft));
    title = `加碼訊號：建議單筆買進約 ${(amount / 10000).toFixed(1)} 萬`;
    why.push(`成立的訊號：${fired.map(s => s.name).join('、')}。投入 1 份預備金（預備金 ÷ 3）。`);
  } else if (fired.length && reserveLeft > 0 && cooling) {
    level = 'wait'; title = `訊號成立，但距離上次加碼才 ${gap} 個交易日`;
    why.push(`兩次加碼至少隔 ${R.gapDays} 個交易日，再等 ${R.gapDays - gap} 個交易日；到時訊號還成立就投入約 ${(round(Math.min(unit, reserveLeft)) / 10000).toFixed(1)} 萬。`);
  } else if (fired.length && reserveLeft <= 0) {
    level = 'full'; title = '訊號成立，但今年單筆預算已用完';
    why.push(`今年已投入 ${(used / 10000).toFixed(1)} 萬，超過或等於預算 ${((B + carry) / 10000).toFixed(1)} 萬。只維持定期定額；有額外資金才考慮，屬於提前動用明年預算。`);
  } else if (today === lastTradeOfYear && remaining > B) {
    level = 'yearend'; amount = round(remaining - B);
    title = `年底投入：建議買進約 ${(amount / 10000).toFixed(1)} 萬`;
    why.push('預備金最多保留一年份預算，超過的部分在 12 月最後一個交易日投入。');
  } else {
    if (dd <= -0.15 && rateShock) { level = 'pause'; title = '升息衝擊中：回檔 15% 訊號暫停'; why.push(`10 年期美債 90 天上升 ${tnx90.toFixed(2)} 個百分點（> 0.75），先暫停一般回檔加碼，只看 VIX 與費半兩個恐慌訊號。`); }
    else why.push('四個加碼訊號都沒有成立。長期來看，等待一般的小回檔反而少賺，所以沒訊號就只做定期定額。');
  }
  if (bias != null && bias > 0.30) why.push(`乖離年線 ${pct(bias)}，市場偏熱；照規則不賣，定期定額照常。`);
  if (dcaToday) why.push(`今天是定期定額扣款日（每月 ${R.dcaDay} 號，假日順延），記得 ${(R.monthly / 10000).toFixed(0)} 萬照常扣款。`);

  // 距離各訊號門檻
  const levels = [-0.08, -0.10, -0.15, -0.20, -0.25, -0.30].map(x => ({ dd: x, price: priceAt(x) }));
  return {
    date: today, trading: isTrade, level, title, amount, why,
    signals: signals.map(s => ({ id: s.id, name: s.name, on: s.on, need: s.need, now: s.now })),
    ind: { close: last.raw || last.c, closeDate: last.d, hi250: priceAt(0), dd, ma240: ma240 ? Math.round(ma240 * ratio * 100) / 100 : null, bias, upTrend: !!upTrend,
      vix, vixMax10, soxChg, tnx, tnx90, rateShock, ndc, ndcMonth: inp.ndc && inp.ndc.month || '' },
    budget: { year, budget: B, carry, core, reserve, unit, used, coreLeft, remaining, reserveLeft, lastLump, gap, cooling,
      unitRange: [R.budgetMin, R.budgetMax].map(b => Math.round((b - Math.round(b * R.coreRatio) + carry) / R.units / 1000) * 1000) },
    levels, dcaToday, dcaDate
  };
}

if (typeof module !== 'undefined') module.exports = { etfAdvise, ETF_RULES };
