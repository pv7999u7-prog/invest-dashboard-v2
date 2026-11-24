// 0050 長期投資：「更多 → 0050 長期投資」頁、通知點進來的彈出視窗、通知開關。
// 資料來自 Apps Script（每個開盤日 8:25～8:35 算一次，存在試算表「0050 建議」分頁）；
// Claude 的雲端排程每天早上把新聞解讀寫在同一列。只有打開這一頁或彈出視窗時才去讀。
'use strict';

const ETF_SCRIPT = 36;               // Apps Script 第幾版起支援
const ETF_STALE_MS = 5 * 60000;
const etfS = { busy: false, at: 0, error: '', pop: false, view: null, sub: null, subChecked: false };
const ETF_TONE = { core: 'up', buy: 'up', yearend: 'up', wait: 'warn', full: 'warn', pause: 'warn', hold: 'mute', closed: 'mute' };
const ETF_NAME = { core: '核心投入', buy: '加碼', wait: '等待間隔', full: '預算用完', yearend: '年底投入', pause: '暫停加碼', hold: '照常', closed: '休市' };

const etfWan = v => (v == null ? '—' : hidden ? MASK : (Math.round(v / 1000) / 10).toLocaleString('en-US', { maximumFractionDigits: 1 }) + ' 萬');
const etfPct = v => (v == null ? '—' : (v > 0 ? '+' : v < 0 ? '−' : '') + Math.abs(v * 100).toFixed(1) + '%');
const etfNum = (v, d = 2) => (v == null ? '—' : Number(v).toFixed(d));

async function etfLoad(action, payload) {
  if (!state.cfg || !state.cfg.url) return;
  if (!action && etfS.busy) return;
  etfS.busy = true; etfS.error = '';
  let msg = '';
  try {
    const j = await post(action ? { etfAction: action, payload: payload || {} } : { etf: true }, '0050');
    if (!j.ok) throw new Error(j.error === 'unauthorized' ? '密碼不對' : (j.error || '讀取失敗'));
    if (state.raw) { state.raw._etf = j.etf; await persist(); }
    msg = j.message || '';
  } catch (e) {
    etfS.error = e.message === 'Failed to fetch' || e.name === 'TypeError' ? '連不上，請檢查網路' : e.message;
    if (action) { etfS.busy = false; etfS.at = Date.now(); throw new Error(etfS.error); }
  }
  etfS.busy = false; etfS.at = Date.now();
  if (route() === 'more/etf') render();
  return msg;
}

// 最新一筆（或正在看的那一天）
function etfRow(E, date) {
  const rows = (E && E.rows || []).filter(r => r.adv);
  if (!rows.length) return null;
  return (date && rows.find(r => r.date === date)) || rows[rows.length - 1];
}

function etfSummary(r, compact) {
  const a = r.adv, tone = ETF_TONE[a.level] || 'mute';
  const news = r.news ? `<div class="grp" style="margin-top:14px">Claude 今日解讀<em>${esc(r.newsAt || '')}</em></div><p class="ptext">${esc(r.news)}</p>` : '';
  return `<div class="card etf-hero ${tone}"><div class="etf-top"><span class="pill ${tone}">${ETF_NAME[a.level] || esc(r.level)}</span><span class="etf-date">${esc(a.date)}${a.at ? '・' + esc(a.at.slice(11)) + ' 分析' : ''}</span></div>
    <div class="etf-title">${esc(a.title)}</div>
    ${a.amount ? `<div class="etf-amt">${etfWan(a.amount)}</div>` : ''}
    <ul class="ben">${(a.why || []).map(w => `<li>${esc(w)}</li>`).join('')}</ul>${news}
    ${compact ? '' : ''}</div>`;
}
function etfSignals(a) {
  return `<div class="card"><h3>加碼訊號<em class="etf-sub">任一個成立就投入 1 份預備金</em></h3>${a.signals.map(s => `<div class="etf-sig ${s.on ? 'on' : ''}">
    <span class="dot">${s.on ? '●' : '○'}</span><div><b>${esc(s.name)}</b><small>${esc(s.now)}</small><small class="need">條件：${esc(s.need)}</small></div></div>`).join('')}</div>`;
}

function pageEtf() {
  let h = head('0050 長期投資');
  if (scriptVersion() < ETF_SCRIPT) return h + `<div class="card"><p class="cap" style="margin:0">這一頁需要 Apps Script 第 ${ETF_SCRIPT} 版，請先更新 Apps Script。</p></div>`;
  const E = state.raw && state.raw._etf;
  if ((!E || Date.now() - etfS.at > ETF_STALE_MS) && !etfS.busy && !(etfS.error && Date.now() - etfS.at < 15000)) setTimeout(() => etfLoad(), 0);
  const err = etfS.error ? `<p class="msg err" style="margin:0 4px 10px">${esc(etfS.error)}</p>` : '';
  if (!E) return h + err + `<div class="empty">${etfS.error ? '讀不到資料' : '正在讀取…'}</div>`;
  const r = etfRow(E, etfS.view), latest = etfRow(E);
  if (!r) h += err + `<div class="card"><h3>還沒有建議</h3><p class="cap" style="margin:0">每個台股開盤日早上 8:25～8:35 會自動分析一次。也可以按下面的「現在分析一次」。</p></div>`;
  else {
    const a = r.adv, b = a.budget, ind = a.ind, old = latest && r.date !== latest.date;
    h += err + (old ? `<button class="act" data-etf-view="" style="margin:0 0 12px">‹ 回到最新（${esc(latest.date)}）</button>` : '');
    h += etfSummary(r) + etfSignals(a);
    // 預算與實際投入
    h += `<div class="card"><h3>${b.year} 年單筆資金<em class="etf-sub">從試算表台股交易表的 0050 買進自動計算</em></h3>
      <div class="kv pol link" data-etf-cfg="預算 ${b.year}"><span>今年單筆預算</span><span class="edv">${etfWan(b.budget)} <i>✎</i></span></div>
      <div class="kv pol"><span>核心 60%（1 月投入）</span><span>${etfWan(b.core)}</span></div>
      <div class="kv pol"><span>預備金${b.carry ? `（含去年滾存 ${etfWan(b.carry)}）` : ''}</span><span>${etfWan(b.reserve)}</span></div>
      <div class="kv pol"><span>每次加碼 1 份</span><span>${etfWan(b.unit)}<small class="etf-note">預算 30 萬時 ${etfWan(b.unitRange[0])}、50 萬時 ${etfWan(b.unitRange[1])}</small></span></div>
      <div class="kv pol tot"><span>今年已投入</span><span>${etfWan(b.used)}</span></div>
      <div class="kv pol tot"><span>今年還能投入</span><span>${etfWan(b.remaining)}</span></div>
      <div class="kv pol"><span>上次單筆買進</span><span>${b.lastLump ? esc(b.lastLump.replace(/-/g, '/')) + (b.gap != null ? `（${b.gap} 個交易日前）` : '') : '—'}</span></div>
      <p class="cap" style="margin:10px 0 0">每月 5 號的定期定額 2 萬另外算，不佔這個預算。你沒照建議買、或買得比較多，隔天的建議會照實際紀錄重算。</p></div>`;
    // 市場指標
    h += `<div class="card"><h3>市場指標<em class="etf-sub">0050 收盤日 ${esc(ind.closeDate || '')}</em></h3>
      <div class="kv pol"><span>0050 收盤</span><span>${etfNum(ind.close)}</span></div>
      <div class="kv pol"><span>近一年最高</span><span>${etfNum(ind.hi250)}</span></div>
      <div class="kv pol"><span>距離高點</span><span>${etfPct(ind.dd)}</span></div>
      <div class="kv pol"><span>年線（240 日均線）</span><span>${etfNum(ind.ma240)}${ind.upTrend ? '・向上' : ''}</span></div>
      <div class="kv pol"><span>乖離年線</span><span>${etfPct(ind.bias)}</span></div>
      <div class="kv pol"><span>VIX（近 10 日最高）</span><span>${etfNum(ind.vix, 1)}（${etfNum(ind.vixMax10, 1)}）</span></div>
      <div class="kv pol"><span>費半前一晚</span><span>${etfPct(ind.soxChg)}</span></div>
      <div class="kv pol"><span>10 年期美債（90 天變化）</span><span>${etfNum(ind.tnx)}%（${ind.tnx90 == null ? '—' : (ind.tnx90 >= 0 ? '+' : '') + ind.tnx90.toFixed(2)}）${ind.rateShock ? '・升息衝擊' : ''}</span></div>
      <div class="kv pol link" data-etf-cfg="景氣燈號分數"><span>景氣燈號分數（${esc(ind.ndcMonth || '月份未知')}）</span><span class="edv">${ind.ndc == null ? '—' : ind.ndc} <i>✎</i></span></div></div>`;
    h += `<div class="card"><h3>參考價位<em class="etf-sub">以近一年最高 ${etfNum(ind.hi250)} 計算</em></h3><div class="rtw"><table class="rt"><thead><tr><th>距離高點</th>${a.levels.map(x => `<th>${Math.round(x.dd * 100)}%</th>`).join('')}</tr></thead>
      <tbody><tr><td>0050 價位</td>${a.levels.map(x => `<td>${x.price.toFixed(1)}</td>`).join('')}</tr></tbody></table></div></div>`;
  }
  // 歷史
  const hist = (E.rows || []).filter(x => x.adv).slice().reverse();
  if (hist.length > 1) h += `<div class="sec">每日建議紀錄<em>最近 ${hist.length} 個開盤日</em></div><div class="list">${hist.map(x => `<div class="row link" data-etf-view="${esc(x.date)}">
    <div class="main"><div class="t1" style="font-size:15px">${esc(x.date.slice(5).replace('-', '/'))}　<span class="pill ${ETF_TONE[x.adv.level] || 'mute'}">${ETF_NAME[x.adv.level] || esc(x.level)}</span></div>
    <div class="t2 wrap">${esc(x.title)}</div></div>${x.amount ? `<div class="r"><div class="t1" style="font-size:15px">${etfWan(x.amount)}</div></div>` : ''}<span class="chev">›</span></div>`).join('')}</div>`;
  // 通知與重新分析
  const subs = E.subs || [];
  h += `<div class="card"><h3>每日通知</h3><p class="cap">每個台股開盤日 8:25～8:35 推播今天的建議，點通知會跳出這份建議。iPhone 要從主畫面的 App 開啟才能收通知（iOS 16.4 以上）。</p>
    <div class="kv pol"><span>這台裝置</span><span id="etfSubState">${etfS.subChecked ? (etfS.sub ? '已開啟' : '未開啟') : '檢查中…'}</span></div>
    <div class="kv pol"><span>已開啟的裝置</span><span>${subs.length ? subs.map(s => esc(s.device || '裝置') + (s.at ? `（${esc(s.at.slice(0, 10))}）` : '')).join('<br>') : '沒有'}</span></div>
    ${E.push ? `<div class="kv pol"><span>上次推播</span><span>${esc(E.push)}</span></div>` : ''}
    ${E.error ? `<p class="msg err" style="margin:8px 0 0">上次自動分析失敗：${esc(E.error)}</p>` : ''}
    <div class="acts wide" style="margin:12px 0 0">${etfS.sub ? '<button class="act" id="etfTest">傳送測試通知</button><button class="act" id="etfOff">關閉這台的通知</button>' : '<button class="act" id="etfOn">開啟每日通知</button>'}<button class="act" id="etfRun">現在分析一次</button></div></div>`;
  h += `<p class="hint" style="margin:14px 4px 8px">規則：每月 5 號定期定額 2 萬；每年單筆預算 30～50 萬，60% 在 1 月投入，40% 當預備金，等四種加碼訊號才投入、每次 1/3、至少隔 10 個交易日；只加碼不賣。這是依回測整理的規則，不是投資建議。${etfS.busy ? '・更新中…' : ''}</p>`;
  return h;
}

/* ---------- 彈出視窗（點通知進來時） ---------- */
function etfPopup() {
  const E = state.raw && state.raw._etf, r = etfRow(E);
  const old = $('#etfPop'); if (old) old.remove();
  const o = document.createElement('div');
  o.className = 'overlay'; o.id = 'etfPop';
  const today = new Date(Date.now() + 8 * 3600e3).toISOString().slice(0, 10);
  const body = r ? etfSummary(r, true) + (r.date !== today ? `<p class="hint" style="margin:0 0 10px">${etfS.busy ? '正在讀取今天的建議…' : `今天的建議還沒產生，這是 ${esc(r.date)} 的。`}</p>` : '')
    : `<div class="empty">${etfS.busy ? '正在讀取今天的建議…' : esc(etfS.error || '還沒有建議')}</div>`;
  o.innerHTML = `<div class="sheet"><div class="sh-head"><h2>0050 今日建議</h2><button type="button" class="x" aria-label="關閉">✕</button></div>${body}
    <button class="btn" type="button" id="etfPopMore">查看完整分析</button></div>`;
  document.body.appendChild(o); document.body.classList.add('noscroll');
  const close = () => { o.remove(); document.body.classList.remove('noscroll'); };
  o.addEventListener('click', e => { if (e.target === o) close(); });
  o.querySelector('.x').addEventListener('click', close);
  o.querySelector('#etfPopMore').addEventListener('click', () => { close(); if (route() !== 'more/etf') location.hash = '#more/etf'; else window.scrollTo(0, 0); });
}
async function etfOpenPopup() {
  etfS.pop = false;
  if (lock.locked) { etfS.pop = true; return; }   // 解鎖後再跳
  etfPopup();
  if (scriptVersion() >= ETF_SCRIPT) { await etfLoad(); if ($('#etfPop')) etfPopup(); }
}

/* ---------- 通知訂閱 ---------- */
const etfKey = s => { const b = atob(s.replace(/-/g, '+').replace(/_/g, '/') + '='.repeat((4 - s.length % 4) % 4)); return Uint8Array.from(b, c => c.charCodeAt(0)); };
async function etfCheckSub() {
  try {
    if (!('serviceWorker' in navigator) || !('PushManager' in window)) { etfS.sub = null; }
    else { const reg = await navigator.serviceWorker.getRegistration(); etfS.sub = reg ? await reg.pushManager.getSubscription() : null; }
  } catch (e) { etfS.sub = null; }
  etfS.subChecked = true;
  const el = $('#etfSubState'); if (el && route() === 'more/etf') render();
}
async function etfSubscribe() {
  if (!('Notification' in window) || !('PushManager' in window) || !('serviceWorker' in navigator)) {
    toast(/iPhone|iPad/.test(navigator.userAgent) && !navigator.standalone ? '請從主畫面的 App 開啟這一頁再開通知（Safari 分頁不能收通知）' : '這個瀏覽器不支援推播通知'); return;
  }
  const perm = await Notification.requestPermission();   // 要在按鈕事件裡第一個呼叫，iOS 才會跳詢問
  if (perm !== 'granted') { toast('沒有允許通知；可以到 iPhone 設定 → 通知 → HK 打開'); return; }
  const E = state.raw && state.raw._etf;
  if (!E || !E.vapid) { toast('還沒讀到通知金鑰，請先下拉更新'); return; }
  try {
    const reg = await navigator.serviceWorker.ready;
    let sub = await reg.pushManager.getSubscription();
    if (!sub) sub = await reg.pushManager.subscribe({ userVisibleOnly: true, applicationServerKey: etfKey(E.vapid) });
    const device = (/iPhone/.test(navigator.userAgent) ? 'iPhone' : /iPad/.test(navigator.userAgent) ? 'iPad' : /Android/.test(navigator.userAgent) ? 'Android' : '電腦') + ' ' + new Date().toLocaleDateString('zh-TW');
    const msg = await etfLoad('sub', { sub: sub.toJSON(), device });
    etfS.sub = sub; toast('✓ ' + msg); render();
  } catch (e) { toast('開啟失敗：' + e.message); }
}
async function etfUnsubscribe() {
  try {
    const reg = await navigator.serviceWorker.ready, sub = await reg.pushManager.getSubscription();
    if (sub) { await etfLoad('unsub', { endpoint: sub.endpoint }); await sub.unsubscribe(); }
    etfS.sub = null; toast('✓ 已關閉這台裝置的通知'); render();
  } catch (e) { toast('關閉失敗：' + e.message); }
}

/* ---------- 設定（預算、景氣燈號） ---------- */
function formEtfCfg(key) {
  const E = state.raw && state.raw._etf, cfg = (E && E.cfg) || {};
  const isBudget = /^預算/.test(key), cur = cfg[key] || (isBudget ? '400000' : '');
  const body = isBudget
    ? fNum('v', key.replace('預算', '') + ' 年單筆預算（萬）', cur ? Number(cur) / 10000 : '', '30～50', '不含每月定期定額') + '<p class="hint">核心 60% 在 1 月投入，其餘當預備金。寫入後會自動用新的預算重新分析。</p>'
    : fNum('v', '景氣對策信號綜合分數', cur, '9～45') + fText('m', '是哪一個月的分數', cfg['景氣燈號月份'] || '', '例如 2026-10') + '<p class="hint">國發會每月 27 日左右公布。Claude 的早晨排程也會自動查最新一期。</p>';
  openSheet({
    title: isBudget ? '修改單筆預算' : '修改景氣燈號', ok: '寫入試算表', noUndo: true,
    send: async (action, payload) => {
      let m = await etfLoad('cfg', payload[0]);
      if (payload[1]) await etfLoad('cfg', payload[1]);
      try { await etfLoad('run'); m += '，已用新的數字重新分析'; } catch (e) { m += '（重新分析失敗：' + e.message + '）'; }
      etfS.view = null;
      return m;
    },
    body,
    submit(f) {
      const v = num(val(f, 'v'));
      if (v == null) throw new Error('請填數字');
      if (isBudget) { if (v < 10 || v > 300) throw new Error('預算請填 10～300 萬'); return ['cfg', [{ key, value: String(Math.round(v * 10000)) }]]; }
      if (v < 9 || v > 45) throw new Error('分數是 9～45');
      const m = val(f, 'm');
      return ['cfg', [{ key, value: String(v) }, m ? { key: '景氣燈號月份', value: m } : null]];
    }
  });
}

/* ---------- 綁定 ---------- */
function bindEtf() {
  if (route() === 'more/etf' && !etfS.subChecked) etfCheckSub();
  document.querySelectorAll('[data-etf-view]').forEach(el => el.addEventListener('click', () => { etfS.view = el.dataset.etfView || null; render(); window.scrollTo(0, 0); }));
  document.querySelectorAll('[data-etf-cfg]').forEach(el => el.addEventListener('click', () => {
    if (hidden) { toast('先按右上角的眼睛顯示數字，再編輯'); return; }
    formEtfCfg(el.dataset.etfCfg);
  }));
  const on = $('#etfOn'); if (on) on.addEventListener('click', etfSubscribe);
  const off = $('#etfOff'); if (off) off.addEventListener('click', etfUnsubscribe);
  const te = $('#etfTest'); if (te) te.addEventListener('click', async () => { te.disabled = true; try { toast('✓ ' + await etfLoad('test')); } catch (e) { toast(e.message); } te.disabled = false; });
  const run = $('#etfRun'); if (run) run.addEventListener('click', async () => {
    run.disabled = true; run.textContent = '分析中…';
    try { toast('✓ ' + await etfLoad('run')); etfS.view = null; } catch (e) { toast(e.message); }
    render();
  });
  if (etfS.pop && !lock.locked) etfOpenPopup();
}

// 通知點下去：App 已經開著時 Service Worker 會傳訊息過來
if ('serviceWorker' in navigator) navigator.serviceWorker.addEventListener('message', e => { if (e.data && e.data.type === 'etfpop') { etfS.pop = true; if (route() !== 'more/etf') location.hash = '#more/etf'; else render(); } });
