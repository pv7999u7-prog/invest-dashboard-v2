// 輸入表單：從 App 寫回試算表。送出後 Apps Script 會回傳最新資料，畫面直接更新。
// 表單需要的現值與位置都來自 Apps Script 的 _edit（state.raw._edit），不寫死在程式裡。
'use strict';

const ed = () => (state.raw && state.raw._edit) || null;
const today = () => { const d = new Date(); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`; };
const slash = s => String(s || '').replace(/-/g, '/');
const val = (f, k) => { const el = f.querySelector(`[name="${k}"]`); return el ? el.value.trim() : ''; };
const nval = (f, k) => { const v = val(f, k).replace(/,/g, ''); return v === '' ? null : Number(v); };
const fmtN = (n, d = 0) => (n == null || isNaN(n) ? '—' : Number(n).toLocaleString('en-US', { maximumFractionDigits: d }));

// 頁面標題旁的按鈕：btns = [[文字, 表單名稱, 參數]]
function head(title, btns = []) {
  if (!ed() || !btns.length) return `<h1 class="title">${title}</h1>`;
  return `<div class="hrow"><h1 class="title">${title}</h1><div class="acts">${btns.map(b =>
    `<button class="act" data-form="${b[1]}"${b[2] != null ? ` data-arg="${esc(String(b[2]))}"` : ''}>${b[0]}</button>`).join('')}</div></div>`;
}

/* ---------- 送出 ---------- */
async function send(action, payload) {
  const res = await fetch(state.cfg.url, { method: 'POST', body: JSON.stringify({ key: state.cfg.key, action, payload }),
    headers: { 'Content-Type': 'text/plain;charset=utf-8' }, cache: 'no-store', redirect: 'follow' });
  const j = await res.json();
  if (!j.ok) throw new Error(j.error === 'unauthorized' ? '密碼不對' : (j.error || '寫入失敗'));
  state.raw = j.data; state.updated = j.updated; state.data = parseAll(j.data); state.lastFetch = Date.now();
  await persist();
  setStatus('ok');
  return j.message || '已寫入試算表';
}

// undoable：寫入成功後的提示多一個「復原」按鈕，6 秒內可以直接按
function toast(msg, undoable) {
  let t = $('#toast');
  if (!t) { t = document.createElement('div'); t.id = 'toast'; document.body.appendChild(t); }
  t.innerHTML = esc(msg) + (undoable ? '<button type="button" class="tundo">復原</button>' : '');
  t.classList.toggle('act', !!undoable);
  const b = t.querySelector('.tundo');
  if (b) b.addEventListener('click', () => { t.classList.remove('show'); doUndo(true); });
  t.classList.add('show');
  clearTimeout(toast.t); toast.t = setTimeout(() => t.classList.remove('show'), undoable ? 6000 : 2600);
}

// 復原上一步（從 App 做的最後一次修改）
async function doUndo(quick) {
  const u = ed() && ed().undo;
  if (!quick && !confirm(`要復原「${(u && u.desc) || '上一步'}」嗎？`)) return;
  toast('復原中…');
  try { const msg = await send('undo', {}); render(); toast('✓ ' + msg); }
  catch (e) { toast(e.message === 'Failed to fetch' || e.name === 'TypeError' ? '連不上，請檢查網路' : e.message); }
}

/* ---------- 表單外框 ---------- */
// spec: { title, body(html), submit(form) → [action, payload] 或 null, after(form), extra: [{label, cls, run(form)}] }
function openSheet(spec) {
  closeSheet();
  const o = document.createElement('div');
  o.className = 'overlay'; o.id = 'sheet';
  o.innerHTML = `<form class="sheet" novalidate><div class="sh-head"><h2>${spec.title}</h2><button type="button" class="x" aria-label="關閉">✕</button></div>
    <div class="sh-body">${spec.body}</div><p class="msg err" id="shErr"></p>
    <button class="btn" type="submit">${spec.ok || '寫入試算表'}</button>
    ${(spec.extra || []).map((x, i) => `<button type="button" class="btn ${x.cls || 'ghost'}" data-extra="${i}">${x.label}</button>`).join('')}</form>`;
  document.body.appendChild(o);
  document.body.classList.add('noscroll');
  const f = o.querySelector('form'), err = o.querySelector('#shErr');
  o.addEventListener('click', e => { if (e.target === o) closeSheet(); });
  o.querySelector('.x').addEventListener('click', closeSheet);
  f.querySelectorAll('.seg').forEach(seg => seg.addEventListener('click', e => {
    const b = e.target.closest('button'); if (!b) return;
    seg.querySelectorAll('button').forEach(x => x.classList.toggle('on', x === b));
    seg.querySelector('input').value = b.dataset.v;
    f.dispatchEvent(new Event('input'));
  }));
  const run = async (getReq) => {
    err.textContent = '';
    let req;
    try { req = getReq(); } catch (e) { err.textContent = e.message; return; }
    if (!req) return;
    const btns = f.querySelectorAll('.btn');
    btns.forEach(b => { b.disabled = true; });
    const first = btns[0], label = first.textContent;
    first.textContent = '寫入中…';
    try {
      const msg = await send(req[0], req[1]);
      closeSheet(); render(); toast('✓ ' + msg, true);
    } catch (e) {
      err.textContent = e.message === 'Failed to fetch' || e.name === 'TypeError' ? '連不上，請檢查網路' : e.message;
      btns.forEach(b => { b.disabled = false; }); first.textContent = label;
    }
  };
  f.addEventListener('submit', e => { e.preventDefault(); run(() => spec.submit(f)); });
  (spec.extra || []).forEach((x, i) => f.querySelector(`[data-extra="${i}"]`).addEventListener('click', () => run(() => x.run(f))));
  if (spec.after) { f.addEventListener('input', () => spec.after(f)); spec.after(f); }
}

function closeSheet() { const o = $('#sheet'); if (o) o.remove(); document.body.classList.remove('noscroll'); }

/* ---------- 欄位 ---------- */
const fDate = (k, label, v) => `<label class="field"><span>${label}</span><input name="${k}" type="date" value="${v || today()}"></label>`;
const fNum = (k, label, v, ph = '', hint = '') => `<label class="field"><span>${label}${hint ? `<em>${hint}</em>` : ''}</span><input name="${k}" inputmode="decimal" value="${v == null ? '' : v}" placeholder="${ph}" autocomplete="off"></label>`;
const fText = (k, label, v = '', ph = '', list = '') => `<label class="field"><span>${label}</span><input name="${k}" value="${esc(v)}" placeholder="${ph}" autocomplete="off"${list ? ` list="${list}"` : ''}></label>`;
const fArea = (k, label, v = '') => `<label class="field"><span>${label}</span><textarea name="${k}" rows="3">${esc(v)}</textarea></label>`;
const fSeg = (k, opts, v) => `<div class="seg">${opts.map(o => `<button type="button" data-v="${o[0]}" class="${o[0] === v ? 'on' : ''}">${o[1]}</button>`).join('')}<input type="hidden" name="${k}" value="${v}"></div>`;
const fSelect = (k, label, opts, v) => `<label class="field"><span>${label}</span><select name="${k}">${opts.map(o => `<option value="${esc(o[0])}"${o[0] === v ? ' selected' : ''}>${esc(o[1])}</option>`).join('')}</select></label>`;
const hint = id => `<p class="hint" id="${id}"></p>`;
const req = (v, label) => { if (v == null || v === '' || isNaN(v)) throw new Error('請填' + label); return v; };

/* ---------- 1. 成本變更 ---------- */
function formCost() {
  const e = ed();
  openSheet({
    title: '成本變更',
    body: fSeg('dir', [['in', '投入'], ['out', '撤出']], 'in') + fDate('date', '日期') +
      fNum('total', '總金額（台幣）', '', '例如 400000') +
      `<div class="grp">分配到哪裡<em>全部填台幣，美股和加密貨幣依匯率換成美金寫入</em></div>` + fNum('tw', '台股', '', '0') +
      `<div class="row2">${fNum('cube', '美股 CUBE', '', '0')}${fNum('ft', '美股 Firstrade', '', '0')}</div>` +
      `<div class="row2">${fNum('crypto', '加密貨幣', '', '0')}${fNum('fx', '匯率', e.fx ? +e.fx.toFixed(2) : '', '', '可改成實際換匯')}</div>` +
      hint('costHint') + fText('note', '備註', '', '例如 7.5W入台股'),
    after(f) {
      const t = nval(f, 'total') || 0, fx = nval(f, 'fx') || 0, parts = ['tw', 'cube', 'ft', 'crypto'].map(k => nval(f, k) || 0);
      const rest = t - parts.reduce((a, b) => a + b, 0);
      const conv = [['CUBE', parts[1]], ['Firstrade', parts[2]], ['加密貨幣', parts[3]]].filter(x => x[1]);
      f.querySelector('#costHint').innerHTML = (fx && conv.length ? conv.map(x => `${x[0]} 約 ${x[0] === '加密貨幣' ? '' : 'US$'}${fmtN(x[1] / fx, 2)}${x[0] === '加密貨幣' ? ' USDT' : ''}`).join('・') + '<br>' : '') +
        (t ? (Math.abs(rest) < 1 ? '已全部分配。' : rest > 0 ? `其餘 NT$${fmtN(rest)} 只記在總成本。` : `分配比總金額多 NT$${fmtN(-rest)}，請檢查。`) : '');
    },
    submit(f) {
      const s = val(f, 'dir') === 'out' ? -1 : 1, t = req(nval(f, 'total'), '總金額');
      const fx = nval(f, 'fx'), usd = k => { const n = nval(f, k) || 0; if (n && !(fx > 0)) throw new Error('請填匯率'); return n ? s * Math.round(n / fx * 100) / 100 : 0; };
      const sum = ['tw', 'cube', 'ft', 'crypto'].reduce((a, k) => a + (nval(f, k) || 0), 0);
      if (sum > t + 0.5) throw new Error('分配的金額比總金額多');
      return ['cost', { date: slash(val(f, 'date')), total: s * t, tw: s * (nval(f, 'tw') || 0),
        us: [{ broker: 'CUBE', usd: usd('cube') }, { broker: 'Firstrade', usd: usd('ft') }], crypto: usd('crypto'), note: val(f, 'note') }];
    }
  });
}

/* ---------- 2. 個股買賣 ---------- */
function formTrade(market) {
  if (market === 'crypto') return formCryptoTrade();
  const e = ed(), tw = market === 'tw';
  const brokers = Object.keys(e.us.brokers);
  const list = tw ? e.tw.holdings.map(h => `<option value="${esc(h.code)}">${esc(h.name)}</option>`)
    : [...new Set(brokers.flatMap(b => e.us.brokers[b].holdings.map(h => h.code)))].map(c => `<option value="${esc(c)}">`);
  openSheet({
    title: tw ? '台股買賣' : '美股買賣',
    body: fSeg('side', [['buy', '買進'], ['sell', '賣出']], 'buy') + fDate('date', '日期') +
      (tw ? '' : fSelect('broker', '券商', brokers.map(b => [b, b]), brokers[brokers.length - 1])) +
      fText('stock', tw ? '股號或股名' : '代號', '', tw ? '例如 2330 或 台積電' : '例如 NVDA', 'stockList') +
      `<datalist id="stockList">${list.join('')}</datalist>` +
      `<div class="row2">${fNum('price', '成交單價', '', tw ? '例如 1050' : '例如 185.5')}${fNum('qty', '股數', '', tw ? '例如 1000' : '例如 10')}</div>` +
      hint('tradeHint') + fText('note', '備註'),
    after(f) {
      const p = nval(f, 'price') || 0, q = nval(f, 'qty') || 0, buy = val(f, 'side') === 'buy';
      const cash = tw ? e.tw.cash : (e.us.brokers[val(f, 'broker')] || {}).cash;
      const amt = p * q, m = n => (n < 0 ? '−' : '') + (tw ? 'NT$' + fmtN(Math.abs(n)) : 'US$' + fmtN(Math.abs(n), 2));
      const code = val(f, 'stock').toUpperCase();
      const h = tw ? e.tw.holdings.find(x => x.code === code || x.name === val(f, 'stock'))
        : (e.us.brokers[val(f, 'broker')] || { holdings: [] }).holdings.find(x => x.code === code);
      let s = amt ? `金額 ${m(amt)}・閒置現金 ${m(cash)} → <b>${m(cash + (buy ? -amt : amt))}</b>` : `閒置現金 ${m(cash)}`;
      if (h) s += `<br>目前持有 ${fmtN(h.qty, 4)} 股，均價 ${fmtN(h.avg, 2)}` + (buy && q ? ` → 買後均價 ${fmtN((h.avg * h.qty + p * q) / (h.qty + q), 2)}` : '');
      else if (code && buy) s += '<br>新股票：會自動加進持股清單和 Price 分頁';
      f.querySelector('#tradeHint').innerHTML = s;
    },
    submit(f) {
      if (!val(f, 'stock')) throw new Error(tw ? '請填股號或股名' : '請填代號');
      return ['trade', { market, broker: tw ? '' : val(f, 'broker'), date: slash(val(f, 'date')), side: val(f, 'side'), stock: val(f, 'stock'),
        price: req(nval(f, 'price'), '成交單價'), qty: req(nval(f, 'qty'), '股數'), note: val(f, 'note') }];
    }
  });
}

/* ---------- 2b. 加密貨幣買賣、出入金（閒置現金＝Binance 的 USDT） ---------- */
function formCryptoTrade() {
  const e = ed(), c = e.crypto, coins = c.coins.filter(x => x.name !== 'USDT');
  const wheres = [...new Set(['Binance'].concat(c.coins.map(x => x.where)))];
  openSheet({
    title: '加密貨幣買賣',
    body: fSeg('side', [['buy', '買進'], ['sell', '賣出']], 'buy') + fDate('date', '日期') +
      `<div class="row2">${fText('coin', '幣名', '', '例如 BTC', 'coinList')}${fSelect('where', '存於', wheres.map(w => [w, w]), 'Binance')}</div>` +
      `<datalist id="coinList">${[...new Set(coins.map(x => x.name))].map(n => `<option value="${esc(n)}">`).join('')}</datalist>` +
      `<div class="row2">${fNum('price', '成交單價（USD）', '', '例如 77000')}${fNum('qty', '顆數', '', '例如 0.05')}</div>` +
      hint('ctHint') + fText('note', '備註'),
    after(f) {
      const p = nval(f, 'price') || 0, q = nval(f, 'qty') || 0, buy = val(f, 'side') === 'buy', amt = p * q;
      const coin = val(f, 'coin').toUpperCase(), h = coins.find(x => x.name === coin && x.where === val(f, 'where'));
      const m = n => (n < 0 ? '−' : '') + fmtN(Math.abs(n), 2);
      let s = amt ? `金額 ${m(amt)} USDT・USDT ${m(c.cash)} → <b>${m(c.cash + (buy ? -amt : amt))}</b>` : `USDT 餘額 ${m(c.cash)}`;
      if (h) s += `<br>${esc(h.where)} 目前 ${fmtN(h.qty, 8)} 顆` + (q ? ` → ${fmtN(h.qty + (buy ? q : -q), 8)} 顆` : '');
      else if (coin && buy) s += '<br>新的幣或存放位置：會自動加進持有清單';
      f.querySelector('#ctHint').innerHTML = s;
    },
    submit(f) {
      if (!val(f, 'coin')) throw new Error('請填幣名');
      return ['trade', { market: 'crypto', broker: val(f, 'where'), date: slash(val(f, 'date')), side: val(f, 'side'), stock: val(f, 'coin'),
        price: req(nval(f, 'price'), '成交單價'), qty: req(nval(f, 'qty'), '顆數'), note: val(f, 'note') }];
    }
  });
}

function formCryptoFlow() {
  const e = ed(), c = e.crypto;
  const srcs = [['fiat', '台幣（新資金／提現）'], ['tw', '台股帳戶'], ...Object.keys(e.us.brokers).map(b => [b, '美股 ' + b])];
  openSheet({
    title: '加密貨幣出入金',
    body: fSeg('dir', [['in', '轉入加密貨幣'], ['out', '轉出加密貨幣']], 'in') + fDate('date', '日期') +
      fSelect('source', '資金來源／去處', srcs, 'fiat') +
      `<div class="row2">${fNum('usdt', 'USDT 數量', '', '例如 3000')}<div id="ntdBox">${fNum('ntd', '台幣金額', '', '例如 98000')}</div></div>` +
      hint('cfHint') + fText('note', '備註'),
    after(f) {
      const src = val(f, 'source'), twd = src === 'fiat' || src === 'tw', u = nval(f, 'usdt') || 0, out = val(f, 'dir') === 'out';
      f.querySelector('#ntdBox').style.visibility = twd ? '' : 'hidden';
      const m = n => (n < 0 ? '−' : '') + fmtN(Math.abs(n), 2);
      let s = `USDT ${m(c.cash)}` + (u ? ` → <b>${m(c.cash + (out ? -u : u))}</b>` : '');
      if (src === 'fiat') s += '<br>算新的投入／撤出：會記在 Total Overall 的總成本';
      else if (src === 'tw') s += `<br>內部轉帳：台股閒置現金 NT$${fmtN(e.tw.cash)} 會${out ? '增加' : '減少'}，總成本不變`;
      else { const b = e.us.brokers[src]; s += `<br>內部轉帳：${src} 閒置現金 US$${fmtN(b.cash, 2)} → ${fmtN(b.cash + (out ? u : -u), 2)}（1 USDT 以 1 USD 計），總成本不變`; }
      const ntd = nval(f, 'ntd');
      if (twd && u && ntd) s += `<br>換算匯率約 ${fmtN(ntd / u, 2)}`;
      f.querySelector('#cfHint').innerHTML = s;
    },
    submit(f) {
      const src = val(f, 'source'), twd = src === 'fiat' || src === 'tw';
      return ['cryptoFlow', { date: slash(val(f, 'date')), dir: val(f, 'dir'), source: src, usdt: req(nval(f, 'usdt'), 'USDT 數量'),
        ntd: twd ? req(nval(f, 'ntd'), '台幣金額') : '', note: val(f, 'note') }];
    }
  });
}

/* ---------- 3. 定期定額 ---------- */
function formDca() {
  const e = ed(), items = e.dca.filter(i => i.monthly || i.shares);
  const ntd = items.reduce((s, i) => s + (i.cur === 'USD' ? Math.round(i.monthly * e.fx / 1000) * 1000 : i.monthly), 0);
  openSheet({
    title: '本月定期定額',
    body: fDate('date', '扣款日') + fNum('total', '本月總投入（台幣）', ntd, '', '記在 Total Overall 的總成本') +
      items.map((i, k) => `<div class="grp">${esc(((i.owner || '') + ' ' + i.name).trim())}<em>累計投入 ${i.cur === 'USD' ? 'US$' : 'NT$'}${fmtN(i.invested, 2)}・${fmtN(i.shares, 4)} 股</em></div>
        <div class="row2">${fNum('add' + k, `本月投入（${i.cur === 'USD' ? '美金' : '台幣'}）`, i.monthly)}${fNum('sh' + k, '目前累計股數', i.shares)}</div>`).join(''),
    submit(f) {
      return ['dca', { date: slash(val(f, 'date')), total: nval(f, 'total') || 0,
        items: items.map((i, k) => ({ id: i.id, add: nval(f, 'add' + k) || 0, shares: val(f, 'sh' + k) === '' ? '' : nval(f, 'sh' + k) })) }];
    }
  });
}

/* ---------- 4. 長期計畫 ---------- */
function formPlan() {
  const p = ed().plan;
  openSheet({
    title: '長期計畫',
    body: `<div class="row2">${fNum('rate', '目標年化報酬率（%）', p.rate)}${fNum('yearly', '每年預設投入（萬）', p.yearly)}</div>
      <div class="grp">各年度<em>投入空白＝用預設值；實際投入在 12/31 會自動記錄，也可以手動改</em></div>
      <div class="ytable"><div class="yh"><span>年度</span><span>計畫投入（萬）</span><span>實際投入（萬）</span></div>` +
      p.years.map((y, k) => `<div class="yr"><span>${y.year}</span><input name="in${k}" inputmode="decimal" value="${y.custom && y.input != null ? y.input : ''}" placeholder="${p.yearly}">
        <input name="ac${k}" inputmode="decimal" value="${y.actual == null ? '' : y.actual}" placeholder="—"></div>`).join('') + '</div>',
    submit(f) {
      const years = [];
      p.years.forEach((y, k) => {
        const inp = val(f, 'in' + k), ac = val(f, 'ac' + k);
        const oldIn = y.custom && y.input != null ? String(y.input) : '', oldAc = y.actual == null ? '' : String(y.actual);
        if (inp !== oldIn || ac !== oldAc) years.push({ year: y.year, input: inp === '' ? '' : Number(inp), actual: ac === '' ? '' : Number(ac) });
      });
      return ['plan', { rate: req(nval(f, 'rate'), '年化報酬率'), yearly: req(nval(f, 'yearly'), '每年投入'), years }];
    }
  });
}

/* ---------- 5. 固定支出（一次編一個月） ---------- */
function formExpense(month) {
  const m = ed().expense.find(x => x.month === Number(month));
  if (!m) return;
  const rows = m.items.concat(Array.from({ length: Math.min(2, m.max - m.items.length) }, () => ({ name: '', amount: null })));
  openSheet({
    title: `${month} 月固定支出`,
    body: `<p class="hint">名稱清空就是刪除；最多 ${m.max} 項。年度分類（房貸、保險、稅金、學費、紅包）會依名稱自動加總。</p>` +
      `<div class="ytable"><div class="yh two"><span>項目</span><span>金額</span></div>` +
      rows.map((x, k) => `<div class="yr two"><input name="n${k}" value="${esc(x.name)}" placeholder="新增項目"><input name="a${k}" inputmode="decimal" value="${x.amount == null ? '' : x.amount}" placeholder="0"></div>`).join('') + '</div>',
    submit(f) {
      return ['expense', { month: Number(month), items: rows.map((x, k) => ({ name: val(f, 'n' + k), amount: nval(f, 'a' + k) })) }];
    }
  });
}

/* ---------- 6. 房貸 ---------- */
// 本息平均攤還的月付金：剩餘本金 × 月利率 ÷ (1 − (1+月利率)^−期數)
const annuity = (bal, ratePct, n) => { const r = ratePct / 100 / 12; return r ? bal * r / (1 - Math.pow(1 + r, -n)) : bal / n; };

function formLoan(name) {
  const h = (state.data.house || []).find(x => x.name === name), l = h && h.loan;
  if (!l) return;
  const paid = l.sched[l.paidN - 1], base = paid ? paid.date : l.base;
  openSheet({
    title: `${esc(name)}・房貸`,
    body: `<p class="hint">從 ${esc(base)} 的剩餘本金 NT$${fmtN(l.balance)} 起算，還剩 ${l.leftN} 期。</p>` +
      `<div class="row2">${fNum('rate', '年利率（%）', +(l.rate * 100).toFixed(4))}${fNum('pmt', '每月還款', '', '空白＝自動估算')}</div>` +
      hint('loanHint') + fText('end', '到期日', l.end || '', '2050/10/26'),
    after(f) {
      const r = nval(f, 'rate'), est = r != null ? Math.round(annuity(l.balance, r, l.leftN)) : null;
      f.querySelector('#loanHint').innerHTML = est ? `依 ${r}% 估算每月還款約 <b>NT$${fmtN(est)}</b>（目前 NT$${fmtN(l.pmt)}）` : '';
    },
    submit(f) {
      const r = req(nval(f, 'rate'), '年利率');
      const pmt = nval(f, 'pmt') || Math.round(annuity(l.balance, r, l.leftN));
      return ['loan', { name, rate: r, pmt, base, balance: Math.round(l.balance), end: slash(val(f, 'end')) }];
    }
  });
}

/* ---------- 6b. 房產估價 ---------- */
function formHouse(name) {
  const x = (state.data.house || []).find(h => h.name === name);
  if (!x) return;
  openSheet({
    title: `${esc(name)}・估價`,
    body: `<p class="hint">估值＝房屋坪數 ${fmtN(x.housePing, 2)} 坪 × 估計單價 ＋ 估計車位價</p>` +
      `<div class="row2">${fNum('unit', '估計單價（萬/坪）', x.estUnit)}${fNum('park', `車位估價（萬，${x.parkN || 0} 個）`, x.estPark)}</div>` + hint('hsHint'),
    after(f) {
      const u = nval(f, 'unit'), p = nval(f, 'park') || 0;
      f.querySelector('#hsHint').innerHTML = u != null ? `估計現值 <b>${fmtN(x.housePing * u + p)} 萬</b>（目前 ${fmtN(x.est)} 萬，購入 ${fmtN(x.price)} 萬）` : '';
    },
    submit: f => ['house', { name, unit: req(nval(f, 'unit'), '估計單價'), park: nval(f, 'park') == null ? '' : nval(f, 'park') }]
  });
}

/* ---------- 7. 里程碑 ---------- */
function formMilestone(row) {
  const m = row ? ed().milestones.find(x => x.row === Number(row)) : null;
  const d = new Date(), ph = `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}`;
  openSheet({
    title: m ? '編輯里程碑' : '新增里程碑',
    body: fText('date', '日期', m ? m.date : ph, '例如 20261105') + fArea('text', '內容', m ? m.text : ''),
    submit(f) {
      if (!val(f, 'date') || !val(f, 'text')) throw new Error('日期和內容都要填');
      return ['milestone', { op: m ? 'set' : 'add', row: m && m.row, date: val(f, 'date'), text: val(f, 'text') }];
    },
    extra: m ? [{ label: '刪除這個里程碑', run: () => (confirm('確定要刪除這個里程碑？') ? ['milestone', { op: 'del', row: m.row }] : null) }] : []
  });
}

/* ---------- 8. 職涯與薪資 ---------- */
function formCareer(kind) {
  const ys = ed().career.years, last = ys[ys.length - 1] || { year: new Date().getFullYear(), title: '', salary: 0 };
  const yearSel = (v) => fSelect('year', '年度', ys.slice().reverse().map(y => [String(y.year), String(y.year)]), String(v));
  const specs = {
    mid: { title: '年中分紅', body: yearSel(last.year) + fNum('amount', '年中分紅金額'),
      submit: f => ['career', { op: 'mid', year: Number(val(f, 'year')), amount: req(nval(f, 'amount'), '年中分紅') }] },
    ye: { title: '年前分紅（農曆年前）', body: yearSel(last.year) + `<div class="row2">${fNum('allowance', '職務加給', '', '0')}${fNum('bonus', '年終（分紅）')}</div>` + hint('yeHint'),
      after: f => { f.querySelector('#yeHint').textContent = `合計 NT$${fmtN((nval(f, 'allowance') || 0) + (nval(f, 'bonus') || 0))}`; },
      submit: f => ['career', { op: 'ye', year: Number(val(f, 'year')), allowance: nval(f, 'allowance') || 0, bonus: req(nval(f, 'bonus'), '年終') }] },
    // 調薪升遷：一個表單四種類型，只顯示該類型要填的欄位
    //   年度調薪升遷：新增下一年度（op 'year'）；結構調薪、自訂項目：寫在那一年底下的新一列（op 'event'，結構調薪名稱自動是「11月結構調薪」）；
    //   改職稱：改某一年的職稱（op 'title'）
    raise: { title: '調薪升遷',
      body: fSeg('type', [['year', '年度調薪'], ['struct', '結構調薪'], ['custom', '自訂項目'], ['title', '改職稱']], 'year') +
        `<div id="rNew">${fNum('ny', '新年度', last.year + 1)}</div><div id="rPick">${yearSel(last.year)}</div>` +
        `<div id="rMonth">${fSelect('month', '月份', Array.from({ length: 12 }, (_, i) => [String(i + 1), (i + 1) + ' 月']), String(new Date().getMonth() + 1))}</div>` +
        `<div id="rLabel">${fText('label', '項目名稱', '', '例如 特別調薪、留任獎金')}</div>` +
        `<div id="rSal">${fNum('salary', '調整後月薪', '', `目前 ${fmtN(last.salary)}`)}${hint('rsHint')}</div>` +
        `<div id="rTitle">${fSeg('promoted', [['0', '沒有升遷'], ['1', '升遷（職稱前加 ↑）']], '0')}${fText('title', '職稱', last.title)}</div>`,
      after: f => {
        const t = val(f, 'type'), show = (id, on) => { f.querySelector(id).style.display = on ? '' : 'none'; };
        show('#rNew', t === 'year'); show('#rPick', t !== 'year'); show('#rMonth', t === 'struct'); show('#rLabel', t === 'custom');
        show('#rSal', t !== 'title'); show('#rTitle', t === 'year' || t === 'title');
        const s = nval(f, 'salary');
        f.querySelector('#rsHint').textContent = (t === 'custom' ? '沒有調薪可以空白。' : '') +
          (s && last.salary ? `月薪 ${s - last.salary >= 0 ? '+' : ''}${fmtN(s - last.salary)}（${((s / last.salary - 1) * 100).toFixed(1)}%）` : '');
      },
      submit: f => {
        const t = val(f, 'type'), sal = nval(f, 'salary'), promoted = val(f, 'promoted') === '1';
        if (t === 'year') return ['career', { op: 'year', year: req(nval(f, 'ny'), '新年度'), salary: req(sal, '調整後月薪'), promoted, title: val(f, 'title') }];
        if (t === 'title') {
          if (!val(f, 'title')) throw new Error('請填職稱');
          return ['career', { op: 'title', year: Number(val(f, 'year')), title: val(f, 'title'), promoted }];
        }
        const label = t === 'struct' ? `${val(f, 'month')}月結構調薪` : val(f, 'label');
        if (!label) throw new Error('請填項目名稱');
        if (t === 'struct' && !sal) throw new Error('請填調整後月薪');
        return ['career', { op: 'event', year: Number(val(f, 'year')), label, salary: sal || '' }];
      } }
  };
  openSheet(specs[kind]);
}

// 時間軸上的項目（結構調薪、自訂項目）：改名稱、月薪，或刪除
function formCareerEvent(row) {
  let ev = null, yr = null;
  ((state.data && state.data.career && state.data.career.years) || []).forEach(y => y.events.forEach(e => { if (e.row === Number(row)) { ev = e; yr = y.year; } }));
  if (!ev) return;
  openSheet({
    title: `${yr} 年・${esc(ev.label)}`, ok: '儲存修改',
    body: fText('label', '項目名稱', ev.label) + fNum('salary', '調整後月薪', ev.salary, '', '沒有調薪可以空白'),
    submit: f => {
      if (!val(f, 'label')) throw new Error('請填項目名稱');
      return ['career', { op: 'eventSet', row: Number(row), expect: ev.label, label: val(f, 'label'), salary: nval(f, 'salary') || '' }];
    },
    extra: [{ label: '刪除這個項目', run: () => (confirm(`確定要刪除「${ev.label}」嗎？刪錯了可以按復原。`) ? ['career', { op: 'eventDel', row: Number(row), expect: ev.label }] : null) }]
  });
}

/* ---------- 9. 明細修改、刪除（資金異動、台股、美股、加密貨幣） ---------- */
const dash = s => String(s || '').replace(/\//g, '-');   // 2026/10/31 → 2026-10-31（日期欄位要這種格式）

// 同一次「成本變更」或「加密貨幣出入金」在其他明細的那幾筆：日期、備註相同的出入金（Apps Script 用同樣的規則）
const isFlowEntry = (k, x) => (k === 'flow' ? !!x.flow : k === 'crypto' ? x.type !== '買進' && x.type !== '賣出' : !!x.flow && !x.code);
function linkedOf(kind, x) {
  const note = String(x.note || '').trim(), L = (state.data && state.data.ledger) || {};
  if (!isFlowEntry(kind, x) || !note) return [];
  const out = [];
  ['flow', 'tw', 'us', 'crypto'].forEach(k => (L[k] || []).forEach(y => {
    if (k === kind && y.row === x.row) return;
    if (y.date === x.date && String(y.note || '').trim() === note && isFlowEntry(k, y)) out.push({ k, y });
  }));
  return out;
}
const linkedName = ({ k, y }) => (k === 'flow' ? `資金異動 ${y.flow < 0 ? '投入' : '撤出'} ${fmtN(Math.abs(y.flow))}`
  : k === 'tw' ? `台股 ${y.flow < 0 ? '入金' : '出金'} ${fmtN(Math.abs(y.flow))}`
  : k === 'us' ? `美股 ${y.b} ${y.flow < 0 ? '入金' : '出金'} US$${fmtN(Math.abs(y.flow), 2)}`
  : `加密貨幣 ${y.type} ${fmtN(Math.abs(y.usdt || 0), 2)} USDT`);
const linkHint = links => (links.length
  ? `<p class="hint">同一次出入金還記在：${links.map(l => esc(linkedName(l))).join('、')}。<br>刪除會一起刪；改日期或備註也會一起改；金額要各自改。</p>`
  : '');

function formLedger(arg) {
  const [kind, r] = String(arg).split(':'), row = Number(r);
  const x = (((state.data && state.data.ledger) || {})[kind] || []).find(i => i.row === row);
  if (!x) { toast('找不到這筆，請下拉更新後再試'); return; }
  const base = { kind, row, expect: { date: x.date, key: x.key } };
  const links = linkedOf(kind, x);
  const ask = links.length
    ? `會一起刪除同一次出入金的另外 ${links.length} 筆：\n${links.map(linkedName).join('\n')}\n\n確定要刪除嗎？刪錯了可以按復原。`
    : '確定要刪除這筆嗎？刪錯了可以在「明細與復原」按復原。';
  const del = { label: links.length ? `刪除這筆（連同另外 ${links.length} 筆）` : '刪除這筆', run: () => (confirm(ask) ? ['ledger', Object.assign({ op: 'del' }, base)] : null) };
  const edit = data => ['ledger', Object.assign({ op: 'edit', data }, base)];

  if (kind === 'flow') {
    openSheet({
      title: '修改資金異動', ok: '儲存修改', extra: [del],
      body: fSeg('dir', [['in', '投入'], ['out', '撤出']], x.flow < 0 ? 'in' : 'out') + fDate('date', '日期', dash(x.date)) +
        fNum('amt', '金額（台幣）', Math.abs(x.flow)) + fText('note', '備註', x.note) + linkHint(links),
      submit: f => { const a = req(nval(f, 'amt'), '金額'); return edit({ date: slash(val(f, 'date')), flow: val(f, 'dir') === 'in' ? -a : a, note: val(f, 'note') }); }
    });
    return;
  }

  if (kind === 'tw' || kind === 'us') {
    const tw = kind === 'tw', brokers = Object.keys(ed().us.brokers);
    const type = x.buy ? 'buy' : x.sell ? 'sell' : x.flow < 0 ? 'in' : 'out';
    openSheet({
      title: tw ? '修改台股明細' : '修改美股明細', ok: '儲存修改', extra: [del],
      body: fSeg('type', [['buy', '買進'], ['sell', '賣出'], ['in', '入金'], ['out', '出金']], type) + fDate('date', '日期', dash(x.date)) +
        (tw ? '' : fSelect('broker', '券商', brokers.map(b => [b, b]), x.b || brokers[0])) +
        `<div id="ldTrade">${fText('stock', tw ? '股號或股名' : '代號', tw ? (x.code || x.b) : x.code)}
          <div class="row2">${fNum('price', '成交單價', x.price)}${fNum('qty', '股數', x.buy || x.sell)}</div></div>` +
        `<div id="ldFlow">${fNum('amount', tw ? '金額（台幣）' : '金額（美金）', x.flow ? Math.abs(x.flow) : '')}</div>` +
        fText('note', '備註', x.note) + hint('ldHint') + linkHint(links),
      after(f) {
        const t = val(f, 'type'), trade = t === 'buy' || t === 'sell';
        f.querySelector('#ldTrade').style.display = trade ? '' : 'none';
        f.querySelector('#ldFlow').style.display = trade ? 'none' : '';
        f.querySelector('#ldHint').textContent = '閒置現金' + (trade ? '和均價' : '') + '會跟著調整';
      },
      submit(f) {
        const t = val(f, 'type'), trade = t === 'buy' || t === 'sell';
        if (trade && !val(f, 'stock')) throw new Error(tw ? '請填股號或股名' : '請填代號');
        return edit({ date: slash(val(f, 'date')), type: t, broker: tw ? '' : val(f, 'broker'), stock: val(f, 'stock'),
          price: trade ? req(nval(f, 'price'), '成交單價') : '', qty: trade ? req(nval(f, 'qty'), '股數') : '',
          amount: trade ? '' : req(nval(f, 'amount'), '金額'), note: val(f, 'note') });
      }
    });
    return;
  }

  // 加密貨幣：買賣可以改幣名、位置、單價、顆數；出入金只改 USDT 數量
  const trade = x.type === '買進' || x.type === '賣出';
  openSheet({
    title: '修改加密貨幣紀錄', ok: '儲存修改', extra: [del],
    body: (trade ? fSeg('side', [['buy', '買進'], ['sell', '賣出']], x.type === '買進' ? 'buy' : 'sell')
      : `<p class="hint" style="margin-top:0">類型：${esc(x.type)}</p>` + fSeg('dir', [['in', '轉入'], ['out', '轉出']], (x.usdt || 0) < 0 ? 'out' : 'in')) +
      fDate('date', '日期', dash(x.date)) +
      (trade ? `<div class="row2">${fText('coin', '幣名', x.coin)}${fText('where', '存於', x.where)}</div>
        <div class="row2">${fNum('price', '成交單價（USD）', x.price)}${fNum('qty', '顆數', x.qty)}</div>`
        : fNum('usdt', 'USDT 數量', Math.abs(x.usdt || 0))) +
      fText('note', '備註', x.note) +
      `<p class="hint">${trade ? '顆數和 USDT 會跟著調整' : 'USDT 會跟著調整'}</p>` + linkHint(links),
    submit: f => edit(trade
      ? { date: slash(val(f, 'date')), side: val(f, 'side'), coin: val(f, 'coin'), where: val(f, 'where'), price: req(nval(f, 'price'), '單價'), qty: req(nval(f, 'qty'), '顆數'), note: val(f, 'note') }
      : { date: slash(val(f, 'date')), usdt: (val(f, 'dir') === 'out' ? -1 : 1) * req(nval(f, 'usdt'), 'USDT 數量'), note: val(f, 'note') })
  });
}

/* ---------- 10. 閒置現金（收到股利、和券商對帳時直接改） ---------- */
function formCash(id) {
  const a = ((ed() && ed().cash) || []).find(x => x.id === id);
  if (!a) return;
  const unit = a.cur === 'NTD' ? '台幣' : a.cur === 'USD' ? '美金' : 'USDT';
  openSheet({
    title: `${esc(a.label)}・閒置現金`,
    body: fNum('amount', `閒置現金（${unit}）`, a.value) + hint('cashHint'),
    after(f) {
      const v = nval(f, 'amount'), d = v != null && a.value != null ? v - a.value : 0;
      f.querySelector('#cashHint').innerHTML = `目前 ${fmtN(a.value, 2)}${Math.abs(d) > 0.004 ? `，${d > 0 ? '增加' : '減少'} ${fmtN(Math.abs(d), 2)}` : ''}<br>` +
        '收到股利或和券商對帳時，直接填券商 App 上的現金餘額。只改現金，不會記成投入成本。';
    },
    submit: f => ['cash', { id, amount: req(nval(f, 'amount'), '閒置現金') }]
  });
}

/* ---------- 綁定 ---------- */
const FORMS = { cost: formCost, trade: formTrade, cryptoFlow: formCryptoFlow, dca: formDca, plan: formPlan, expense: formExpense, loan: formLoan, house: formHouse, milestone: formMilestone, career: formCareer, careerEvent: formCareerEvent, ledger: formLedger, cash: formCash };

function bindEdit() {
  // ＋ 開頭的按鈕是「新增」：加號換成金色
  document.querySelectorAll('.act').forEach(b => {
    if (b.classList.contains('add') || !/^＋/.test(b.textContent)) return;
    b.classList.add('add');
    b.innerHTML = '<span class="plus">＋</span>' + b.innerHTML.replace(/^＋/, '');
  });
  document.querySelectorAll('[data-form]').forEach(el => el.addEventListener('click', ev => {
    ev.stopPropagation();
    if (!ed()) { toast('請先更新 Apps Script 到最新版'); return; }
    if (hidden) { toast('先按右上角的眼睛顯示數字，再編輯'); return; }
    FORMS[el.dataset.form](el.dataset.arg);
  }));
  // 明細頁：切換種類、復原上一步
  const seg = $('#ledgerSeg');
  if (seg) seg.addEventListener('click', e => { const b = e.target.closest('button'); if (!b) return; ledgerKind = b.dataset.k; render(); });
  const ub = $('#undoBtn');
  if (ub) ub.addEventListener('click', () => doUndo(false));
}
