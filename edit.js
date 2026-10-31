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

function toast(msg) {
  let t = $('#toast');
  if (!t) { t = document.createElement('div'); t.id = 'toast'; document.body.appendChild(t); }
  t.textContent = msg; t.classList.add('show');
  clearTimeout(toast.t); toast.t = setTimeout(() => t.classList.remove('show'), 2600);
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
      closeSheet(); render(); toast('✓ ' + msg);
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
      `<div class="grp">分配到哪裡</div>` + fNum('tw', '台股（台幣）', '', '0') +
      `<div class="row2">${fNum('cube', '美股 CUBE（美金）', '', '0')}${fNum('ft', '美股 Firstrade（美金）', '', '0')}</div>` +
      hint('costHint') + fText('note', '備註', '', '例如 7.5W入台股'),
    after(f) {
      const t = nval(f, 'total') || 0, tw = nval(f, 'tw') || 0, usd = (nval(f, 'cube') || 0) + (nval(f, 'ft') || 0);
      const rest = t - tw - usd * (e.fx || 0);
      f.querySelector('#costHint').innerHTML = t ? `美金以匯率 ${fmtN(e.fx, 2)} 換算約 NT$${fmtN(usd * e.fx)}；` +
        (Math.abs(rest) < 1000 ? '已全部分配。' : `其餘 NT$${fmtN(rest)} 只記在總成本（例如虛擬幣）。`) : '';
    },
    submit(f) {
      const s = val(f, 'dir') === 'out' ? -1 : 1, t = req(nval(f, 'total'), '總金額');
      return ['cost', { date: slash(val(f, 'date')), total: s * t, tw: s * (nval(f, 'tw') || 0),
        us: [{ broker: 'CUBE', usd: s * (nval(f, 'cube') || 0) }, { broker: 'Firstrade', usd: s * (nval(f, 'ft') || 0) }], note: val(f, 'note') }];
    }
  });
}

/* ---------- 2. 個股買賣 ---------- */
function formTrade(market) {
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
    year: { title: '新年度（調薪／升遷）', body: fNum('year', '年度', last.year + 1) + fNum('salary', '調薪後月薪', '', `目前 ${fmtN(last.salary)}`) + hint('rsHint') +
        fSeg('promoted', [['0', '沒有升遷'], ['1', '升遷']], '0') + fText('title', '職稱', last.title),
      after: f => { const s = nval(f, 'salary'); f.querySelector('#rsHint').textContent = s && last.salary ? `調薪 ${s - last.salary >= 0 ? '+' : ''}${fmtN(s - last.salary)}（${((s / last.salary - 1) * 100).toFixed(1)}%）` : ''; },
      submit: f => ['career', { op: 'year', year: req(nval(f, 'year'), '年度'), salary: req(nval(f, 'salary'), '調薪後月薪'), promoted: val(f, 'promoted') === '1', title: val(f, 'title') }] },
    title: { title: '修改職稱', body: yearSel(last.year) + fText('title', '職稱', last.title) + fSeg('promoted', [['0', '一般'], ['1', '升遷（加 ↑）']], '0'),
      submit: f => ['career', { op: 'title', year: Number(val(f, 'year')), title: val(f, 'title'), promoted: val(f, 'promoted') === '1' }] }
  };
  openSheet(specs[kind]);
}

/* ---------- 綁定 ---------- */
const FORMS = { cost: formCost, trade: formTrade, dca: formDca, plan: formPlan, expense: formExpense, loan: formLoan, milestone: formMilestone, career: formCareer };

function bindEdit() {
  document.querySelectorAll('[data-form]').forEach(el => el.addEventListener('click', ev => {
    ev.stopPropagation();
    if (!ed()) { toast('請先更新 Apps Script 到最新版'); return; }
    if (hidden) { toast('先按右上角的眼睛顯示數字，再編輯'); return; }
    FORMS[el.dataset.form](el.dataset.arg);
  }));
}
