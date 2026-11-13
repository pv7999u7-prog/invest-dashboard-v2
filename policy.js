// 人壽保單：固定支出頁的第二個分頁。資料在另一份試算表（保單試算表），由 Apps Script 讀寫。
// 只有打開這個分頁才會去讀；讀到的內容放在 state.raw._policy，跟其他快取一起存在這台裝置上。
// 欄位用試算表第一列的標題找，App 裡不寫死任何人名或金額。
'use strict';

const POLICY_SCRIPT = 31;          // Apps Script 第幾版起支援保單
const POLICY_STALE_MS = 5 * 60000; // 超過這麼久沒讀，打開分頁時再讀一次
let expTab = 'month';              // 固定支出頁目前的分頁：month 每月支出／policy 人壽保單
let polView = null;                // 保單分頁裡正在看的：null 總覽、'p:保單號碼' 一張保單、's:姓名' 一個人的理賠總覽
const polState = { busy: false, at: 0, error: '' };

const polNo = v => String(v == null ? '' : v).replace(/[^0-9A-Za-z]/g, '').replace(/^0+/, '');
// 一段文字裡（例如繳費日期的「遠雄 1020210385」）有沒有這個保單號碼
const polIn = (text, no) => String(text == null ? '' : text).split(/[^0-9A-Za-z]+/).some(t => t && polNo(t) === no);
// 第一列是標題：每一列變成 { 標題: 內容, _row: 試算表列號 }
function polRows(grid) {
  if (!grid || !grid.length) return [];
  const head = grid[0].map(h => String(h).trim());
  return grid.slice(1).map((r, i) => { const o = { _row: i + 2 }; head.forEach((h, c) => { if (h) o[h] = String(r[c] == null ? '' : r[c]).trim(); }); return o; });
}
const polKey = (o, re) => Object.keys(o).find(k => re.test(k));
// 「12/16」→ 下一次繳費日（今天也算）和還有幾天
function polNext(md) {
  const m = String(md || '').match(/(\d{1,2})\s*[\/月.-]\s*(\d{1,2})/);
  if (!m) return null;
  const now = new Date(), today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  let d = new Date(today.getFullYear(), Number(m[1]) - 1, Number(m[2]));
  if (d < today) d = new Date(today.getFullYear() + 1, Number(m[1]) - 1, Number(m[2]));
  return { date: d, days: Math.round((d - today) / 86400000), month: Number(m[1]), day: Number(m[2]) };
}
const polDays = n => (n === 0 ? '今天' : n === 1 ? '明天' : n + ' 天後');

function parsePolicy(P) {
  const cover = polRows(P.cover).filter(c => c['保單號碼']);
  const pay = polRows(P.pay).filter(x => x['項目']);
  const list = polRows(P.list).filter(p => p['保單號碼']).map(p => {
    const no = polNo(p['保單號碼']);
    const pr = pay.find(x => polIn(x['項目'], no));
    return Object.assign(p, { no, premium: num(p['年繳保費']), next: polNext(p['每年繳費日']), payRow: pr || null,
      cover: cover.filter(c => polNo(c['保單號碼']) === no) });
  });
  const people = [];
  list.forEach(p => { if (people.indexOf(p['被保險人']) < 0) people.push(p['被保險人']); });
  return { list, people, sums: parsePolicySums(P.sum) };
}

// 理賠總表：每個人一段（姓名一列，接著「保險公司」「保單號碼」…），取「合計」那一欄
function parsePolicySums(grid) {
  const out = [];
  let cur = null, col = -1;
  (grid || []).forEach((r, i) => {
    const a = String(r[0] || '').trim(), nextA = String(((grid[i + 1] || [])[0]) || '').trim();
    if (a && nextA === '保險公司' && r.slice(1).every(v => String(v || '').trim() === '')) { cur = { name: a, rows: [] }; col = -1; out.push(cur); return; }
    if (!cur || !a) return;
    if (a === '險種') { col = r.findIndex(v => String(v).trim() === '合計'); return; }
    if (col < 0 || ['保險公司', '保單號碼', '投保日期', '投保年齡', '約別', '繳費年期', '保障期'].indexOf(a) >= 0) return;
    const v = String(r[col] == null ? '' : r[col]).trim();
    if (v && v !== '-') cur.rows.push({ k: a, v });
  });
  return out;
}

/* ---------- 讀取與寫入 ---------- */
async function loadPolicy() {
  if (polState.busy || !state.cfg || !state.cfg.url) return;
  polState.busy = true; polState.error = '';
  try {
    const j = await post({ policy: true }, '保單');
    if (!j.ok) throw new Error(j.error === 'unauthorized' ? '密碼不對' : (j.error || '讀取失敗'));
    if (state.raw) { state.raw._policy = j.policy; await persist(); }
  } catch (e) {
    polState.error = e.message === 'Failed to fetch' || e.name === 'TypeError' ? '連不上，請檢查網路' : e.message;
  }
  polState.busy = false; polState.at = Date.now();
  if (route() === 'more/expense' && expTab === 'policy') render();
}

async function sendPolicy(action, payload) {
  const j = await post({ policyAction: action, payload }, '保單');
  if (!j.ok) throw new Error(j.error === 'unauthorized' ? '密碼不對' : (j.error || '寫入失敗'));
  if (state.raw) { state.raw._policy = j.policy; await persist(); }
  polState.at = Date.now();
  return j.message || '已寫入保單試算表';
}

/* ---------- 畫面 ---------- */
const expSeg = () => `<div class="seg" id="expSeg"><button data-k="month" class="${expTab === 'month' ? 'on' : ''}">每月支出</button><button data-k="policy" class="${expTab === 'policy' ? 'on' : ''}">人壽保單</button></div>`;
const polMoney = v => (v ? ntd(v) : '—');
// 理賠內容很長：照「；」「。」拆成一點一點
const polBullets = t => { const a = String(t || '').split(/[；。]/).map(s => s.trim()).filter(Boolean); return a.length ? `<ul class="ben">${a.map(s => `<li>${esc(s)}</li>`).join('')}</ul>` : ''; };

function pagePolicy() {
  let h = head('固定支出') + expSeg();
  if (scriptVersion() < POLICY_SCRIPT) return h + `<div class="card"><h3>人壽保單</h3><p class="cap" style="margin:0">這個分頁需要 Apps Script 第 ${POLICY_SCRIPT} 版，請先更新 Apps Script 再回來。</p></div>`;
  const P = state.raw && state.raw._policy;
  if ((!P || Date.now() - polState.at > POLICY_STALE_MS) && !polState.busy && !(polState.error && Date.now() - polState.at < 15000)) setTimeout(loadPolicy, 0);
  const err = polState.error ? `<p class="msg err" style="margin:0 4px 10px">${esc(polState.error)}</p>` : '';
  if (!P) return h + err + `<div class="empty">${polState.error ? '讀不到保單資料' : '正在讀取保單…'}</div>${polState.error ? '<button class="btn ghost" id="polReload">再試一次</button>' : ''}`;
  if (!P.linked) return h + err + `<div class="card"><h3>連結保單試算表</h3><p class="cap">貼上雲端硬碟裡保單試算表的網址（要有「人壽保單」分頁）。只要設定一次，網址會記在你的 Apps Script 裡。</p>
    <label class="field"><span>試算表網址</span><input id="polUrl" type="url" placeholder="https://docs.google.com/spreadsheets/d/…" autocomplete="off" autocapitalize="off"></label>
    <button class="btn" id="polLink">連結</button></div>`;
  const D = parsePolicy(P);
  if (polView && polView[0] === 'p') { const p = D.list.find(x => 'p:' + x.no === polView); if (p) return h + err + policyDetail(p); polView = null; }
  if (polView && polView[0] === 's') { const s = D.sums.find(x => 's:' + x.name === polView); if (s) return h + err + policySum(s); polView = null; }

  const total = D.list.reduce((t, p) => t + (p.premium || 0), 0);
  const sched = D.list.filter(p => p.next).sort((a, b) => a.next.date - b.next.date), nx = sched[0];
  h += err + `<div class="card hero"><div class="lbl">全年保費</div><div class="big">${ntd(total)}</div>
    <div class="meta">${D.list.length} 張保單・${D.people.length} 人${nx ? `・下一筆 ${nx.next.month}/${nx.next.day}（${polDays(nx.next.days)}）` : ''}</div></div>`;
  h += `<div class="sec">繳費時程<em>從最近的開始</em></div><div class="list">${sched.map(p => `<div class="row link" data-pol="p:${esc(p.no)}">
    <div class="main"><div class="t1" style="font-size:15px">${p.next.month}/${p.next.day}　${esc(p['被保險人'])}</div>
      <div class="t2 wrap"><span class="nb">${esc(p['保險公司'])} ${esc(p['保單號碼'])}</span><wbr><span class="nb">・${esc(p['繳費方式'] || '—')}</span></div></div>
    <div class="r"><div class="t1" style="font-size:15px">${polMoney(p.premium)}</div><div class="t2">${polDays(p.next.days)}</div></div><span class="chev">›</span></div>`).join('')}</div>`;
  D.people.forEach(name => {
    const mine = D.list.filter(p => p['被保險人'] === name), sum = D.sums.find(s => s.name.indexOf(name) === 0);
    h += `<div class="sec">${esc(name)}<em>${mine.length} 張・年繳 ${hidden ? '••••' : int(mine.reduce((t, p) => t + (p.premium || 0), 0))}</em></div><div class="list">` +
      mine.map(p => `<div class="row link" data-pol="p:${esc(p.no)}"><div class="main"><div class="t1" style="font-size:15px">${esc(p['保險公司'])}　${esc(p['保單號碼'])}</div>
        <div class="t2 wrap">${esc(p['險種內容（主約＋附約）'] || p[polKey(p, /險種/)] || '')}</div></div>
        <div class="r"><div class="t1" style="font-size:15px">${polMoney(p.premium)}</div><div class="t2">${esc(p['每年繳費日'] || '')}</div></div><span class="chev">›</span></div>`).join('') +
      (sum ? `<div class="row link" data-pol="s:${esc(sum.name)}"><div class="main"><div class="t1" style="font-size:15px">理賠總覽</div><div class="t2">全部保單加起來，各種情況能領多少</div></div><span class="chev">›</span></div>` : '') + '</div>';
  });
  h += `<p class="hint" style="margin:14px 4px 8px">資料來自「${esc(P.title || '保單試算表')}」${P.at ? '，讀取時間 ' + esc(P.at) : ''}${polState.busy ? '・更新中…' : ''}。在這裡修改會寫回那份試算表。</p>
    <div class="acts wide"><button class="act" id="polReload">重新整理</button>${P.url ? `<a class="act" href="${esc(P.url)}" target="_blank" rel="noopener">開啟試算表</a>` : ''}</div>`;
  return h;
}

function policyDetail(p) {
  const kv = (k, v) => (v ? `<div class="kv"><span>${k}</span><span>${v}</span></div>` : '');
  const diff = p.payRow && num(p.payRow['金額']) && p.premium && Math.abs(num(p.payRow['金額']) - p.premium) > 0.5;
  let h = `<button class="act" data-pol="" style="margin:0 0 12px">‹ 全部保單</button>
  <div class="card"><h3>${esc(p['保險公司'])}　${esc(p['保單號碼'])}</h3><p class="cap">${esc(p['被保險人'])}</p>
    ${kv('每年繳費日', esc(p['每年繳費日']) + (p.next ? `（${polDays(p.next.days)}）` : ''))}
    ${kv('年繳保費', polMoney(p.premium))}
    ${kv('扣款方式', esc(p['繳費方式']))}
    ${kv('投保／繳費期間', esc(p['投保／繳費期間']))}
    ${diff ? `<p class="hint" style="margin:8px 0 0;color:var(--neg, #ff8a8a)">「繳費日期」分頁寫的金額是 ${hidden ? '••••' : esc(p.payRow['金額'])}，和這裡不一樣。</p>` : ''}
    <div class="grp" style="margin-top:14px">險種內容</div><p class="ptext">${esc(p['險種內容（主約＋附約）'] || '')}</p>
    ${p['備註'] ? `<div class="grp">備註</div><p class="ptext">${esc(p['備註'])}</p>` : ''}
    <button class="btn ghost" data-form="policy" data-arg="${esc(p.no)}" style="margin-top:12px">修改繳費與保單資料</button></div>
  <div class="sec">保障與理賠<em>${p.cover.length} 個險種・點一下可以修改</em></div>`;
  h += p.cover.map(c => { const pk = polKey(c, /^年繳保費/);
    return `<div class="card link" data-form="policyCover" data-arg="${c._row}"><h3 style="margin-bottom:2px">${esc(c['險種'])}</h3>
      <p class="cap">${[c['類別'], c['保額／計畫']].filter(Boolean).map(esc).join('・')}</p>
      ${kv('繳費期／保障期', [c['繳費期'], c['保障期']].filter(Boolean).map(esc).join('／'))}
      ${kv('年繳保費', pk && num(c[pk]) ? ntd(num(c[pk])) : '')}
      ${polBullets(c['理賠內容'])}</div>`; }).join('') || '<div class="empty">「保障與理賠」分頁沒有這張保單的資料</div>';
  return h;
}

function policySum(s) {
  return `<button class="act" data-pol="" style="margin:0 0 12px">‹ 全部保單</button>
  <div class="card"><h3>${esc(s.name)}</h3><p class="cap">理賠總覽：名下全部保單加起來（來自「理賠總表」分頁，只能看，要改請到試算表）</p>
    ${s.rows.map(r => `<div class="kv pol"><span>${esc(r.k)}</span><span>${hidden && /^[\d,]+$/.test(r.v) ? MASK : esc(r.v)}</span></div>`).join('')}</div>`;
}

/* ---------- 表單 ---------- */
// 只送有改的欄位；fields = [[標題, 輸入類型]]
function policyForm(title, row, base, fields) {
  openSheet({
    title, ok: '寫入保單試算表', send: sendPolicy, noUndo: true,
    body: fields.map(([k, type], i) => (type === 'area' ? fArea('f' + i, k, row[k] || '') : fText('f' + i, k, row[k] || ''))).join(''),
    submit(f) {
      const values = {};
      fields.forEach(([k], i) => { const v = val(f, 'f' + i); if (v !== String(row[k] || '').trim()) values[k] = v; });
      if (!Object.keys(values).length) throw new Error('沒有改動');
      return ['edit', Object.assign({ row: row._row, values }, base)];
    }
  });
}
function formPolicy(no) {
  const P = state.raw && state.raw._policy, p = P && parsePolicy(P).list.find(x => x.no === no);
  if (!p) return;
  policyForm('修改保單', p, { tab: 'list', no: p['保單號碼'] },
    [['每年繳費日', 'text'], ['年繳保費', 'text'], ['繳費方式', 'text'], ['投保／繳費期間', 'text'], ['險種內容（主約＋附約）', 'area'], ['備註', 'area']].filter(x => x[0] in p));
}
function formPolicyCover(rowNo) {
  const P = state.raw && state.raw._policy, c = P && polRows(P.cover).find(x => x._row === Number(rowNo));
  if (!c) return;
  const pk = polKey(c, /^年繳保費/);
  policyForm('修改險種', c, { tab: 'cover', no: c['保單號碼'], name: c['險種'] },
    [['險種', 'text'], ['類別', 'text'], ['保額／計畫', 'text'], ['繳費期', 'text'], ['保障期', 'text']].concat(pk ? [[pk, 'text']] : []).concat([['理賠內容', 'area']]).filter(x => x[0] in c));
}

// 這個檔在 edit.js 之後才載入，所以表單在這裡登記
FORMS.policy = formPolicy;
FORMS.policyCover = formPolicyCover;

/* ---------- 綁定（每次畫面重畫後） ---------- */
function bindPolicy() {
  const seg = $('#expSeg');
  if (seg) seg.addEventListener('click', e => { const b = e.target.closest('button'); if (!b) return; expTab = b.dataset.k; polView = null; render(); window.scrollTo(0, 0); });
  document.querySelectorAll('[data-pol]').forEach(el => el.addEventListener('click', () => { polView = el.dataset.pol || null; render(); window.scrollTo(0, 0); }));
  const rl = $('#polReload');
  if (rl) rl.addEventListener('click', () => { polState.error = ''; loadPolicy(); render(); });
  const lk = $('#polLink');
  if (lk) lk.addEventListener('click', async () => {
    const url = $('#polUrl').value.trim();
    if (!url) { toast('請貼上試算表網址'); return; }
    lk.disabled = true; lk.textContent = '連結中…';
    try { const msg = await sendPolicy('link', { url }); toast('✓ ' + msg); }
    catch (e) { toast(e.message === 'Failed to fetch' || e.name === 'TypeError' ? '連不上，請檢查網路' : e.message); }
    render();
  });
}
