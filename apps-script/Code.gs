// 投資總覽 v2 的資料介面：App 可以讀，也可以寫回試算表（雙向）
// 安裝方式：試算表 → 擴充功能 → Apps Script → 把這整段貼上 → 改下面的 KEY → 部署為網頁應用程式
//   執行身分：我　／　誰可以存取：所有人
// App 每次都帶著密碼來；密碼不對就什麼都拿不到，也不能寫。
// 每次從 App 寫入，都會在「App 紀錄」分頁留一筆；改到的格子另外記在隱藏的「App 復原」分頁，App 可以「復原上一步」。

// 只要改這一行：密碼放在兩個單引號中間，例如 const KEY = 'abc123';
const KEY = '請改成你自己的密碼';

// 程式碼版本（App 設定頁會顯示，用來確認部署的是不是最新版；這行不用改）
const SCRIPT_VERSION = 36;

// 要讀的工作表與範圍（試算表版面有大改時，才需要調整這裡）
const RANGES = {
  total:     ['Total Overall', 'A1:Z300'],
  tw:        ['台股', 'A1:Q400'],
  us:        ['美股', 'A1:Q600'],
  crypto:    ['幣 ', 'A1:AF200'],
  myDca:     ['My 定期定額 (台幣帳戶)', 'A1:J40'],
  kidsUs:    ['橙&澈 美股', 'A1:J60'],
  kidsTw:    ['橙&澈 定期定額台股', 'A1:J80'],
  plan:      ['Plan', 'A1:I40'],
  milestone: ['Milestone', 'A1:B200'],
  lifeMs:    ['人生里程碑', 'A1:B200'],
  expense:   ['固定支出', 'A1:L60'],
  house:     ['House', 'A1:L10'],
  bb:        ['BB', 'A1:AD60'],
  retire:    ['退休金', 'A1:F120']
};

const LOG_SHEET = 'App 紀錄';
const LIFE_SHEET = '人生里程碑';   // 和 Milestone 分頁同樣格式：A 日期（文字）B 內容
function lifeSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(LIFE_SHEET) || ss.insertSheet(LIFE_SHEET);
}
const BROKERS = { CUBE: 3, Firstrade: 4 };   // 美股分頁：券商 → 摘要列（C 欄是現金）
const CRYPTO = '幣 ';                          // 加密貨幣分頁（名稱後面有一個空白）
const CRYPTO_CASH = ['USDT', 'Binance'];        // 加密貨幣的閒置現金：Binance 上的 USDT
// 固定支出：1～6 月的項目在第 2～8 列，7～12 月在第 10～17 列；每個月兩欄（項目、金額）
const EXP_BLOCKS = [{ months: [1, 2, 3, 4, 5, 6], from: 2, to: 8 }, { months: [7, 8, 9, 10, 11, 12], from: 10, to: 17 }];

// App 用 POST 把密碼放在內容裡送過來，不會出現在網址上
function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents) || {}; } catch (err) {}
  // 用字串比對，KEY 就算忘了加引號寫成數字也能對上
  const expected = String(KEY).trim();
  if (!expected || expected.startsWith('請改成') || String(body.key || '').trim() !== expected) {
    return json({ ok: false, error: 'unauthorized' });
  }
  MS = { t0: Date.now() };
  // 人壽保單（另一份試算表）：只讀寫保單，不讀投資的資料
  if (body.policy || body.policyAction) return policyRequest(body);
  // 0050 長期投資：每日建議、設定、通知訂閱（不讀投資的資料）
  if (body.etf || body.etfAction) return etfRequest(body);
  ensureTrigger();
  try { ensureLiveTrigger(); } catch (err) {}
  // 即時價會讓整份試算表重算，很花時間：寫入時不抓（背景每 10 分鐘會抓）；
  // 讀取時只有 App 另外要求（live）才抓，不然只在背景超過 15 分鐘沒跑時補抓
  if (!body.action) { try { refreshLive(false, body.live ? 0 : 15 * 60 * 1000); } catch (err) {} }
  mark('live');
  if (body.action) {
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const fn = ACTIONS[body.action];
      if (!fn) throw new Error('不認得的動作：' + body.action);
      if (body.action !== 'undo') REC = { desc: '', ops: [] };
      const msg = fn(body.payload || {});
      SpreadsheetApp.flush();
      mark('write');
      const data = readAll();
      // 復原紀錄和 App 紀錄等讀完才寫：每次「寫了再讀」試算表都要整份重算一次（好幾秒），所以寫的動作集中放最後
      const desc = msg || body.action, saved = !!(REC && REC.ops.length), at = now();
      if (REC) { REC.desc = desc; recSave(); }
      log(body.action, body.payload);
      if (saved && data._edit && data._edit.undo) data._edit.undo = { count: Math.min((data._edit.undo.count || 0) + 1, UNDO_KEEP), time: at, desc };
      mark('log');
      return json({ ok: true, message: msg || '已寫入試算表', updated: at, data, ms: msOut() });
    } catch (err) {
      // 寫到一半失敗：照記下的資料倒回去，不留半套
      let rolled = false;
      if (REC && REC.ops.length) {
        try { SpreadsheetApp.flush(); } catch (e) {}
        try { applyUndo(REC.ops); SpreadsheetApp.flush(); rolled = true; } catch (e) {}
      }
      REC = null; RC = null;
      return json({ ok: false, error: String(err.message || err) + (rolled ? '（已自動還原，試算表沒有被改動）' : '') });
    } finally {
      lock.releaseLock();
    }
  }
  const data = readAll();
  return json({ ok: true, updated: now(), data, ms: msOut() });
}

// 各階段花的時間（毫秒），App 設定頁會顯示，用來看慢在哪裡
let MS = { t0: 0 };
function mark(name) { const t = Date.now(); MS[name] = t - (MS.last || MS.t0); MS.last = t; }
function msOut() {
  const o = { total: Date.now() - MS.t0 };
  ['live', 'write', 'ranges', 'edit', 'daily', 'log'].forEach(k => { if (MS[k] != null) o[k] = MS[k]; });
  return o;
}

// 直接用瀏覽器打開網址時，只確認服務有在運作、是第幾版，不給資料
function doGet() {
  return json({ ok: false, error: 'use the app', version: SCRIPT_VERSION });
}

/* ================================================================
   讀取
   ================================================================ */
// 每碰一次試算表大約要 0.15～0.3 秒，所以讀取期間每個分頁只整張讀一次（RC），後面都從記憶體拿
function readAll() {
  RC = {}; BG = null;
  try { return readAllCached(); } finally { RC = null; BG = null; }
}

function readAllCached() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const data = {};
  BG = batchGrids(Object.keys(RANGES).map(k => RANGES[k][0]).concat(['Price']));
  data._batch = BATCH;
  const isDateValue = v => Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime());
  Object.keys(RANGES).forEach(name => {
    const [sheetName, a1] = RANGES[name];
    const sheet = cachedSheet(sheetName);
    if (!sheet) { data[name] = []; return; }
    const rows = sheet.getRange(a1).getValues().map(row => row.map(v =>
      isDateValue(v) ? Utilities.formatDate(v, tz, 'yyyy/MM/dd') : v));
    // 去掉尾端的空白列，減少傳輸量
    let last = rows.length;
    while (last > 0 && rows[last - 1].every(v => v === '')) last--;
    data[name] = rows.slice(0, last);
  });
  data._sheetUrl = ss.getUrl();   // App「更多 → 原始資料」用來打開這份試算表
  data._scriptVersion = SCRIPT_VERSION;
  mark('ranges');
  try { data._edit = editInfo(); } catch (err) { data._editError = String(err.message || err); }
  mark('edit');
  try { data._daily = dailyChanges(); } catch (err) {}
  mark('daily');
  try { data._live = liveInfo(); } catch (err) {}
  return data;
}

// App 輸入表單要用的現值與位置
function editInfo() {
  const tw = sh('台股'), us = sh('美股');
  const twH = holdings(tw);
  const info = { tw: { cash: num(tw.getRange('C3').getValue()), holdings: twH,
    prices: twPrices(tw, twH) },
    us: { brokers: {} }, dca: dcaItems() };
  Object.keys(BROKERS).forEach(b => {
    info.us.brokers[b] = { cash: num(us.getRange('C' + BROKERS[b]).getValue()), holdings: usHoldings(b) };
  });
  info.fx = num(sh('Total Overall').getRange('B2').getValue());
  info.total = num(sh('Total Overall').getRange('G2').getValue());
  info.expense = expenseItems();
  info.milestones = sh('Milestone').getRange('A1:B200').getDisplayValues()
    .map((r, i) => ({ row: i + 1, date: r[0], text: r[1] })).filter(m => m.date || m.text);
  const life = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LIFE_SHEET);
  info.lifeMilestones = life ? sh(LIFE_SHEET).getRange('A1:B200').getDisplayValues()
    .map((r, i) => ({ row: i + 1, date: r[0], text: r[1] })).filter(m => m.date || m.text) : [];
  info.plan = planInfo();
  info.loans = loanRows();
  info.career = careerInfo();
  const coins = cryptoRows(), usdt = coins.find(x => x.name === CRYPTO_CASH[0] && x.where === CRYPTO_CASH[1]);
  info.crypto = { coins, cash: usdt ? usdt.qty : 0 };
  info.house = houseRows();
  info.undo = undoInfo();
  info.cash = cashAccounts();
  return info;
}

/* ================================================================
   寫入動作
   ================================================================ */
const ACTIONS = {
  // 成本變更：total 正數＝投入、負數＝撤出；tw（台幣）進台股；us: [{broker, usd}] 進美股
  cost(p) {
    const total = need(p.total, '總金額'), date = toDate(p.date);
    const note = p.note || (total >= 0 ? '資金投入' : '資金撤出');
    appendTotalOverall(date, -total, note);
    if (num(p.tw)) {
      const tw = sh('台股'), last = lastRowOf(tw, 'A', '日期');
      appendRow(tw, last, 1, 10, { 1: date, 2: '', 3: -num(p.tw), 4: '', 5: '', 6: '', 7: '', 8: '', 9: '', 10: note });
      addTo(tw.getRange('C3'), num(p.tw));
    }
    (p.us || []).filter(u => num(u.usd)).forEach(u => {
      if (!BROKERS[u.broker]) throw new Error('不認得的券商：' + u.broker);
      const us = sh('美股'), last = lastRowOf(us, 'A', '日期');
      appendRow(us, last, 1, 10, { 1: date, 2: u.broker, 3: -num(u.usd), 4: '', 5: '', 6: '', 7: '', 8: '', 9: '', 10: note });
      addTo(us.getRange('C' + BROKERS[u.broker]), num(u.usd));
    });
    if (num(p.crypto)) {
      const c = cryptoCash();
      addTo(sh(CRYPTO).getRange(c.row, 3), num(p.crypto));
      cryptoLog(date, num(p.crypto) > 0 ? '入金（台幣）' : '出金（台幣）', 'USDT', CRYPTO_CASH[1], '', '', num(p.crypto), note);
    }
    return `已記錄${total >= 0 ? '投入' : '撤出'} ${fmt(Math.abs(total))}`;
  },

  // 個股買賣：market 'tw' 或 'us'；stock 可以是股號或股名；side 'buy' 或 'sell'
  trade(p) {
    const date = toDate(p.date), price = need(p.price, '單價'), qty = need(p.qty, '數量');
    if (price <= 0 || qty <= 0) throw new Error('單價和數量要大於 0');
    const buy = p.side === 'buy', amt = round2(price * qty);
    if (p.market === 'tw') {
      const s = sh('台股'), st = resolveTw(String(p.stock || '').trim());
      let h = holdings(s).find(x => x.code === st.code);
      const manual = num(p.manualPrice);   // 上櫃等抓不到價格的股票，手動填的現價
      if (!h) { if (!buy) throw new Error('台股沒有持有 ' + st.code); h = addTwHolding(st, price, manual); }
      else if (manual > 0) put(twPriceCell(s, h), manual);
      if (!buy && qty > h.qty) throw new Error(`${st.name} 只有 ${h.qty} 股，不能賣 ${qty} 股`);
      if (buy) setAvg(s.getRange(h.row, 15), h.avg, h.qty, price, qty);
      const last = lastRowOf(s, 'A', '日期');
      appendRow(s, last, 1, 10, { 1: date, 2: st.name, 3: '', 4: st.code, 5: price, 6: buy ? qty : '', 7: buy ? '' : qty, 8: r => (buy ? `=-E${r}*F${r}` : ''), 9: r => (buy ? '' : `=G${r}*E${r}`), 10: p.note || '' });
      addTo(s.getRange('C3'), buy ? -amt : amt);
      return `已記錄台股${buy ? '買進' : '賣出'} ${st.name} ${qty} 股`;
    }
    if (p.market === 'crypto') return cryptoTrade(p, date, price, qty, buy, amt);
    const broker = p.broker;
    if (!BROKERS[broker]) throw new Error('請選券商');
    const s = sh('美股'), code = String(p.stock || '').trim().toUpperCase();
    if (!/^[A-Z][A-Z0-9.\-]*$/.test(code)) throw new Error('美股請輸入代號，例如 QQQ');
    let h = usHoldings(broker).find(x => x.code === code);
    if (!h) { if (!buy) throw new Error(broker + ' 沒有持有 ' + code); h = addUsHolding(broker, code, price); }
    if (!buy && qty > h.qty + 1e-9) throw new Error(`${code} 只有 ${h.qty} 股，不能賣 ${qty} 股`);
    if (buy) setAvg(s.getRange(h.row, 14), h.avg, h.qty, price, qty);
    const last = lastRowOf(s, 'A', '日期');
    appendRow(s, last, 1, 10, { 1: date, 2: broker, 3: '', 4: code, 5: price, 6: buy ? qty : '', 7: buy ? '' : qty, 8: r => (buy ? `=-E${r}*F${r}` : ''), 9: r => (buy ? '' : `=G${r}*E${r}`), 10: p.note || '' });
    addTo(s.getRange('C' + BROKERS[broker]), buy ? -amt : amt);
    return `已記錄美股${buy ? '買進' : '賣出'} ${code} ${qty} 股`;
  },

  // 定期定額：total（台幣）記在 Total Overall；items: [{id, add, shares}] 累加總投入、更新累計股數
  dca(p) {
    const date = toDate(p.date), items = dcaItems();
    if (num(p.total)) appendTotalOverall(date, -num(p.total), p.note || '定期定額');
    (p.items || []).forEach(x => {
      const it = items.find(i => i.id === x.id);
      if (!it) throw new Error('找不到定期定額項目：' + x.id);
      const s = sh(it.sheet);
      if (num(x.add)) put(s.getRange(it.row, it.col + 5), round2(it.invested + num(x.add)));
      if (x.shares !== '' && x.shares != null) put(s.getRange(it.row, it.col + 6), num(x.shares));
    });
    return '已更新定期定額';
  },

  // 長期計畫：rate 年化報酬率（%）、yearly 每年預設投入（萬）、years: [{year, input, actual}]
  plan(p) {
    const s = sh('Plan'), info = planInfo();
    if (p.rate != null && p.rate !== '') put(s.getRange('C2'), num(p.rate));
    if (p.yearly != null && p.yearly !== '') put(s.getRange('F3'), num(p.yearly));
    (p.years || []).forEach(y => {
      const r = info.years.find(x => x.year === num(y.year));
      if (!r) return;
      if (y.input === '' || y.input == null) putF(s.getRange(r.row, 6), '=$F$3');
      else put(s.getRange(r.row, 6), num(y.input));
      if (y.actual !== undefined) put(s.getRange(r.row, 8), y.actual === '' ? '' : num(y.actual));
    });
    return '已更新長期計畫';
  },

  // 固定支出：一次存整個月；month 1～12，items: [{name, amount}]（名稱空白的項目會被刪掉）
  expense(p) {
    const s = sh('固定支出'), m = num(p.month);
    const blk = EXP_BLOCKS.find(b => b.months.indexOf(m) >= 0);
    if (!blk) throw new Error('月份要是 1～12');
    const col = blk.months.indexOf(m) * 2 + 1, n = blk.to - blk.from + 1;
    const rg = s.getRange(blk.from, col, n, 2);
    const list = (p.items || []).filter(x => String(x.name || '').trim() !== '')
      .map(x => [String(x.name).trim(), need(x.amount, `「${String(x.name).trim()}」的金額`)]);
    if (list.length > n) throw new Error(`${m} 月最多 ${n} 項`);
    while (list.length < n) list.push(['', '']);
    putVs(rg, list);
    return `已更新 ${m} 月固定支出`;
  },

  // 房貸：改利率、每月還款；改利率時用今天的剩餘本金重新起算
  loan(p) {
    const r = loanRows().find(x => x.name === p.name);
    if (!r) throw new Error('找不到房貸：' + p.name);
    const s = sh('House');
    if (p.base) put(s.getRange(r.row, 2), toDate(p.base));
    if (p.balance != null && p.balance !== '') put(s.getRange(r.row, 3), num(p.balance));
    if (p.rate != null && p.rate !== '') fmtCell(put(s.getRange(r.row, 4), num(p.rate) / 100), '0.00%');
    if (p.pmt != null && p.pmt !== '') put(s.getRange(r.row, 5), num(p.pmt));
    if (p.day) put(s.getRange(r.row, 6), num(p.day));
    if (p.end) put(s.getRange(r.row, 7), toDate(p.end));
    return '已更新房貸';
  },

  // 房產估價：unit 估計單價（萬/坪）、park 估計車位價（萬）
  house(p) {
    const r = houseRows().find(x => x.name === p.name);
    if (!r) throw new Error('找不到房產：' + p.name);
    const s = sh('House');
    put(s.getRange(r.row, 12), need(p.unit, '估計單價'));
    if (p.park != null && p.park !== '') put(s.getRange(r.row, 11), num(p.park));
    return `已更新 ${p.name} 估價`;
  },

  // 加密貨幣出入金：dir 'in' 轉進加密貨幣、'out' 轉出；source 'fiat'（台幣新資金）、'tw'（台股）、'CUBE'／'Firstrade'（美股）
  // usdt 是加密貨幣這邊增減的 USDT；ntd 是台幣那邊的金額（台幣、台股才需要）
  cryptoFlow(p) {
    const date = toDate(p.date), sign = p.dir === 'out' ? -1 : 1, usdt = need(p.usdt, 'USDT 數量');
    // 各表用同一個備註，之後在明細刪除或改日期時，才找得到同一次出入金的其他筆
    const note = p.note || { fiat: sign > 0 ? '加密貨幣入金' : '加密貨幣出金' }[p.source] || (sign > 0 ? '轉到加密貨幣' : '從加密貨幣轉回');
    if (usdt <= 0) throw new Error('USDT 數量要大於 0');
    const c = cryptoCash(), inOut = sign > 0 ? '轉入' : '轉出';
    if (sign < 0 && usdt > c.qty + 1e-9) throw new Error(`USDT 只有 ${c.qty}，不能轉出 ${usdt}`);
    let label;
    if (p.source === 'fiat') {
      const ntd = need(p.ntd, '台幣金額');
      appendTotalOverall(date, -sign * ntd, note);
      label = sign > 0 ? '入金（台幣）' : '出金（台幣）';
    } else if (p.source === 'tw') {
      const ntd = need(p.ntd, '台幣金額'), tw = sh('台股'), last = lastRowOf(tw, 'A', '日期');
      appendRow(tw, last, 1, 10, { 1: date, 2: '', 3: sign * ntd, 4: '', 5: '', 6: '', 7: '', 8: '', 9: '', 10: note });
      addTo(tw.getRange('C3'), -sign * ntd);
      label = '台股' + inOut;
    } else if (BROKERS[p.source]) {
      const us = sh('美股'), last = lastRowOf(us, 'A', '日期');
      appendRow(us, last, 1, 10, { 1: date, 2: p.source, 3: sign * usdt, 4: '', 5: '', 6: '', 7: '', 8: '', 9: '', 10: note });
      addTo(us.getRange('C' + BROKERS[p.source]), -sign * usdt);
      label = p.source + inOut;
    } else throw new Error('請選資金來源');
    addTo(sh(CRYPTO).getRange(c.row, 3), sign * usdt);
    cryptoLog(date, label, 'USDT', CRYPTO_CASH[1], '', '', sign * usdt, note);
    return `已記錄加密貨幣${sign > 0 ? '轉入' : '轉出'} ${usdt} USDT`;
  },

  // 里程碑：op 'add'（加在最後）、'set'（改）、'del'（刪除）；kind 'life' 是人生里程碑，其他是投資里程碑（Milestone 分頁）
  milestone(p) {
    const s = p.kind === 'life' ? lifeSheet() : sh('Milestone');
    if (p.op === 'add') {
      const vals = s.getRange('A1:B300').getValues();
      let r = vals.length; while (r > 0 && vals[r - 1][0] === '' && vals[r - 1][1] === '') r--;
      // 日期欄存成文字（例如 20261105），避免被當成數字或日期改掉格式
      put(fmtCell(s.getRange(r + 1, 1), '@'), need(p.date, '日期', true));
      put(s.getRange(r + 1, 2), need(p.text, '內容', true));
      return '已新增里程碑';
    }
    const row = num(p.row);
    if (!(row >= 1)) throw new Error('找不到這個里程碑');
    if (p.op === 'del') { snap(s.getRange(row, 1, 1, 2)); s.deleteRow(row); recRow('rowDel', s, row); return '已刪除里程碑'; }
    put(fmtCell(s.getRange(row, 1), '@'), String(need(p.date, '日期', true)));
    put(s.getRange(row, 2), need(p.text, '內容', true));
    return '已更新里程碑';
  },

  // 職涯：op 'mid' 年中分紅、'ye' 年前分紅（職務加給＋年終）、'year' 新增年度（調薪／升遷）、'title' 改職稱、
  // 'ot' 某年的加班費（U 欄 Overtime；Y 欄另一組計算的 Overtime 也填同一個數字）、
  // 'event' 在某年底下加一列項目（例如「11月結構調薪」或自訂名稱，可以帶調整後月薪）、'eventSet' 改、'eventDel' 刪
  career(p) {
    const s = sh('BB'), info = careerInfo(), y = num(p.year);
    if (p.op === 'eventSet' || p.op === 'eventDel') {
      const row = num(p.row), v = row >= 1 ? s.getRange(row, 1, 1, 2).getValues()[0] : null;
      const isYear = v && num(v[0]) > 1990 && num(v[0]) < 2100;
      if (!v || isYear || String(v[1]).trim() !== String(p.expect || '').trim()) throw new Error('這個項目在試算表上已經變動，請下拉更新後再試');
      if (p.op === 'eventDel') {
        snap(s.getRange(row, 1, 1, s.getLastColumn()));
        s.deleteRow(row);
        recRow('rowDel', s, row);
        return `已刪除「${v[1]}」`;
      }
      const label = need(p.label, '項目名稱', true), sal = num(p.salary);
      put(s.getRange(row, 2), label);
      put(s.getRange(row, 6), sal || '');
      put(s.getRange(row, 13), sal || '');
      return `已更新「${label}」`;
    }
    const r = info.years.find(x => x.year === y);
    if (p.op === 'year') {
      if (r) throw new Error(y + ' 年已經有了');
      const prev = info.years[info.years.length - 1];
      if (!prev) throw new Error('BB 分頁沒有任何年度');
      const row = prev.lastRow + 1;
      s.insertRowAfter(prev.lastRow);
      recRow('rowIns', s, row);
      const salary = need(p.salary, '調薪後月薪');
      const title = (p.promoted ? '↑ ' : '') + (p.title || prev.title);
      const pr = prev.row;
      s.getRange(row, 1, 1, 5).setValues([[y, title, num(s.getRange(pr, 3).getValue()) + 1, num(s.getRange(pr, 4).getValue()) + 1, num(s.getRange(pr, 5).getValue()) + 1]]);
      s.getRange(row, 6).setValue(prev.salary);
      s.getRange(row, 10).setFormula(`=IFERROR(I${row}/F${row},"")`);
      s.getRange(row, 11).setFormula(`=IFERROR(G${row}/F${row},"")`);
      s.getRange(row, 13).setValue(salary);
      s.getRange(row, 15).setFormula(`=IFERROR(N${row}/M${row},"")`);
      s.getRange(row, 17).setFormula(`=M${row}-M${prev.lastSalaryRow}`);
      s.getRange(row, 18).setFormula(`=Q${row}/M${prev.lastSalaryRow}`);
      return `已新增 ${y} 年`;
    }
    if (!r) throw new Error('BB 分頁找不到 ' + y + ' 年');
    if (p.op === 'event') {
      const label = need(p.label, '項目名稱', true), sal = num(p.salary), row = r.lastRow + 1;
      s.insertRowAfter(r.lastRow);
      recRow('rowIns', s, row);
      s.getRange(row, 2).setValue(label);
      if (sal) { s.getRange(row, 6).setValue(sal); s.getRange(row, 13).setValue(sal); }
      return `已新增 ${y} 年「${label}」`;
    }
    if (p.op === 'ot') {
      const amt = need(p.amount, '加班費');
      put(s.getRange(r.row, 21), amt);
      put(s.getRange(r.row, 25), amt);
      return `已記錄 ${y} 年加班費`;
    }
    if (p.op === 'mid') { put(s.getRange(r.row, 14), need(p.amount, '年中分紅')); putF(s.getRange(r.row, 15), `=IFERROR(N${r.row}/M${r.row},"")`); return `已記錄 ${y} 年中分紅`; }
    if (p.op === 'ye') {
      const a = num(p.allowance) || 0, b = need(p.bonus, '年終');
      put(s.getRange(r.row, 8), a); put(s.getRange(r.row, 9), b);
      putF(s.getRange(r.row, 7), `=H${r.row}+I${r.row}`);
      putF(s.getRange(r.row, 10), `=IFERROR(I${r.row}/F${r.row},"")`);
      putF(s.getRange(r.row, 11), `=IFERROR(G${r.row}/F${r.row},"")`);
      return `已記錄 ${y} 年前分紅`;
    }
    if (p.op === 'title') { put(s.getRange(r.row, 2), (p.promoted ? '↑ ' : '') + need(p.title, '職稱', true)); return '已更新職稱'; }
    throw new Error('不認得的職涯動作');
  },

  // 明細的修改與刪除：kind 'flow'（Total Overall 資金異動）、'tw'（台股交易表）、'us'（美股交易表）、'crypto'（幣分頁交易紀錄）
  // op 'edit'（data 是新內容）或 'del'；expect 是 App 看到的日期與金額，用來確認試算表上還是同一筆
  ledger(p) {
    const L = LEDGERS[p.kind];
    if (!L) throw new Error('不認得的明細種類');
    const s = sh(L.sheet), head = findRow(s, L.col, L.header), last = lastRowOf(s, L.col, L.header);
    let row = num(p.row);
    if (!(row > head && row <= last)) throw new Error('找不到這筆資料，請下拉更新後再試');
    const old = ledgerRead(p.kind, s, row);
    if (p.expect && (old.date !== dkey(p.expect.date) || Math.abs((old.key || 0) - (num(p.expect.key) || 0)) > 0.005))
      throw new Error('這筆資料在試算表上已經變動，請下拉更新後再試');
    // 同一次出入金在其他表的那幾筆（p.linked === false 才不連動）
    const group = p.linked === false ? [] : linkedRows(p.kind, row, old);
    if (p.op === 'del') {
      // 全部一起刪；同一張表從下往上刪，前面的列號才不會跑掉
      const all = group.concat([{ k: p.kind, r: row, x: old }]).sort((a, b) => (a.k < b.k ? -1 : a.k > b.k ? 1 : b.r - a.r));
      all.forEach(g => {
        const L2 = LEDGERS[g.k], s2 = sh(L2.sheet);
        L2.unapply(s2, g.x);
        deleteLedgerRow(s2, g.r, L2.c1, L2.n, lastRowOf(s2, L2.col, L2.header));
      });
      return group.length ? `已刪除這筆${L.label}，和同一次出入金的另外 ${group.length} 筆` : '已刪除這筆' + L.label;
    }
    // 資金異動依各帳戶金額重新分配（一天一筆入金，當天加碼就改這一筆）：
    // p.split = { dir: 'in'|'out', date, note, total（台幣）, tw（台幣）, CUBE、Firstrade（美金）, crypto（USDT） }
    // 帳戶沒給＝不動（只同步日期、備註）；給 0＝刪掉那個帳戶的那筆；原本沒有、給了金額＝新增一筆
    if (p.op === 'split') {
      if (p.kind !== 'flow') throw new Error('只有資金異動可以重新分配');
      const sp = p.split || {}, sign = sp.dir === 'out' ? -1 : 1;   // 1＝投入
      const date = sp.date || old.date, note = sp.note != null ? String(sp.note) : old.note;
      const total = need(sp.total, '總金額');
      if (total <= 0) throw new Error('總金額要大於 0');
      L.write(s, row, { date, flow: -sign * total, note });
      const nd = toDate(date), touched = {}, dels = [];
      const sync = g => {   // 只同步日期、備註
        const L2 = LEDGERS[g.k], s2 = sh(L2.sheet);
        if (dkey(nd) !== g.x.date) put(s2.getRange(g.r, L2.dateCol), nd);
        if (note.trim() !== String(g.x.note).trim()) put(s2.getRange(g.r, L2.noteCol), note);
      };
      const accts = [['tw', g => g.k === 'tw'], ['CUBE', g => g.k === 'us' && g.x.b === 'CUBE'], ['Firstrade', g => g.k === 'us' && g.x.b === 'Firstrade'], ['crypto', g => g.k === 'crypto']];
      accts.forEach(([a, is]) => {
        const g = group.find(is);
        if (g) touched[g.k + ':' + g.r] = true;
        const v = sp[a];
        if (v === undefined || v === null || v === '') { if (g) sync(g); return; }
        const amt = num(v);
        if (amt == null || amt < 0) throw new Error('金額不能是負的');
        if (g && amt === 0) { dels.push(g); return; }
        if (g) {
          const L2 = LEDGERS[g.k], s2 = sh(L2.sheet);
          L2.unapply(s2, g.x);
          if (g.k === 'crypto') L2.write(s2, g.r, { date, usdt: sign * amt, note }, g.x);
          else L2.write(s2, g.r, { date, type: sign > 0 ? 'in' : 'out', broker: g.x.b, amount: amt, note }, g.x);
          return;
        }
        if (!amt) return;
        // 原本沒有這個帳戶：照「成本變更」的方式新增一筆
        if (a === 'tw') {
          const tw = sh('台股');
          appendRow(tw, lastRowOf(tw, 'A', '日期'), 1, 10, { 1: nd, 2: '', 3: -sign * amt, 4: '', 5: '', 6: '', 7: '', 8: '', 9: '', 10: note });
          addTo(tw.getRange('C3'), sign * amt);
        } else if (a === 'crypto') {
          addTo(sh(CRYPTO).getRange(cryptoCash().row, 3), sign * amt);
          cryptoLog(nd, sign > 0 ? '入金（台幣）' : '出金（台幣）', 'USDT', CRYPTO_CASH[1], '', '', sign * amt, note);
        } else {
          const us = sh('美股');
          appendRow(us, lastRowOf(us, 'A', '日期'), 1, 10, { 1: nd, 2: a, 3: -sign * amt, 4: '', 5: '', 6: '', 7: '', 8: '', 9: '', 10: note });
          addTo(us.getRange('C' + BROKERS[a]), sign * amt);
        }
      });
      group.forEach(g => { if (!touched[g.k + ':' + g.r] && g.k !== 'flow') sync(g); });
      // 改成 0 的帳戶最後刪；同一張表從下往上刪
      dels.sort((a, b) => (a.k < b.k ? -1 : a.k > b.k ? 1 : b.r - a.r)).forEach(g => {
        const L2 = LEDGERS[g.k], s2 = sh(L2.sheet);
        L2.unapply(s2, g.x);
        deleteLedgerRow(s2, g.r, L2.c1, L2.n, lastRowOf(s2, L2.col, L2.header));
      });
      return `已更新這筆資金異動（${sign > 0 ? '投入' : '撤出'} ${fmt(total)}）和各帳戶的金額`;
    }
    if (p.op !== 'edit') throw new Error('不認得的明細動作');
    const d = p.data || {};
    const shift = L.unapply(s, old) || 0;   // 加新股票時，持股清單多一列，交易表會往下移
    row += shift;
    const msg = L.write(s, row, d, old) || '已更新這筆' + L.label;
    // 同一組其他帳戶的金額：App 在資金異動表單一起填的（linkAmounts: [{k, r, amount}]，金額正數，方向照原本）
    const nd = d.date ? toDate(d.date) : null, nn = d.note != null ? String(d.note) : null;
    let synced = 0;
    const done = {};
    (p.linkAmounts || []).forEach(la => {
      const g = group.find(x => x.k === la.k && x.r === num(la.r));
      if (!g) return;
      const amt = need(la.amount, '金額');
      if (amt <= 0) throw new Error('金額要大於 0');
      const L2 = LEDGERS[g.k], s2 = sh(L2.sheet);
      let r2 = g.r;
      if (shift && L2.sheet === L.sheet && r2 > row - shift) r2 += shift;
      const sign = (g.k === 'crypto' ? g.x.usdt : g.x.flow) < 0 ? -1 : 1;
      const date = d.date || g.x.date, note = nn != null ? nn : g.x.note;   // 文字日期，write 會自己轉
      L2.unapply(s2, g.x);
      if (g.k === 'crypto') L2.write(s2, r2, { date, usdt: sign * amt, note }, g.x);
      else if (g.k === 'flow') L2.write(s2, r2, { date, flow: sign * amt, note }, g.x);
      else L2.write(s2, r2, { date, type: sign < 0 ? 'in' : 'out', broker: g.x.b, amount: amt, note }, g.x);
      done[g.k + ':' + g.r] = true;
      synced++;
    });
    // 日期、備註改了，同一組的其他筆一起改
    group.forEach(g => {
      if (done[g.k + ':' + g.r]) return;
      const L2 = LEDGERS[g.k], s2 = sh(L2.sheet);
      let r2 = g.r;
      if (shift && L2.sheet === L.sheet && r2 > row - shift) r2 += shift;
      const changed = (nd && dkey(nd) !== g.x.date) || (nn != null && nn.trim() !== String(g.x.note).trim());
      if (!changed) return;
      if (nd) put(s2.getRange(r2, L2.dateCol), nd);
      if (nn != null) put(s2.getRange(r2, L2.noteCol), nn);
      synced++;
    });
    return synced ? `${msg}；同一次出入金的另外 ${synced} 筆也一起改了` : msg;
  },

  // 台股現價：price 手動填（上櫃等抓不到的股票），auto: true 改回自動抓價（公式）
  price(p) {
    const s = sh('台股'), h = holdings(s).find(x => x.code === normCode(p.code || ''));
    if (!h) throw new Error('台股沒有持有 ' + p.code);
    const cell = twPriceCell(s, h);
    if (p.auto) {
      if (cell.getSheet().getName() !== 'Price') throw new Error(`${h.name} 的現價沒有連到 Price 分頁，請在試算表調整`);
      putF(cell, twPriceFormula(cell.getRow()));
      return `${h.name} 改回自動抓價`;
    }
    const v = need(p.price, '現價');
    if (v <= 0) throw new Error('現價要大於 0');
    put(cell, v);
    return `已更新 ${h.name} 現價 ${v}`;
  },

  // 退休金分頁的自訂退休年齡與試算參數：只能改 B 欄不是公式的格子，寫成文字（試算用 VALUE() 讀）；
  // label 是 A 欄的名稱，用來確認還是同一列。寫完試算表會自己重算，回傳的資料就是新的結果
  retire(p) {
    const s = sh('退休金'), row = num(p.row), v = String(p.value == null ? '' : p.value).trim();
    if (!(row >= 2)) throw new Error('找不到這一列');
    if (v === '') throw new Error('請填數值');
    if (!/^-?[\d,]+(\.\d+)?%?$/.test(v)) throw new Error('請填數字（可以有千分位或 %）');
    const a = String(s.getRange(row, 1).getDisplayValues()[0][0]).trim(), cell = s.getRange(row, 2);
    if (a !== String(p.label || '').trim()) throw new Error('這一列在試算表上已經變動，請下拉更新後再試');
    if (cell.getFormula()) throw new Error('這一格是自動計算的，不能直接改');
    put(fmtCell(cell, '@'), v);
    return `已更新「${a.replace(/（[^）]*可以改[^）]*）/, '')}」`;
  },

  // 閒置現金直接改成某個數字（例如收到股利、手續費和券商對帳）：id 見 cashAccounts
  cash(p) {
    const a = cashAccounts().find(x => x.id === p.id);
    if (!a) throw new Error('找不到這個帳戶');
    put(cashCell(a.id), need(p.amount, '閒置現金'));
    return `已更新 ${a.label} 閒置現金`;
  },

  // 復原上一步：把最後一次從 App 寫入（或修改、刪除）的內容倒回去
  undo() {
    return '已復原：' + undoLast();
  }
};

/* ================================================================
   自動紀錄：每天晚上跑一次；週三記總資產、12/31 記年底結算
   ================================================================ */
// 定時工作都用台灣時間：每天 22:00 看是不是週三（記總資產）；每天 06:00 看是不是 1/1（台股、美股 12/31 都收盤了，記年底）
function ensureTrigger() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('trigger-v2')) return;
  ScriptApp.getProjectTriggers().filter(t => ['dailyJob', 'yearEndJob'].indexOf(t.getHandlerFunction()) >= 0).forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('dailyJob').timeBased().everyDays(1).atHour(22).inTimezone('Asia/Taipei').create();
  ScriptApp.newTrigger('yearEndJob').timeBased().everyDays(1).atHour(6).inTimezone('Asia/Taipei').create();
  props.setProperty('trigger-v2', '1');
}

// 即時價每 10 分鐘自動更新一次（不開 App 時試算表也是新的）
function ensureLiveTrigger() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('trigger-live-v1')) return;
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'liveJob').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('liveJob').timeBased().everyMinutes(10).create();
  props.setProperty('trigger-live-v1', '1');
}
function liveJob() {
  try { etfJob(false); } catch (err) {}   // 0050 每日建議：開盤日 8:25 後第一次跑時做
  refreshLive(true);
}

function dailyJob() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), tz = ss.getSpreadsheetTimeZone();
  const d = new Date(), dow = Utilities.formatDate(d, tz, 'u'), md = Utilities.formatDate(d, tz, 'MM/dd');
  if (dow === '3') weeklyRecord();
}

// 1/1 早上：記前一年 12/31 的年底總資產（當年最後總資產，也是隔年的初始金額）
function yearEndJob() {
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), d = new Date();
  if (Utilities.formatDate(d, tz, 'MM/dd') === '01/01') yearEndRecord(Number(Utilities.formatDate(d, tz, 'yyyy')) - 1);
}

// 每週三：在 Total Overall 的歷史表加一列，記下當天總資產（資金異動空白）
function weeklyRecord() {
  const s = sh('Total Overall'), last = lastRowOf(s, 'A', 'Date');
  const today = toDate(), lastDate = s.getRange(last, 1).getValue();
  if (lastDate instanceof Date && sameDay(lastDate, today)) return;   // 今天已經有紀錄（例如剛記過出入金）
  appendTotalOverall(today, '', '每週自動紀錄');
  log('weekly', {});
}

// 12/31：Plan 填當年實際資產（萬）與實際投入（萬）；Total Overall 年底總結表填當天快照
function yearEndRecord(year) {
  const t = sh('Total Overall'), tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  year = year || Number(Utilities.formatDate(new Date(), tz, 'yyyy'));
  const total = num(t.getRange('G2').getValue());
  // 歷史表加一列 12/31 的年底紀錄（App 用它當隔年 IRR 的初始金額）；已經有那天的紀錄就不重複加
  const dec31 = new Date(year, 11, 31), lastRow = lastRowOf(t, 'A', 'Date'), lastDate = t.getRange(lastRow, 1).getValue();
  if (lastDate instanceof Date && dkey(lastDate) === dkey(dec31)) {
    // 12/31 剛好是週三，晚上已經記過一筆（美股還沒收盤）：更新成收盤後的總資產
    t.getRange(lastRow, 2).setValue(Math.round(total));
    if (!String(t.getRange(lastRow, 10).getValue()).trim() || String(t.getRange(lastRow, 10).getValue()) === '每週自動紀錄') t.getRange(lastRow, 10).setValue('年底紀錄（當年最後總資產）');
  } else appendTotalOverall(dec31, '', '年底紀錄（當年最後總資產）');
  // 當年投入＝歷史表裡當年的資金異動合計（負數是投入）
  const last = lastRowOf(t, 'A', 'Date'), head = findRow(t, 'A', 'Date');
  const rows = t.getRange(head + 1, 1, last - head, 4).getValues();
  const input = -rows.filter(r => r[0] instanceof Date && dkey(r[0]).slice(0, 4) === String(year)).reduce((s, r) => s + (num(r[3]) || 0), 0);
  const plan = planInfo().years.find(y => y.year === year);
  if (plan) {
    const s = sh('Plan');
    s.getRange(plan.row, 7).setValue(Math.round(total / 1000) / 10);
    s.getRange(plan.row, 8).setValue(Math.round(input / 1000) / 10);
  }
  // 年底總結表（K 欄「日期」底下）：找到這一年的列，填快照
  const k = t.getRange('K1:K120').getDisplayValues();
  const i = k.findIndex(r => r[0] === year + '/12/31');
  if (i >= 0) {
    const g = c => num(t.getRange(c).getValue());
    // L 匯率 M 總資產 N IRR O 美股合計 P CUBE Q 小孩美股 R 定期定額美股 S Firstrade T 虛擬幣 U 放貸 V 台股合計 W 我的台股 X 小孩台股
    t.getRange(i + 1, 12, 1, 13).setValues([[g('B2'), total, g('I2'), g('H10'), g('G11'), g('G9'), g('G10'), g('G12'), g('H4'), g('G6'), g('G7') + g('G8'), g('G7'), g('G8')]]);
  }
  log('yearEnd', { year, total, input });
}

/* ================================================================
   試算表結構
   ================================================================ */
// Total Overall 歷史表加一列：A 日期 B 當下總資產 D 資金異動 G 匯率 J 備註（C、E、F、H 是公式，跟著上一列）
function appendTotalOverall(date, flow, note) {
  const s = sh('Total Overall'), head = findRow(s, 'A', 'Date'), last = lastRowOf(s, 'A', 'Date');
  const total = num(s.getRange('G2').getValue()), fx = num(s.getRange('B2').getValue()), first = head + 1;
  appendRow(s, last, 1, 10, { 1: date, 2: Math.round(total), 3: r => `=-SUM($D$${first}:D${r})`, 4: flow, 5: r => `=B${r}-C${r}`,
    6: r => `=E${r}/C${r}`, 7: fx, 8: r => `=B${r}/G${r}`, 10: note });
}

// 在表格最後一列之後加一列。為了讓 Google 表格（Table）自動延伸，先在最後一列「之前」插入，
// 再把原本最後一列搬回去，新資料寫在真正的最後一列。values 用欄號（1 起算）指定要寫的格子。
// 注意：表格裡有指定類型的欄位（例如日期欄）不准改數字格式，所以這裡只複製內容與公式、不碰格式，
// 日期欄會自動套用表格的日期格式。寫入失敗時把插入的列刪掉，不留半套資料。
// values 的值可以是函式 row => 值（例如要寫 =-E153*F153 這種跟列號有關的公式）。
function appendRow(s, last, c1, n, values) {
  s.insertRowBefore(last);
  unsnap(s);
  const row = last + 1;
  // 第一步：把原本的最後一列搬到插入的空白列（失敗就刪掉那列空白列，原本的資料沒動）
  try {
    s.getRange(row, c1, 1, n).copyTo(s.getRange(last, c1, 1, n), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
    SpreadsheetApp.flush();   // 讓試算表現在就回報錯誤，才能在這裡還原
  } catch (err) {
    try { s.deleteRow(last); } catch (e) {}
    throw err;
  }
  // 第二步：新資料寫在真正的最後一列（失敗就刪掉這一列，原本的最後一列已經在上一列）
  try {
    Object.keys(values).forEach(c => { const v = values[c]; s.getRange(row, Number(c)).setValue(typeof v === 'function' ? v(row) : v); });
    SpreadsheetApp.flush();
  } catch (err) {
    try { s.deleteRow(row); } catch (e) {}
    throw err;
  }
  recRow('rowIns', s, row);   // 復原時刪掉這一列，原本的最後一列已經在上一列
  return row;
}

// 某欄標題列之後、第一個空白之前的最後一列
function lastRowOf(s, col, header) {
  const head = findRow(s, col, header);
  const vals = s.getRange(`${col}${head + 1}:${col}${head + 2000}`).getValues();
  let i = 0; while (i < vals.length && vals[i][0] !== '') i++;
  if (!i) throw new Error(`「${s.getName()}」的表格是空的`);
  return head + i;
}

function findRow(s, col, text) {
  const vals = s.cached ? s.getRange(`${col}1:${col}400`).getValues().map(r => [typeof r[0] === 'string' ? r[0] : ''])
    : s.getRange(`${col}1:${col}400`).getDisplayValues();
  const i = vals.findIndex(r => r[0].trim() === text);
  if (i < 0) throw new Error(`「${s.getName()}」找不到「${text}」`);
  return i + 1;
}

// 台股持股：I 股號 J 股名 K 庫存股數 N 現價 O 均價（從第 3 列的標題往下）
function holdings(s) {
  const head = findRow(s, 'I', '股號');
  const vals = s.getRange(head + 1, 9, 30, 8).getValues();
  const out = [];
  for (let i = 0; i < vals.length && vals[i][0] !== ''; i++) {
    out.push({ row: head + 1 + i, code: normCode(vals[i][0]), name: String(vals[i][1]), qty: num(vals[i][2]) || 0, avg: num(vals[i][6]) || 0 });
  }
  return out;
}

// 美股持股：每個券商一段（I 欄寫券商名，下一列是「代號」標題）；J 庫存股數 M 現價 N 均價
function usHoldings(broker) {
  const s = sh('美股'), vals = s.getRange('I1:N60').getValues();
  const b = vals.findIndex(r => String(r[0]).trim() === broker);
  if (b < 0) throw new Error('美股分頁找不到 ' + broker);
  const out = [];
  for (let i = b + 2; i < vals.length && vals[i][0] !== ''; i++) {
    out.push({ row: i + 1, code: String(vals[i][0]).trim().toUpperCase(), qty: num(vals[i][1]) || 0, avg: num(vals[i][5]) || 0 });
  }
  return out;
}

// 新的台股：Price 分頁加一列抓價格，持股清單加一列（公式照上一列）
// 台股現價的公式：先試 TPE:代號，再試前面補 0（00878 寫成 0878），最後試 Yahoo 的上櫃頁面
function twPriceFormula(row) {
  return liveWrap(`"TW:"&A${row}`, `IFERROR(GOOGLEFINANCE("TPE:"&A${row}),IFERROR(GOOGLEFINANCE("TPE:0"&A${row}),IFERROR(VALUE(IMPORTXML("https://tw.stock.yahoo.com/quote/"&A${row}&".TWO","//span[contains(@class,'Fz(32px)')]")),"")))`);
}

// 某一檔台股的現價格子：持股清單 N 欄是 =Price!C5 這種公式就指到 Price 分頁那格，不然就是 N 欄本身
function twPriceCell(s, h) {
  const m = String(s.getRange(h.row, 14).getFormula()).match(/^=\s*Price!\$?C\$?(\d+)\s*$/i);
  return m ? sh('Price').getRange(Number(m[1]), 3) : s.getRange(h.row, 14);
}

// 每一檔台股的現價與是不是手動填的：整批讀（持股清單 N 欄＋Price 分頁 C 欄），不要一檔一檔讀
function twPrices(s, hs) {
  if (!hs.length) return [];
  const r0 = hs[0].row, own = s.getRange(r0, 14, hs[hs.length - 1].row - r0 + 1, 1), ownF = own.getFormulas(), ownV = own.getValues();
  const pr = sh('Price').getRange('C1:C400'), pF = pr.getFormulas(), pV = pr.getValues();
  return hs.map(h => {
    const i = h.row - r0, m = String(ownF[i][0]).match(/^=\s*Price!\$?C\$?(\d+)\s*$/i), k = m ? Number(m[1]) - 1 : -1;
    if (m && k >= pV.length) { const c = twPriceCell(s, h); return { code: h.code, name: h.name, price: num(c.getValue()), manual: !c.getFormula() }; }
    return { code: h.code, name: h.name, price: num(m ? pV[k][0] : ownV[i][0]), manual: !(m ? pF[k][0] : ownF[i][0]) };
  });
}

// manual：手動現價（上櫃等抓不到的），有填就寫數字，不放公式
function addTwHolding(st, price, manual) {
  const p = sh('Price'), pr = firstEmpty(p, 'A', 2);
  putVs(fmtCell(p.getRange(pr, 1, 1, 2), '@'), [[st.code, st.name]]);
  if (manual > 0) put(p.getRange(pr, 3), manual);
  else putF(p.getRange(pr, 3), twPriceFormula(pr));
  const s = sh('台股'), list = holdings(s), last = list[list.length - 1].row;
  s.insertRowBefore(last);
  s.getRange(last + 1, 9, 1, 9).copyTo(s.getRange(last, 9, 1, 9));
  const row = last + 1;
  recRow('rowIns', s, row);
  fmtCell(s.getRange(row, 9, 1, 2), '@').setValues([[st.code, st.name]]);
  s.getRange(row, 14).setFormula(`=Price!C${pr}`);
  s.getRange(row, 15).setValue(price);
  SpreadsheetApp.flush();
  return { row, code: st.code, name: st.name, qty: 0, avg: 0 };
}

function addUsHolding(broker, code, price) {
  const p = sh('Price');
  let pr = p.getRange('E2:E200').getValues().findIndex(r => String(r[0]).trim().toUpperCase() === code) + 2;
  if (pr < 2) {
    pr = firstEmpty(p, 'E', 2);
    put(p.getRange(pr, 5), code);
    putF(p.getRange(pr, 6), liveWrap(`"US:"&UPPER(E${pr})`, `GOOGLEFINANCE("${code}")`));
  }
  const s = sh('美股'), list = usHoldings(broker), last = list[list.length - 1].row;
  s.insertRowBefore(last);
  s.getRange(last + 1, 9, 1, 9).copyTo(s.getRange(last, 9, 1, 9));
  const row = last + 1;
  recRow('rowIns', s, row);
  s.getRange(row, 9).setValue(code);
  s.getRange(row, 13).setFormula(`=Price!$F$${pr}`);
  s.getRange(row, 14).setValue(price);
  SpreadsheetApp.flush();
  return { row, code, qty: 0, avg: 0 };
}

// 股號或股名 → {code, name}：先找持股，再找 Price 分頁；新股票一定要給股號
function resolveTw(q) {
  if (!q) throw new Error('請輸入股號或股名');
  const nq = normCode(q);
  const known = holdings(sh('台股')).map(h => ({ code: h.code, name: h.name }))
    .concat(sh('Price').getRange('A2:B200').getValues().filter(r => r[0] !== '').map(r => ({ code: normCode(r[0]), name: String(r[1]) })));
  const hit = known.find(k => k.code === nq) || known.find(k => k.name === q) || known.find(k => k.name.indexOf(q) >= 0);
  if (hit) return hit;
  if (/^\d{4,6}[A-Z]?$/.test(nq)) return { code: nq, name: nq };
  throw new Error(`找不到「${q}」，第一次買的股票請輸入股號`);
}

// 定期定額：每張表的標題列有「累計股數」；資料列 A 股名、B 每月金額、F 總投入、G 累計股數、H 目前股價
function dcaItems() {
  const out = [];
  const scan = (sheetName, owner) => {
    const s = sh(sheetName), vals = s.getRange('A1:U40').getValues();
    vals.forEach((r, i) => {
      for (let c = 0; c < r.length; c++) {
        if (String(r[c]).trim() !== '累計股數') continue;
        const col = c - 6;   // 這張表的第一欄（股名）：股名、金額/月、起始日、迄今、月數、總投入、累計股數，累計股數往左 6 欄
        const cur = /USD/.test(String(vals[i][col + 1])) ? 'USD' : 'NTD';
        for (let j = i + 1; j < vals.length && vals[j][col] !== '' && String(vals[j][col]).trim() !== 'Total'; j++) {
          const v = vals[j];
          out.push({ id: `${sheetName}!${j + 1}!${col + 1}`, sheet: sheetName, row: j + 1, col: col + 1, owner: tableOwner(vals, i, col, owner),
            name: String(v[col]).trim(), cur, monthly: num(v[col + 1]) || 0, invested: num(v[col + 5]) || 0, shares: num(v[col + 6]) || 0, price: num(v[col + 7]) || 0 });
        }
      }
    });
  };
  scan('My 定期定額 (台幣帳戶)', '我');
  scan('橙&澈 定期定額台股', '');
  return out;
}

// 小孩的兩張表並排：用上面摘要表（Name 欄）的名字依左右順序對上
function tableOwner(vals, headRow, col, fallback) {
  if (fallback) return fallback;
  const names = [];
  for (let i = 0; i < headRow; i++) if (String(vals[i][0]).trim() === 'Name') {
    for (let j = i + 1; j < headRow && vals[j][0] !== '' && String(vals[j][0]).trim() !== 'Total'; j++) names.push(String(vals[j][0]).trim());
  }
  return names[col === 0 ? 0 : 1] || '';
}

function expenseItems() {
  const s = sh('固定支出'), out = [];
  EXP_BLOCKS.forEach(b => {
    const vals = s.getRange(b.from, 1, b.to - b.from + 1, 12).getValues();
    b.months.forEach((m, k) => {
      const items = vals.map(r => ({ name: String(r[k * 2]).trim(), amount: num(r[k * 2 + 1]) })).filter(x => x.name || x.amount);
      out.push({ month: m, max: b.to - b.from + 1, items });
    });
  });
  return out;
}

function planInfo() {
  const s = sh('Plan'), vals = s.getRange('A1:I60').getValues();
  const head = vals.findIndex(r => String(r[1]).trim() === 'Year');
  const years = [];
  for (let i = head + 1; i < vals.length && num(vals[i][1]); i++) {
    years.push({ row: i + 1, year: num(vals[i][1]), input: vals[i][5] === '' ? null : num(vals[i][5]), actual: vals[i][7] === '' ? null : num(vals[i][7]), total: vals[i][6] === '' ? null : num(vals[i][6]) });
  }
  const f = s.getRange(head + 2, 6, years.length, 1).getFormulas();
  years.forEach((y, i) => { y.custom = !f[i][0]; });
  return { rate: num(vals[1][2]), yearly: num(vals[2][5]), years };
}

function loanRows() {
  const s = sh('House'), vals = s.getRange('A1:J20').getValues();
  const head = vals.findIndex(r => String(r[0]).trim() === '房貸');
  if (head < 0) return [];
  const out = [];
  for (let i = head + 1; i < vals.length && vals[i][0] !== ''; i++) out.push({ row: i + 1, name: String(vals[i][0]).trim() });
  return out;
}

// BB：A 欄有年份的是年度列；之後沒有年份的列屬於同一年
function careerInfo() {
  const s = sh('BB'), vals = s.getRange('A1:R80').getValues(), years = [];
  vals.forEach((r, i) => {
    const y = num(r[0]);
    if (y > 1990 && y < 2100 && Math.floor(y) === y) years.push({ year: y, row: i + 1, lastRow: i + 1, title: String(r[1]).replace(/^↑\s*/, ''), salary: num(r[12]) || num(r[5]), lastSalaryRow: i + 1 });
    else if (years.length && r.some(v => v !== '')) {
      const cur = years[years.length - 1];
      cur.lastRow = i + 1;
      if (num(r[12])) { cur.salary = num(r[12]); cur.lastSalaryRow = i + 1; }
    }
  });
  return { years };
}

// 加密貨幣分頁：第 4 列起 A 幣名 B 存於 C 顆數（手打）D 現價 E 總價USD F 總價NTD
function cryptoRows() {
  const vals = sh(CRYPTO).getRange('A4:C200').getValues(), out = [];
  vals.forEach((r, i) => { if (String(r[0]).trim()) out.push({ row: i + 4, name: normCode(r[0]), where: String(r[1]).trim(), qty: num(r[2]) || 0 }); });
  return out;
}

function cryptoCash() {
  const hit = cryptoRows().find(x => x.name === CRYPTO_CASH[0] && x.where === CRYPTO_CASH[1]);
  return hit || addCryptoRow(CRYPTO_CASH[0], CRYPTO_CASH[1], null);
}

// 新的幣（或新的存放位置）：持有清單加一列；價格接 Price 分頁 H/I 欄（GOOGLEFINANCE 抓，抓不到的幣用成交價）
function addCryptoRow(coin, where, price) {
  const s = sh(CRYPTO), row = firstEmpty(s, 'A', 4);
  let priceCell = 1;
  if (coin !== CRYPTO_CASH[0]) {
    const p = sh('Price');
    let pr = p.getRange('H2:H100').getValues().findIndex(r => normCode(r[0]) === coin) + 2;
    if (pr < 2) {
      pr = firstEmpty(p, 'H', 2);
      put(p.getRange(pr, 8), coin);
      putF(p.getRange(pr, 9), liveWrap(`"C:"&UPPER(H${pr})`, `IFERROR(GOOGLEFINANCE("CURRENCY:"&H${pr}&"USD"),${price})`));
    }
    priceCell = `=Price!I${pr}`;
  }
  putVs(s.getRange(row, 1, 1, 3), [[coin, where, 0]]);
  put(s.getRange(row, 4), priceCell);
  putF(s.getRange(row, 5), `=D${row}*C${row}`);
  putF(s.getRange(row, 6), `=E${row}*$J$2`);
  SpreadsheetApp.flush();
  return { row, name: coin, where, qty: 0 };
}

// 加密貨幣買賣：用 Binance 的 USDT 付款／收款
function cryptoTrade(p, date, price, qty, buy, amt) {
  const coin = normCode(p.stock || ''), where = String(p.broker || CRYPTO_CASH[1]).trim();
  if (!/^[A-Z0-9]{2,10}$/.test(coin)) throw new Error('請輸入幣名，例如 BTC');
  if (coin === CRYPTO_CASH[0]) throw new Error('USDT 請用「出入金」');
  const s = sh(CRYPTO), cash = cryptoCash();
  let h = cryptoRows().find(x => x.name === coin && x.where === where);
  if (!h) { if (!buy) throw new Error(`${where} 沒有 ${coin}`); h = addCryptoRow(coin, where, price); }
  if (!buy && qty > h.qty + 1e-12) throw new Error(`${where} 的 ${coin} 只有 ${h.qty} 顆，不能賣 ${qty} 顆`);
  put(s.getRange(h.row, 3), Math.round((h.qty + (buy ? qty : -qty)) * 1e8) / 1e8);
  addTo(s.getRange(cash.row, 3), buy ? -amt : amt);
  cryptoLog(date, buy ? '買進' : '賣出', coin, where, price, qty, buy ? -amt : amt, p.note || '');
  return `已記錄加密貨幣${buy ? '買進' : '賣出'} ${coin} ${qty} 顆`;
}

// 加密貨幣分頁 Y～AF 欄的交易紀錄（L～U 欄是原本的已兌現表和備註，不要動）
function cryptoLog(date, type, coin, where, price, qty, usdt, note) {
  const s = sh(CRYPTO), row = firstEmpty(s, 'Y', 2);
  put(fmtCell(s.getRange(row, 25), 'yyyy/mm/dd'), date);
  putVs(s.getRange(row, 26, 1, 7), [[type, coin, where, price, qty, round2(usdt), note]]);
}

// House 上半部的房產列（「房貸」那列以前）：A 名稱 K 估計車位（萬）L 估計單價（萬/坪）
function houseRows() {
  const vals = sh('House').getRange('A1:L20').getValues(), out = [];
  for (let i = 2; i < vals.length && String(vals[i][0]).trim() !== '房貸'; i++) {
    if (String(vals[i][0]).trim()) out.push({ row: i + 1, name: String(vals[i][0]).trim(), unit: num(vals[i][11]), park: num(vals[i][10]) });
  }
  return out;
}

/* ================================================================
   即時價：台股用證交所公開報價、美股用 Finnhub（要金鑰）、加密貨幣用 Kraken
   抓到的價格寫在隱藏分頁「App 即時價」（代號、現價、昨收、抓取時間）。
   Price 分頁的價格公式會包一層：有即時價就用即時價，沒有就用原本的公式（GOOGLEFINANCE 等）；超過一小時沒抓到的即時價會被清掉。
   每 10 分鐘自動更新；App 先讀到資料後，價格超過 2 分鐘沒更新會再另外要求抓一次（不擋住讀取和寫入）。
   Finnhub 金鑰放在 Apps Script「專案設定 → 指令碼屬性」的 FINNHUB_KEY，不寫在程式碼裡。
   ================================================================ */
const LIVE_SHEET = 'App 即時價';

// 價格公式外面包一層即時價：keyExpr 是算出代號的公式片段，orig 是原本的公式（不含 =）
function liveWrap(keyExpr, orig) {
  const ref = `'${LIVE_SHEET}'!$A:$D`;
  // 不在公式裡用 NOW() 判斷新舊：NOW() 會讓整份試算表每次有任何變動都重算。太舊的即時價改由 refreshLive 清掉
  return `=LET(k,${keyExpr},v,IFERROR(VLOOKUP(k,${ref},2,FALSE),""),IF(AND(ISNUMBER(v),v>0),v,${orig}))`;
}

function liveSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(LIVE_SHEET);
  if (!s) {
    s = ss.insertSheet(LIVE_SHEET);
    s.getRange(1, 1, 1, 4).setValues([['代號（App 用，請勿修改）', '現價', '昨收', '抓取時間']]);
    s.hideSheet();
  }
  return s;
}

// { 'TW:2330': {price, prev, at(毫秒)} }
function liveRows() {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LIVE_SHEET), out = {};
  if (!s) return out;
  s.getDataRange().getValues().slice(1).forEach(r => {
    if (String(r[0]) !== '') out[String(r[0])] = { price: num(r[1]), prev: num(r[2]), at: r[3] instanceof Date ? r[3].getTime() : 0 };
  });
  return out;
}

// maxAge：上次抓的時間在這麼多毫秒以內就不重抓（沒給就是 1 分鐘）
function refreshLive(force, maxAge) {
  const props = PropertiesService.getScriptProperties();
  if (!force && Date.now() - Number(props.getProperty('live-at') || 0) < (maxAge || 60 * 1000)) return;
  props.setProperty('live-at', String(Date.now()));
  const p = sh('Price'), uniq = a => a.filter((x, i) => x && a.indexOf(x) === i);
  const tw = uniq(p.getRange('A2:A200').getValues().map(r => String(r[0]).trim()));
  const us = uniq(p.getRange('E2:E200').getValues().map(r => normCode(r[0])));
  const cr = uniq(p.getRange('H2:H100').getValues().map(r => normCode(r[0]))).filter(c => c !== CRYPTO_CASH[0]);
  let rows = [];
  try { rows = rows.concat(liveTw(tw)); } catch (err) {}
  try { rows = rows.concat(liveUs(us, props.getProperty('FINNHUB_KEY'))); } catch (err) {}
  try { rows = rows.concat(liveCrypto(cr)); } catch (err) {}
  const s = liveSheet(), now = new Date();
  const last = s.getLastRow(), keys = last > 1 ? s.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0])) : [];
  const vals = last > 1 ? s.getRange(2, 1, last - 1, 4).getValues() : [];
  rows.forEach(x => {
    const i = keys.indexOf(x.key), row = [x.key, x.price, x.prev || '', now];
    if (i >= 0) vals[i] = row; else { keys.push(x.key); vals.push(row); }
  });
  // 這次沒抓到、而且超過一小時沒更新的：把現價清掉，Price 的公式就會退回原本的 GOOGLEFINANCE
  vals.forEach(r => { if (r[3] !== now && !(r[3] instanceof Date && now.getTime() - r[3].getTime() < 3600 * 1000)) r[1] = ''; });
  if (vals.length) s.getRange(2, 1, vals.length, 4).setValues(vals);
  wrapPriceFormulas();
}

// 證交所 MIS：每個代號同時查上市（tse）和上櫃（otc）；試算表裡 00878 寫成 0878，也補 0 查
function liveTw(codes) {
  if (!codes.length) return [];
  const want = [];
  codes.forEach(c => { [c].concat(/^0\d{3}$/.test(c) ? ['0' + c] : []).forEach(v => { want.push(`tse_${v}.tw`, `otc_${v}.tw`); }); });
  const out = {};
  for (let i = 0; i < want.length; i += 40) {
    const url = 'https://mis.twse.com.tw/stock/api/getStockInfo.jsp?json=1&delay=0&ex_ch=' + encodeURIComponent(want.slice(i, i + 40).join('|'));
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' } });
    if (res.getResponseCode() !== 200) continue;
    (JSON.parse(res.getContentText()).msgArray || []).forEach(m => {
      if (!m.c) return;
      const code = codes.find(c => c === m.c || '0' + c === m.c);
      if (!code) return;
      const bid = num(String(m.b || '').split('_')[0]);
      const price = num(m.z) || bid || num(m.y);   // 剛好沒成交時 z 是 -，用買進價，再不行用昨收
      if (price > 0) out[code] = { key: 'TW:' + code, price, prev: num(m.y) };
    });
  }
  return Object.keys(out).map(k => out[k]);
}

// Finnhub：每個代號一次；沒有金鑰就不抓（維持原本的 GOOGLEFINANCE）
function liveUs(codes, key) {
  if (!key || !codes.length) return [];
  const res = UrlFetchApp.fetchAll(codes.map(c => ({ url: `https://finnhub.io/api/v1/quote?symbol=${encodeURIComponent(c)}&token=${encodeURIComponent(key)}`, muteHttpExceptions: true })));
  const out = [];
  res.forEach((r, i) => {
    try { const j = JSON.parse(r.getContentText()); if (num(j.c) > 0) out.push({ key: 'US:' + codes[i], price: num(j.c), prev: num(j.pc) }); } catch (err) {}
  });
  return out;
}

// Kraken：現價 c[0]；「昨收」用當天開盤價 o（加密貨幣沒有收盤）
function liveCrypto(coins) {
  if (!coins.length) return [];
  const res = UrlFetchApp.fetchAll(coins.map(c => ({ url: 'https://api.kraken.com/0/public/Ticker?pair=' + c + 'USD', muteHttpExceptions: true })));
  const out = [];
  res.forEach((r, i) => {
    try {
      const j = JSON.parse(r.getContentText()), k = j.result && Object.keys(j.result)[0];
      if (k && num(j.result[k].c[0]) > 0) out.push({ key: 'C:' + coins[i], price: num(j.result[k].c[0]), prev: num(j.result[k].o) });
    } catch (err) {}
  });
  return out;
}

// Price 分頁的價格公式還沒包即時價的，包上去（手動填的數字不動）
function wrapPriceFormulas() {
  const p = sh('Price');
  [['A', 3, 'TW:', c => `A${c}`], ['E', 6, 'US:', c => `UPPER(E${c})`], ['H', 9, 'C:', c => `UPPER(H${c})`]].forEach(([col, pc, prefix, keyOf]) => {
    const codes = p.getRange(`${col}2:${col}200`).getValues(), f = p.getRange(2, pc, 199, 1).getFormulas();
    codes.forEach((r, i) => {
      const row = i + 2, old = f[i][0];
      if (!String(r[0]).trim() || !old) return;
      if (old.indexOf(LIVE_SHEET) >= 0) {
        // 舊版包的公式裡有 NOW()：改成沒有 NOW() 的寫法（只做一次）
        if (!/NOW\(\)/i.test(old)) return;
        const fixed = old.replace(/,\s*t\s*,\s*IFERROR\(VLOOKUP\(k,[^()]*,4,FALSE\),0\)/i, '').replace(/,\s*NOW\(\)\s*-\s*t\s*<\s*1\/24/i, '');
        if (fixed !== old && !/NOW\(\)/i.test(fixed) && !/[(,]t[,)]/.test(fixed.slice(0, fixed.indexOf('IF(AND(')))) p.getRange(row, pc).setFormula(fixed);
        return;
      }
      p.getRange(row, pc).setFormula(liveWrap(`"${prefix}"&${keyOf(row)}`, old.replace(/^=/, '')));
    });
  });
}

function liveInfo() {
  const props = PropertiesService.getScriptProperties();
  return { at: Number(props.getProperty('live-at') || 0), finnhub: !!props.getProperty('FINNHUB_KEY') };
}

/* ================================================================
   今日漲跌：隱藏的「App 漲跌」分頁，用 GOOGLEFINANCE 抓 Price 分頁每一檔的現價與昨收
   回傳 { 'TW:2330': 0.012, 'US:QQQ': -0.004, 'C:BTC': 0.02 }（漲跌幅）；App 再用各帳戶持股現值算金額
   第一次加進來的代號要等 GOOGLEFINANCE 抓到價格，下一次打開 App 才會有
   ================================================================ */
const DAILY_SHEET = 'App 漲跌';
function dailyChanges() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), p = sh('Price'), keys = [];
  p.getRange('A2:A200').getValues().forEach(r => { const c = String(r[0]).trim(); if (c) keys.push('TW:' + c); });
  p.getRange('E2:E200').getValues().forEach(r => { const c = normCode(r[0]); if (c) keys.push('US:' + c); });
  p.getRange('H2:H100').getValues().forEach(r => { const c = normCode(r[0]); if (c && c !== CRYPTO_CASH[0]) keys.push('C:' + c); });
  let s = ss.getSheetByName(DAILY_SHEET);
  if (!s) {
    s = ss.insertSheet(DAILY_SHEET);
    s.getRange(1, 1, 1, 3).setValues([['代號（App 用，請勿修改）', '現價', '昨收']]);
    s.hideSheet();
  }
  let all = s.getDataRange().getValues();
  const have = all.slice(1).map(r => String(r[0]));
  const add = keys.filter((k, i) => keys.indexOf(k) === i && have.indexOf(k) < 0);
  if (add.length) {
    // 台股：先試 TPE:代號，不行再補一個 0（試算表裡 00878 寫成 0878）；美股直接用代號；加密貨幣用前一天的歷史價格當昨收
    const f = (k, r) => {
      const code = `MID(A${r},${k.indexOf(':') + 2},20)`;
      if (k.indexOf('TW:') === 0) return [`=IFERROR(GOOGLEFINANCE("TPE:"&${code}),IFERROR(GOOGLEFINANCE("TPE:0"&${code}),""))`,
        `=IFERROR(GOOGLEFINANCE("TPE:"&${code},"closeyest"),IFERROR(GOOGLEFINANCE("TPE:0"&${code},"closeyest"),""))`];
      if (k.indexOf('US:') === 0) return [`=IFERROR(GOOGLEFINANCE(${code}),"")`, `=IFERROR(GOOGLEFINANCE(${code},"closeyest"),"")`];
      return [`=IFERROR(GOOGLEFINANCE("CURRENCY:"&${code}&"USD"),"")`, `=IFERROR(INDEX(GOOGLEFINANCE("CURRENCY:"&${code}&"USD","price",TODAY()-1),2,2),"")`];
    };
    const start = s.getLastRow() + 1;
    s.getRange(start, 1, add.length, 3).setValues(add.map((k, i) => [k].concat(f(k, start + i))));
    SpreadsheetApp.flush();
    all = s.getDataRange().getValues();
  }
  const out = {};
  all.slice(1).forEach(r => { const now = num(r[1]), prev = num(r[2]); if (now > 0 && prev > 0) out[String(r[0])] = now / prev - 1; });
  // 即時價（證交所、Finnhub、Kraken）一小時內抓過的，優先用它的現價和昨收
  const live = liveRows(), fresh = Date.now() - 3600 * 1000;
  Object.keys(live).forEach(k => { const x = live[k]; if (x.at >= fresh && x.price > 0 && x.prev > 0) out[k] = x.price / x.prev - 1; });
  return out;
}

/* ================================================================
   明細的修改與刪除：資金異動（Total Overall）、台股交易表、美股交易表、加密貨幣交易紀錄
   每種明細：在哪個分頁、哪一欄有標題、從第幾欄起佔幾欄；
   read 讀一列；unapply 把這筆對閒置現金、均價、顆數的影響拿掉；write 寫入新內容並套用新的影響
   ================================================================ */
const LEDGERS = {
  flow: {
    sheet: 'Total Overall', col: 'A', header: 'Date', c1: 1, n: 10, label: '資金異動', dateCol: 1, noteCol: 10,
    parse: v => ({ date: dkey(v[0]), key: num(v[3]) || 0, flow: num(v[3]), fx: num(v[6]), note: String(v[9]) }),
    isFlow: x => !!x.flow,
    unapply() { return 0; },
    write(s, row, d) {
      put(s.getRange(row, 1), toDate(d.date));
      put(s.getRange(row, 4), need(d.flow, '金額'));
      put(s.getRange(row, 10), String(d.note || ''));
      return '已更新這筆資金異動';
    }
  },
  tw: {
    sheet: '台股', col: 'A', header: '日期', c1: 1, n: 10, label: '台股明細', dateCol: 1, noteCol: 10,
    parse: tradeParse, isFlow: x => !!x.flow && !x.code,
    unapply(s, old) { return tradeUnapply('tw', s, old); },
    write(s, row, d, old) { return tradeWrite('tw', s, row, d, old); }
  },
  us: {
    sheet: '美股', col: 'A', header: '日期', c1: 1, n: 10, label: '美股明細', dateCol: 1, noteCol: 10,
    parse: tradeParse, isFlow: x => !!x.flow && !x.code,
    unapply(s, old) { return tradeUnapply('us', s, old); },
    write(s, row, d, old) { return tradeWrite('us', s, row, d, old); }
  },
  crypto: {
    sheet: CRYPTO, col: 'Y', header: '日期', c1: 25, n: 8, label: '加密貨幣紀錄', dateCol: 25, noteCol: 32,
    parse: v => ({ date: dkey(v[0]), type: String(v[1]).trim(), coin: normCode(v[2]), where: String(v[3]).trim(), price: num(v[4]),
      qty: num(v[5]) || 0, usdt: num(v[6]) || 0, key: num(v[6]) || 0, note: String(v[7]) }),
    isFlow: x => x.type !== '買進' && x.type !== '賣出',
    unapply(s, old) { cryptoEffect(s, old, -1); return 0; },
    write(s, row, d, old) {
      const date = toDate(d.date), note = String(d.note || '');
      if (old.type === '買進' || old.type === '賣出') {
        const buy = d.side ? d.side === 'buy' : old.type === '買進';
        const coin = normCode(d.coin || old.coin), where = String(d.where || old.where).trim();
        const price = need(d.price, '單價'), qty = need(d.qty, '顆數');
        if (price <= 0 || qty <= 0) throw new Error('單價和顆數要大於 0');
        const amt = round2(price * qty), rec = { type: buy ? '買進' : '賣出', coin, where, qty, usdt: buy ? -amt : amt };
        cryptoEffect(s, rec, 1, price);
        putVs(s.getRange(row, 25, 1, 8), [[date, rec.type, coin, where, price, qty, rec.usdt, note]]);
      } else {
        const usdt = need(d.usdt, 'USDT 數量');
        cryptoEffect(s, { type: old.type, usdt }, 1);
        putVs(s.getRange(row, 25, 1, 8), [[date, old.type, old.coin, old.where, '', '', usdt, note]]);
      }
      return '已更新這筆加密貨幣紀錄';
    }
  }
};

// 台股、美股交易表的一列：A 日期 B 股名（美股是券商）C 出入金 D 股號 E 價格 F 買入股數 G 賣出股數 H 買入金額 I 賣出金額 J 備註
function tradeParse(v) {
  return { date: dkey(v[0]), b: String(v[1]).trim(), flow: num(v[2]), code: v[3] === '' ? '' : normCode(v[3]), price: num(v[4]),
    buy: num(v[5]), sell: num(v[6]), h: num(v[7]), i: num(v[8]), note: String(v[9]),
    key: num(v[2]) || num(v[5]) || num(v[6]) || 0 };
}

function tradeCash(m, s, broker) {
  if (m === 'tw') return s.getRange('C3');
  return BROKERS[broker] ? s.getRange('C' + BROKERS[broker]) : null;
}

// 讀一列明細
function ledgerRead(k, s, row) {
  const L = LEDGERS[k];
  return L.parse(s.getRange(row, L.c1, 1, L.n).getValues()[0]);
}

// 同一次「成本變更」或「加密貨幣出入金」會在資金異動、台股、美股、加密貨幣各記一筆出入金，日期和備註相同。
// 找出跟這筆同一組的其他筆（備註空白的不算，避免誤抓）
function linkedRows(kind, row, x) {
  const note = String(x.note || '').trim();
  if (!LEDGERS[kind].isFlow(x) || !note) return [];
  const out = [];
  Object.keys(LEDGERS).forEach(k => {
    const L = LEDGERS[k], s = sh(L.sheet);
    let head, last;
    try { head = findRow(s, L.col, L.header); last = lastRowOf(s, L.col, L.header); } catch (err) { return; }
    s.getRange(head + 1, L.c1, last - head, L.n).getValues().forEach((v, i) => {
      const r = head + 1 + i, y = L.parse(v);
      if (k === kind && r === row) return;
      if (y.date === x.date && String(y.note).trim() === note && L.isFlow(y)) out.push({ k, r, x: y });
    });
  });
  return out;
}

// 拿掉一列交易的影響：閒置現金退回；買進的話均價倒算回買之前（庫存股數是公式，刪掉或改掉那列就會自己更新）
function tradeUnapply(m, s, old) {
  const cash = tradeCash(m, s, old.b);
  const eff = -(old.flow || 0) + (old.h || 0) + (old.i || 0);   // 這列讓閒置現金變多少
  if (cash && eff) addTo(cash, -eff);
  if (old.buy && old.price && old.code) {
    const list = m === 'tw' ? holdings(s) : (BROKERS[old.b] ? usHoldings(old.b) : []);
    const h = list.find(x => x.code === old.code);
    if (h) {
      const rest = h.qty - old.buy;
      put(s.getRange(h.row, m === 'tw' ? 15 : 14), rest > 0 ? round2((h.avg * h.qty - old.price * old.buy) / rest) : h.avg);
    }
  }
  return 0;
}

// d: {date, type: 'buy'|'sell'|'in'|'out', stock, broker, price, qty, amount, note}
function tradeWrite(m, s, row, d, old) {
  const tw = m === 'tw', date = toDate(d.date), note = String(d.note || '');
  const b = tw ? '' : String(d.broker || '');
  if (!tw && !BROKERS[b]) throw new Error('請選券商');
  if (d.type === 'in' || d.type === 'out') {
    const amt = need(d.amount, '金額');
    if (amt <= 0) throw new Error('金額要大於 0');
    const flow = d.type === 'in' ? -amt : amt;   // 試算表的寫法：負數是存進這個帳戶
    putVs(s.getRange(row, 1, 1, 10), [[date, b, flow, '', '', '', '', '', '', note]]);
    addTo(tradeCash(m, s, b), -flow);
    return `已更新這筆${tw ? '台股' : '美股'}${d.type === 'in' ? '入金' : '出金'}`;
  }
  const price = need(d.price, '單價'), qty = need(d.qty, '股數'), buy = d.type === 'buy';
  if (price <= 0 || qty <= 0) throw new Error('單價和股數要大於 0');
  let st, h;
  if (tw) { st = resolveTw(String(d.stock || '').trim()); h = holdings(s).find(x => x.code === st.code); }
  else {
    const code = normCode(d.stock || '');
    if (!/^[A-Z][A-Z0-9.\-]*$/.test(code)) throw new Error('美股請輸入代號，例如 QQQ');
    st = { code, name: code }; h = usHoldings(b).find(x => x.code === code);
  }
  if (!h) {
    if (!buy) throw new Error('沒有持有 ' + st.code);
    h = tw ? addTwHolding(st, price) : addUsHolding(b, st.code, price);
    row += 1;   // 持股清單多一列，下面的交易表整個往下移一列
  }
  // 庫存股數是公式，目前還算著這列的舊內容；先扣掉舊的才是「不含這筆」的庫存
  const same = old && old.code === st.code && (tw || old.b === b);
  const held = h.qty - (same ? (old.buy || 0) - (old.sell || 0) : 0);
  if (!buy && qty > held + 1e-9) throw new Error(`${st.name} 只有 ${held} 股，不能賣 ${qty} 股`);
  if (buy) setAvg(s.getRange(h.row, tw ? 15 : 14), h.avg, held, price, qty);
  putVs(s.getRange(row, 1, 1, 10), [[date, tw ? st.name : b, '', st.code, price, buy ? qty : '', buy ? '' : qty,
    buy ? `=-E${row}*F${row}` : '', buy ? '' : `=G${row}*E${row}`, note]]);
  addTo(tradeCash(m, s, b), buy ? -round2(price * qty) : round2(price * qty));
  return `已更新這筆${tw ? '台股' : '美股'}${buy ? '買進' : '賣出'}`;
}

// 加密貨幣紀錄的影響：買賣改顆數與 USDT；出入金只改 USDT。sign 1 套用、-1 拿掉
function cryptoEffect(s, rec, sign, price) {
  if (rec.type === '買進' || rec.type === '賣出') {
    const dq = (rec.type === '買進' ? rec.qty : -rec.qty) * sign;
    let h = cryptoRows().find(x => x.name === rec.coin && x.where === rec.where);
    if (!h && sign > 0) {
      if (dq < 0) throw new Error(`${rec.where} 沒有 ${rec.coin}`);
      h = addCryptoRow(rec.coin, rec.where, price);
    }
    if (h) {
      const nq = Math.round((h.qty + dq) * 1e8) / 1e8;
      if (nq < -1e-9) throw new Error(`${rec.where} 的 ${rec.coin} 只有 ${h.qty} 顆`);
      put(s.getRange(h.row, 3), nq);
    }
  }
  if (rec.usdt) addTo(s.getRange(cryptoCash().row, 3), sign * rec.usdt);
}

// 刪掉明細的一列：只動這張表的欄（c1 起 n 欄），下面的往上補；
// 從 A 欄開始的表（資金異動、台股、美股）最後一列整列空了才刪整列；加密貨幣紀錄旁邊還有別的表，只清空不刪列
function deleteLedgerRow(s, row, c1, n, last) {
  snap(s.getRange(row, c1, last - row + 1, n));
  if (row < last) s.getRange(row + 1, c1, last - row, n).copyTo(s.getRange(row, c1, last - row, n), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
  s.getRange(last, c1, 1, n).clearContent();
  SpreadsheetApp.flush();
  if (c1 === 1 && s.getRange(last, 1, 1, s.getLastColumn()).getValues()[0].every(v => v === '')) { s.deleteRow(last); recRow('rowDel', s, last); }
}

// 各帳戶的閒置現金：id、名稱、幣別、在 App 哪一頁、目前金額
// 小孩美股帳戶從「橙&澈 美股」的 Name 表讀名字（D 欄是現金餘額 USD），不寫死在程式裡
function cashAccounts() {
  const out = [], tw = sh('台股'), us = sh('美股');
  out.push({ id: 'tw', label: '台股 CUBE', cur: 'NTD', page: 'tw', value: num(tw.getRange('C3').getValue()) });
  Object.keys(BROKERS).forEach(b => out.push({ id: 'us:' + b, label: '美股 ' + b, cur: 'USD', page: 'us', value: num(us.getRange('C' + BROKERS[b]).getValue()) }));
  try {
    const k = sh('橙&澈 美股'), h = findRow(k, 'A', 'Name'), v = k.getRange(h + 1, 1, 10, 4).getValues();
    for (let i = 0; i < v.length && v[i][0] !== '' && String(v[i][0]).trim() !== 'Total'; i++) {
      out.push({ id: 'kid:' + (h + 1 + i), label: String(v[i][0]).trim() + ' 美股', cur: 'USD', page: 'more/kids', value: num(v[i][3]) });
    }
  } catch (err) {}
  const c = cryptoRows().find(x => x.name === CRYPTO_CASH[0] && x.where === CRYPTO_CASH[1]);
  out.push({ id: 'usdt', label: '加密貨幣 Binance', cur: 'USDT', page: 'crypto', value: c ? c.qty : 0 });
  return out;
}

function cashCell(id) {
  if (id === 'tw') return sh('台股').getRange('C3');
  if (id.indexOf('us:') === 0 && BROKERS[id.slice(3)]) return sh('美股').getRange('C' + BROKERS[id.slice(3)]);
  if (id.indexOf('kid:') === 0) return sh('橙&澈 美股').getRange(Number(id.slice(4)), 4);
  if (id === 'usdt') return sh(CRYPTO).getRange(cryptoCash().row, 3);
  throw new Error('找不到這個帳戶');
}

// 日期統一成 yyyy/MM/dd 字串，用來比對 App 看到的是不是同一筆
function dkey(v) {
  if (v instanceof Date) return Utilities.formatDate(v, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy/MM/dd');
  const m = String(v == null ? '' : v).match(/(\d{4})\D(\d{1,2})\D(\d{1,2})/);
  return m ? `${m[1]}/${('0' + m[2]).slice(-2)}/${('0' + m[3]).slice(-2)}` : String(v == null ? '' : v);
}

/* ================================================================
   復原：每次從 App 寫入時，先記下會被改到的格子（原本的值或公式）和插入、刪除的列，
   存在隱藏的「App 復原」分頁；「復原上一步」照相反順序倒回去。寫到一半失敗也用同一份資料自動還原。
   ================================================================ */
const UNDO_SHEET = 'App 復原';
const UNDO_KEEP = 30;   // 最多保留幾步
let REC = null;         // 這次動作的紀錄 {desc, ops}；null 表示不記（例如自動紀錄、復原本身）

// 復原要記下每一格寫入前的內容。一格一格去讀的話，每次都是「寫了再讀」，試算表要整份重算一次（好幾秒）；
// 所以每個分頁在這次動作第一次要記的時候整張讀一次（SNAP），之後都從記憶體拿。
// 同一格在這次動作裡寫了兩次，第二次記到的仍是動作前的內容，倒回去的結果一樣是動作前的樣子。
// 插入或刪除列之後列號會變，所以 recRow 會丟掉那個分頁的 SNAP，下次再重讀。
// SNAP 掛在這次動作的紀錄（REC.snap）上，換一個動作就是新的，不會拿到上一個動作留下的舊資料
function snapGrid(s) {
  const name = s.getName(), m = REC.snap || (REC.snap = {});
  if (!m[name]) { const d = s.getDataRange(); m[name] = { v: d.getValues(), f: d.getFormulas() }; }
  return m[name];
}
function unsnap(s) { if (REC && REC.snap) delete REC.snap[s.getName()]; }
function snap(range) {
  if (REC) {
    const g = snapGrid(range.getSheet()), r0 = range.getRow(), c0 = range.getColumn(), nr = range.getNumRows(), nc = range.getNumColumns(), d = [];
    for (let i = 0; i < nr; i++) {
      const fr = g.f[r0 - 1 + i] || [], vr = g.v[r0 - 1 + i] || [], row = [];
      for (let j = 0; j < nc; j++) {
        const f = fr[c0 - 1 + j], x = vr[c0 - 1 + j];
        row.push(f || (x instanceof Date ? { $d: x.getTime() } : x === undefined ? '' : x));
      }
      d.push(row);
    }
    REC.ops.push({ t: 'cells', sh: range.getSheet().getName(), r: r0, c: c0, d });
  }
  return range;
}
function recRow(t, s, row) { unsnap(s); if (REC) REC.ops.push({ t, sh: s.getName(), row }); }
function put(range, v) { snap(range); range.setValue(v); return range; }
function putF(range, f) { snap(range); range.setFormula(f); return range; }
function putVs(range, vals) { snap(range); range.setValues(vals); return range; }

function applyUndo(ops) {
  if (REC) REC.snap = {};
  ops.slice().reverse().forEach(op => {
    const t = sh(op.sh);
    if (op.t === 'cells') t.getRange(op.r, op.c, op.d.length, op.d[0].length).setValues(op.d.map(r => r.map(x => (x && x.$d != null ? new Date(x.$d) : x))));
    else if (op.t === 'rowIns') t.deleteRow(op.row);
    else if (op.t === 'rowDel') t.insertRowBefore(op.row);
  });
}

function undoSheet(create) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(UNDO_SHEET);
  if (!s && create) {
    s = ss.insertSheet(UNDO_SHEET);
    s.getRange(1, 1, 1, 3).setValues([['時間', '說明', '還原資料（App 用，請勿修改）']]);
    s.setFrozenRows(1);
    s.hideSheet();
  }
  return s;
}

function recSave() {
  if (!REC || !REC.ops.length) { REC = null; return; }
  const s = undoSheet(true), j = JSON.stringify(REC.ops), parts = [];
  // 一格最多 5 萬字，切段存；每段前面加 ~，避免開頭是 = 被當成公式
  for (let i = 0; i < j.length; i += 40000) parts.push('~' + j.slice(i, i + 40000));
  const row = s.getLastRow() + 1;
  s.getRange(row, 1, 1, 2 + parts.length).setValues([[now(), REC.desc].concat(parts)]);
  const extra = row - 1 - UNDO_KEEP;
  if (extra > 0) s.deleteRows(2, extra);
  REC = null;
}

function undoLast() {
  const s = undoSheet(false);
  if (!s || s.getLastRow() < 2) throw new Error('沒有可以復原的動作');
  const r = s.getLastRow(), v = s.getRange(r, 1, 1, s.getLastColumn()).getValues()[0];
  const ops = JSON.parse(v.slice(2).filter(x => x !== '').map(x => String(x).replace(/^~/, '')).join(''));
  applyUndo(ops);
  s.deleteRow(r);
  return v[1];
}

function undoInfo() {
  const s = undoSheet(false);
  if (!s) return { count: 0 };
  const rows = s.getRange(1, 1, UNDO_KEEP + 10, 2).getDisplayValues();
  let r = rows.length; while (r > 0 && rows[r - 1][0] === '' && rows[r - 1][1] === '') r--;
  if (r < 2) return { count: 0 };
  return { count: r - 1, time: rows[r - 1][0], desc: rows[r - 1][1] };
}

/* ================================================================
   小工具
   ================================================================ */
function sh(name) {
  const s = RC ? cachedSheet(name) : SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!s) throw new Error('找不到分頁：' + name);
  return s;
}

/* ================================================================
   讀取用的記憶體分頁：整張的值只讀一次，getRange(...).getValues() 從記憶體切出來（超出範圍補空白）。
   公式和顯示文字用到的地方很少，照樣去讀真的試算表。只在 readAll 期間使用（RC 不是 null），寫入一律用真的分頁。
   ================================================================ */
let RC = null;

/* ================================================================
   一次讀完：用 Google Sheets 進階服務（Apps Script 編輯器左邊「服務」→ 加入 Google Sheets API）
   一次把要用的分頁全部讀回來。一個分頁一個分頁讀，每次大約 0.3～0.4 秒；一次讀完只要一兩次連線。
   沒有加入這個服務就照原本的方式讀，功能一樣，只是比較慢。
   進階服務回傳的日期是數字（序號），所以讀兩次：一次要數字、一次要日期變文字，
   「第一次是數字、第二次是文字」的格子就是日期，換成和原本讀法一樣的 Date。
   每一版第一次用的時候會和原本的讀法整份比對，不一樣就退回原本的讀法（App 設定頁會顯示原因）。
   ================================================================ */
let BG = null;       // { 分頁名稱: 二維陣列 }；null 表示這次不用
let BATCH = 'off';   // 回報給 App：'off' 沒加入服務、'on' 使用中、其他是退回原本讀法的原因

function batchGrids(names) {
  BATCH = 'off';
  if (typeof Sheets === 'undefined' || !Sheets.Spreadsheets) return null;
  const props = PropertiesService.getScriptProperties(), key = 'batch-v' + SCRIPT_VERSION, st = props.getProperty(key) || '';
  if (st.indexOf('不一致') === 0) { BATCH = st; return null; }
  try {
    const grids = batchRead(names);
    if (st !== '1') {
      const diff = batchDiff(grids);
      if (diff) {
        // 價格剛好在兩次讀取之間跳動也會不一樣，所以連續三次都不一樣才停用
        const n = (Number(st.slice(1)) || 0) + 1;
        props.setProperty(key, n >= 3 ? ('不一致：' + diff).slice(0, 300) : 'f' + n);
        BATCH = '不一致：' + diff;
        return null;
      }
      props.setProperty(key, '1');
    }
    BATCH = 'on';
    return grids;
  } catch (err) {
    BATCH = '錯誤：' + String(err.message || err).slice(0, 200);
    return null;
  }
}

function batchRead(names) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), tz = ss.getSpreadsheetTimeZone(), id = ss.getId();
  const have = ss.getSheets().map(s => s.getName());
  const want = names.filter((n, i) => names.indexOf(n) === i && have.indexOf(n) >= 0);
  const ranges = want.map(n => "'" + n.replace(/'/g, "''") + "'");
  const get = dt => Sheets.Spreadsheets.Values.batchGet(id, { ranges, valueRenderOption: 'UNFORMATTED_VALUE', dateTimeRenderOption: dt }).valueRanges;
  const nums = get('SERIAL_NUMBER'), strs = get('FORMATTED_STRING'), out = {}, offs = {};
  want.forEach((n, i) => {
    const a = nums[i].values || [], b = strs[i].values || [];
    out[n] = a.map((row, r) => row.map((v, c) => {
      if (typeof v !== 'number') return v == null ? '' : v;
      return typeof (b[r] || [])[c] === 'string' ? serialToDate(v, tz, offs) : v;
    }));
  });
  return out;
}

// 試算表的日期序號（1899/12/30 起算的天數，試算表時區的時間）→ Date
function serialToDate(serial, tz, offs) {
  const wall = Math.round((serial - 25569) * 86400000), day = Math.floor(serial);
  if (offs[day] === undefined) {
    const m = String(Utilities.formatDate(new Date(wall), tz, 'Z')).match(/^([+-])(\d\d):?(\d\d)$/);
    offs[day] = m ? (m[1] === '-' ? -1 : 1) * (Number(m[2]) * 3600000 + Number(m[3]) * 60000) : 0;
  }
  return new Date(wall - offs[day]);
}

// 和原本的讀法比對，回傳第一個不一樣的地方（都一樣回傳空字串）。數字差 5% 以內算一樣（價格會跳動）
function batchDiff(grids) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), isD = v => Object.prototype.toString.call(v) === '[object Date]';
  const show = v => (isD(v) ? '日期 ' + v.toISOString() : typeof v + ' ' + String(v)).slice(0, 40);
  const same = (x, y) => {
    if (isD(x) || isD(y)) return isD(x) && isD(y) && Math.abs(x.getTime() - y.getTime()) < 1000;
    if (typeof x === 'number' && typeof y === 'number') return Math.abs(x - y) <= 0.05 * Math.max(Math.abs(x), Math.abs(y)) + 1e-9;
    return x === y;
  };
  for (const name of Object.keys(grids)) {
    const a = grids[name], b = ss.getSheetByName(name).getDataRange().getValues();
    const rows = Math.max(a.length, b.length);
    for (let r = 0; r < rows; r++) {
      const ra = a[r] || [], rb = b[r] || [], cols = Math.max(ra.length, rb.length);
      for (let c = 0; c < cols; c++) {
        const x = ra[c] === undefined ? '' : ra[c], y = rb[c] === undefined ? '' : rb[c];
        if (!same(x, y)) return `${name} 第 ${r + 1} 列第 ${c + 1} 欄：${show(x)} ≠ ${show(y)}`;
      }
    }
  }
  return '';
}

function cachedSheet(name) {
  if (name in RC) return RC[name];
  const real = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!real) return (RC[name] = null);
  const grid = BG && BG[name] ? BG[name] : real.getDataRange().getValues();
  const a1 = a => {
    const m = String(a).match(/^\$?([A-Za-z]+)\$?(\d+)(?::\$?([A-Za-z]+)\$?(\d+))?$/);
    if (!m) throw new Error('看不懂的範圍：' + a);
    const col = t => t.toUpperCase().split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
    const r = Number(m[2]), c = col(m[1]);
    return m[3] ? [r, c, Number(m[4]) - r + 1, col(m[3]) - c + 1] : [r, c, 1, 1];
  };
  const sheet = {
    cached: true,
    getName: () => name,
    getLastRow: () => (grid.length === 1 && grid[0].every(v => v === '') ? 0 : grid.length),
    getRange(r, c, nr, nc) {
      if (typeof r === 'string') [r, c, nr, nc] = a1(r);
      nr = nr || 1; nc = nc || 1;
      const realRange = () => real.getRange(r, c, nr, nc);
      const getValues = () => {
        const out = [];
        for (let i = 0; i < nr; i++) {
          const src = grid[r - 1 + i] || [], row = [];
          for (let j = 0; j < nc; j++) { const v = src[c - 1 + j]; row.push(v === undefined ? '' : v); }
          out.push(row);
        }
        return out;
      };
      return { getValues, getValue: () => getValues()[0][0],
        getFormulas: () => realRange().getFormulas(), getFormula: () => realRange().getFormula(),
        getDisplayValues: () => realRange().getDisplayValues(),
        getRow: () => r, getColumn: () => c, getSheet: () => real };
    }
  };
  return (RC[name] = sheet);
}

function num(v) {
  if (typeof v === 'number') return v;
  const s = String(v == null ? '' : v).replace(/[,\s]|NT\$|US\$|\$/g, '');
  if (s === '') return null;
  const n = Number(s.replace(/%$/, ''));
  if (isNaN(n)) return null;
  return /%$/.test(s) ? n / 100 : n;
}

function need(v, label, text) {
  if (text) { if (String(v == null ? '' : v).trim() === '') throw new Error('請填' + label); return String(v).trim(); }
  const n = num(v);
  if (n == null) throw new Error('請填' + label);
  return n;
}

function normCode(v) { return String(v).trim().toUpperCase(); }

function toDate(s) {
  if (!s) { const d = new Date(); return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
  const m = String(s).match(/^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})/);
  if (!m) throw new Error('日期格式要像 2026/11/05');
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

function sameDay(a, b) { return a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate(); }
function round2(n) { return Math.round(n * 100) / 100; }
function fmt(n) { return Math.round(n).toLocaleString('en-US'); }

// 均價（移動平均）：買進後＝（原均價×原股數＋買價×買股數）÷ 新股數；賣出不改均價
function setAvg(cell, avg, held, price, qty) {
  const total = held + qty;
  put(cell, round2(held > 0 ? (avg * held + price * qty) / total : price));
}

// 現金這類手打的數字：在原本的值上加減（格子是公式就不動）
function addTo(cell, delta) {
  if (cell.getFormula()) return;
  put(cell, round2((num(cell.getValue()) || 0) + delta));
}

// 設定數字格式；Google 表格（Table）裡有指定類型的欄位不准改格式，那種欄位本來就會自動套格式，略過即可
function fmtCell(range, f) {
  if (REC) snapGrid(range.getSheet());
  try { range.setNumberFormat(f); } catch (err) {}
  return range;
}

function firstEmpty(s, col, from) {
  const vals = s.getRange(`${col}${from}:${col}${from + 500}`).getValues();
  let last = -1;
  vals.forEach((r, i) => { if (r[0] !== '') last = i; });
  return from + last + 1;
}

function log(action, payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(LOG_SHEET);
  if (!s) { s = ss.insertSheet(LOG_SHEET); s.appendRow(['時間', '動作', '內容']); s.setFrozenRows(1); }
  s.appendRow([now(), action, JSON.stringify(payload || {})]);
}

function now() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy/MM/dd HH:mm');
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ================================================================
   人壽保單：資料在另一份試算表（雲端硬碟裡的保單試算表），App「固定支出 → 人壽保單」讀寫它。
   那份試算表的 ID 存在指令碼屬性 POLICY_ID（第一次在 App 貼上網址就會記下來），不寫在程式碼裡。
   分頁：人壽保單（每張保單一列）、保障與理賠（每個險種一列）、繳費日期（每張保單一列）、理賠總表（只讀）。
   欄位用第一列的標題找，不靠第幾欄；App 只有打開保單頁時才讀，平常更新不會多花時間。
   保單的修改不走「復原上一步」（那個只管這份試算表）；寫入前會確認那一列還是同一張保單。
   ================================================================ */
// acc 意外險：一年一約，每年續保新增一列（保單年度、續保自＝前一年的保單號碼）
const POLICY_TABS = { list: '人壽保單', cover: '保障與理賠', pay: '繳費日期', sum: '理賠總表', acc: '意外險' };
// 保單號碼只留英數、去掉開頭的 0（「全球人壽 0027108250」和 27108250 算同一張）
const policyNo = v => String(v == null ? '' : v).replace(/[^0-9A-Za-z]/g, '').replace(/^0+/, '');
// 比對用的號碼：整個號碼，和最後一段（「0123-15ABC00001」在繳費日期寫成「15ABC00001」也算同一張）
const policyKeys = no => [policyNo(no), policyNo(String(no == null ? '' : no).split(/[^0-9A-Za-z]+/).filter(Boolean).pop())].filter(Boolean);
const policyTok = (t, no) => !!t && policyKeys(no).indexOf(policyNo(t)) >= 0;
// 一段文字裡（例如繳費日期的「遠雄 1020210385」）有沒有這個保單號碼
const policyIn = (text, no) => String(text == null ? '' : text).split(/[^0-9A-Za-z]+/).some(t => policyTok(t, no));

function policyBook() {
  const id = PropertiesService.getScriptProperties().getProperty('POLICY_ID');
  if (!id) return null;
  try { return SpreadsheetApp.openById(id); }
  catch (err) {
    if (/permission|授權|authoriz/i.test(String(err.message || err))) throw new Error('Apps Script 還沒有權限開啟保單試算表：請到 Apps Script 編輯器選 authorizePolicy 按「執行」並允許，再重新部署一次');
    throw new Error('打不開保單試算表：' + String(err.message || err));
  }
}
// 在 Apps Script 編輯器手動執行一次，用來取得開啟另一份試算表的權限
function authorizePolicy() { const b = policyBook(); return b ? b.getName() : '尚未連結保單試算表'; }

function policyRead() {
  const ss = policyBook();
  if (!ss) return { linked: false };
  const out = { linked: true, url: ss.getUrl(), title: ss.getName(), at: now() };
  Object.keys(POLICY_TABS).forEach(k => {
    const s = ss.getSheetByName(POLICY_TABS[k]), rows = s ? s.getDataRange().getDisplayValues() : [];
    let last = rows.length;
    while (last > 0 && rows[last - 1].every(v => v === '')) last--;
    out[k] = rows.slice(0, last);
  });
  return out;
}

// 寫一列裡的幾個欄位：values = { 標題: 新內容 }。原本是數字的格子，新內容像數字就寫數字（保留千分位格式）
function policyWrite(s, row, values, skip) {
  const head = s.getRange(1, 1, 1, s.getLastColumn()).getDisplayValues()[0].map(h => String(h).trim());
  Object.keys(values).forEach(h => {
    if (skip.indexOf(h) >= 0) return;
    const c = head.indexOf(h) + 1;
    if (!c) throw new Error(`「${s.getName()}」找不到「${h}」這一欄`);
    const cell = s.getRange(row, c), v = String(values[h] == null ? '' : values[h]).trim();
    cell.setValue(/^-?[\d,]+(\.\d+)?$/.test(v) && (typeof cell.getValue() === 'number' || /保費|金額/.test(h)) ? Number(v.replace(/,/g, '')) : v);
  });
}

const POLICY_ACTIONS = {
  // 連結保單試算表：貼網址或 ID
  link(p) {
    const m = String(p.url || '').match(/\/d\/([A-Za-z0-9_-]{20,})/) || String(p.url || '').trim().match(/^([A-Za-z0-9_-]{20,})$/);
    if (!m) throw new Error('請貼上保單試算表的網址');
    let ss;
    try { ss = SpreadsheetApp.openById(m[1]); }
    catch (err) { throw new Error(/permission|授權|authoriz/i.test(String(err.message || err)) ? 'Apps Script 還沒有權限開啟別的試算表：請到 Apps Script 編輯器選 authorizePolicy 按「執行」並允許，再重新部署一次' : '打不開這份試算表，請確認網址和權限'); }
    if (!ss.getSheetByName(POLICY_TABS.list)) throw new Error(`這份試算表沒有「${POLICY_TABS.list}」分頁`);
    PropertiesService.getScriptProperties().setProperty('POLICY_ID', m[1]);
    return `已連結「${ss.getName()}」`;
  },

  // 修改一列：tab 'list'（人壽保單）或 'cover'（保障與理賠）；row 列號、no 保單號碼、name 險種（cover 用）是用來確認還是同一列
  edit(p) {
    const ss = policyBook();
    if (!ss) throw new Error('還沒有連結保單試算表');
    if (p.tab !== 'list' && p.tab !== 'cover' && p.tab !== 'acc') throw new Error('不認得的分頁');
    const s = ss.getSheetByName(POLICY_TABS[p.tab]), row = num(p.row), values = p.values || {};
    if (!s || !(row >= 2)) throw new Error('找不到這一列');
    if (!Object.keys(values).length) throw new Error('沒有要修改的內容');
    const head = s.getRange(1, 1, 1, s.getLastColumn()).getDisplayValues()[0].map(h => String(h).trim());
    const cur = s.getRange(row, 1, 1, head.length).getDisplayValues()[0], at = h => String(cur[head.indexOf(h)] || '').trim();
    if (!policyNo(p.no) || policyNo(at('保單號碼')) !== policyNo(p.no) || (p.tab === 'cover' && at('險種') !== String(p.name || '').trim())) {
      throw new Error('這一列在試算表上已經變動，請按「重新整理」後再試');
    }
    policyWrite(s, row, values, ['被保險人', '保險公司', '保單號碼']);
    // 繳費日、保費、扣款方式有改：繳費日期分頁同一張保單的那一列也一起改
    if (p.tab === 'list' || p.tab === 'acc') policyPaySync(ss, p.no, values, null);
    if (p.tab === 'acc' && values['年繳保費'] != null) accExpenseSync(at('被保險人'), values['年繳保費']);
    return p.tab === 'cover' ? `已更新「${String(p.name || '').trim()}」` : '已更新保單';
  },

  // 意外險續保：照前一年那一列（row、no 確認）新增下一年的一列；繳費日期分頁那一列換成新保單；固定支出的金額一起改
  accRenew(p) {
    const ss = policyBook();
    if (!ss) throw new Error('還沒有連結保單試算表');
    const s = ss.getSheetByName(POLICY_TABS.acc), row = num(p.row), values = p.values || {};
    if (!s || !(row >= 2)) throw new Error('找不到前一年的意外險');
    const head = s.getRange(1, 1, 1, s.getLastColumn()).getDisplayValues()[0].map(h => String(h).trim());
    const cur = s.getRange(row, 1, 1, head.length).getDisplayValues()[0], at = h => String(cur[head.indexOf(h)] || '').trim();
    if (!policyNo(p.no) || policyNo(at('保單號碼')) !== policyNo(p.no)) throw new Error('這一列在試算表上已經變動，請按「重新整理」後再試');
    const newNo = String(values['保單號碼'] || '').trim();
    if (!policyNo(newNo)) throw new Error('請填新的保單號碼');
    if (policyNo(newNo) === policyNo(p.no)) throw new Error('新的保單號碼和前一年一樣');
    const all = s.getDataRange().getDisplayValues(), ic = head.indexOf('保單號碼');
    if (all.some((r, i) => i > 0 && policyNo(r[ic]) === policyNo(newNo))) throw new Error('這個保單號碼已經登記過了');
    // 新的一列：沒填的欄位沿用前一年（被保險人、保險公司一定沿用），續保自＝前一年的保單號碼
    const out = {};
    head.forEach(h => { if (h) out[h] = at(h); });
    Object.keys(values).forEach(h => { if (h in out && ['被保險人', '保險公司'].indexOf(h) < 0) out[h] = String(values[h] == null ? '' : values[h]).trim(); });
    if ('續保自' in out) out['續保自'] = at('保單號碼');
    const at2 = s.getLastRow() + 1;
    s.getRange(row, 1, 1, head.length).copyFormatToRange(s, 1, head.length, at2, at2);
    head.forEach((h, c) => {
      if (!h) return;
      const v = out[h];
      s.getRange(at2, c + 1).setValue(/^-?[\d,]+(\.\d+)?$/.test(v) && /保費|年度/.test(h) ? Number(v.replace(/,/g, '')) : v);
    });
    policyPaySync(ss, p.no, out, newNo);
    accExpenseSync(out['被保險人'], out['年繳保費']);
    return `已新增 ${out['保單年度'] || ''} 年度意外險`;
  }
};

// 繳費日期分頁：找到「項目」裡有這張保單號碼的那一列，同步日期／月份／金額／扣款方式；newNo 有給就把項目裡的號碼換成新的
function policyPaySync(ss, no, values, newNo) {
  const pay = ss.getSheetByName(POLICY_TABS.pay), sync = {};
  if (!pay) return;
  if (values['每年繳費日'] != null) {
    sync['日期'] = values['每年繳費日'];
    const mm = String(values['每年繳費日']).match(/^\s*(\d{1,2})\s*[\/月]/);
    if (mm) sync['月份'] = Number(mm[1]) + '月';
  }
  if (values['年繳保費'] != null) sync['金額'] = values['年繳保費'];
  if (values['繳費方式'] != null) sync['扣款方式'] = values['繳費方式'];
  const g = pay.getDataRange().getDisplayValues(), ph = g[0].map(h => String(h).trim()), ic = ph.indexOf('項目');
  const r = ic < 0 ? -1 : g.findIndex((x, i) => i > 0 && policyIn(x[ic], no));
  if (r <= 0) return;
  if (newNo) {
    const old = String(g[r][ic]), toks = old.split(/([^0-9A-Za-z]+)/);
    sync['項目'] = toks.map(t => (policyTok(t, no) ? String(newNo).replace(/^0*\d+-/, '') : t)).join('');
  }
  Object.keys(sync).forEach(h => { if (ph.indexOf(h) < 0) delete sync[h]; });
  if (Object.keys(sync).length) policyWrite(pay, r + 1, sync, []);
}

// 固定支出（這份投資試算表）：名稱有「意外」又有被保險人名字（取最後兩個字）的項目，金額改成新保費
function accExpenseSync(person, amount) {
  const amt = num(String(amount || '').replace(/,/g, '')), nm = String(person || '').trim().slice(-2);
  if (!amt || !nm) return;
  let s;
  try { s = sh('固定支出'); } catch (e) { return; }
  EXP_BLOCKS.forEach(b => {
    const rg = s.getRange(b.from, 1, b.to - b.from + 1, 12), v = rg.getValues();
    v.forEach((r, i) => b.months.forEach((m, k) => {
      const name = String(r[k * 2] || '');
      if (/意外/.test(name) && name.indexOf(nm) >= 0) s.getRange(b.from + i, k * 2 + 2).setValue(amt);
    }));
  });
}

// 保單的讀寫：不讀投資的資料，回得比較快
function policyRequest(body) {
  try {
    let message = '';
    if (body.policyAction) {
      const lock = LockService.getScriptLock();
      lock.waitLock(20000);
      try {
        const fn = POLICY_ACTIONS[body.policyAction];
        if (!fn) throw new Error('不認得的動作：' + body.policyAction);
        message = fn(body.payload || {});
        SpreadsheetApp.flush();
        log('policy.' + body.policyAction, body.payload);
      } finally { lock.releaseLock(); }
    }
    return json({ ok: true, message, policy: policyRead(), ms: msOut() });
  } catch (err) {
    return json({ ok: false, error: String(err.message || err) });
  }
}

// ===== 0050 長期投資（由 etf/build.py 從 etf/*.js 合併，要改請改 etf/ 裡的檔案）=====
// ---- etf/engine.js ----
// 0050 長期投資：每日建議的計算規則（純計算，不含任何個人資料；資料由呼叫的人傳進來）
// 規則來自 2026-11 的回測討論：
//   定期定額每月 5 號 2 萬永不停；每年單筆預算 30～50 萬，核心 60% 在 1 月投入，其餘是預備金；
//   預備金只在四種訊號出現時加碼，每次 1 份＝預備金 ÷ 3，兩次至少隔 10 個交易日；
//   只加碼不賣；預備金可滾到明年但最多一年份預算，超過的部分 12 月最後一個交易日投入。
// 「實際投入」用試算表的紀錄算，所以沒照建議買（或買得比建議多）時，隔天的建議會跟著調整。

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

// ---- etf/webpush.js ----
// Web Push（RFC 8291 aes128gcm＋RFC 8292 VAPID）純 JavaScript 版，給 Apps Script 用（它沒有 ECDH、AES-GCM）。
// P-256 用 BigInt；AES-128-GCM 自己實作；HMAC／SHA-256／亂數由呼叫的人提供（Apps Script 用 Utilities）。
// 位元組一律用 0～255 的一般陣列。

const WP_P = 0xffffffff00000001000000000000000000000000ffffffffffffffffffffffffn;
const WP_N = 0xffffffff00000000ffffffffffffffffbce6faada7179e84f3b9cac2fc632551n;
const WP_B = 0x5ac635d8aa3a93e7b3ebbd55769886bc651d06b0cc53b0f63bce3c3e27d2604bn;
const WP_G = [0x6b17d1f2e12c4247f8bce6e563a440f277037d812deb33a0f4a13945d898c296n, 0x4fe342e2fe1a7f9b8ee7eb4a7c0f9e162bce33576b315ececbb6406837bf51f5n];

const wpMod = (a, m) => { const r = a % m; return r < 0n ? r + m : r; };
function wpInv(a, m) { // 擴展歐幾里得
  let [r0, r1, s0, s1] = [wpMod(a, m), m, 1n, 0n];
  while (r1) { const q = r0 / r1; [r0, r1] = [r1, r0 - q * r1]; [s0, s1] = [s1, s0 - q * s1]; }
  return wpMod(s0, m);
}
// Jacobian 座標的點：[X, Y, Z]；無窮遠點 Z = 0
function wpDbl(Pt) {
  const [X, Y, Z] = Pt, p = WP_P;
  if (!Z || !Y) return [0n, 1n, 0n];
  const YY = Y * Y % p, S = 4n * X * YY % p, ZZ = Z * Z % p;
  const M = 3n * (X - ZZ) * (X + ZZ) % p;   // a = -3
  const X3 = wpMod(M * M - 2n * S, p), Y3 = wpMod(M * (S - X3) - 8n * YY * YY, p), Z3 = 2n * Y * Z % p;
  return [X3, Y3, Z3];
}
function wpAdd(P1, P2) {
  if (!P1[2]) return P2; if (!P2[2]) return P1;
  const p = WP_P, [X1, Y1, Z1] = P1, [X2, Y2, Z2] = P2;
  const Z1Z1 = Z1 * Z1 % p, Z2Z2 = Z2 * Z2 % p;
  const U1 = X1 * Z2Z2 % p, U2 = X2 * Z1Z1 % p, S1 = Y1 * Z2 * Z2Z2 % p, S2 = Y2 * Z1 * Z1Z1 % p;
  if (U1 === U2) return S1 === S2 ? wpDbl(P1) : [0n, 1n, 0n];
  const H = wpMod(U2 - U1, p), R = wpMod(S2 - S1, p), HH = H * H % p, HHH = H * HH % p, V = U1 * HH % p;
  const X3 = wpMod(R * R - HHH - 2n * V, p), Y3 = wpMod(R * (V - X3) - S1 * HHH, p), Z3 = Z1 * Z2 % p * H % p;
  return [X3, Y3, Z3];
}
function wpMul(k, Pt) {
  let R = [0n, 1n, 0n], Q = [Pt[0], Pt[1], 1n];
  for (let i = BigInt(k.toString(2).length) - 1n; i >= 0n; i--) { R = wpDbl(R); if ((k >> i) & 1n) R = wpAdd(R, Q); }
  if (!R[2]) throw new Error('point at infinity');
  const zi = wpInv(R[2], WP_P), zi2 = zi * zi % WP_P;
  return [R[0] * zi2 % WP_P, R[1] * zi2 % WP_P * zi % WP_P];
}
const wpToBig = b => b.reduce((t, x) => (t << 8n) | BigInt(x & 255), 0n);
const wpFromBig = (n, len) => { const out = []; for (let i = 0; i < len; i++) { out.unshift(Number(n & 255n)); n >>= 8n; } return out; };
const wpPointBytes = Pt => [4].concat(wpFromBig(Pt[0], 32), wpFromBig(Pt[1], 32));
function wpPoint(bytes) {
  if (bytes.length !== 65 || bytes[0] !== 4) throw new Error('bad public key');
  const x = wpToBig(bytes.slice(1, 33)), y = wpToBig(bytes.slice(33));
  if (wpMod(y * y - (x * x * x - 3n * x + WP_B), WP_P) !== 0n) throw new Error('point not on curve');
  return [x, y];
}
function wpScalar(random) { // 1 ≤ k < n
  for (;;) { const k = wpToBig(random(32)); if (k > 0n && k < WP_N) return k; }
}

/* ---------- AES-128 與 GCM ---------- */
const WP_SBOX = (() => {
  const s = new Array(256), inv = new Array(256); let p = 1, q = 1;
  do {
    p = p ^ ((p << 1) & 255) ^ (p & 128 ? 0x1b : 0);
    q ^= q << 1; q ^= q << 2; q ^= q << 4; q &= 255; if (q & 128) q ^= 0x09;
    const x = q ^ ((q << 1 | q >> 7) & 255) ^ ((q << 2 | q >> 6) & 255) ^ ((q << 3 | q >> 5) & 255) ^ ((q << 4 | q >> 4) & 255) ^ 0x63;
    s[p] = x;
  } while (p !== 1);
  s[0] = 0x63; return s;
})();
const wpX2 = b => ((b << 1) ^ (b & 128 ? 0x1b : 0)) & 255;
function wpAesKey(key) {
  const w = key.slice(); let rc = 1;
  for (let i = 16; i < 176; i += 4) {
    let t = w.slice(i - 4, i);
    if (i % 16 === 0) { t = [WP_SBOX[t[1]] ^ rc, WP_SBOX[t[2]], WP_SBOX[t[3]], WP_SBOX[t[0]]]; rc = wpX2(rc); }
    for (let j = 0; j < 4; j++) w[i + j] = w[i - 16 + j] ^ t[j];
  }
  return w;
}
function wpAesBlock(w, input) {
  let s = input.map((b, i) => b ^ w[i]);
  for (let r = 1; r <= 10; r++) {
    const t = s.map(b => WP_SBOX[b]);
    const sh = new Array(16);
    for (let c = 0; c < 4; c++) for (let rr = 0; rr < 4; rr++) sh[c * 4 + rr] = t[((c + rr) % 4) * 4 + rr];
    if (r < 10) for (let c = 0; c < 4; c++) {
      const a = sh.slice(c * 4, c * 4 + 4), x = a[0] ^ a[1] ^ a[2] ^ a[3];
      for (let k = 0; k < 4; k++) sh[c * 4 + k] = a[k] ^ x ^ wpX2(a[k] ^ a[(k + 1) % 4]);
    }
    s = sh.map((b, i) => b ^ w[r * 16 + i]);
  }
  return s;
}
function wpGmul(X, Y) { // GF(2^128)，GCM 的位元順序
  const R = 0xe1n << 120n; let Z = 0n, V = Y;
  for (let i = 127; i >= 0; i--) { if ((X >> BigInt(i)) & 1n) Z ^= V; V = V & 1n ? (V >> 1n) ^ R : V >> 1n; }
  return Z;
}
function wpGcmEncrypt(key, iv, pt) { // 沒有 AAD；回傳 密文＋16 位元組 tag
  const w = wpAesKey(key), H = wpToBig(wpAesBlock(w, new Array(16).fill(0)));
  const J0 = iv.concat([0, 0, 0, 1]), ct = [];
  for (let i = 0, ctr = 2; i < pt.length; i += 16, ctr++) {
    const cb = iv.concat(wpFromBig(BigInt(ctr), 4)), ks = wpAesBlock(w, cb);
    for (let j = 0; j < 16 && i + j < pt.length; j++) ct.push(pt[i + j] ^ ks[j]);
  }
  let g = 0n;
  for (let i = 0; i < ct.length; i += 16) { const blk = ct.slice(i, i + 16); while (blk.length < 16) blk.push(0); g = wpGmul(g ^ wpToBig(blk), H); }
  g = wpGmul(g ^ BigInt(ct.length * 8), H);   // len(A)=0 || len(C)
  const tag = wpFromBig(g, 16).map((b, i) => b ^ wpAesBlock(w, J0)[i]);
  return ct.concat(tag);
}

/* ---------- HKDF、編碼 ---------- */
const wpUtf8 = s => Array.from(unescape(encodeURIComponent(s)), c => c.charCodeAt(0));
function wpHkdf(prim, salt, ikm, info, len) {
  const prk = prim.hmac(salt, ikm);
  return prim.hmac(prk, info.concat([1])).slice(0, len);   // len ≤ 32，一輪就夠
}
const WP_B64 = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_';
function wpB64u(bytes) {
  let s = '';
  for (let i = 0; i < bytes.length; i += 3) {
    const n = (bytes[i] << 16) | ((bytes[i + 1] || 0) << 8) | (bytes[i + 2] || 0);
    s += WP_B64[n >> 18 & 63] + WP_B64[n >> 12 & 63] + (i + 1 < bytes.length ? WP_B64[n >> 6 & 63] : '') + (i + 2 < bytes.length ? WP_B64[n & 63] : '');
  }
  return s;
}
function wpUnb64u(s) {
  s = String(s).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
  const out = []; let buf = 0, bits = 0;
  for (const ch of s) { const v = WP_B64.indexOf(ch); if (v < 0) throw new Error('bad base64'); buf = (buf << 6) | v; bits += 6; if (bits >= 8) { bits -= 8; out.push((buf >> bits) & 255); } }
  return out;
}

/* ---------- 對外 ---------- */
// 產生 VAPID 金鑰（第一次用時呼叫一次，存起來）
function wpVapidKeys(prim) {
  const d = wpScalar(prim.random);
  return { publicKey: wpB64u(wpPointBytes(wpMul(d, WP_G))), privateKey: wpB64u(wpFromBig(d, 32)) };
}
function wpVapidAuth(endpoint, vapid, prim, subject) {
  const aud = endpoint.match(/^https?:\/\/[^/]+/)[0];
  const head = wpB64u(wpUtf8(JSON.stringify({ typ: 'JWT', alg: 'ES256' })));
  const claims = wpB64u(wpUtf8(JSON.stringify({ aud, exp: Math.floor(Date.now() / 1000) + 12 * 3600, sub: subject })));
  const z = wpToBig(prim.sha256(wpUtf8(head + '.' + claims))) % WP_N, d = wpToBig(wpUnb64u(vapid.privateKey));
  let r, s;
  do {
    const k = wpScalar(prim.random);
    r = wpMul(k, WP_G)[0] % WP_N;
    s = wpInv(k, WP_N) * (z + r * d) % WP_N;
  } while (!r || !s);
  return `vapid t=${head}.${claims}.${wpB64u(wpFromBig(r, 32).concat(wpFromBig(s, 32)))}, k=${vapid.publicKey}`;
}
// sub = {endpoint, p256dh, auth}；回傳要 POST 的 body（位元組陣列）
function wpEncrypt(sub, payload, prim) {
  const ua = wpUnb64u(sub.p256dh), auth = wpUnb64u(sub.auth), uaPt = wpPoint(ua);
  const e = wpScalar(prim.random), as = wpPointBytes(wpMul(e, WP_G));
  const shared = wpFromBig(wpMul(e, uaPt)[0], 32);
  const ikm = wpHkdf(prim, auth, shared, wpUtf8('WebPush: info\0').concat(ua, as), 32);
  const salt = prim.random(16);
  const cek = wpHkdf(prim, salt, ikm, wpUtf8('Content-Encoding: aes128gcm\0'), 16);
  const nonce = wpHkdf(prim, salt, ikm, wpUtf8('Content-Encoding: nonce\0'), 12);
  const ct = wpGcmEncrypt(cek, nonce, wpUtf8(payload).concat([2]));
  return salt.concat([0, 0, 16, 0], [65], as, ct);   // rs = 4096
}
function wpRequest(sub, payload, vapid, prim, subject) {
  return {
    url: sub.endpoint, method: 'post', contentType: 'application/octet-stream', muteHttpExceptions: true,
    headers: { Authorization: wpVapidAuth(sub.endpoint, vapid, prim, subject), 'Content-Encoding': 'aes128gcm', TTL: '43200', Urgency: 'high' },
    payload: wpEncrypt(sub, payload, prim)
  };
}

// ---- etf/gas.js ----
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
// ===== 0050 長期投資 結束 =====
