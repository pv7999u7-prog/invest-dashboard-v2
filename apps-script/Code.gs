// 投資總覽 v2 的資料介面：App 可以讀，也可以寫回試算表（雙向）
// 安裝方式：試算表 → 擴充功能 → Apps Script → 把這整段貼上 → 改下面的 KEY → 部署為網頁應用程式
//   執行身分：我　／　誰可以存取：所有人
// App 每次都帶著密碼來；密碼不對就什麼都拿不到，也不能寫。
// 每次從 App 寫入，都會在「App 紀錄」分頁留一筆，打錯時可以對照還原。

// 只要改這一行：密碼放在兩個單引號中間，例如 const KEY = 'abc123';
const KEY = '請改成你自己的密碼';

// 程式碼版本（App 設定頁會顯示，用來確認部署的是不是最新版；這行不用改）
const SCRIPT_VERSION = 12;

// 要讀的工作表與範圍（試算表版面有大改時，才需要調整這裡）
const RANGES = {
  total:     ['Total Overall', 'A1:Z300'],
  tw:        ['台股', 'A1:Q400'],
  us:        ['美股', 'A1:Q600'],
  crypto:    ['幣 ', 'A1:N60'],
  myDca:     ['My 定期定額 (台幣帳戶)', 'A1:J40'],
  kidsUs:    ['橙&澈 美股', 'A1:J60'],
  kidsTw:    ['橙&澈 定期定額台股', 'A1:J80'],
  plan:      ['Plan', 'A1:I40'],
  milestone: ['Milestone', 'A1:B200'],
  expense:   ['固定支出', 'A1:L60'],
  house:     ['House', 'A1:L10'],
  bb:        ['BB', 'A1:AD60']
};

const LOG_SHEET = 'App 紀錄';
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
  ensureTrigger();
  try { ensurePriceTrigger(); } catch (err) {}
  refreshCryptoPricesIfStale();
  if (body.action) {
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const fn = ACTIONS[body.action];
      if (!fn) throw new Error('不認得的動作：' + body.action);
      const msg = fn(body.payload || {});
      log(body.action, body.payload);
      SpreadsheetApp.flush();
      return json({ ok: true, message: msg || '已寫入試算表', updated: now(), data: readAll() });
    } catch (err) {
      return json({ ok: false, error: String(err.message || err) });
    } finally {
      lock.releaseLock();
    }
  }
  return json({ ok: true, updated: now(), data: readAll() });
}

// 直接用瀏覽器打開網址時，只確認服務有在運作、是第幾版，不給資料
function doGet() {
  return json({ ok: false, error: 'use the app', version: SCRIPT_VERSION });
}

/* ================================================================
   讀取
   ================================================================ */
function readAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const data = {};
  const isDateValue = v => Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime());
  Object.keys(RANGES).forEach(name => {
    const [sheetName, a1] = RANGES[name];
    const sheet = ss.getSheetByName(sheetName);
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
  try { data._edit = editInfo(); } catch (err) { data._editError = String(err.message || err); }
  return data;
}

// App 輸入表單要用的現值與位置
function editInfo() {
  const tw = sh('台股'), us = sh('美股');
  const info = { tw: { cash: num(tw.getRange('C3').getValue()), holdings: holdings(tw) }, us: { brokers: {} }, dca: dcaItems() };
  Object.keys(BROKERS).forEach(b => {
    info.us.brokers[b] = { cash: num(us.getRange('C' + BROKERS[b]).getValue()), holdings: usHoldings(b) };
  });
  info.fx = num(sh('Total Overall').getRange('B2').getValue());
  info.total = num(sh('Total Overall').getRange('G2').getValue());
  info.expense = expenseItems();
  info.milestones = sh('Milestone').getRange('A1:B200').getDisplayValues()
    .map((r, i) => ({ row: i + 1, date: r[0], text: r[1] })).filter(m => m.date || m.text);
  info.plan = planInfo();
  info.loans = loanRows();
  info.career = careerInfo();
  info.crypto = { coins: cryptoRows(), cash: cryptoCash().qty };
  info.house = houseRows();
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
      if (!h) { if (!buy) throw new Error('台股沒有持有 ' + st.code); h = addTwHolding(st, price); }
      if (!buy && qty > h.qty) throw new Error(`${st.name} 只有 ${h.qty} 股，不能賣 ${qty} 股`);
      if (buy) setAvg(s.getRange(h.row, 15), h.avg, h.qty, price, qty);
      const last = lastRowOf(s, 'A', '日期');
      appendRow(s, last, 1, 10, { 1: date, 2: st.name, 3: '', 4: st.code, 5: price, 6: buy ? qty : '', 7: buy ? '' : qty, 8: buy ? -amt : '', 9: buy ? '' : amt, 10: p.note || '' });
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
    appendRow(s, last, 1, 10, { 1: date, 2: broker, 3: '', 4: code, 5: price, 6: buy ? qty : '', 7: buy ? '' : qty, 8: buy ? -amt : '', 9: buy ? '' : amt, 10: p.note || '' });
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
      if (num(x.add)) s.getRange(it.row, it.col + 5).setValue(round2(it.invested + num(x.add)));
      if (x.shares !== '' && x.shares != null) s.getRange(it.row, it.col + 6).setValue(num(x.shares));
    });
    return '已更新定期定額';
  },

  // 長期計畫：rate 年化報酬率（%）、yearly 每年預設投入（萬）、years: [{year, input, actual}]
  plan(p) {
    const s = sh('Plan'), info = planInfo();
    if (p.rate != null && p.rate !== '') s.getRange('C2').setValue(num(p.rate));
    if (p.yearly != null && p.yearly !== '') s.getRange('F3').setValue(num(p.yearly));
    (p.years || []).forEach(y => {
      const r = info.years.find(x => x.year === num(y.year));
      if (!r) return;
      if (y.input === '' || y.input == null) s.getRange(r.row, 6).setFormula('=$F$3');
      else s.getRange(r.row, 6).setValue(num(y.input));
      if (y.actual !== undefined) s.getRange(r.row, 8).setValue(y.actual === '' ? '' : num(y.actual));
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
    rg.setValues(list);
    return `已更新 ${m} 月固定支出`;
  },

  // 房貸：改利率、每月還款；改利率時用今天的剩餘本金重新起算
  loan(p) {
    const r = loanRows().find(x => x.name === p.name);
    if (!r) throw new Error('找不到房貸：' + p.name);
    const s = sh('House');
    if (p.base) s.getRange(r.row, 2).setValue(toDate(p.base));
    if (p.balance != null && p.balance !== '') s.getRange(r.row, 3).setValue(num(p.balance));
    if (p.rate != null && p.rate !== '') s.getRange(r.row, 4).setValue(num(p.rate) / 100).setNumberFormat('0.00%');
    if (p.pmt != null && p.pmt !== '') s.getRange(r.row, 5).setValue(num(p.pmt));
    if (p.day) s.getRange(r.row, 6).setValue(num(p.day));
    if (p.end) s.getRange(r.row, 7).setValue(toDate(p.end));
    return '已更新房貸';
  },

  // 房產估價：unit 估計單價（萬/坪）、park 估計車位價（萬）
  house(p) {
    const r = houseRows().find(x => x.name === p.name);
    if (!r) throw new Error('找不到房產：' + p.name);
    const s = sh('House');
    s.getRange(r.row, 12).setValue(need(p.unit, '估計單價'));
    if (p.park != null && p.park !== '') s.getRange(r.row, 11).setValue(num(p.park));
    return `已更新 ${p.name} 估價`;
  },

  // 加密貨幣出入金：dir 'in' 轉進加密貨幣、'out' 轉出；source 'fiat'（台幣新資金）、'tw'（台股）、'CUBE'／'Firstrade'（美股）
  // usdt 是加密貨幣這邊增減的 USDT；ntd 是台幣那邊的金額（台幣、台股才需要）
  cryptoFlow(p) {
    const date = toDate(p.date), sign = p.dir === 'out' ? -1 : 1, usdt = need(p.usdt, 'USDT 數量');
    if (usdt <= 0) throw new Error('USDT 數量要大於 0');
    const c = cryptoCash(), inOut = sign > 0 ? '轉入' : '轉出';
    if (sign < 0 && usdt > c.qty + 1e-9) throw new Error(`USDT 只有 ${c.qty}，不能轉出 ${usdt}`);
    let label;
    if (p.source === 'fiat') {
      const ntd = need(p.ntd, '台幣金額');
      appendTotalOverall(date, -sign * ntd, p.note || (sign > 0 ? '加密貨幣入金' : '加密貨幣出金'));
      label = sign > 0 ? '入金（台幣）' : '出金（台幣）';
    } else if (p.source === 'tw') {
      const ntd = need(p.ntd, '台幣金額'), tw = sh('台股'), last = lastRowOf(tw, 'A', '日期');
      appendRow(tw, last, 1, 10, { 1: date, 2: '', 3: sign * ntd, 4: '', 5: '', 6: '', 7: '', 8: '', 9: '', 10: p.note || (sign > 0 ? '轉到加密貨幣' : '從加密貨幣轉回') });
      addTo(tw.getRange('C3'), -sign * ntd);
      label = '台股' + inOut;
    } else if (BROKERS[p.source]) {
      const us = sh('美股'), last = lastRowOf(us, 'A', '日期');
      appendRow(us, last, 1, 10, { 1: date, 2: p.source, 3: sign * usdt, 4: '', 5: '', 6: '', 7: '', 8: '', 9: '', 10: p.note || (sign > 0 ? '轉到加密貨幣' : '從加密貨幣轉回') });
      addTo(us.getRange('C' + BROKERS[p.source]), -sign * usdt);
      label = p.source + inOut;
    } else throw new Error('請選資金來源');
    addTo(sh(CRYPTO).getRange(c.row, 3), sign * usdt);
    cryptoLog(date, label, 'USDT', CRYPTO_CASH[1], '', '', sign * usdt, p.note || '');
    return `已記錄加密貨幣${sign > 0 ? '轉入' : '轉出'} ${usdt} USDT`;
  },

  // 里程碑：op 'add'（加在最後）、'set'（改）、'del'（刪除）
  milestone(p) {
    const s = sh('Milestone');
    if (p.op === 'add') {
      const vals = s.getRange('A1:B300').getValues();
      let r = vals.length; while (r > 0 && vals[r - 1][0] === '' && vals[r - 1][1] === '') r--;
      // 日期欄存成文字（例如 20261105），避免被當成數字或日期改掉格式
      s.getRange(r + 1, 1).setNumberFormat('@').setValue(need(p.date, '日期', true));
      s.getRange(r + 1, 2).setValue(need(p.text, '內容', true));
      return '已新增里程碑';
    }
    const row = num(p.row);
    if (!(row >= 1)) throw new Error('找不到這個里程碑');
    if (p.op === 'del') { s.deleteRow(row); return '已刪除里程碑'; }
    s.getRange(row, 1).setNumberFormat('@').setValue(String(need(p.date, '日期', true)));
    s.getRange(row, 2).setValue(need(p.text, '內容', true));
    return '已更新里程碑';
  },

  // 職涯：op 'mid' 年中分紅、'ye' 年前分紅（職務加給＋年終）、'year' 新增年度（調薪／升遷）、'title' 改職稱
  career(p) {
    const s = sh('BB'), info = careerInfo(), y = num(p.year);
    const r = info.years.find(x => x.year === y);
    if (p.op === 'year') {
      if (r) throw new Error(y + ' 年已經有了');
      const prev = info.years[info.years.length - 1];
      if (!prev) throw new Error('BB 分頁沒有任何年度');
      const row = prev.lastRow + 1;
      s.insertRowAfter(prev.lastRow);
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
    if (p.op === 'mid') { s.getRange(r.row, 14).setValue(need(p.amount, '年中分紅')); s.getRange(r.row, 15).setFormula(`=IFERROR(N${r.row}/M${r.row},"")`); return `已記錄 ${y} 年中分紅`; }
    if (p.op === 'ye') {
      const a = num(p.allowance) || 0, b = need(p.bonus, '年終');
      s.getRange(r.row, 8).setValue(a); s.getRange(r.row, 9).setValue(b);
      s.getRange(r.row, 7).setFormula(`=H${r.row}+I${r.row}`);
      s.getRange(r.row, 10).setFormula(`=IFERROR(I${r.row}/F${r.row},"")`);
      s.getRange(r.row, 11).setFormula(`=IFERROR(G${r.row}/F${r.row},"")`);
      return `已記錄 ${y} 年前分紅`;
    }
    if (p.op === 'title') { s.getRange(r.row, 2).setValue((p.promoted ? '↑ ' : '') + need(p.title, '職稱', true)); return '已更新職稱'; }
    throw new Error('不認得的職涯動作');
  }
};

/* ================================================================
   自動紀錄：每天晚上跑一次；週三記總資產、12/31 記年底結算
   ================================================================ */
function ensureTrigger() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('trigger-v1')) return;
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'dailyJob').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('dailyJob').timeBased().everyDays(1).atHour(22).create();
  props.setProperty('trigger-v1', '1');
}

function dailyJob() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), tz = ss.getSpreadsheetTimeZone();
  const d = new Date(), dow = Utilities.formatDate(d, tz, 'u'), md = Utilities.formatDate(d, tz, 'MM/dd');
  if (dow === '3') weeklyRecord();
  if (md === '12/31') yearEndRecord();
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
function yearEndRecord() {
  const t = sh('Total Overall'), year = Number(Utilities.formatDate(new Date(), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy'));
  const total = num(t.getRange('G2').getValue());
  // 當年投入＝歷史表裡當年的資金異動合計（負數是投入）
  const last = lastRowOf(t, 'A', 'Date'), head = findRow(t, 'A', 'Date');
  const rows = t.getRange(head + 1, 1, last - head, 4).getValues();
  const input = -rows.filter(r => r[0] instanceof Date && r[0].getFullYear() === year).reduce((s, r) => s + (num(r[3]) || 0), 0);
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
  const s = sh('Total Overall'), last = lastRowOf(s, 'A', 'Date');
  const total = num(s.getRange('G2').getValue()), fx = num(s.getRange('B2').getValue());
  appendRow(s, last, 1, 10, { 1: date, 2: Math.round(total), 4: flow, 7: fx, 10: note });
}

// 在表格最後一列之後加一列。為了讓 Google 表格（Table）自動延伸，先在最後一列「之前」插入，
// 再把原本最後一列搬回去，新資料寫在真正的最後一列。values 用欄號（1 起算）指定要寫的格子。
function appendRow(s, last, c1, n, values) {
  s.insertRowBefore(last);
  s.getRange(last + 1, c1, 1, n).copyTo(s.getRange(last, c1, 1, n));
  const row = last + 1;
  Object.keys(values).forEach(c => {
    const v = values[c], cell = s.getRange(row, Number(c));
    if (v instanceof Date) cell.setValue(v).setNumberFormat('yyyy/mm/dd');
    else cell.setValue(v);
  });
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
  const vals = s.getRange(`${col}1:${col}400`).getDisplayValues();
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
function addTwHolding(st, price) {
  const p = sh('Price'), pr = firstEmpty(p, 'A', 2);
  p.getRange(pr, 1, 1, 2).setNumberFormat('@').setValues([[st.code, st.name]]);
  p.getRange(pr, 3).setFormula(`=IFERROR(GOOGLEFINANCE("TPE:"&A${pr}),IFERROR(VALUE(IMPORTXML("https://tw.stock.yahoo.com/quote/"&A${pr}&".TWO","//span[contains(@class,'Fz(32px)')]")),""))`);
  const s = sh('台股'), list = holdings(s), last = list[list.length - 1].row;
  s.insertRowBefore(last);
  s.getRange(last + 1, 9, 1, 9).copyTo(s.getRange(last, 9, 1, 9));
  const row = last + 1;
  s.getRange(row, 9, 1, 2).setNumberFormat('@').setValues([[st.code, st.name]]);
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
    p.getRange(pr, 5).setValue(code);
    p.getRange(pr, 6).setFormula(`=GOOGLEFINANCE("${code}")`);
  }
  const s = sh('美股'), list = usHoldings(broker), last = list[list.length - 1].row;
  s.insertRowBefore(last);
  s.getRange(last + 1, 9, 1, 9).copyTo(s.getRange(last, 9, 1, 9));
  const row = last + 1;
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
        const col = c - 5;   // 這張表的第一欄（股名）
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

// 新的幣（或新的存放位置）：持有清單加一列；價格接 Price 分頁 H/I 欄（沒有的幣先寫成交價）
function addCryptoRow(coin, where, price) {
  const s = sh(CRYPTO), row = firstEmpty(s, 'A', 4);
  let priceCell = 1;
  if (coin !== CRYPTO_CASH[0]) {
    const p = sh('Price');
    let pr = p.getRange('H2:H100').getValues().findIndex(r => normCode(r[0]) === coin) + 2;
    if (pr < 2) { pr = firstEmpty(p, 'H', 2); p.getRange(pr, 8, 1, 2).setValues([[coin, price]]); }
    priceCell = `=Price!I${pr}`;
  }
  s.getRange(row, 1, 1, 3).setValues([[coin, where, 0]]);
  s.getRange(row, 4).setValue(priceCell);
  s.getRange(row, 5).setFormula(`=D${row}*C${row}`);
  s.getRange(row, 6).setFormula(`=E${row}*$J$2`);
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
  s.getRange(h.row, 3).setValue(Math.round((h.qty + (buy ? qty : -qty)) * 1e8) / 1e8);
  addTo(s.getRange(cash.row, 3), buy ? -amt : amt);
  cryptoLog(date, buy ? '買進' : '賣出', coin, where, price, qty, buy ? -amt : amt, p.note || '');
  return `已記錄加密貨幣${buy ? '買進' : '賣出'} ${coin} ${qty} 顆`;
}

// 加密貨幣分頁 P～W 欄的交易紀錄
function cryptoLog(date, type, coin, where, price, qty, usdt, note) {
  const s = sh(CRYPTO), row = firstEmpty(s, 'P', 2);
  s.getRange(row, 16).setValue(date).setNumberFormat('yyyy/mm/dd');
  s.getRange(row, 17, 1, 7).setValues([[type, coin, where, price, qty, round2(usdt), note]]);
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
   小工具
   ================================================================ */
function sh(name) {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!s) throw new Error('找不到分頁：' + name);
  return s;
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
  cell.setValue(round2(held > 0 ? (avg * held + price * qty) / total : price));
}

// 現金這類手打的數字：在原本的值上加減（格子是公式就不動）
function addTo(cell, delta) {
  if (cell.getFormula()) return;
  cell.setValue(round2((num(cell.getValue()) || 0) + delta));
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
   幣價：從 Kraken 抓 Price 分頁 H 欄每個幣的美元價格，寫到 I 欄
   每小時自動跑一次；App 打開時如果超過 10 分鐘沒更新，也會先更新再回傳
   （取代原本 KRAKEN 檔的 updatePrices，舊的定時觸發會換成這個）
   ================================================================ */
function refreshCryptoPrices() {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Price');
  if (!s) return;
  const coins = s.getRange('H2:H40').getValues().map(r => String(r[0]).trim().toUpperCase());
  const list = coins.map((c, i) => ({ c, row: i + 2 })).filter(x => x.c && x.c !== 'USDT' && x.c !== 'USDC');
  if (!list.length) return;
  const res = UrlFetchApp.fetchAll(list.map(x => ({ url: 'https://api.kraken.com/0/public/Ticker?pair=' + x.c + 'USD', muteHttpExceptions: true })));
  res.forEach((r, i) => {
    try {
      const j = JSON.parse(r.getContentText()), k = j.result && Object.keys(j.result)[0];
      const price = k ? parseFloat(j.result[k].c[0]) : NaN;
      if (price > 0) s.getRange(list[i].row, 9).setValue(price);   // 抓不到的幣（Kraken 沒有）保留原本的價格
    } catch (err) {}
  });
  PropertiesService.getScriptProperties().setProperty('price-at', String(Date.now()));
}

function refreshCryptoPricesIfStale() {
  const at = Number(PropertiesService.getScriptProperties().getProperty('price-at') || 0);
  if (Date.now() - at < 10 * 60 * 1000) return;
  try { refreshCryptoPrices(); SpreadsheetApp.flush(); } catch (err) {}
}

// 每小時更新幣價的定時觸發：只建立一次；原本 KRAKEN 檔 updatePrices 的觸發一併換掉，避免兩個同時寫
function ensurePriceTrigger() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('price-trigger-v1')) return;
  ScriptApp.getProjectTriggers().filter(t => ['updatePrices', 'refreshCryptoPrices'].indexOf(t.getHandlerFunction()) >= 0).forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('refreshCryptoPrices').timeBased().everyHours(1).create();
  props.setProperty('price-trigger-v1', '1');
}
