// 投資總覽 v2 的資料介面：App 可以讀，也可以寫回試算表（雙向）
// 安裝方式：試算表 → 擴充功能 → Apps Script → 把這整段貼上 → 改下面的 KEY → 部署為網頁應用程式
//   執行身分：我　／　誰可以存取：所有人
// App 每次都帶著密碼來；密碼不對就什麼都拿不到，也不能寫。
// 每次從 App 寫入，都會在「App 紀錄」分頁留一筆；改到的格子另外記在隱藏的「App 復原」分頁，App 可以「復原上一步」。

// 只要改這一行：密碼放在兩個單引號中間，例如 const KEY = 'abc123';
const KEY = '請改成你自己的密碼';

// 程式碼版本（App 設定頁會顯示，用來確認部署的是不是最新版；這行不用改）
const SCRIPT_VERSION = 19;

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
  if (body.action) {
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const fn = ACTIONS[body.action];
      if (!fn) throw new Error('不認得的動作：' + body.action);
      if (body.action !== 'undo') REC = { desc: '', ops: [] };
      const msg = fn(body.payload || {});
      SpreadsheetApp.flush();
      if (REC) { REC.desc = msg || body.action; recSave(); }
      log(body.action, body.payload);
      return json({ ok: true, message: msg || '已寫入試算表', updated: now(), data: readAll() });
    } catch (err) {
      // 寫到一半失敗：照記下的資料倒回去，不留半套
      let rolled = false;
      if (REC && REC.ops.length) {
        try { SpreadsheetApp.flush(); } catch (e) {}
        try { applyUndo(REC.ops); SpreadsheetApp.flush(); rolled = true; } catch (e) {}
      }
      REC = null;
      return json({ ok: false, error: String(err.message || err) + (rolled ? '（已自動還原，試算表沒有被改動）' : '') });
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
      if (!h) { if (!buy) throw new Error('台股沒有持有 ' + st.code); h = addTwHolding(st, price); }
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

  // 里程碑：op 'add'（加在最後）、'set'（改）、'del'（刪除）
  milestone(p) {
    const s = sh('Milestone');
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
    if (p.op !== 'edit') throw new Error('不認得的明細動作');
    const d = p.data || {};
    const shift = L.unapply(s, old) || 0;   // 加新股票時，持股清單多一列，交易表會往下移
    row += shift;
    const msg = L.write(s, row, d, old) || '已更新這筆' + L.label;
    // 日期、備註改了，同一組的其他筆一起改（金額怎麼分到各帳戶沒辦法自動推算，不連動）
    const nd = d.date ? toDate(d.date) : null, nn = d.note != null ? String(d.note) : null;
    let synced = 0;
    group.forEach(g => {
      const L2 = LEDGERS[g.k], s2 = sh(L2.sheet);
      let r2 = g.r;
      if (shift && L2.sheet === L.sheet && r2 > row - shift) r2 += shift;
      const changed = (nd && dkey(nd) !== g.x.date) || (nn != null && nn.trim() !== String(g.x.note).trim());
      if (!changed) return;
      if (nd) put(s2.getRange(r2, L2.dateCol), nd);
      if (nn != null) put(s2.getRange(r2, L2.noteCol), nn);
      synced++;
    });
    return synced ? `${msg}；同一次出入金的另外 ${synced} 筆也改了日期、備註` : msg;
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
  putVs(fmtCell(p.getRange(pr, 1, 1, 2), '@'), [[st.code, st.name]]);
  putF(p.getRange(pr, 3), `=IFERROR(GOOGLEFINANCE("TPE:"&A${pr}),IFERROR(VALUE(IMPORTXML("https://tw.stock.yahoo.com/quote/"&A${pr}&".TWO","//span[contains(@class,'Fz(32px)')]")),""))`);
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
    putF(p.getRange(pr, 6), `=GOOGLEFINANCE("${code}")`);
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
      putF(p.getRange(pr, 9), `=IFERROR(GOOGLEFINANCE("CURRENCY:"&H${pr}&"USD"),${price})`);
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
   明細的修改與刪除：資金異動（Total Overall）、台股交易表、美股交易表、加密貨幣交易紀錄
   每種明細：在哪個分頁、哪一欄有標題、從第幾欄起佔幾欄；
   read 讀一列；unapply 把這筆對閒置現金、均價、顆數的影響拿掉；write 寫入新內容並套用新的影響
   ================================================================ */
const LEDGERS = {
  flow: {
    sheet: 'Total Overall', col: 'A', header: 'Date', c1: 1, n: 10, label: '資金異動', dateCol: 1, noteCol: 10,
    parse: v => ({ date: dkey(v[0]), key: num(v[3]) || 0, flow: num(v[3]), note: String(v[9]) }),
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

function snap(range) {
  if (REC) {
    const f = range.getFormulas(), v = range.getValues();
    REC.ops.push({ t: 'cells', sh: range.getSheet().getName(), r: range.getRow(), c: range.getColumn(),
      d: v.map((row, i) => row.map((x, j) => f[i][j] || (x instanceof Date ? { $d: x.getTime() } : x))) });
  }
  return range;
}
function recRow(t, s, row) { if (REC) REC.ops.push({ t, sh: s.getName(), row }); }
function put(range, v) { snap(range); range.setValue(v); return range; }
function putF(range, f) { snap(range); range.setFormula(f); return range; }
function putVs(range, vals) { snap(range); range.setValues(vals); return range; }

function applyUndo(ops) {
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
  const extra = s.getLastRow() - 1 - UNDO_KEEP;
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
  if (!s || s.getLastRow() < 2) return { count: 0 };
  const r = s.getLastRow(), v = s.getRange(r, 1, 1, 2).getDisplayValues()[0];
  return { count: r - 1, time: v[0], desc: v[1] };
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
  put(cell, round2(held > 0 ? (avg * held + price * qty) / total : price));
}

// 現金這類手打的數字：在原本的值上加減（格子是公式就不動）
function addTo(cell, delta) {
  if (cell.getFormula()) return;
  put(cell, round2((num(cell.getValue()) || 0) + delta));
}

// 設定數字格式；Google 表格（Table）裡有指定類型的欄位不准改格式，那種欄位本來就會自動套格式，略過即可
function fmtCell(range, f) {
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
