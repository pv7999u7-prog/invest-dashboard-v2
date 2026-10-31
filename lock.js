// App 鎖定：Face ID（WebAuthn 通行密鑰）或 Apps Script 密碼都能解鎖。
// 啟用後，設定與快取的財務資料都用一把隨機資料金鑰（AES-GCM）加密後才存進 localStorage。
// 這把資料金鑰另外用兩種方式各包一份：
//   1. 用 Apps Script 密碼（PBKDF2）包 → 輸入密碼解鎖
//   2. 用 Face ID 通行密鑰的 PRF 輸出包 → Face ID 解鎖（iOS 18 以上支援）
// 裝置不支援 PRF 時，Face ID 只當作畫面鎖（資料金鑰改為不加密地存放），設定頁會註明。
'use strict';

const LOCK_KEY = 'inv2-lock-v1';
const ENC_KEY = 'inv2-enc-v1';
const RELOCK_SEC = 60;   // App 在背景超過幾秒，回來時要重新解鎖
const PBKDF2_ITER = 210000;

const textEnc = new TextEncoder(), textDec = new TextDecoder();
const b64 = buf => btoa(String.fromCharCode(...new Uint8Array(buf)));
const unb64 = s => Uint8Array.from(atob(s), c => c.charCodeAt(0));
const rand = n => crypto.getRandomValues(new Uint8Array(n));

const lock = {
  rec: null,        // localStorage 裡的鎖定設定
  dk: null,         // 解鎖後放在記憶體裡的資料金鑰
  locked: false,
  hiddenAt: 0,
  error: '',
  busy: false
};

function loadLockRecord() {
  try { lock.rec = JSON.parse(localStorage.getItem(LOCK_KEY) || 'null'); } catch (e) { lock.rec = null; }
  lock.locked = !!lock.rec;
}
const lockEnabled = () => !!lock.rec;
const bioEnabled = () => !!(lock.rec && lock.rec.bio);

async function aesKey(raw) {
  return crypto.subtle.importKey('raw', raw, 'AES-GCM', false, ['encrypt', 'decrypt']);
}
async function aesEncrypt(key, bytes) {
  const iv = rand(12);
  const ct = await crypto.subtle.encrypt({ name: 'AES-GCM', iv }, key, bytes);
  return { iv: b64(iv), ct: b64(ct) };
}
async function aesDecrypt(key, box) {
  return new Uint8Array(await crypto.subtle.decrypt({ name: 'AES-GCM', iv: unb64(box.iv) }, key, unb64(box.ct)));
}
async function passwordKey(password, salt) {
  const base = await crypto.subtle.importKey('raw', textEnc.encode(password), 'PBKDF2', false, ['deriveKey']);
  return crypto.subtle.deriveKey({ name: 'PBKDF2', salt, iterations: PBKDF2_ITER, hash: 'SHA-256' }, base,
    { name: 'AES-GCM', length: 256 }, false, ['encrypt', 'decrypt']);
}

async function bioAvailable() {
  try {
    return !!(window.PublicKeyCredential && await PublicKeyCredential.isUserVerifyingPlatformAuthenticatorAvailable());
  } catch (e) { return false; }
}

// 用通行密鑰做一次 Face ID 驗證，回傳 PRF 輸出（不支援時為 null）
async function bioAssert(credId, prfSalt) {
  const cred = await navigator.credentials.get({ publicKey: {
    challenge: rand(32), timeout: 60000, userVerification: 'required',
    allowCredentials: [{ type: 'public-key', id: unb64(credId) }],
    extensions: { prf: { eval: { first: unb64(prfSalt) } } }
  } });
  const ext = cred.getClientExtensionResults ? cred.getClientExtensionResults() : {};
  return ext.prf && ext.prf.results && ext.prf.results.first ? new Uint8Array(ext.prf.results.first) : null;
}

// 啟用鎖定。withBio=true 時同時註冊 Face ID
async function enableLock(withBio) {
  const password = state.cfg && state.cfg.key;
  if (!password) throw new Error('請先設定 Apps Script 網址與密碼');
  const dk = rand(32), pwSalt = rand(16);
  const rec = { v: 1, pwSalt: b64(pwSalt), wrapPw: await aesEncrypt(await passwordKey(password, pwSalt), dk), bio: null };
  if (withBio) {
    const prfSalt = rand(32);
    const cred = await navigator.credentials.create({ publicKey: {
      challenge: rand(32), timeout: 60000,
      rp: { name: '投資總覽' },
      user: { id: rand(16), name: '投資總覽', displayName: '投資總覽' },
      pubKeyCredParams: [{ type: 'public-key', alg: -7 }, { type: 'public-key', alg: -257 }],
      authenticatorSelection: { authenticatorAttachment: 'platform', userVerification: 'required', residentKey: 'preferred' },
      extensions: { prf: { eval: { first: prfSalt } } }
    } });
    const credId = b64(cred.rawId);
    const ext = cred.getClientExtensionResults ? cred.getClientExtensionResults() : {};
    let prfOut = ext.prf && ext.prf.results && ext.prf.results.first ? new Uint8Array(ext.prf.results.first) : null;
    // 有些裝置註冊時不直接給 PRF 結果，要再驗證一次才拿得到
    if (!prfOut && ext.prf && ext.prf.enabled) prfOut = await bioAssert(credId, b64(prfSalt));
    rec.bio = prfOut
      ? { mode: 'prf', credId, prfSalt: b64(prfSalt), wrap: await aesEncrypt(await aesKey(prfOut), dk) }
      : { mode: 'gate', credId, prfSalt: b64(prfSalt), dk: b64(dk) };
  }
  lock.rec = rec; lock.dk = dk; lock.locked = false;
  localStorage.setItem(LOCK_KEY, JSON.stringify(rec));
  await persist();
  localStorage.removeItem(LS.cfg); localStorage.removeItem(LS.cache);
}

// 關閉鎖定：資料改回不加密存放
async function disableLock() {
  lock.rec = null; lock.dk = null; lock.locked = false;
  localStorage.removeItem(LOCK_KEY); localStorage.removeItem(ENC_KEY);
  await persist();
}

// 密碼改了（設定頁重新儲存）時，重新包一份密碼鑰匙
async function rewrapPassword(password) {
  if (!lock.rec || !lock.dk) return;
  const pwSalt = rand(16);
  lock.rec.pwSalt = b64(pwSalt);
  lock.rec.wrapPw = await aesEncrypt(await passwordKey(password, pwSalt), lock.dk);
  localStorage.setItem(LOCK_KEY, JSON.stringify(lock.rec));
}

// 把設定與快取存起來：有鎖就加密，沒鎖就照舊
async function persist() {
  const cache = state.raw ? { raw: state.raw, updated: state.updated } : null;
  try {
    if (lockEnabled()) {
      if (!lock.dk) return;
      const box = await aesEncrypt(await aesKey(lock.dk), textEnc.encode(JSON.stringify({ cfg: state.cfg, cache })));
      localStorage.setItem(ENC_KEY, JSON.stringify(box));
    } else {
      if (state.cfg) localStorage.setItem(LS.cfg, JSON.stringify(state.cfg)); else localStorage.removeItem(LS.cfg);
      if (cache) localStorage.setItem(LS.cache, JSON.stringify(cache)); else localStorage.removeItem(LS.cache);
    }
  } catch (e) { console.error(e); }
}

async function openWithKey(dk) {
  const box = JSON.parse(localStorage.getItem(ENC_KEY) || 'null');
  let payload = { cfg: null, cache: null };
  if (box) payload = JSON.parse(textDec.decode(await aesDecrypt(await aesKey(dk), box)));
  lock.dk = dk; lock.locked = false; lock.error = '';
  state.cfg = payload.cfg;
  if (payload.cache) { state.raw = payload.cache.raw; state.updated = payload.cache.updated; state.data = parseAll(payload.cache.raw); }
  render();
  refresh();
}

async function unlockWithBio() {
  if (!bioEnabled() || lock.busy) return;
  lock.busy = true; lock.error = '';
  try {
    const b = lock.rec.bio;
    const prfOut = await bioAssert(b.credId, b.prfSalt);
    let dk;
    if (b.mode === 'prf') {
      if (!prfOut) throw new Error('這次 Face ID 沒有回傳金鑰，請改用密碼');
      dk = await aesDecrypt(await aesKey(prfOut), b.wrap);
    } else dk = unb64(b.dk);
    await openWithKey(dk);
  } catch (e) {
    lock.error = e.name === 'NotAllowedError' ? '' : (e.message || 'Face ID 解鎖失敗');
    render();
  } finally { lock.busy = false; }
}

async function unlockWithPassword(pw) {
  try {
    const dk = await aesDecrypt(await passwordKey(pw, unb64(lock.rec.pwSalt)), lock.rec.wrapPw);
    await openWithKey(dk);
  } catch (e) {
    lock.error = '密碼不對';
    render();
  }
}

function relock() {
  if (!lockEnabled()) return;
  lock.dk = null; lock.locked = true; lock.error = '';
  state.cfg = null; state.raw = null; state.data = null; state.updated = '';
  render();
}

function pageLock() {
  const bio = bioEnabled();
  return `<div style="text-align:center;padding:48px 8px 8px">
    <div class="badge" style="width:64px;height:64px;border-radius:18px;margin:0 auto 14px;background:var(--tint);font-size:30px">
      <svg viewBox="0 0 24 24" width="32" height="32" fill="none" stroke="#fff" stroke-width="2" stroke-linecap="round"><rect x="5" y="11" width="14" height="10" rx="2"/><path d="M8 11V8a4 4 0 0 1 8 0v3"/></svg></div>
    <h1 class="title" style="margin-bottom:4px">投資總覽已鎖定</h1>
    <p class="msg" style="margin:0 0 22px">${bio ? '用 Face ID 或密碼解鎖' : '輸入密碼解鎖'}</p>
  </div>
  <div class="card">
    ${bio ? `<button class="btn" id="unlockBio" style="margin-bottom:14px">
      <span style="display:inline-flex;align-items:center;gap:8px"><svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="#fff" stroke-width="1.8" stroke-linecap="round"><path d="M4 8V6a2 2 0 0 1 2-2h2M16 4h2a2 2 0 0 1 2 2v2M20 16v2a2 2 0 0 1-2 2h-2M8 20H6a2 2 0 0 1-2-2v-2"/><path d="M9 9v1.5M15 9v1.5M12 9v4h-1M9.5 16a3.5 3.5 0 0 0 5 0"/></svg>使用 Face ID 解鎖</span></button>` : ''}
    <label class="field"><span>密碼</span><input id="unlockPw" type="password" autocomplete="current-password"></label>
    <button class="btn" id="unlockPwBtn" style="${bio ? 'background:var(--card-2);color:var(--tint)' : ''}">用密碼解鎖</button>
    ${lock.error ? `<p class="msg err" style="margin:10px 0 0">${esc(lock.error)}</p>` : ''}
  </div>`;
}

function bindLock() {
  const b = document.getElementById('unlockBio');
  if (b) b.addEventListener('click', unlockWithBio);
  const go = () => { const v = document.getElementById('unlockPw').value; if (v) unlockWithPassword(v); };
  const p = document.getElementById('unlockPwBtn');
  if (p) p.addEventListener('click', go);
  const i = document.getElementById('unlockPw');
  if (i) i.addEventListener('keydown', e => { if (e.key === 'Enter') go(); });
}

// 設定頁的「安全」區塊
function securitySection(canBio) {
  if (!state.cfg || !state.cfg.url) return '';
  let body;
  if (!lockEnabled()) {
    body = `<p class="cap">開啟後，每次打開 App 都要先解鎖，存在這台裝置上的資料也會加密。</p>
      ${canBio ? '<button class="btn" id="lockBio">啟用 Face ID＋密碼解鎖</button>' : ''}
      <button class="btn" id="lockPw" style="margin-top:10px;background:var(--card-2);color:var(--tint)">只用密碼解鎖</button>
      ${canBio ? '' : '<p class="msg" style="margin:10px 0 0">這個瀏覽器不支援 Face ID，可以先用密碼鎖。</p>'}`;
  } else {
    const b = lock.rec.bio;
    body = `<div class="kv"><span>解鎖方式</span><span>${b ? 'Face ID＋密碼' : '密碼'}</span></div>
      <div class="kv"><span>資料加密</span><span>${!b || b.mode === 'prf' ? '已加密' : 'Face ID 僅當畫面鎖'}</span></div>
      <div class="kv"><span>自動鎖定</span><span>離開 ${RELOCK_SEC} 秒後</span></div>
      ${b && b.mode === 'gate' ? '<p class="msg" style="margin:6px 0 10px">這台裝置的 Face ID 不支援產生加密金鑰（需要 iOS 18 以上），所以 Face ID 只擋畫面；用密碼解鎖的資料仍是加密的。</p>' : ''}
      ${!b && canBio ? '<button class="btn" id="lockBio" style="margin-top:10px">加上 Face ID</button>' : ''}
      <button class="btn ghost" id="lockOff">關閉鎖定</button>`;
  }
  return `<div class="card"><h3>安全</h3>${body}<p class="msg err" id="lockMsg" style="margin:10px 0 0"></p></div>`;
}

function bindSecurity() {
  const msg = document.getElementById('lockMsg');
  const run = async fn => {
    try { await fn(); render(); }
    catch (e) { if (msg) msg.textContent = e.name === 'NotAllowedError' ? '已取消 Face ID 設定' : (e.message || '設定失敗'); }
  };
  const bio = document.getElementById('lockBio');
  if (bio) bio.addEventListener('click', () => run(() => enableLock(true)));
  const pw = document.getElementById('lockPw');
  if (pw) pw.addEventListener('click', () => run(() => enableLock(false)));
  const off = document.getElementById('lockOff');
  if (off) off.addEventListener('click', () => { if (confirm('關閉後，App 打開時不用解鎖，資料也改為不加密存放。確定嗎？')) run(disableLock); });
}
