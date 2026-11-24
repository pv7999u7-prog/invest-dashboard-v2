// Web Push（RFC 8291 aes128gcm＋RFC 8292 VAPID）純 JavaScript 版，給 Apps Script 用（它沒有 ECDH、AES-GCM）。
// P-256 用 BigInt；AES-128-GCM 自己實作；HMAC／SHA-256／亂數由呼叫的人提供（Apps Script 用 Utilities）。
// 位元組一律用 0～255 的一般陣列。
'use strict';

// Apps Script 的編輯器不認得 123n 這種 BigInt 寫法，所以常數都用 BigInt() 建立；
// 萬一執行環境沒有 BigInt，這裡不能出錯（不然整個 Apps Script 都不能跑），只有推播會失敗
const wpBig = x => (typeof BigInt === 'function' ? BigInt(x) : null);
const WN0 = wpBig(0), WN1 = wpBig(1), WN2 = wpBig(2), WN3 = wpBig(3), WN4 = wpBig(4), WN8 = wpBig(8), WN120 = wpBig(120), WN255 = wpBig(255);

const WP_P = wpBig('0xffffffff00000001000000000000000000000000ffffffffffffffffffffffff');
const WP_N = wpBig('0xffffffff00000000ffffffffffffffffbce6faada7179e84f3b9cac2fc632551');
const WP_B = wpBig('0x5ac635d8aa3a93e7b3ebbd55769886bc651d06b0cc53b0f63bce3c3e27d2604b');
const WP_G = [wpBig('0x6b17d1f2e12c4247f8bce6e563a440f277037d812deb33a0f4a13945d898c296'), wpBig('0x4fe342e2fe1a7f9b8ee7eb4a7c0f9e162bce33576b315ececbb6406837bf51f5')];

const wpMod = (a, m) => { const r = a % m; return r < WN0 ? r + m : r; };
function wpInv(a, m) { // 擴展歐幾里得
  let [r0, r1, s0, s1] = [wpMod(a, m), m, WN1, WN0];
  while (r1) { const q = r0 / r1; [r0, r1] = [r1, r0 - q * r1]; [s0, s1] = [s1, s0 - q * s1]; }
  return wpMod(s0, m);
}
// Jacobian 座標的點：[X, Y, Z]；無窮遠點 Z = 0
function wpDbl(Pt) {
  const [X, Y, Z] = Pt, p = WP_P;
  if (!Z || !Y) return [WN0, WN1, WN0];
  const YY = Y * Y % p, S = WN4 * X * YY % p, ZZ = Z * Z % p;
  const M = WN3 * (X - ZZ) * (X + ZZ) % p;   // a = -3
  const X3 = wpMod(M * M - WN2 * S, p), Y3 = wpMod(M * (S - X3) - WN8 * YY * YY, p), Z3 = WN2 * Y * Z % p;
  return [X3, Y3, Z3];
}
function wpAdd(P1, P2) {
  if (!P1[2]) return P2; if (!P2[2]) return P1;
  const p = WP_P, [X1, Y1, Z1] = P1, [X2, Y2, Z2] = P2;
  const Z1Z1 = Z1 * Z1 % p, Z2Z2 = Z2 * Z2 % p;
  const U1 = X1 * Z2Z2 % p, U2 = X2 * Z1Z1 % p, S1 = Y1 * Z2 * Z2Z2 % p, S2 = Y2 * Z1 * Z1Z1 % p;
  if (U1 === U2) return S1 === S2 ? wpDbl(P1) : [WN0, WN1, WN0];
  const H = wpMod(U2 - U1, p), R = wpMod(S2 - S1, p), HH = H * H % p, HHH = H * HH % p, V = U1 * HH % p;
  const X3 = wpMod(R * R - HHH - WN2 * V, p), Y3 = wpMod(R * (V - X3) - S1 * HHH, p), Z3 = Z1 * Z2 % p * H % p;
  return [X3, Y3, Z3];
}
function wpMul(k, Pt) {
  let R = [WN0, WN1, WN0], Q = [Pt[0], Pt[1], WN1];
  for (let i = wpBig(k.toString(2).length) - WN1; i >= WN0; i--) { R = wpDbl(R); if ((k >> i) & WN1) R = wpAdd(R, Q); }
  if (!R[2]) throw new Error('point at infinity');
  const zi = wpInv(R[2], WP_P), zi2 = zi * zi % WP_P;
  return [R[0] * zi2 % WP_P, R[1] * zi2 % WP_P * zi % WP_P];
}
const wpToBig = b => b.reduce((t, x) => (t << WN8) | wpBig(x & 255), WN0);
const wpFromBig = (n, len) => { const out = []; for (let i = 0; i < len; i++) { out.unshift(Number(n & WN255)); n >>= WN8; } return out; };
const wpPointBytes = Pt => [4].concat(wpFromBig(Pt[0], 32), wpFromBig(Pt[1], 32));
function wpPoint(bytes) {
  if (bytes.length !== 65 || bytes[0] !== 4) throw new Error('bad public key');
  const x = wpToBig(bytes.slice(1, 33)), y = wpToBig(bytes.slice(33));
  if (wpMod(y * y - (x * x * x - WN3 * x + WP_B), WP_P) !== WN0) throw new Error('point not on curve');
  return [x, y];
}
function wpScalar(random) { // 1 ≤ k < n
  for (;;) { const k = wpToBig(random(32)); if (k > WN0 && k < WP_N) return k; }
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
  const R = wpBig('0xe1') << WN120; let Z = WN0, V = Y;
  for (let i = 127; i >= 0; i--) { if ((X >> wpBig(i)) & WN1) Z ^= V; V = V & WN1 ? (V >> WN1) ^ R : V >> WN1; }
  return Z;
}
function wpGcmEncrypt(key, iv, pt) { // 沒有 AAD；回傳 密文＋16 位元組 tag
  const w = wpAesKey(key), H = wpToBig(wpAesBlock(w, new Array(16).fill(0)));
  const J0 = iv.concat([0, 0, 0, 1]), ct = [];
  for (let i = 0, ctr = 2; i < pt.length; i += 16, ctr++) {
    const cb = iv.concat(wpFromBig(wpBig(ctr), 4)), ks = wpAesBlock(w, cb);
    for (let j = 0; j < 16 && i + j < pt.length; j++) ct.push(pt[i + j] ^ ks[j]);
  }
  let g = WN0;
  for (let i = 0; i < ct.length; i += 16) { const blk = ct.slice(i, i + 16); while (blk.length < 16) blk.push(0); g = wpGmul(g ^ wpToBig(blk), H); }
  g = wpGmul(g ^ wpBig(ct.length * 8), H);   // len(A)=0 || len(C)
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
  if (typeof BigInt !== 'function') throw new Error('這個 Apps Script 執行環境不支援 BigInt，無法送推播');
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
  if (typeof BigInt !== 'function') throw new Error('這個 Apps Script 執行環境不支援 BigInt，無法送推播');
  return {
    url: sub.endpoint, method: 'post', contentType: 'application/octet-stream', muteHttpExceptions: true,
    headers: { Authorization: wpVapidAuth(sub.endpoint, vapid, prim, subject), 'Content-Encoding': 'aes128gcm', TTL: '43200', Urgency: 'high' },
    payload: wpEncrypt(sub, payload, prim)
  };
}

if (typeof module !== 'undefined') module.exports = { wpVapidKeys, wpVapidAuth, wpEncrypt, wpRequest, wpGcmEncrypt, wpB64u, wpUnb64u, wpMul, WP_G };
