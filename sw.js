// 離線快取：只快取 App 本身的檔案；試算表資料（script.google.com）一律不經過快取
// 改了程式後，把 VERSION 改一下，下次連網開啟時就會換成新版
const PREFIX = 'inv2-';
const VERSION = PREFIX + 'v65';
const FILES = ['./', 'index.html', 'lock.js', 'app.js', 'edit.js', 'policy.js', 'etf-page.js', 'manifest.webmanifest', 'icon-v9-180.png', 'icon-v9-192.png', 'icon-v9-512.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSION).then(c => c.addAll(FILES)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys()
    .then(keys => Promise.all(keys.filter(k => k.startsWith(PREFIX) && k !== VERSION).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});

// 先用網路拿最新版，沒網路時才用快取
// cache: 'no-cache' 讓瀏覽器每次都跟 GitHub 確認，不會拿到 10 分鐘內的舊檔
self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET' || new URL(e.request.url).origin !== location.origin) return;
  if (e.request.destination === 'image') return;   // 圖示交給瀏覽器直接抓，加入主畫面時才拿得到
  e.respondWith(
    fetch(e.request, { cache: 'no-cache' })
      .then(res => {
        const copy = res.clone();
        caches.open(VERSION).then(c => c.put(e.request, copy));
        return res;
      })
      .catch(() => caches.match(e.request, { ignoreSearch: true }))
  );
});

// 0050 每日建議的推播：Apps Script 每個開盤日早上送來；點通知打開 App 並跳出今日建議
self.addEventListener('push', e => {
  let m = {};
  try { m = e.data ? e.data.json() : {}; } catch (err) { m = { body: e.data ? e.data.text() : '' }; }
  e.waitUntil(self.registration.showNotification(m.title || '0050 今日建議', {
    body: m.body || '打開看今天的建議', tag: m.tag || 'etf', data: { url: m.url || './#etfpop' },
    icon: 'icon-v9-192.png', badge: 'icon-v9-192.png'
  }));
});
self.addEventListener('notificationclick', e => {
  e.notification.close();
  const url = new URL((e.notification.data && e.notification.data.url) || './#etfpop', self.registration.scope).href;
  e.waitUntil(self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then(list => {
    const c = list.find(x => x.url.startsWith(self.registration.scope));
    if (c) { c.postMessage({ type: 'etfpop' }); return c.focus(); }
    return self.clients.openWindow(url);
  }));
});
