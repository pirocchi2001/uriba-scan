// オフライン対応：アプリ本体を端末にキャッシュする。
// 更新を配信するときは CACHE の番号を上げる。
const CACHE = 'uriba-scan-v3';
const ASSETS = [
  './',
  'index.html',
  'style.css',
  'app.js',
  'master-parse.js',
  'manifest.webmanifest',
  'icons/icon.svg',
  'icons/icon-180.png',
  'icons/icon-192.png',
  'icons/icon-512.png',
  'vendor/xlsx.core.min.js',
  'vendor/barcode-detector-ponyfill.js',
  'vendor/zxing_reader.wasm',
  'master.enc',
];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(ASSETS)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

// ネットワーク優先（最新版を取りに行き、電波がなければキャッシュを使う）
self.addEventListener('fetch', (e) => {
  const url = new URL(e.request.url);
  if (e.request.method !== 'GET' || url.origin !== location.origin) return;
  // 動画は分割取得(Range)されるので扱わない（キャッシュから返すとiPhoneで再生できなくなる）
  if (url.pathname.endsWith('.mp4')) return;
  e.respondWith(
    fetch(e.request)
      .then((res) => {
        if (res.ok) {
          const copy = res.clone();
          caches.open(CACHE).then((c) => c.put(e.request, copy));
        }
        return res;
      })
      .catch(() => caches.match(e.request, { ignoreSearch: true }))
  );
});
