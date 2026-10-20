// Sổ lớp · bộ nhớ ngoại tuyến. Đổi VERSION mỗi lần cập nhật app để máy người dùng nhận bản mới.
const VERSION = '2.1.0';
const CACHE = 'so-lop-' + VERSION;
const ASSETS = ['./', './app.js', './core.js', './report.js', './sync.js', './xlsx.js', './style.css', './manifest.webmanifest', './icon.svg', './icon-192.png', './icon-512.png', './icon-maskable.png', './apple-touch-icon.png'];

self.addEventListener('install', event => {
  event.waitUntil((async () => {
    const cache = await caches.open(CACHE);
    await Promise.all(ASSETS.map(async path => {
      const res = await fetch(path, { cache: 'reload' });
      if (!res.ok) throw Error('Không tải được ' + path);
      await cache.put(path, res.redirected ? new Response(await res.blob(), { headers: res.headers }) : res);
    }));
  })());
});
self.addEventListener('message', event => { if (event.data === 'skip') self.skipWaiting(); });
self.addEventListener('activate', event => {
  event.waitUntil((async () => {
    for (const key of await caches.keys()) if (key.startsWith('so-lop-') && key !== CACHE) await caches.delete(key);
    await self.clients.claim();
  })());
});
self.addEventListener('fetch', event => {
  const req = event.request, url = new URL(req.url);
  if (req.method !== 'GET' || url.origin !== self.location.origin) return;
  event.respondWith((async () => {
    const cache = await caches.open(CACHE);
    if (req.mode === 'navigate') return (await cache.match('./')) || fetch(req);
    return (await cache.match(req, { ignoreSearch: true })) || fetch(req);
  })());
});
