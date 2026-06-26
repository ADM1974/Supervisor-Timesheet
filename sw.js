/* Service worker — makes the app installable and openable offline, WITHOUT
   trapping users on a stale version.

   Strategy:
   - API calls (/.netlify/functions/*): never touched — always go to network.
   - Page + assets (GET): NETWORK-FIRST. Online users always get the latest deploy;
     the cache is only a fallback when there's no signal. This avoids the classic
     PWA trap where a cache-first shell serves an old build after a new deploy. */
const CACHE = 'supervisor-shell-v1';
const SHELL = ['./', './index.html', './manifest.webmanifest', './icons/icon-192.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});
self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  if (url.pathname.includes('/.netlify/functions/')) return;   // API: always network
  if (e.request.method !== 'GET') return;
  e.respondWith(
    fetch(e.request)
      .then(res => {
        const copy = res.clone();
        caches.open(CACHE).then(c => c.put(e.request, copy)).catch(() => {});
        return res;
      })
      .catch(() => caches.match(e.request).then(hit => hit || caches.match('./index.html')))
  );
});
