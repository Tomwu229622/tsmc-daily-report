/* TSMC 日報 service worker
 * - App shell (index, web app manifest, icons) is precached; the report list
 *   (manifest.json) is also fetched at install so the app works offline
 *   right after the first visit.
 * - index, manifest.json and report pages (*.html) are network-first with a
 *   timeout, so the newest daily report always wins when online; the last
 *   cached copy is served when offline or when the network is too slow.
 * - Report pages cache is trimmed to the most recent MAX_REPORTS entries.
 * - Bump VERSION whenever icons or app.webmanifest change (they are cache-first).
 */
const VERSION = 'v1';
const SHELL_CACHE = 'tsmc-shell-' + VERSION;
const PAGE_CACHE = 'tsmc-pages-' + VERSION;
const MAX_REPORTS = 30;
const NETWORK_TIMEOUT_MS = 5000;

const SHELL = [
  './',
  './index.html',
  './app.webmanifest',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/icon-maskable-512.png',
  './icons/apple-touch-icon.png',
  './icons/favicon-32.png',
];

function stripQuery(u) {
  const url = new URL(u, self.location.href);
  url.search = '';
  return url.toString();
}

self.addEventListener('install', event => {
  event.waitUntil((async () => {
    const shell = await caches.open(SHELL_CACHE);
    await shell.addAll(SHELL);
    try {
      const res = await fetch('./manifest.json', { cache: 'no-store' });
      if (res.ok) await (await caches.open(PAGE_CACHE)).put(stripQuery('./manifest.json'), res);
    } catch (e) { /* offline during install: list will be cached on next online load */ }
    await self.skipWaiting();
  })());
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys => Promise.all(
      keys.filter(k => k.startsWith('tsmc-') && k !== SHELL_CACHE && k !== PAGE_CACHE)
          .map(k => caches.delete(k))
    )).then(() => self.clients.claim())
  );
});

async function trimPages() {
  const cache = await caches.open(PAGE_CACHE);
  const keys = await cache.keys();
  // keep manifest.json; trim oldest report pages beyond the limit
  const pages = keys.filter(r => !new URL(r.url).pathname.endsWith('/manifest.json'));
  for (let i = 0; i < pages.length - MAX_REPORTS; i++) await cache.delete(pages[i]);
}

async function storeCopy(cacheName, key, res) {
  try {
    const cache = await caches.open(cacheName);
    await cache.delete(key);   // re-insert so key order = recency
    await cache.put(key, res);
    if (cacheName === PAGE_CACHE) await trimPages();
  } catch (e) { /* quota or opaque/partial response: keep serving, skip caching */ }
}

function offlinePage() {
  return new Response(
    '<!doctype html><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<body style="font-family:Segoe UI,Microsoft JhengHei,sans-serif;padding:24px;color:#0D1B2A;background:#F0F4F8">' +
    '<h3>目前離線</h3><p>這份報告尚未存到本機，請連上網路後再開啟。</p></body>',
    { status: 503, headers: { 'Content-Type': 'text/html; charset=utf-8' } });
}

async function networkFirst(event, cacheName, isIndex) {
  const request = event.request;
  const key = stripQuery(request.url);
  const network = fetch(request, { cache: 'no-store' }).then(res => {
    if (res && res.ok && res.status === 200) {
      event.waitUntil(storeCopy(cacheName, key, res.clone()));
    }
    return res;
  });
  const timeout = new Promise(resolve => setTimeout(() => resolve(null), NETWORK_TIMEOUT_MS));

  let res = null;
  try {
    res = await Promise.race([network, timeout]);
  } catch (e) { res = null; }
  if (res) return res;

  const hit = await caches.match(key);
  if (hit) return hit;
  // slow network but nothing cached: keep waiting for the network
  try {
    return await network;
  } catch (e) {
    if (isIndex) {
      const shell = await caches.match(stripQuery('./index.html'));
      if (shell) return shell;
    }
    if (request.mode === 'navigate' || request.destination === 'iframe') return offlinePage();
    throw e;
  }
}

self.addEventListener('fetch', event => {
  const req = event.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return;   // leave external requests alone

  const path = url.pathname;
  const isIndex = req.mode === 'navigate' && req.destination !== 'iframe' &&
                  (path.endsWith('/') || path.endsWith('/index.html'));
  if (isIndex) {
    event.respondWith(networkFirst(event, SHELL_CACHE, true));
  } else if (path.endsWith('/manifest.json') || path.endsWith('.html')) {
    event.respondWith(networkFirst(event, PAGE_CACHE, false));
  } else {
    // static assets (icons, web app manifest): cache-first
    event.respondWith(caches.match(req).then(hit => hit || fetch(req)));
  }
});
