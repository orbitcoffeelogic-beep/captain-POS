/**
 * CAPTAIN COFFEE POS & KDS — OFFLINE-FIRST PWA SERVICE WORKER (sw.js)
 * Mengunci file aplikasi utama, Tailwind CSS CDN, dan Google Fonts di memori Tablet Android
 * sehingga aplikasi tetap dapat dibuka 100% normal meskipun Wi-Fi mati & tablet di-restart.
 */
const CACHE_NAME = 'captain-pos-offline-v5.13';
const CORE_ASSETS = [
  './',
  './index.html',
  './manifest.json',
  'https://cdn.tailwindcss.com',
  'https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&family=JetBrains+Mono:ital,wght@0,400;0,600;0,700;0,800;1,400&family=League+Spartan:wght@600;700;800;900&family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap'
];

self.addEventListener('install', (event) => {
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_NAME).then(async (cache) => {
      for (const url of CORE_ASSETS) {
        try {
          const req = new Request(url, { mode: url.startsWith('http') ? 'no-cors' : 'same-origin' });
          const res = await fetch(req);
          await cache.put(url, res);
        } catch (err) {
          console.warn('[SW] Gagal pre-cache aset:', url, err);
        }
      }
    })
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(
        keys.map((key) => {
          if (key !== CACHE_NAME) {
            return caches.delete(key);
          }
        })
      )
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.method !== 'GET') return;

  const url = new URL(req.url);

  // Jangan pernah cache panggilan API Google Apps Script (biarkan mesin SyncQueue & fetchOrdersFromSheets yang mengatur)
  if (url.hostname.includes('script.google.com') || url.hostname.includes('script.googleusercontent.com')) {
    return;
  }

  // Strategi Network-First untuk navigasi HTML utama (agar update terbaru selalu masuk saat online, tapi 100% jalan dari cache saat offline)
  if (req.mode === 'navigate' || url.pathname.endsWith('.html') || url.pathname.endsWith('/')) {
    event.respondWith(
      fetch(req)
        .then((networkRes) => {
          const clone = networkRes.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(req, clone));
          return networkRes;
        })
        .catch(async () => {
          const cached = await caches.match(req);
          if (cached) return cached;
          return caches.match('./index.html');
        })
    );
    return;
  }

  // Strategi Cache-First dengan Stale-While-Revalidate untuk Tailwind CDN, Google Fonts, & aset statis
  event.respondWith(
    caches.match(req).then((cachedRes) => {
      const fetchPromise = fetch(req)
        .then((networkRes) => {
          if (networkRes && (networkRes.status === 200 || networkRes.type === 'opaque')) {
            const clone = networkRes.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(req, clone));
          }
          return networkRes;
        })
        .catch(() => cachedRes);

      return cachedRes || fetchPromise;
    })
  );
});
