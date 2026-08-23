const CACHE_NAME = 'react-sheets-shell';
const SHELL_URLS = ['/', '/index.html', '/manifest.webmanifest'];

async function cacheUrls(urls) {
  const cache = await caches.open(CACHE_NAME);
  await Promise.all(urls.map(async (url) => {
    try {
      const request = new Request(url, { credentials: 'same-origin' });
      const response = await fetch(request);
      if (response.ok) await cache.put(request, response.clone());
    } catch {
      // A failed optional resource must never prevent the installed shell
      // from becoming available offline.
    }
  }));
}

self.addEventListener('install', (event) => {
  event.waitUntil(cacheUrls(SHELL_URLS).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (event) => {
  event.waitUntil((async () => {
    const names = await caches.keys();
    await Promise.all(names.filter((name) => name !== CACHE_NAME).map((name) => caches.delete(name)));
    await self.clients.claim();
  })());
});

self.addEventListener('message', (event) => {
  if (event.data?.type !== 'react-sheets.precache' || !Array.isArray(event.data.urls)) return;
  const urls = event.data.urls.filter((url) => typeof url === 'string' && new URL(url, self.location.origin).origin === self.location.origin);
  event.waitUntil(cacheUrls(urls));
});

self.addEventListener('fetch', (event) => {
  if (event.request.method !== 'GET') return;
  const url = new URL(event.request.url);
  if (url.origin !== self.location.origin || url.pathname.startsWith('/api/')) return;

  if (event.request.mode === 'navigate') {
    event.respondWith((async () => {
      try {
        const response = await fetch(event.request);
        const cache = await caches.open(CACHE_NAME);
        await cache.put('/index.html', response.clone());
        return response;
      } catch {
        return (await caches.match('/index.html')) || (await caches.match('/')) || Response.error();
      }
    })());
    return;
  }

  event.respondWith((async () => {
    const cached = await caches.match(event.request);
    if (cached) {
      event.waitUntil(fetch(event.request).then(async (response) => {
        if (response.ok) {
          const cache = await caches.open(CACHE_NAME);
          await cache.put(event.request, response.clone());
        }
      }).catch(() => undefined));
      return cached;
    }
    const response = await fetch(event.request);
    if (response.ok) {
      const cache = await caches.open(CACHE_NAME);
      await cache.put(event.request, response.clone());
    }
    return response;
  })());
});
