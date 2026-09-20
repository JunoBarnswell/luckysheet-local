// Retires already-installed shells during an explicit browser application upgrade.
const SHELL_URLS = [];
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', event => event.waitUntil((async () => {
  await caches.delete('react-sheets-shell');
  await self.registration.unregister();
  await self.clients.claim();
})()));
