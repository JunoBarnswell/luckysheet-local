/** One-time retirement of the old offline shell; workbook pages now require Java. */
export function registerOfflineShell(): void {
  if (!('serviceWorker' in navigator)) return;
  void navigator.serviceWorker.getRegistrations().then(async registrations => {
    for (const registration of registrations) {
      const script = registration.active?.scriptURL ?? registration.waiting?.scriptURL;
      if (script && new URL(script).pathname === '/sw.js') await registration.unregister();
    }
    if ('caches' in window) await caches.delete('react-sheets-shell');
  }).catch(error => console.error('Unable to retire previous offline shell', error));
}
