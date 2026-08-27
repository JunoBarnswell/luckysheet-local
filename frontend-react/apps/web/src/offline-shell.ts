function sameOriginResourceUrls(): string[] {
  const urls = new Set<string>(['/', '/index.html', '/manifest.webmanifest']);
  for (const entry of performance.getEntriesByType('resource')) {
    if (!(entry instanceof PerformanceResourceTiming)) continue;
    const url = new URL(entry.name, window.location.origin);
    if (url.origin === window.location.origin && !url.pathname.startsWith('/api/')) {
      urls.add(url.pathname + url.search);
    }
  }
  return [...urls];
}

/** Registers the offline application shell. Workbook data is intentionally
 * owned by the canonical local workspace store, never by this cache. */
export function registerOfflineShell(): void {
  if (import.meta.env.DEV) {
    if ('serviceWorker' in navigator) {
      void navigator.serviceWorker.getRegistrations().then((registrations) =>
        Promise.all(registrations.map((registration) => registration.unregister())),
      );
    }
    return;
  }
  if (!('serviceWorker' in navigator) || !window.isSecureContext) return;
  window.addEventListener('load', () => {
    void navigator.serviceWorker.register('/sw.js', { scope: '/' })
      .then(async (registration) => {
        await navigator.serviceWorker.ready;
        registration.active?.postMessage({ type: 'react-sheets.precache', urls: sameOriginResourceUrls() });
      })
      .catch(() => undefined);
  }, { once: true });
}
