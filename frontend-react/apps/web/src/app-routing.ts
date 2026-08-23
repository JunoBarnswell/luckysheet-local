import { useCallback, useSyncExternalStore } from 'react';

export type ApplicationRoute =
  | { kind: 'auth-callback' | 'auth-silent-renew' | 'hub' }
  | { kind: 'not-found'; pathname: string }
  | { kind: 'workbook'; unitId: string };

function normalizedPathname(pathname: string): string {
  if (pathname.length > 1 && pathname.endsWith('/')) return pathname.slice(0, -1);
  return pathname || '/';
}

export function parseApplicationRoute(pathname = typeof window === 'undefined' ? '/workbooks' : window.location.pathname): ApplicationRoute {
  const normalized = normalizedPathname(pathname);
  if (normalized === '/' || normalized === '/workbooks') return { kind: 'hub' };
  if (normalized === '/auth/callback') return { kind: 'auth-callback' };
  if (normalized === '/auth/silent-renew') return { kind: 'auth-silent-renew' };
  const workbook = /^\/workbooks\/([^/]+)$/.exec(normalized);
  if (workbook?.[1]) return { kind: 'workbook', unitId: decodeURIComponent(workbook[1]) };
  return { kind: 'not-found', pathname: normalized };
}

function subscribe(listener: () => void): () => void {
  window.addEventListener('popstate', listener);
  return () => window.removeEventListener('popstate', listener);
}

function currentPathname(): string {
  return window.location.pathname;
}

export function useApplicationRoute(): ApplicationRoute {
  const pathname = useSyncExternalStore(subscribe, currentPathname, () => '/workbooks');
  return parseApplicationRoute(pathname);
}

export function navigate(pathname: string, { replace = false }: { replace?: boolean } = {}): void {
  const target = pathname.startsWith('/') ? pathname : `/${pathname}`;
  if (window.location.pathname === target) return;
  if (replace) window.history.replaceState({}, '', target);
  else window.history.pushState({}, '', target);
  window.dispatchEvent(new PopStateEvent('popstate'));
}

export function useNavigate(): (pathname: string, options?: { replace?: boolean }) => void {
  return useCallback((pathname: string, options?: { replace?: boolean }) => navigate(pathname, options), []);
}
