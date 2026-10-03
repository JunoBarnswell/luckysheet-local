import assert from 'node:assert/strict';
import test from 'node:test';
import { AuthDomain } from '../auth/domain';
import { ApplicationRuntime } from './runtime';
import { createSpreadsheetSdk } from '../sdk';
import { SdkError } from '../error';

function authPort() {
  let subject: string | null = null;
  return async (input: string | URL | Request, init?: RequestInit): Promise<Response> => {
    const path = String(input);
    if (path === '/api/auth/config') return Response.json({ mode: 'local' });
    if (path === '/api/auth/login') subject = 'user-a';
    if (path === '/api/auth/logout') subject = null;
    if (path === '/api/auth/session') return Response.json({ authenticated: Boolean(subject), subject, displayName: subject, admin: false, bootstrapRequired: false, csrfToken: 'csrf' });
    if (path.endsWith('/workbooks')) return Response.json({ items: [], nextCursor: null });
    return Response.json({});
  };
}

test('runtime scopes cache and catalog owners to auth subject and retires old actions', async () => {
  const auth = new AuthDomain({ fetch: authPort() });
  const runtime = new ApplicationRuntime(auth, undefined, authPort());
  await auth.session.initialize();
  const anonymousCatalog = runtime.catalog;
  await assert.rejects(anonymousCatalog.list(), /unavailable/);
  await auth.session.authenticate('a', 'password');
  const firstCatalog = runtime.catalog;
  assert.notEqual(firstCatalog, anonymousCatalog);
  await runtime.ensureStorageReady();
  assert.equal(runtime.getSnapshot().state, 'ready');
  await auth.session.signOut();
  assert.notEqual(runtime.catalog, firstCatalog);
  await assert.rejects(firstCatalog.list(), /unavailable/);
  assert.equal(runtime.getSnapshot().state, 'warming');
  await runtime.dispose();
  auth.dispose();
});

test('StrictMode effect cleanup is cancelable, while final release replaces workspace ownership', async () => {
  const auth = new AuthDomain({ fetch: authPort() });
  const runtime = new ApplicationRuntime(auth);
  const initial = runtime.catalog;
  const release = runtime.acquire();
  release();
  const releaseRemount = runtime.acquire();
  await new Promise(resolve => setTimeout(resolve, 10));
  assert.equal(runtime.catalog, initial, 'StrictMode remount must retain the owner');
  releaseRemount();
  await new Promise(resolve => setTimeout(resolve, 10));
  assert.notEqual(runtime.catalog, initial, 'final unmount must release the old cache owner');
  const nextRelease = runtime.acquire();
  await runtime.ensureStorageReady();
  assert.equal(runtime.getSnapshot().state, 'ready', 'a fresh cache must reopen without disposed coordinator fallback');
  nextRelease();
  await runtime.dispose();
  auth.dispose();
});

test('SDK exposes named actions without protocol/persistence owners and fails after disposal', async () => {
  const sdk = createSpreadsheetSdk({ fetch: authPort() });
  for (const forbidden of ['request', 'remote', 'persistence', 'resolver', 'getAccessToken', 'getCsrfToken']) {
    assert.equal(forbidden in sdk.auth, false);
    assert.equal(forbidden in sdk.workbooks, false);
  }
  await sdk.dispose();
  await assert.rejects(sdk.auth.signOut(), (error: unknown) => error instanceof SdkError && error.code === 'RUNTIME_DISPOSED');
  await assert.rejects(sdk.workbooks.list(), /unavailable/);
});
