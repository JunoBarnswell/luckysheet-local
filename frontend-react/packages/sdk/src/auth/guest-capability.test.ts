import assert from 'node:assert/strict';
import test from 'node:test';
import { createSpreadsheetSdk } from '../sdk';
import { runtimeFor } from '../sdk';
import { AuthDomain } from './domain';
import { SdkError } from '../error';

const anonymous = { context: null, authenticated: false, subject: null, displayName: null, admin: false, bootstrapRequired: false, csrfToken: 'test-csrf' };
function browserHost(initial: string) {
  const original = Object.getOwnPropertyDescriptor(globalThis, 'window');
  let location = new URL(initial);
  const values = new Map<string, string>();
  Object.defineProperty(globalThis, 'window', { configurable: true, value: {
    get location() { return location; },
    history: { state: null, replaceState(_state: unknown, _title: string, url: string) { location = new URL(url, location); } },
    sessionStorage: { getItem: (key: string) => values.get(key) ?? null, setItem: (key: string, value: string) => { values.set(key, value); }, removeItem: (key: string) => { values.delete(key); } },
  } });
  return { location: () => location, values, navigate: (url: string) => { location = new URL(url); }, restore: () => {
    if (original) Object.defineProperty(globalThis, 'window', original); else Reflect.deleteProperty(globalThis, 'window');
  } };
}
const fetchPort: typeof fetch = async input => Response.json(String(input).endsWith('/config') ? { mode: 'local' } : anonymous);

test('SDK initialization captures, clears and privately retains the workbook-scoped fragment capability', async () => {
  const host = browserHost('https://app.test/workbooks/book?share=legacy&keep=yes#share=secret&keep=anchor');
  const sdk = createSpreadsheetSdk({ fetch: fetchPort });
  try {
    assert.equal(host.location().hash, '#keep=anchor');
    assert.equal(host.location().search, '?keep=yes');
    assert.equal(host.values.get('share:/workbooks/book'), 'secret');
    await sdk.auth.initialize();
    assert.equal(sdk.auth.getSnapshot().phase, 'guest');
    assert.equal(sdk.identity.getSnapshot().phase, 'unverified', 'a captured capability is not a verified principal');
    assert.equal(sdk.auth.getSnapshot().capabilities.canManageUsers, false);
    assert.equal(JSON.stringify(sdk.auth.getSnapshot()).includes('secret'), false);
    assert.equal(JSON.stringify(sdk.identity.getSnapshot()).includes('secret'), false);
    assert.equal('shareTokenProvider' in sdk.auth, false);
    await assert.rejects(sdk.users.listUsers(), (error: unknown) => error instanceof SdkError && error.code === 'FORBIDDEN');
    const current = runtimeFor(sdk).catalog;
    host.navigate('https://app.test/workbooks/other');
    await assert.rejects(current.list(), (error: unknown) => error instanceof SdkError && error.code === 'STALE_OPERATION');
    assert.equal(host.values.has('share:/workbooks/book'), false);
    assert.equal(sdk.auth.getSnapshot().phase, 'anonymous');
    assert.notEqual(runtimeFor(sdk).catalog, current, 'route retirement releases the guest workspace owner');
  } finally { await sdk.dispose(); host.restore(); }
});

test('guest refresh uses the tab credential owner and sign-out revokes old contextual transport', async () => {
  const host = browserHost('https://app.test/workbooks/book#share=secret');
  const first = new AuthDomain({ fetch: fetchPort });
  first.dispose();
  const reloaded = new AuthDomain({ fetch: fetchPort });
  try {
    await reloaded.session.initialize();
    const bound = reloaded.createTransport();
    assert.equal(bound.shareTokenProvider(), 'secret');
    await reloaded.session.signOut();
    assert.equal(reloaded.session.getSnapshot().phase, 'anonymous');
    assert.equal(host.values.has('share:/workbooks/book'), false);
    assert.throws(bound.shareTokenProvider, (error: unknown) => error instanceof SdkError && error.code === 'STALE_OPERATION');
    assert.throws(() => bound.csrfTokenProvider(), (error: unknown) => error instanceof SdkError && error.code === 'STALE_OPERATION');
  } finally { reloaded.dispose(); host.restore(); }
});

test('legacy query credentials are erased and never accepted as a share authority', async () => {
  const host = browserHost('https://app.test/workbooks/book?share=legacy');
  const domain = new AuthDomain({ fetch: fetchPort });
  try {
    await domain.session.initialize();
    assert.equal(host.location().search, '');
    assert.equal(domain.session.getSnapshot().phase, 'anonymous');
    assert.equal(domain.createTransport().shareTokenProvider(), null);
  } finally { domain.dispose(); host.restore(); }
});
