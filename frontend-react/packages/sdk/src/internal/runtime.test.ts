import assert from 'node:assert/strict';
import test from 'node:test';
import { createHash } from 'node:crypto';
import { loadOpcPackageGraph } from '@react-sheets/exchange-excel-ooxml';
import { WorkbookModel } from '@react-sheets/core-model';
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

test('a host without browser Worker fails closed instead of opening a partially initialized editor', async () => {
  const auth = new AuthDomain({ fetch: authPort() });
  const runtime = new ApplicationRuntime(auth);
  const snapshot = new WorkbookModel('loading-boundary', 'Loading boundary').snapshot();
  assert.throws(() => runtime.createSession({ schema: 'WorkbookResolution', unitId: snapshot.unitId, snapshot,
    revision: 0, lifecycle: 'active', source: 'remote', mode: 'remote', localRecord: null,
    binding: { location: 'remote', syncMode: 'remote' },
    access: { unitId: snapshot.unitId, role: 'owner', accessRevision: 0, regions: [] } }), (error: unknown) => error instanceof SdkError && error.code === 'UNSUPPORTED_FEATURE');
  await runtime.dispose();
  auth.dispose();
});


test('SDK catalog export resolves canonical remote assets and rejects missing or corrupt bytes without partial state', async () => {
  const png = Uint8Array.from(Buffer.from('iVBORw0KGgoAAAANSUhEUgAAADAAAAAgCAIAAADbtmxLAAAAN0lEQVR4nO3OQQ0AMAgEMFRgFNOTMRccjyYV0Op5p1R8ICQkJJQeCAkJCaUHQkJCQumBkJDQsg8dwKZ5fgcr3gAAAABJRU5ErkJggg==', 'base64'));
  const hash = createHash('sha256').update(png).digest('hex');
  const asset = { schema: 'AssetRef' as const, assetId: `asset-${hash}`, contentHash: hash, mimeType: 'image/png', byteLength: png.byteLength };
  const workbook = new WorkbookModel('export-assets', 'Remote assets');
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.cells.set(0, 0, { value: null, presentation: { kind: 'image', asset, fit: 'contain' } });
  const snapshot = workbook.snapshot();
  const before = structuredClone(snapshot);
  let assetMode: 'success' | 'missing' | 'metadata' | 'hash' = 'success';
  let assetReads = 0;
  const auth = authPort();
  const fetchPort: typeof fetch = async (input, init) => {
    const path = String(input);
    if (path === '/api/workbooks/export-assets/snapshot') return Response.json({ snapshot, revision: 0 });
    if (path === '/api/workbooks/export-assets/access') return Response.json({ unitId: snapshot.unitId, role: 'owner', accessRevision: 0, regions: [] });
    if (path === '/api/workbooks/export-assets') return Response.json({ unitId: snapshot.unitId, name: snapshot.name });
    if (path.startsWith('/api/workbooks/export-assets/assets/')) {
      assert.equal(path, `/api/workbooks/export-assets/assets/${asset.assetId}`);
      assert.equal(init?.method ?? 'GET', 'GET');
      assetReads++;
      if (assetMode === 'missing') return Response.json({ code: 'NOT_FOUND', message: 'Asset missing' }, { status: 404 });
      const bytes = png.slice();
      if (assetMode === 'hash') bytes[bytes.length - 1] = bytes[bytes.length - 1]! ^ 1;
      return new Response(bytes, { headers: { 'content-type': 'image/png', 'content-length': String(bytes.byteLength),
        'x-content-sha256': assetMode === 'metadata' ? '0'.repeat(64) : hash } });
    }
    return auth(input, init);
  };
  const sdk = createSpreadsheetSdk({ fetch: fetchPort });
  try {
    await sdk.auth.initialize(); await sdk.auth.authenticate('a', 'password');
    const result = await sdk.workbooks.exportWorkbook(snapshot.unitId, { execution: 'inline-test' });
    assert.deepEqual(loadOpcPackageGraph(result.buffer).files[`xl/media/${asset.assetId}.png`], png);
    assert.equal(assetReads, 1);
    for (const mode of ['missing', 'metadata', 'hash'] as const) {
      assetMode = mode;
      await assert.rejects(sdk.workbooks.exportWorkbook(snapshot.unitId, { execution: 'inline-test' }), (error: unknown) =>
        error instanceof SdkError && error.code === 'REQUEST_REJECTED' && error.operation === 'workbooks.export'
        && error.message.includes(snapshot.unitId) && Boolean(error.recovery) && error.cause instanceof Error);
      assert.deepEqual(snapshot, before);
    }
    assetMode = 'success';
    const retried = await sdk.workbooks.exportWorkbook(snapshot.unitId, { execution: 'inline-test' });
    assert.deepEqual(loadOpcPackageGraph(retried.buffer).files[`xl/media/${asset.assetId}.png`], png);
  } finally { await sdk.dispose(); }
});
