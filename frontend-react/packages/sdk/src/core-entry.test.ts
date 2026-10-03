import assert from 'node:assert/strict';
import { registerHooks } from 'node:module';
import test from 'node:test';

test('the real SDK core entry initializes without resolving React and exposes separate identity and users', async () => {
  const hook = registerHooks({ resolve(specifier, context, next) {
    if (/^react(?:-dom)?(?:\/|$)/.test(specifier)) throw new Error(`SDK core attempted to load ${specifier}`);
    return next(specifier, context);
  } });
  try {
    const { createSpreadsheetSdk } = await import('./index');
    const sdk = createSpreadsheetSdk({ fetch: async input => Response.json(String(input).endsWith('/config') ? { mode: 'local' } : {
      context: null, authenticated: false, subject: null, displayName: null, admin: false, bootstrapRequired: false, csrfToken: 'test-csrf',
    }) });
    try {
      await sdk.auth.initialize();
      assert.deepEqual(sdk.identity.getSnapshot(), { phase: 'unverified', context: null, displayName: null });
      assert.deepEqual(Object.keys(sdk.identity).sort(), ['getSnapshot', 'subscribe']);
      assert.deepEqual(Object.keys(sdk.users).sort(), ['createUser', 'listUsers', 'resetPassword', 'setUserEnabled']);
      assert.equal('useWorkbook' in await import('./index'), false);
      assert.equal('useSdkServices' in await import('./index'), false);
      await assert.rejects(sdk.users.listUsers(), (error: unknown) => error instanceof Error && 'code' in error && error.code === 'FORBIDDEN');
    } finally { await sdk.dispose(); }
    assert.equal(sdk.identity.getSnapshot().phase, 'retired');
  } finally { hook.deregister(); }
});
