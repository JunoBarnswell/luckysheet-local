import assert from 'node:assert/strict';
import test from 'node:test';
import { createSpreadsheetSdk } from '../sdk';
import { SdkError } from '../error';

const context = { authority: 'local', subject: 'admin-1', principal: 'admin-1', scopeId: 'local', sessionId: 'public-nonce', tenantId: null, appCode: null, employmentId: null, contextVersion: 0, contextId: 'context-admin-1' };
const session = { context, authenticated: true, subject: 'admin-1', displayName: 'Admin', admin: true, bootstrapRequired: false, csrfToken: 'private-csrf' };
const anonymous = { context: null, authenticated: false, subject: null, displayName: null, admin: false, bootstrapRequired: false, csrfToken: 'anonymous-csrf' };

test('identity publishes verified context and retirement while administration stays in users', async () => {
  let active = session as typeof session | typeof anonymous;
  const sdk = createSpreadsheetSdk({ fetch: async input => Response.json(String(input).endsWith('/config') ? { mode: 'local' } : active) });
  let notifications = 0;
  const unsubscribe = sdk.identity.subscribe(() => { notifications++; });
  try {
    await sdk.auth.initialize();
    const identity = sdk.identity.getSnapshot();
    assert.equal(identity.phase, 'verified');
    assert.equal(identity.context, sdk.auth.getSnapshot().context);
    assert.ok(Object.isFrozen(identity)); assert.ok(Object.isFrozen(identity.context));
    await sdk.auth.refresh();
    assert.equal(sdk.identity.getSnapshot(), identity, 'renewing the same identity retains the published projection');
    assert.equal(notifications, 1);
    active = anonymous; await sdk.auth.refresh();
    assert.equal(sdk.identity.getSnapshot().context, null);
    assert.equal(notifications, 2);
    await assert.rejects(sdk.users.listUsers(), (error: unknown) => error instanceof SdkError && error.code === 'FORBIDDEN');
    assert.equal('createUser' in sdk.identity, false); assert.equal('getSnapshot' in sdk.users, false);
  } finally { unsubscribe(); await sdk.dispose(); }
});

test('invalid administrative inputs are rejected before network and wire failures retain SdkError status', async () => {
  let reads = 0, status = 403;
  const sdk = createSpreadsheetSdk({ fetch: async input => {
    const path = String(input);
    if (path.endsWith('/config')) return Response.json({ mode: 'local' });
    if (path.endsWith('/session')) return Response.json(session);
    reads++; return Response.json({}, { status });
  } });
  try {
    await sdk.auth.initialize();
    for (const action of [() => sdk.users.createUser({ username: '', displayName: 'New', password: 'private' }),
      () => sdk.users.setUserEnabled('', false), () => sdk.users.resetPassword('user-1', '')]) {
      await assert.rejects(action(), (error: unknown) => error instanceof SdkError && error.code === 'INVALID_ARGUMENT');
    }
    assert.equal(reads, 0);
    await assert.rejects(sdk.users.listUsers(), (error: unknown) => error instanceof SdkError && error.code === 'FORBIDDEN' && error.status === 403);
    status = 503;
    await assert.rejects(sdk.users.listUsers(), (error: unknown) => error instanceof SdkError && error.code === 'SERVICE_UNAVAILABLE' && error.status === 503);
  } finally { await sdk.dispose(); }
});

test('retired administrative response bodies cannot publish under another identity', async () => {
  let reply!: (response: Response) => void, started!: () => void;
  const requested = new Promise<void>(resolve => { started = resolve; });
  let active = session as typeof session | typeof anonymous;
  const sdk = createSpreadsheetSdk({ fetch: async input => {
    const path = String(input);
    if (path.endsWith('/config')) return Response.json({ mode: 'local' });
    if (path.endsWith('/session')) return Response.json(active);
    started(); return new Promise<Response>(resolve => { reply = resolve; });
  } });
  try {
    await sdk.auth.initialize();
    const users = sdk.users.listUsers();
    await requested;
    active = anonymous; await sdk.auth.refresh();
    reply(Response.json([{ id: 'secret-user', username: 'hidden', displayName: 'Private', enabled: true, admin: false }]));
    await assert.rejects(users, (error: unknown) => error instanceof SdkError && error.code === 'STALE_OPERATION');
  } finally { await sdk.dispose(); }
});
