import assert from 'node:assert/strict';
import test from 'node:test';
import { AuthDomain } from './domain';
import { SdkError } from '../error';

const anonymous = { context: null, authenticated: false, subject: null, displayName: null, admin: false, bootstrapRequired: false, csrfToken: 'anonymous-csrf' };
const context = { authority: 'local', subject: 'admin-1', principal: 'admin-1', scopeId: 'local', sessionId: 'public-nonce', tenantId: null, appCode: null, employmentId: null, contextVersion: 0, contextId: 'context-admin-1' };
const admin = { context, authenticated: true, subject: 'admin-1', displayName: 'Admin', admin: true, bootstrapRequired: false, csrfToken: 'authenticated-csrf' };
function json(value: unknown, status = 200) { return new Response(JSON.stringify(value), { status, headers: { 'Content-Type': 'application/json' } }); }
function port(responses: Response[]) {
  const requests: { path: string; init?: RequestInit }[] = [];
  const fetch: typeof globalThis.fetch = async (input, init) => {
    requests.push({ path: String(input), init });
    const next = responses.shift();
    assert.ok(next, `unexpected request ${input}`);
    return next;
  };
  return { fetch, requests };
}

test('local login rotates CSRF internally and never publishes credentials', async () => {
  const transport = port([json({ mode: 'local' }), json(anonymous), json({}), json(admin), json([]), json({}), json(anonymous)]);
  const domain = new AuthDomain({ fetch: transport.fetch });
  await Promise.all([domain.session.initialize(), domain.session.initialize()]);
  assert.equal(transport.requests.length, 2, 'initialization must be single-flight');
  await domain.session.authenticate('admin', 'private-password');
  assert.equal(domain.session.getSnapshot().subject, 'admin-1');
  assert.equal(domain.session.getSnapshot().capabilities.canManageUsers, true);
  assert.equal(new Headers(transport.requests[2]?.init?.headers).get('X-CSRF-TOKEN'), 'anonymous-csrf');
  await domain.users.listUsers();
  await domain.session.signOut();
  assert.equal(new Headers(transport.requests[5]?.init?.headers).get('X-CSRF-TOKEN'), 'authenticated-csrf');
  const snapshot = domain.session.getSnapshot();
  assert.equal(snapshot.subject, null);
  assert.equal(snapshot.capabilities.canManageUsers, false);
  for (const forbidden of ['accessToken', 'csrfToken', 'password', 'request', 'getAccessToken', 'getCsrfToken']) {
    assert.equal(forbidden in snapshot, false);
    assert.equal(forbidden in domain.session, false);
  }
  assert.ok(Object.isFrozen(snapshot));
  assert.ok(Object.isFrozen(snapshot.capabilities));
  domain.dispose();
});

test('failed login is typed and cannot grant an identity or admin capability', async () => {
  const transport = port([json({ mode: 'local' }), json(anonymous), json({ message: 'Rejected' }, 401)]);
  const domain = new AuthDomain({ fetch: transport.fetch });
  await domain.session.initialize();
  await assert.rejects(domain.session.authenticate('admin', 'wrong'), (error: unknown) => error instanceof SdkError && error.code === 'UNAUTHENTICATED' && error.status === 401);
  assert.equal(domain.session.getSnapshot().phase, 'anonymous');
  assert.equal(domain.session.getSnapshot().subject, null);
  assert.equal(domain.session.getSnapshot().capabilities.canManageUsers, false);
  domain.dispose();
});

test('bootstrap credentials remain in the one request and never enter the public snapshot', async () => {
  const transport = port([json({ mode: 'local' }), json({ ...anonymous, bootstrapRequired: true }), json({}), json(admin)]);
  const domain = new AuthDomain({ fetch: transport.fetch });
  await domain.session.initialize();
  await domain.session.bootstrap('one-time-token', 'admin', 'private-password', 'Admin');
  assert.equal(transport.requests[2]?.path, '/api/auth/bootstrap');
  assert.equal(domain.session.getSnapshot().phase, 'authenticated');
  assert.equal(JSON.stringify(domain.session.getSnapshot()).includes('one-time-token'), false);
  assert.equal(JSON.stringify(domain.session.getSnapshot()).includes('private-password'), false);
  domain.dispose();
});

test('malformed session revokes existing identity and can recover by reinitialization', async () => {
  const transport = port([json({ mode: 'local' }), json(admin), json({ ...admin, subject: null }), json({ mode: 'local' }), json(anonymous)]);
  const domain = new AuthDomain({ fetch: transport.fetch });
  await domain.session.initialize();
  await assert.rejects(domain.session.signIn(), (error: unknown) => error instanceof SdkError && error.code === 'CONTRACT_INVALID');
  assert.equal(domain.session.getSnapshot().subject, null);
  assert.equal(domain.getCsrfToken(), null);
  // Initialization failures and expired sessions both need a usable retry boundary.
  await domain.session.initialize();
  assert.equal(domain.session.getSnapshot().phase, 'anonymous');
  domain.dispose();
});

test('unknown configuration and invalid local session fail closed and retry succeeds', async () => {
  for (const response of [json({ mode: 'unknown' }), json({ mode: 'local' })]) {
    const transport = port([response, ...(response.clone().status === 200 ? [json({ ...anonymous, admin: true }), json({ mode: 'local' }), json(anonymous)] : [])]);
    const domain = new AuthDomain({ fetch: transport.fetch });
    await domain.session.initialize();
    assert.equal(domain.session.getSnapshot().phase, 'error');
    assert.equal(domain.session.getSnapshot().capabilities.canManageUsers, false);
    assert.equal(domain.getCsrfToken(), null);
    domain.dispose();
  }
  const transport = port([json({ mode: 'unknown' }), json({ mode: 'local' }), json(anonymous)]);
  const domain = new AuthDomain({ fetch: transport.fetch });
  await domain.session.initialize();
  await domain.session.initialize();
  assert.equal(domain.session.getSnapshot().phase, 'anonymous');
  domain.dispose();
});

test('ordinary users cannot dispatch identity mutations or arbitrary HTTP requests', async () => {
  const transport = port([json({ mode: 'local' }), json({ ...admin, admin: false })]);
  const domain = new AuthDomain({ fetch: transport.fetch });
  await domain.session.initialize();
  for (const action of [() => domain.users.listUsers(), () => domain.users.createUser({ username: 'new', displayName: 'New', password: 'private' }),
    () => domain.users.setUserEnabled('other', false), () => domain.users.resetPassword('other', 'private')]) {
    await assert.rejects(action(), (error: unknown) => error instanceof SdkError && error.code === 'FORBIDDEN');
  }
  assert.equal(transport.requests.length, 2);
  assert.equal('request' in domain.session, false);
  domain.dispose();
});

test('identity response is validated and projected without additional secret fields', async () => {
  const user = { id: 'user-1', username: 'user', displayName: 'User', enabled: true, admin: false };
  const transport = port([json({ mode: 'local' }), json(admin), json([{ ...user, passwordHash: 'secret' }]), json([{ ...user, enabled: 'yes' }])]);
  const domain = new AuthDomain({ fetch: transport.fetch });
  await domain.session.initialize();
  assert.deepEqual(await domain.users.listUsers(), [user]);
  await assert.rejects(domain.users.listUsers(), (error: unknown) => error instanceof SdkError && error.code === 'CONTRACT_INVALID');
  domain.dispose();
});

test('network error is observable without granting identity; disposed callbacks cannot publish', async () => {
  const domain = new AuthDomain({ fetch: async () => { throw new Error('network down'); } });
  await domain.session.initialize();
  assert.equal(domain.session.getSnapshot().error?.code, 'SERVICE_UNAVAILABLE');
  assert.ok(domain.session.getSnapshot().error?.recovery);
  domain.dispose();
  await assert.rejects(domain.session.signOut(), (error: unknown) => error instanceof SdkError && error.code === 'RUNTIME_DISPOSED');
});

function verified(principal: string, contextId: string) {
  return { ...admin, admin: false, csrfToken: '', subject: principal, context: { ...context, authority: 'https://issuer.test',
    subject: 'trusted-user', principal, scopeId: 'scope-a', sessionId: 'sso-session', contextId } };
}
function credentialOwner() {
  let credential = { token: 'credential-a', expiresAt: Date.now() + 60000 } as { token: string; expiresAt: number } | null;
  let listener: ((event: import('./contract').CredentialEvent) => void) | null = null;
  let acquisitions = 0;
  return {
    source: { acquire: async () => { acquisitions++; return credential; },
      subscribe: (next: (event: import('./contract').CredentialEvent) => void) => { listener = next; return () => { listener = null; }; },
      signOut: async () => { credential = null; listener?.('signed-out'); } },
    set: (next: typeof credential) => { credential = next; },
    emit: (event: import('./contract').CredentialEvent) => listener?.(event),
    acquisitions: () => acquisitions,
  };
}

test('external bearer is verified by the server and renewal preserves the same context', async () => {
  const owner = credentialOwner();
  const wire = port([json({ mode: 'oidc' }), json(verified('principal-a', 'context-a')), json(verified('principal-a', 'context-a')), json({ done: true })]);
  const domain = new AuthDomain({ source: { kind: 'external-bearer', credentials: owner.source }, fetch: wire.fetch });
  await Promise.all([domain.session.initialize(), domain.session.initialize()]);
  assert.equal(wire.requests.length, 2);
  const bound = domain.createTransport();
  owner.set({ token: 'credential-renewed', expiresAt: Date.now() + 60000 });
  const tokens = await Promise.all([bound.authTokenProvider(), bound.authTokenProvider()]);
  assert.deepEqual(tokens, ['credential-renewed', 'credential-renewed']);
  assert.equal(wire.requests.length, 3, 'concurrent acquisition/verification must be single-flight');
  const response = await bound.fetchImpl('/api/workbooks');
  assert.deepEqual(await response.json(), { done: true });
  assert.equal(new Headers(wire.requests[3]?.init?.headers).get('X-Spreadsheet-Context'), 'context-a');
  assert.equal(wire.requests[1]?.init?.credentials, 'omit');
  assert.equal(domain.session.getSnapshot().subject, 'principal-a', 'caller did not choose the server principal');
  assert.equal(JSON.stringify(domain.session.getSnapshot()).includes('credential-'), false);
  domain.dispose();
});

test('same subject in a changed workspace retires old credential and fetch closures before dispatch', async () => {
  const owner = credentialOwner();
  const wire = port([json({ mode: 'oidc' }), json(verified('principal-a', 'context-a')), json(verified('principal-a', 'context-b'))]);
  const domain = new AuthDomain({ source: { kind: 'external-bearer', credentials: owner.source }, fetch: wire.fetch });
  await domain.session.initialize();
  const old = domain.createTransport();
  owner.set({ token: 'credential-b', expiresAt: Date.now() + 60000 });
  await assert.rejects(old.authTokenProvider(), (error: unknown) => error instanceof SdkError && error.code === 'STALE_OPERATION');
  assert.equal(domain.session.getSnapshot().subject, 'principal-a');
  assert.equal(domain.session.getSnapshot().context?.contextId, 'context-b');
  await assert.rejects(old.fetchImpl('/api/workbooks', { method: 'POST' }), (error: unknown) => error instanceof SdkError && error.code === 'STALE_OPERATION');
  assert.equal(wire.requests.length, 3);
  domain.dispose();
});

test('retired context cannot publish a response body already returned by fetch', async () => {
  const owner = credentialOwner();
  const wire = port([json({ mode: 'oidc' }), json(verified('principal-a', 'context-a')), json({ privateValue: 42 })]);
  const domain = new AuthDomain({ source: { kind: 'external-bearer', credentials: owner.source }, fetch: wire.fetch });
  await domain.session.initialize();
  const response = await domain.createTransport().fetchImpl('/api/workbooks');
  owner.emit('signed-out');
  await assert.rejects(response.json(), (error: unknown) => error instanceof SdkError && error.code === 'STALE_OPERATION');
  assert.equal(domain.session.getSnapshot().context, null);
  domain.dispose();
});

test('expired or rejected external credentials clear identity without dispatching business writes', async () => {
  for (const failure of ['expired', 'server-rejected'] as const) {
    const owner = credentialOwner();
    const wire = port([json({ mode: 'oidc' }), json(verified('principal-a', 'context-a')), ...(failure === 'server-rejected' ? [json({}, 401)] : [])]);
    const domain = new AuthDomain({ source: { kind: 'external-bearer', credentials: owner.source }, fetch: wire.fetch });
    await domain.session.initialize();
    const bound = domain.createTransport();
    owner.set({ token: 'credential-b', expiresAt: Date.now() + (failure === 'expired' ? -1 : 60000) });
    await assert.rejects(bound.authTokenProvider(), (error: unknown) => error instanceof SdkError && error.code === 'UNAUTHENTICATED');
    assert.equal(domain.session.getSnapshot().subject, null);
    assert.equal(domain.getCsrfToken(), null);
    assert.equal(wire.requests.some(request => request.path === '/api/workbooks'), false);
    domain.dispose();
  }
});

test('host-session preserves Gateway prefix, uses host CSRF and never accepts host-selected subject', async () => {
  const wire = port([json({ mode: 'delegated' }), json(verified('server-principal', 'server-context')), json({})]);
  const domain = new AuthDomain({ baseUrl: 'https://gateway.test/api/workspaces/tenant/app/excel',
    collaborationUrl: 'wss://gateway.test/api/workspaces/tenant/app/excel/ws', fetch: wire.fetch,
    source: { kind: 'host-session', session: { getCsrfToken: async () => 'gateway-csrf', subscribe: () => () => {}, signOut: async () => {} } } });
  await domain.session.initialize();
  assert.equal(wire.requests[0]?.path, 'https://gateway.test/api/workspaces/tenant/app/excel/api/auth/config');
  assert.equal(domain.session.getSnapshot().subject, 'server-principal');
  assert.equal(domain.getCsrfToken(), 'gateway-csrf');
  assert.equal(wire.requests[1]?.init?.credentials, 'include');
  assert.equal(new Headers(wire.requests[1]?.init?.headers).has('Authorization'), false);
  assert.equal(new URL(domain.collaborationUrl!).searchParams.get('context'), 'server-context');
  const bound = domain.createTransport();
  await assert.rejects(bound.fetchImpl('https://untrusted.test/api/workbooks'), (error: unknown) => error instanceof SdkError && error.code === 'FORBIDDEN');
  assert.equal(wire.requests.length, 2);
  await assert.rejects(domain.session.authenticate('spoofed', 'secret'), (error: unknown) => error instanceof SdkError && error.code === 'AUTH_CONFIGURATION_ERROR');
  domain.dispose();
});

test('credential source mismatch and malformed verified context fail closed', async () => {
  const owner = credentialOwner();
  const mismatched = new AuthDomain({ source: { kind: 'external-bearer', credentials: owner.source }, fetch: port([json({ mode: 'local' })]).fetch });
  await mismatched.session.initialize();
  assert.equal(mismatched.session.getSnapshot().error?.code, 'AUTH_CONFIGURATION_ERROR');
  mismatched.dispose();
  const wire = port([json({ mode: 'oidc' }), json({ ...verified('principal-a', 'context-a'), context: { ...context, principal: 'different' } })]);
  const domain = new AuthDomain({ source: { kind: 'external-bearer', credentials: owner.source }, fetch: wire.fetch });
  await domain.session.initialize();
  assert.equal(domain.session.getSnapshot().error?.code, 'CONTRACT_INVALID');
  assert.equal(domain.session.getSnapshot().context, null);
  domain.dispose();
});
