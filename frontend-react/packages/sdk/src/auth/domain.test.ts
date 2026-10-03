import assert from 'node:assert/strict';
import test from 'node:test';
import { AuthDomain } from './domain';
import { SdkError } from '../error';

const anonymous = { authenticated: false, subject: null, displayName: null, admin: false, bootstrapRequired: false, csrfToken: 'anonymous-csrf' };
const admin = { authenticated: true, subject: 'admin-1', displayName: 'Admin', admin: true, bootstrapRequired: false, csrfToken: 'authenticated-csrf' };
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
  await domain.identity.listUsers();
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
  for (const action of [() => domain.identity.listUsers(), () => domain.identity.createUser({ username: 'new', displayName: 'New', password: 'private' }),
    () => domain.identity.setUserEnabled('other', false), () => domain.identity.resetPassword('other', 'private')]) {
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
  assert.deepEqual(await domain.identity.listUsers(), [user]);
  await assert.rejects(domain.identity.listUsers(), (error: unknown) => error instanceof SdkError && error.code === 'CONTRACT_INVALID');
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
