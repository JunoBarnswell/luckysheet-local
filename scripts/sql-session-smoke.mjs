import { createHash, generateKeyPairSync, sign } from 'node:crypto';
import { createServer } from 'node:http';

const baseUrl = process.env.BASE_URL ?? 'http://127.0.0.1:8082';
const issuer = process.env.AUTH_ISSUER ?? 'http://127.0.0.1:8090';
const audience = process.env.AUTH_AUDIENCE ?? 'ci-sql-session';
const { privateKey, publicKey } = generateKeyPairSync('rsa', { modulusLength: 2048 });
const jwk = publicKey.export({ format: 'jwk' });
jwk.kid = 'ci-sql-session'; jwk.alg = 'RS256'; jwk.use = 'sig';
const jwks = JSON.stringify({ keys: [jwk] });
const jwksServer = createServer((request, response) => {
  if (request.url === '/.well-known/jwks.json') { response.writeHead(200, { 'content-type': 'application/json' }); response.end(jwks); return; }
  response.writeHead(404); response.end();
});
await new Promise((resolve) => jwksServer.listen(8090, '127.0.0.1', resolve));

const encoded = (value) => Buffer.from(JSON.stringify(value)).toString('base64url');
const now = Math.floor(Date.now() / 1000);
const header = encoded({ alg: 'RS256', typ: 'JWT', kid: 'ci-sql-session' });
const claims = encoded({ iss: issuer, aud: audience, sub: 'ci-owner', iat: now, exp: now + 300 });
const input = `${header}.${claims}`;
const token = `${input}.${sign('RSA-SHA256', Buffer.from(input), privateKey).toString('base64url')}`;
const auth = { authorization: `Bearer ${token}`, 'content-type': 'application/json' };

async function waitForServer() {
  for (let attempt = 0; attempt < 60; attempt += 1) {
    try { const response = await fetch(`${baseUrl}/health`); if (response.status === 200) return; } catch { /* server is still starting */ }
    await new Promise((resolve) => setTimeout(resolve, 2000));
  }
  throw new Error('SQL session server did not become ready at /health');
}
await waitForServer();
const unitId = `ci-sql-${Date.now()}`;

const create = await fetch(`${baseUrl}/api/workbooks`, { method: 'POST', headers: auth, body: JSON.stringify({ unitId, name: 'CI SQL Session' }) });
if (create.status !== 201) throw new Error(`Authenticated workbook create failed: ${create.status} ${await create.text()}`);
const read = await fetch(`${baseUrl}/api/workbooks/${unitId}/manifest`, { headers: auth });
if (read.status !== 200) throw new Error(`Authenticated manifest read failed: ${read.status} ${await read.text()}`);
const operation = { schema: 'OperationEnvelope', operationId: `${unitId}-op`, unitId, clientSequence: 1, baseRevision: 0, createdAt: new Date().toISOString(), mutations: [{ id: 'cell.set', sheetId: 'sheet-1', params: { sheetId: 'sheet-1', row: 0, column: 0, value: { value: 42 } } }] };
const commit = await fetch(`${baseUrl}/api/workbooks/${unitId}/operations`, { method: 'POST', headers: auth, body: JSON.stringify(operation) });
if (![200, 201].includes(commit.status)) throw new Error(`Authenticated operation commit failed: ${commit.status} ${await commit.text()}`);
const committed = await fetch(`${baseUrl}/api/workbooks/${unitId}/manifest?revision=1`, { headers: auth });
if (committed.status !== 200) throw new Error(`Committed manifest read failed: ${committed.status} ${await committed.text()}`);
const committedManifest = await committed.json();
if (committedManifest.revision !== 1 || !Array.isArray(committedManifest.pages) || committedManifest.pages.length !== 1) {
  throw new Error(`Canonical commit did not publish one revision-one page: ${JSON.stringify(committedManifest)}`);
}
const anonymous = await fetch(`${baseUrl}/api/workbooks/${unitId}/manifest`);
if (anonymous.status !== 401) throw new Error(`Anonymous rejection path expected 401, got ${anonymous.status}`);
console.log(`SQL session verified: create=${create.status} manifest=${read.status} commit=${commit.status} revision=${committedManifest.revision} anonymous=${anonymous.status}`);
jwksServer.close();
