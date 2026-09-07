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
const snapshot = { schema: 'WorkbookSnapshot', version: 10, unitId, name: 'CI SQL Session', dimensionMetrics: { normalFontFamily: 'Calibri', normalFontSizePx: 14.6666666667, maximumDigitWidthPx: 7 }, calculationSettings: { mode: 'automatic', iterativeCalculation: false, maximumIterations: 100, maximumChange: 0.001, precisionAsDisplayed: false, calculateBeforeSave: true, fullCalculationOnLoad: false }, editingOptions: { allowEditDirectly: true, moveAfterEnter: true, enterDirection: 'down', formulaAutoComplete: true, valueAutoComplete: true, fixedDecimalPlaces: null }, definedNameModels: [], dataModel: { sources: [], tables: [], relationships: [], views: [] }, sheets: [{ kind: 'worksheet', id: 'sheet-1', name: 'Sheet1', rowCount: 1000, columnCount: 26, cells: {}, merges: [], pane: { kind: 'none' }, defaultRowHeightPx: 20, defaultColumnWidthPx: 64, pivots: [], sparklines: [], drawings: [], drawingPayloads: {}, review: { notesByCell: {}, notesById: {}, threadIdsByCell: {}, threadsById: {} } }] };

const create = await fetch(`${baseUrl}/api/workbooks`, { method: 'POST', headers: auth, body: JSON.stringify({ unitId, name: 'CI SQL Session', snapshot }) });
if (create.status !== 201) throw new Error(`Authenticated workbook create failed: ${create.status} ${await create.text()}`);
const read = await fetch(`${baseUrl}/api/workbooks/${unitId}/snapshot`, { headers: auth });
if (read.status !== 200) throw new Error(`Authenticated snapshot read failed: ${read.status} ${await read.text()}`);
const operation = { schema: 'OperationEnvelope', operationId: `${unitId}-op`, unitId, clientSequence: 1, baseRevision: 0, createdAt: new Date().toISOString(), mutations: [{ id: 'sheet.extent.grow', sheetId: 'sheet-1', params: { rowCount: 1000, columnCount: 26 } }] };
const commit = await fetch(`${baseUrl}/api/workbooks/${unitId}/operations`, { method: 'POST', headers: auth, body: JSON.stringify(operation) });
if (![200, 201].includes(commit.status)) throw new Error(`Authenticated operation commit failed: ${commit.status} ${await commit.text()}`);
const anonymous = await fetch(`${baseUrl}/api/workbooks/${unitId}/snapshot`);
if (anonymous.status !== 401) throw new Error(`Anonymous rejection path expected 401, got ${anonymous.status}`);
console.log(`SQL session verified: create=${create.status} read=${read.status} commit=${commit.status} anonymous=${anonymous.status}`);
jwksServer.close();
