import { expect, test, type JSHandle, type Page } from '@playwright/test';
import { readFile, writeFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import { importOoxmlDocument } from '@react-sheets/exchange-excel-ooxml';
import type { SpreadsheetSdk } from '@react-sheets/sdk';
import { installBrowserDiagnostics } from './support/workbook-fixtures';
import { FINANCIAL_FUNCTION_CORPUS } from '../packages/formula-engine/src/fixtures/financial-corpus';

test.use({ trace: 'retain-on-failure', screenshot: 'only-on-failure' });
test.skip(!process.env.SDK_UAT_ENABLED, 'Requires the isolated real Java/H2 SDK UAT service');
test.setTimeout(120_000);
const entryPath = fileURLToPath(import.meta.resolve('@react-sheets/sdk')).replaceAll('\\', '/');
const entryUrl = `/@fs${entryPath.startsWith('/') ? entryPath : `/${entryPath}`}`;
// A real Vite HTML entry runs its normal React transform without mounting Web.
const consumerUrl = '/sdk-consumer.uat.html';
const password = 'Uat-Private-Password-2026';

async function publicSdk(page: Page): Promise<JSHandle<SpreadsheetSdk>> {
  await page.goto(consumerUrl);
  return page.evaluateHandle(async (entryUrl) => {
    const { createSpreadsheetSdk } = await import(/* @vite-ignore */ entryUrl) as typeof import('@react-sheets/sdk');
    const sdk = createSpreadsheetSdk();
    await sdk.auth.initialize();
    return sdk;
  }, entryUrl);
}
async function ownerSdk(page: Page): Promise<JSHandle<SpreadsheetSdk>> {
  const sdk = await publicSdk(page);
  if (await sdk.evaluate(sdk => sdk.auth.getSnapshot().bootstrapRequired)) {
    const token = (await readFile(process.env.SDK_UAT_BOOTSTRAP_FILE!, 'utf8')).trim();
    await sdk.evaluate((sdk, input) => sdk.auth.bootstrap(input.token, 'uat-admin', input.password, 'SDK UAT Admin'), { token, password });
  } else await sdk.evaluate((sdk, password) => sdk.auth.authenticate('uat-admin', password), password);
  return sdk;
}
async function calculations(sdk: JSHandle<SpreadsheetSdk>, id: string) {
  return sdk.evaluate(async (sdk, id) => {
    const workbook = await sdk.workbooks.open(id);
    const links = await workbook.externalLinks.refresh();
    const cells = await Promise.all(['B1', 'B2', 'B3', 'B4', 'B5'].map(address => workbook.worksheets.at(0).cells.get(address).read()));
    return { links, values: cells.map(cell => cell.calculatedValue), formulas: cells.map(cell => cell.formula) };
  }, id);
}

test('O1.2-a/b/c/d: canonical range cut rewrites moved inputs, undoes overwritten values and survives native reimport', async ({ page }) => {
  const diagnostics = installBrowserDiagnostics(page), sdk = await ownerSdk(page);
  try {
    const result = await sdk.evaluate(async sdk => {
      const entry = await sdk.workbooks.create({ name: `Range cut ${Date.now()}` });
      let workbook = await sdk.workbooks.open(entry.unitId), sheet = workbook.worksheets.at(0);
      const dependent = await workbook.worksheets.add({ name: 'Dependent', rowCount: 20, columnCount: 20 }); await workbook.flush();
      await sheet.ranges.get('A1:B1').setInputs([[{ kind: 'value', value: 7 }, { kind: 'formula', formula: '=A1+$A$1+H1' }]]);
      await sheet.cells.get('H1').setValue(11); await sheet.cells.get('F1').setFormula('=$A$1'); await sheet.cells.get('G1').setFormula('=B1');
      await sheet.ranges.get('C3:D3').setValues([[88, 99]]); await dependent.cells.get('A1').setFormula(`=${sheet.name}!B1`); await workbook.flush();
      const read = async () => ({ source: await sheet.ranges.get('A1:B1').readValues(), target: await sheet.ranges.get('C3:D3').readValues(),
        formula: (await sheet.cells.get('D3').read()).formula, external: await sheet.ranges.get('F1:H1').readValues(),
        absolute: (await sheet.cells.get('F1').read()).formula, other: await dependent.cells.get('A1').read() });
      const before = await read();
      await sheet.ranges.get('A1:B1').moveTo(sheet.ranges.get('C3:D3')); await workbook.flush(); const moved = await read();
      await workbook.undo(); await workbook.flush(); const undone = await read();
      await workbook.redo(); await workbook.flush(); const redone = await read();
      await sheet.ranges.get('D3').moveTo(sheet.ranges.get('J1')); await workbook.flush();
      const formulaOnly = await sheet.cells.get('J1').read(), formulaOnlyExternal = await sheet.cells.get('G1').read();
      await workbook.undo(); await workbook.flush();
      const rejected: string[] = [];
      for (const action of [
        () => sheet.ranges.get('C3:D3').moveTo(sheet.ranges.get('E5')),
        () => sheet.ranges.get('C3:D3').moveTo(dependent.ranges.get('C3:D3')),
        () => sheet.ranges.get('C3:D3').moveTo(sheet.ranges.get('D3:E3')),
      ]) { try { await action(); rejected.push('unexpected-success'); } catch (error) { rejected.push((error as { code?: string }).code ?? ''); } }
      const afterRejected = await read(); await workbook.save(); const ids = { sheet: sheet.id, dependent: dependent.id }; workbook.close();
      workbook = await sdk.workbooks.open(entry.unitId); sheet = workbook.worksheets.byId(ids.sheet);
      const reopened = { target: await sheet.ranges.get('C3:D3').readValues(), formula: (await sheet.cells.get('D3').read()).formula,
        external: await sheet.ranges.get('F1:H1').readValues(), other: await workbook.worksheets.byId(ids.dependent).cells.get('A1').read() };
      const output = await sdk.workbooks.exportWorkbook(entry.unitId, { fileName: 'sdk-range-cut.xlsx' });
      return { before, moved, undone, redone, formulaOnly, formulaOnlyExternal, rejected, afterRejected, reopened, bytes: Array.from(new Uint8Array(output.buffer)) };
    });
    expect(result.before.source).toEqual([[7, 25]]); expect(result.before.target).toEqual([[88, 99]]);
    expect(result.moved.source).toEqual([[null, null]]); expect(result.moved.target).toEqual([[7, 25]]);
    expect(result.moved.formula).toBe('=C3+$C$3+H1'); expect(result.moved.absolute).toBe('=$C$3');
    expect(result.moved.external).toEqual([[7, 25, 11]]); expect(result.moved.other.formula).toBe('=Sheet1!D3'); expect(result.moved.other.calculatedValue).toBe(25);
    expect(result.undone).toEqual(result.before); expect(result.redone).toEqual(result.moved);
    expect(result.formulaOnly.formula).toBe('=C3+$C$3+H1'); expect(result.formulaOnly.calculatedValue).toBe(25);
    expect(result.formulaOnlyExternal.formula).toBe('=J1'); expect(result.formulaOnlyExternal.calculatedValue).toBe(25);
    expect(result.rejected).toEqual(['INVALID_ARGUMENT', 'UNSUPPORTED_FEATURE', 'REQUEST_REJECTED']); expect(result.afterRejected).toEqual(result.moved);
    expect(result.reopened.target).toEqual(result.moved.target); expect(result.reopened.formula).toBe(result.moved.formula);
    expect(result.reopened.external).toEqual(result.moved.external); expect(result.reopened.other.formula).toBe(result.moved.other.formula); expect(result.reopened.other.calculatedValue).toBe(25);
    const bytes = Uint8Array.from(result.bytes); await writeFile(path.join(process.env.SDK_UAT_EVIDENCE_DIR!, 'sdk-range-cut.xlsx'), bytes);
    const native = await importOoxmlDocument({ fileName: 'sdk-range-cut.xlsx', buffer: bytes.buffer, options: { compatibilityTarget: 'B' } });
    expect(native.snapshot.sheets[0]!.cells['2']!['3']!.formula).toBe('=C3+$C$3+H1');
    expect(native.snapshot.sheets[0]!.cells['0']!['5']!.formula).toBe('=$C$3'); expect(native.snapshot.sheets[1]!.cells['0']!['0']!.formula).toBe('=Sheet1!D3');
    const reimported = await sdk.evaluate(async (sdk, bytes) => {
      const imported = await sdk.workbooks.importWorkbook({ fileName: 'sdk-range-cut.xlsx', buffer: Uint8Array.from(bytes).buffer, options: { compatibilityTarget: 'B' } });
      const workbook = await sdk.workbooks.open(imported.entry.unitId);
      return { target: await workbook.worksheets.at(0).ranges.get('C3:D3').readValues(), other: (await workbook.worksheets.at(1).cells.get('A1').read()).calculatedValue };
    }, result.bytes);
    expect(reimported.target).toEqual([[7, 25]]); expect(reimported.other).toBe(25); diagnostics.assertClean();
  } finally { await sdk.evaluate(sdk => sdk.dispose()); }
});

test('F1.1-e: twelve financial functions save, reopen, export and reimport through public SDK; source faults remain visible', async ({ page }) => {
  const diagnostics = installBrowserDiagnostics(page), sdk = await ownerSdk(page);
  const fixtures = Object.values(FINANCIAL_FUNCTION_CORPUS), formulas = fixtures.map(fixture => fixture.formula);
  try {
    const result = await sdk.evaluate(async (sdk, formulas) => {
      const entry = await sdk.workbooks.create({ name: `Financial functions ${Date.now()}` });
      let workbook = await sdk.workbooks.open(entry.unitId), sheet = workbook.worksheets.at(0);
      await sheet.cells.get('C1').setValue(77); await workbook.flush();
      await sheet.ranges.get(`A1:A${formulas.length}`).setFormulas(formulas.map(formula => [formula])); await workbook.flush();
      const initial = await sheet.ranges.get(`A1:A${formulas.length}`).readValues();
      await workbook.save(); workbook.close();
      workbook = await sdk.workbooks.open(entry.unitId); sheet = workbook.worksheets.at(0);
      const reopened = await sheet.ranges.get(`A1:A${formulas.length}`).readValues(), unaffected = (await sheet.cells.get('C1').read()).value;
      const output = await sdk.workbooks.exportWorkbook(entry.unitId, { fileName: 'sdk-financial.xlsx' });
      return { id: entry.unitId, initial, reopened, unaffected, bytes: Array.from(new Uint8Array(output.buffer)) };
    }, formulas);
    const assertNumbers = (values: readonly (readonly unknown[])[]) => {
      expect(values).toHaveLength(fixtures.length);
      values.forEach((row, index) => { expect(row).toHaveLength(1); expect(typeof row[0]).toBe('number'); expect(row[0] as number).toBeCloseTo(fixtures[index]!.expected, 10); });
    };
    assertNumbers(result.initial); assertNumbers(result.reopened); expect(result.unaffected).toBe(77);
    const bytes = Uint8Array.from(result.bytes);
    await writeFile(path.join(process.env.SDK_UAT_EVIDENCE_DIR!, 'sdk-financial.xlsx'), bytes);
    const native = await importOoxmlDocument({ fileName: 'sdk-financial.xlsx', buffer: bytes.buffer, options: { compatibilityTarget: 'B' } });
    for (let row = 0; row < formulas.length; row++) expect(native.snapshot.sheets[0]!.cells[String(row)]!['0']!.formula).toBe(formulas[row]);
    const reimported = await sdk.evaluate(async (sdk, input) => {
      const imported = await sdk.workbooks.importWorkbook({ fileName: 'sdk-financial.xlsx', buffer: Uint8Array.from(input.bytes).buffer, options: { compatibilityTarget: 'B' } });
      const workbook = await sdk.workbooks.open(imported.entry.unitId);
      return workbook.worksheets.at(0).ranges.get(`A1:A${input.count}`).readValues();
    }, { bytes: result.bytes, count: fixtures.length });
    assertNumbers(reimported);
    const linked = await sdk.evaluate(async (sdk, id) => {
      const entry = await sdk.workbooks.create({ name: `Rates ${Date.now()}` }), source = await sdk.workbooks.open(entry.unitId), rates = source.worksheets.at(0);
      await rates.cells.get('A1').setValue(0.5); await source.flush();
      const target = await sdk.workbooks.open(id); await target.externalLinks.bind(source, 'Rates.xlsx');
      const formula = `=PMT('[Rates.xlsx]${rates.name}'!A1,2,250)`;
      await target.worksheets.at(0).ranges.get('B1:B2').setFormulas([[formula], [`=IFERROR(${formula.slice(1)},7)`]]); await target.flush();
      const initial = await target.worksheets.at(0).ranges.get('B1:B2').readValues();
      await rates.cells.get('A1').setValue(0); await source.flush(); await target.externalLinks.refresh();
      const updated = await target.worksheets.at(0).ranges.get('B1:B2').readValues();
      await sdk.workbooks.moveToTrash(source.id); await target.externalLinks.refresh();
      const unavailable = await target.worksheets.at(0).ranges.get('B1:B2').readValues();
      await sdk.workbooks.restore(entry.unitId); await target.externalLinks.refresh();
      const restored = await target.worksheets.at(0).ranges.get('B1:B2').readValues();
      return { sourceId: entry.unitId, initial, updated, unavailable, restored };
    }, result.id);
    expect(linked.initial).toEqual([[-225], [-225]]); expect(linked.updated).toEqual([[-125], [-125]]);
    for (const row of linked.unavailable) expect(row[0]).toMatchObject({ kind: 'error', code: '#REF!', inputFault: { source: linked.sourceId } });
    expect(linked.restored).toEqual([[-125], [-125]]);
    diagnostics.assertClean();
  } finally { await sdk.evaluate(sdk => sdk.dispose()); }
});

test('MWB-02.c/d: five cross-workbook functions refresh real authorized versions, revoke and recover source access', async ({ page, browser }) => {
  const ownerDiagnostics = installBrowserDiagnostics(page);
  const owner = await ownerSdk(page);
  const readerContext = await browser.newContext({ baseURL: 'http://127.0.0.1:4180' });
  const readerPage = await readerContext.newPage();
  const readerDiagnostics = installBrowserDiagnostics(readerPage);
  const rejected: { path: string; status: number }[] = [];
  readerPage.on('response', response => { if (response.status() >= 400) rejected.push({ path: new URL(response.url()).pathname, status: response.status() }); });
  let reader: JSHandle<SpreadsheetSdk> | undefined;
  try {
    const setup = await owner.evaluate(async (sdk, input) => {
      await sdk.identity.createUser({ username: input.username, displayName: 'Workbook reader', password: input.password });
      const user = (await sdk.identity.listUsers()).find(user => user.username === input.username)!;
      const sourceEntry = await sdk.workbooks.create({ name: `MWB source ${input.username}` });
      const targetEntry = await sdk.workbooks.create({ name: `MWB target ${input.username}` });
      const [source, target] = await Promise.all([sdk.workbooks.open(sourceEntry.unitId), sdk.workbooks.open(targetEntry.unitId)]);
      const sheet = source.worksheets.at(0);
      await sheet.cells.get('A1').setValue(10); await sheet.cells.get('A2').setValue(20);
      await target.externalLinks.bind(source, 'Source.xlsx');
      const formulas = ['SUM', 'AVERAGE', 'COUNT', 'MIN', 'MAX'].map(fn => `=${fn}('[Source.xlsx]${sheet.name}'!A1:A2)`);
      for (let row = 0; row < formulas.length; row++) await target.worksheets.at(0).cells.get(`B${row + 1}`).setFormula(formulas[row]!);
      await target.save();
      await sdk.workbooks.grantAccess(source.id, user.id, 'viewer');
      await sdk.workbooks.grantAccess(target.id, user.id, 'editor');
      return { sourceId: source.id, targetId: target.id, readerId: user.id, username: input.username, formulas };
    }, { username: `mwb-reader-${Date.now()}`, password });
    reader = await publicSdk(readerPage);
    await reader.evaluate((sdk, input) => sdk.auth.authenticate(input.username, input.password), { username: setup.username, password });
    const initial = await calculations(reader, setup.targetId);
    expect(initial.values).toEqual([30, 15, 2, 10, 20]); expect(initial.formulas).toEqual(setup.formulas);
    await owner.evaluate(async (sdk, id) => { const source = await sdk.workbooks.open(id); await source.worksheets.at(0).cells.get('A2').setValue(40); await source.flush(); }, setup.sourceId);
    const updated = await calculations(reader, setup.targetId);
    expect(updated.values).toEqual([50, 25, 2, 10, 40]);
    expect(updated.links[0]!.sourceRevision).toBeGreaterThan(initial.links[0]!.sourceRevision);
    await owner.evaluate((sdk, input) => sdk.workbooks.revokeAccess(input.sourceId, input.readerId), setup);
    const deniedGraphResponse = readerPage.waitForResponse(response => response.ok() && new URL(response.url()).pathname === `/api/workbooks/${setup.targetId}/external-calculation/inputs`);
    const denied = await calculations(reader, setup.targetId);
    const deniedGraph = await (await deniedGraphResponse).json();
    const deniedNode = deniedGraph.nodes.find((node: { unitId: string }) => node.unitId === setup.sourceId);
    expect(deniedNode).toMatchObject({ state: 'denied', error: { code: 'FORBIDDEN' } });
    expect(deniedNode).not.toHaveProperty('snapshot');
    expect(denied.links[0]!.state).toBe('denied'); expect(denied.formulas).toEqual(setup.formulas);
    for (const value of denied.values) expect(value).toMatchObject({ kind: 'error', code: '#BLOCKED!' });
    await owner.evaluate((sdk, input) => sdk.workbooks.grantAccess(input.sourceId, input.readerId, 'viewer'), setup);
    const restored = await calculations(reader, setup.targetId);
    expect(restored.links[0]!.state).toBe('connected'); expect(restored.values).toEqual([50, 25, 2, 10, 40]);
    expect(rejected).toEqual([]);
    expect(readerDiagnostics.pageErrors).toEqual([]); expect(readerDiagnostics.requestFailures).toEqual([]);
    expect(readerDiagnostics.consoleErrors).toEqual([]);
    ownerDiagnostics.assertClean();
  } finally { if (reader) await reader.evaluate(sdk => sdk.dispose()); await readerContext.close(); await owner.evaluate(sdk => sdk.dispose()); }
});

test('MWB-01.b: real SDK subject switch and active-workbook disposal retire old objects', async ({ page, browser }) => {
  const ownerDiagnostics = installBrowserDiagnostics(page);
  const owner = await ownerSdk(page);
  const readerContext = await browser.newContext({ baseURL: 'http://127.0.0.1:4180' });
  const readerPage = await readerContext.newPage();
  const readerDiagnostics = installBrowserDiagnostics(readerPage);
  let reader: JSHandle<SpreadsheetSdk> | undefined;
  try {
    const setup = await owner.evaluate(async (sdk, input) => {
      const usernames = [`scope-a-${input.id}`, `scope-b-${input.id}`];
      for (const username of usernames) await sdk.identity.createUser({ username, displayName: username, password: input.password });
      const users = await sdk.identity.listUsers();
      const entry = await sdk.workbooks.create({ name: `Scope retirement ${input.id}` });
      const workbook = await sdk.workbooks.open(entry.unitId);
      await workbook.worksheets.at(0).cells.get('A1').setValue(42); await workbook.flush();
      for (const username of usernames) await sdk.workbooks.grantAccess(workbook.id, users.find(user => user.username === username)!.id, 'editor');
      return { unitId: workbook.id, usernames };
    }, { id: Date.now(), password });
    reader = await publicSdk(readerPage);
    const result = await reader.evaluate(async (sdk, input) => {
      await sdk.auth.authenticate(input.usernames[0]!, input.password);
      const first = await sdk.workbooks.open(input.unitId), cell = first.worksheets.at(0).cells.get('A1');
      await cell.setValue(43); await first.flush();
      const firstSubject = sdk.auth.getSnapshot().subject;
      await sdk.auth.authenticate(input.usernames[1]!, input.password);
      const secondSubject = sdk.auth.getSnapshot().subject;
      let subjectRetired = '';
      try { await cell.read(); } catch (cause) { subjectRetired = (cause as { code?: string }).code ?? ''; }
      const second = await sdk.workbooks.open(input.unitId), secondCell = second.worksheets.at(0).cells.get('A1');
      const value = (await secondCell.read()).value;
      const freshOwner = first !== second;
      await sdk.dispose();
      let sdkRetired = '';
      try { await secondCell.read(); } catch (cause) { sdkRetired = (cause as { code?: string }).code ?? ''; }
      return { firstSubject, secondSubject, subjectRetired, sdkRetired, value, freshOwner };
    }, { ...setup, password });
    expect(result.firstSubject).not.toBe(result.secondSubject); expect(result.firstSubject).toBeTruthy(); expect(result.secondSubject).toBeTruthy();
    expect(result.subjectRetired).toBe('RUNTIME_DISPOSED'); expect(result.sdkRetired).toBe('RUNTIME_DISPOSED');
    expect(result.freshOwner).toBe(true); expect(result.value).toBe(43);
    ownerDiagnostics.assertClean(); readerDiagnostics.assertClean();
  } finally { if (reader) await reader.evaluate(sdk => sdk.dispose()); await readerContext.close(); await owner.evaluate(sdk => sdk.dispose()); }
});

test('MWB-03.a/b/c: three-workbook graph propagates real source commits and authorization without manual refresh', async ({ page, browser }) => {
  const ownerDiagnostics = installBrowserDiagnostics(page);
  const owner = await ownerSdk(page);
  const readerContext = await browser.newContext({ baseURL: 'http://127.0.0.1:4180' });
  const readerPage = await readerContext.newPage();
  const readerDiagnostics = installBrowserDiagnostics(readerPage);
  const subscriptions: string[] = [];
  readerPage.on('websocket', socket => socket.on('framesent', frame => {
    const message = JSON.parse(String(frame.payload)) as { type?: string; unitId?: string };
    if (message.type === 'calculation.subscribe' && message.unitId) subscriptions.push(message.unitId);
  }));
  const rejected: { path: string; status: number }[] = [];
  page.on('response', response => { if (response.status() >= 400) rejected.push({ path: new URL(response.url()).pathname, status: response.status() }); });
  let reader: JSHandle<SpreadsheetSdk> | undefined;
  try {
    const setup = await owner.evaluate(async (sdk, input) => {
      await sdk.identity.createUser({ username: input.username, displayName: 'Dependency reader', password: input.password });
      const user = (await sdk.identity.listUsers()).find(user => user.username === input.username)!;
      const entries = await Promise.all(['Leaf', 'Middle', 'Root'].map(name => sdk.workbooks.create({ name: `${name} ${input.username}` })));
      const [a, b, c] = await Promise.all(entries.map(entry => sdk.workbooks.open(entry.unitId)));
      await a!.worksheets.at(0).cells.get('A1').setValue(10); await a!.flush();
      await b!.externalLinks.bind(a!, 'A.xlsx');
      await b!.worksheets.at(0).cells.get('A1').setFormula(`='[A.xlsx]${a!.worksheets.at(0).name}'!A1*2`); await b!.flush();
      await c!.externalLinks.bind(b!, 'B.xlsx');
      const reference = `'[B.xlsx]${b!.worksheets.at(0).name}'!A1`;
      const formulas = [`=SUM(${reference})`, `=COUNT(${reference})`, `=IFERROR(${reference},0)`];
      for (let index = 0; index < formulas.length; index++) await c!.worksheets.at(0).cells.get(`B${index + 1}`).setFormula(formulas[index]!);
      await c!.save();
      for (const workbook of [a!, b!]) await sdk.workbooks.grantAccess(workbook.id, user.id, 'viewer');
      await sdk.workbooks.grantAccess(c!.id, user.id, 'editor');
      const result = { a: a!.id, b: b!.id, c: c!.id, readerId: user.id, username: input.username, formulas };
      await a!.close(); await b!.close();
      return result;
    }, { username: `dag-reader-${Date.now()}`, password });
    reader = await publicSdk(readerPage);
    await reader.evaluate((sdk, input) => sdk.auth.authenticate(input.username, input.password), { username: setup.username, password });
    const graphPath = `/api/workbooks/${setup.c}/external-calculation/inputs`;
    const initialGraphResponse = readerPage.waitForResponse(response => response.ok() && new URL(response.url()).pathname === graphPath);
    const readRoot = () => reader!.evaluate(async (sdk, id) => {
      const workbook = await sdk.workbooks.open(id);
      const cells = await Promise.all(['B1', 'B2', 'B3'].map(address => workbook.worksheets.at(0).cells.get(address).read()));
      return { values: cells.map(cell => cell.calculatedValue), formulas: cells.map(cell => cell.formula) };
    }, setup.c);
    expect((await readRoot()).values).toEqual([20, 1, 20]);
    const initialGraph = await (await initialGraphResponse).json();
    expect(initialGraph.nodes).toHaveLength(3);
    const leafRevision = initialGraph.nodes.find((node: { unitId: string }) => node.unitId === setup.a).revision as number;
    await expect.poll(() => subscriptions.includes(setup.c)).toBe(true);

    // This listener must observe the server-triggered request before any Cell.read,
    // because a public read independently verifies fresh graph inputs.
    const automaticUpdate = readerPage.waitForResponse(async response => {
      if (!response.ok() || new URL(response.url()).pathname !== graphPath) return false;
      const graph = await response.json();
      return graph.nodes.some((node: { unitId: string; state: string; revision: number }) => node.unitId === setup.a && node.state === 'connected' && node.revision > leafRevision);
    });
    await owner.evaluate(async (sdk, id) => {
      const source = await sdk.workbooks.open(id);
      await source.worksheets.at(0).cells.get('A1').setValue(40); await source.flush(); await source.close();
    }, setup.a);
    await automaticUpdate;
    const updated = await readRoot(); expect(updated.values).toEqual([80, 1, 80]); expect(updated.formulas).toEqual(setup.formulas);

    const automaticRevoke = readerPage.waitForResponse(async response => {
      if (!response.ok() || new URL(response.url()).pathname !== graphPath) return false;
      return (await response.json()).nodes.some((node: { unitId: string; state: string }) => node.unitId === setup.a && node.state === 'denied');
    });
    await owner.evaluate((sdk, input) => sdk.workbooks.revokeAccess(input.a, input.readerId), setup);
    const revoked = await (await automaticRevoke).json();
    expect(revoked.nodes.find((node: { unitId: string }) => node.unitId === setup.a)).not.toHaveProperty('snapshot');
    for (const value of (await readRoot()).values) expect(value).toMatchObject({ kind: 'error', code: '#BLOCKED!' });

    const automaticRestore = readerPage.waitForResponse(async response => {
      if (!response.ok() || new URL(response.url()).pathname !== graphPath) return false;
      return (await response.json()).nodes.some((node: { unitId: string; state: string }) => node.unitId === setup.a && node.state === 'connected');
    });
    await owner.evaluate((sdk, input) => sdk.workbooks.grantAccess(input.a, input.readerId, 'viewer'), setup);
    await automaticRestore;
    expect((await readRoot()).values).toEqual([80, 1, 80]);
    await reader.evaluate(async (sdk, id) => { const workbook = await sdk.workbooks.open(id); await workbook.save(); await workbook.close(); }, setup.c);
    expect((await readRoot()).values).toEqual([80, 1, 80]);

    const before = await (await page.request.get(`/api/workbooks/${setup.a}/snapshot`)).json();
    const cycle = await owner.evaluate(async (sdk, input) => {
      const [a, c] = await Promise.all([sdk.workbooks.open(input.a), sdk.workbooks.open(input.c)]);
      try { await a.externalLinks.bind(c, 'C.xlsx'); return 'unexpected-success'; }
      catch (cause) { return (cause as { code?: string }).code; }
    }, setup);
    expect(cycle).toBe('CIRCULAR_DEPENDENCY');
    expect(await (await page.request.get(`/api/workbooks/${setup.a}/snapshot`)).json()).toEqual(before);
    expect((await readRoot()).values).toEqual([80, 1, 80]);
    const oldSourceCell = await owner.evaluateHandle(async (sdk, id) => (await sdk.workbooks.open(id)).worksheets.at(0).cells.get('A1'), setup.a);
    const oldSourceCode = () => oldSourceCell.evaluate(async cell => { try { await cell.read(); return 'alive'; } catch (cause) { return (cause as { code?: string }).code; } });

    const waitLeafState = (state: 'broken' | 'connected') => readerPage.waitForResponse(async response => {
      if (!response.ok() || new URL(response.url()).pathname !== graphPath) return false;
      return (await response.json()).nodes.some((node: { unitId: string; state: string }) => node.unitId === setup.a && node.state === state);
    });
    const automaticTrash = waitLeafState('broken');
    await owner.evaluate((sdk, id) => sdk.workbooks.moveToTrash(id), setup.a);
    const trashed = await (await automaticTrash).json();
    await expect.poll(oldSourceCode).toBe('RUNTIME_DISPOSED');
    expect(trashed.nodes.find((node: { unitId: string }) => node.unitId === setup.a)).not.toHaveProperty('snapshot');
    for (const value of (await readRoot()).values) expect(value).toMatchObject({ kind: 'error', code: '#REF!' });
    const automaticUntrash = waitLeafState('connected');
    await owner.evaluate((sdk, id) => sdk.workbooks.restore(id), setup.a); await automaticUntrash;
    expect((await readRoot()).values).toEqual([80, 1, 80]);
    expect(await owner.evaluate(async (sdk, id) => (await (await sdk.workbooks.open(id)).worksheets.at(0).cells.get('A1').read()).value, setup.a)).toBe(40);
    expect(await oldSourceCode()).toBe('RUNTIME_DISPOSED');
    const secondTrash = waitLeafState('broken');
    await owner.evaluate((sdk, id) => sdk.workbooks.moveToTrash(id), setup.a); await secondTrash;
    const automaticPurge = readerPage.waitForResponse(async response => {
      if (!response.ok() || new URL(response.url()).pathname !== graphPath) return false;
      return (await response.json()).nodes.some((node: { unitId: string; error?: { code: string } }) => node.unitId === setup.a && node.error?.code === 'NOT_FOUND');
    });
    await owner.evaluate((sdk, id) => sdk.workbooks.purge(id), setup.a); await automaticPurge;
    for (const value of (await readRoot()).values) expect(value).toMatchObject({ kind: 'error', code: '#REF!' });
    expect(rejected).toEqual([{ path: `/api/workbooks/${setup.a}/external-calculation/binding-validation`, status: 422 }]);
    expect(ownerDiagnostics.consoleErrors.filter(message => !/Failed to load resource:.*422/.test(message))).toEqual([]);
    expect(ownerDiagnostics.pageErrors).toEqual([]); expect(ownerDiagnostics.requestFailures).toEqual([]);
    readerDiagnostics.assertClean();
  } finally { if (reader) await reader.evaluate(sdk => sdk.dispose()); await readerContext.close(); await owner.evaluate(sdk => sdk.dispose()); }
});

test('O1.1-a/b/c/g: public Range matrix, formulas, history, formatting, immutable reads and value copy save/reopen', async ({ page }) => {
  const diagnostics = installBrowserDiagnostics(page), sdk = await ownerSdk(page);
  const rejected: { path: string; status: number }[] = [];
  page.on('response', response => { if (response.status() >= 400) rejected.push({ path: new URL(response.url()).pathname, status: response.status() }); });
  try {
    const result = await sdk.evaluate(async sdk => {
      const entry = await sdk.workbooks.create({ name: `Object ranges ${Date.now()}` });
      let workbook = await sdk.workbooks.open(entry.unitId), sheet = workbook.worksheets.at(0);
      const range = sheet.ranges.get('D8:E9');
      await range.setStyle({ bold: true }, { numberFormat: '0.00' }); await workbook.flush();
      await range.setValues([[1, '=literal'], [true, null]]); await workbook.flush();
      const values = await range.readValues(), presentation = (await range.read())[0]![0]!;
      await workbook.undo(); await workbook.flush(); const undo = await range.readValues();
      await workbook.redo(); await workbook.flush(); const redo = await range.readValues();
      await sheet.ranges.get('G1:G2').setFormulas([['=SUM(D8:E8)'], ['=COUNT(D8:E9)']]); await workbook.flush();
      const formulas = await sheet.ranges.get('G1:G2').readValues();
      const beforeInvalid = await range.readValues(), failures: string[] = [];
      for (const action of [() => range.setValues([[9, 9]]), () => range.setValues([[9, 9], [9, Infinity]]), () => range.setFormulas([['=1', '=2'], ['=3', '4']])]) {
        try { await action(); throw new Error('Invalid matrix unexpectedly accepted'); } catch (error) { const code = (error as { code?: string }).code; if (!code) throw error; failures.push(code); }
      }
      const afterInvalid = await range.readValues();
      const immutable = Object.isFrozen(await range.read()) && Object.isFrozen(presentation.style);
      await range.clear('contents'); await workbook.flush(); const cleared = await range.readValues(), clearStyle = (await range.read())[0]![0]!.style;
      await workbook.undo(); await workbook.flush();
      const targetEntry = await sdk.workbooks.create({ name: `Copy target ${Date.now()}` }), target = await sdk.workbooks.open(targetEntry.unitId);
      await sheet.ranges.get('G1:G2').copyValuesTo(target.worksheets.at(0).ranges.get('B2:B3')); await target.save();
      const copied = await target.worksheets.at(0).ranges.get('B2:B3').read();
      await workbook.save(); workbook.close(); target.close();
      workbook = await sdk.workbooks.open(entry.unitId); sheet = workbook.worksheets.at(0);
      const reopened = await sheet.ranges.get('D8:E9').readValues(), reopenedStyle = (await sheet.ranges.get('D8:E9').read())[0]![0]!.style;
      const targetReload = await sdk.workbooks.open(targetEntry.unitId), copyReload = await targetReload.worksheets.at(0).ranges.get('B2:B3').readValues();
      await sheet.ranges.get('D8:E9').clear('formats'); await workbook.flush(); const formatsCleared = await sheet.ranges.get('D8:E9').read();
      return { values, presentation, undo, redo, formulas, failures, beforeInvalid, afterInvalid, immutable, cleared, clearStyle, copied, reopened, reopenedStyle, copyReload, formatsCleared };
    });
    expect(result.values).toEqual([[1, '=literal'], [true, null]]); expect(result.undo).toEqual([[null, null], [null, null]]); expect(result.redo).toEqual(result.values);
    expect(result.presentation.style?.bold).toBe(true); expect(result.presentation.numberFormat).toBe('0.00'); expect(result.formulas).toEqual([[1], [1]]);
    expect(result.failures).toEqual(['INVALID_ARGUMENT', 'INVALID_ARGUMENT', 'INVALID_ARGUMENT']); expect(result.afterInvalid).toEqual(result.beforeInvalid); expect(result.immutable).toBe(true);
    expect(result.cleared).toEqual([[null, null], [null, null]]); expect(result.clearStyle?.bold).toBe(true);
    expect(result.copied.map(row => row[0]!.value)).toEqual([1, 1]); expect(result.copied.every(row => row[0]!.formula === undefined)).toBe(true);
    expect(result.reopened).toEqual(result.values); expect(result.reopenedStyle?.bold).toBe(true); expect(result.copyReload).toEqual([[1], [1]]);
    expect(result.formatsCleared[0]![0]!.value).toBe(1); expect(result.formatsCleared[0]![0]!.style).toBeUndefined(); expect(result.formatsCleared[0]![0]!.numberFormat).toBeUndefined();
    expect(rejected).toEqual([]); diagnostics.assertClean();
  } finally { await sdk.evaluate(sdk => sdk.dispose()); }
});

test('O1.1-d/e/f: real Java structural planner owns Worksheet identities, references, row/column structure, fill and merge', async ({ page }) => {
  const diagnostics = installBrowserDiagnostics(page), sdk = await ownerSdk(page);
  const rejected: { path: string; status: number }[] = [];
  page.on('response', response => { if (response.status() >= 400) rejected.push({ path: new URL(response.url()).pathname, status: response.status() }); });
  try {
    const result = await sdk.evaluate(async sdk => {
      const entry = await sdk.workbooks.create({ name: `Worksheet objects ${Date.now()}` });
      let workbook = await sdk.workbooks.open(entry.unitId);
      const first = workbook.worksheets.at(0), added = await workbook.worksheets.add({ name: 'Inputs', rowCount: 20, columnCount: 10 }); await workbook.flush();
      await added.ranges.get('A1:B2').setValues([[10, 20], [30, 40]]); await workbook.flush();
      await first.cells.get('A1').setFormula('=Inputs!A1'); await workbook.flush();
      await added.rename('Values'); await workbook.flush(); const renamedFormula = await first.cells.get('A1').read();
      await added.reorder(0); await workbook.flush(); const stable = workbook.worksheets.at(0) === added;
      await added.rows.insert(0, 1); await workbook.flush(); const inserted = await first.cells.get('A1').read();
      await added.rows.delete(0, 1); await workbook.flush();
      await added.columns.insert(0, 1); await workbook.flush(); const columnInserted = await first.cells.get('A1').read();
      await added.columns.delete(0, 1); await workbook.flush();
      await added.rows.setPixels([2], 30); await added.columns.setPixels([3], 90); await workbook.flush();
      await added.rows.setHidden([2], true); await added.columns.setHidden([3], true); await workbook.flush();
      await added.cells.get('D3').setValue(42); await added.setPane({ kind: 'frozen', xSplit: 1, ySplit: 2, startRow: 2, startColumn: 1, state: 'frozen' }); await workbook.flush();
      const dimensions = added.snapshot(), hiddenValue = (await added.cells.get('D3').read()).value;
      const seed = added.ranges.get('F1:F2'), fill = added.ranges.get('F1:F4'); await seed.setValues([[1], [2]]); await workbook.flush();
      await fill.fillFrom(seed, 'down', 'series'); await workbook.flush(); const filled = await fill.readValues();
      await workbook.undo(); await workbook.flush(); const fillUndone = await fill.readValues();
      const merge = added.ranges.get('H1:I1'); await merge.setValues([[10, 20]]); await workbook.flush();
      let mergeRejected = false; try { await merge.merge(); } catch (error) { mergeRejected = (error as { code?: string }).code === 'REQUEST_REJECTED'; }
      await merge.merge({ confirmDataLoss: true }); await workbook.flush(); const merged = await merge.readValues(), merges = added.snapshot().merges;
      await workbook.undo(); await workbook.flush(); const mergeUndo = await merge.readValues();
      const duplicate = await added.duplicate('Copied'); await workbook.flush(); const copied = (await duplicate.cells.get('D3').read()).value;
      const oldRange = duplicate.ranges.get('D3'); await duplicate.remove(); await workbook.flush();
      let deleted = ''; try { await oldRange.read(); } catch (error) { deleted = (error as { code?: string }).code ?? ''; }
      await workbook.undo(); await workbook.flush(); const revived = workbook.worksheets.byId(duplicate.id) === duplicate && (await oldRange.readValues())[0]![0] === 42;
      await workbook.redo(); await workbook.flush(); let removedAgain = ''; try { await oldRange.read(); } catch (error) { removedAgain = (error as { code?: string }).code ?? ''; }
      await workbook.undo(); await workbook.flush(); const revivedAgain = (await oldRange.readValues())[0]![0] === 42;
      await workbook.save(); const ids = { added: added.id, duplicate: duplicate.id, first: first.id }; workbook.close();
      workbook = await sdk.workbooks.open(entry.unitId);
      const reopened = workbook.worksheets.byId(ids.added), reopenedSnapshot = reopened.snapshot(), reopenedValue = (await reopened.cells.get('D3').read()).value;
      return { stable, renamedFormula, inserted, columnInserted, dimensions, hiddenValue, filled, fillUndone, mergeRejected, merged, merges, mergeUndo, copied, deleted, revived, removedAgain, revivedAgain, reopenedSnapshot, reopenedValue, reopenedReference: await workbook.worksheets.byId(ids.first).cells.get('A1').read() };
    });
    expect(result.stable).toBe(true); expect(result.renamedFormula.formula).toBe('=Values!A1'); expect(result.renamedFormula.calculatedValue).toBe(10);
    expect(result.inserted.formula).toBe('=Values!A2'); expect(result.inserted.calculatedValue).toBe(10); expect(result.columnInserted.formula).toBe('=Values!B1'); expect(result.columnInserted.calculatedValue).toBe(10);
    expect(result.dimensions.rowHeightsPx[2]).toBe(30); expect(result.dimensions.columnWidthsPx[3]).toBe(90); expect(result.dimensions.hiddenRows).toContain(2); expect(result.dimensions.hiddenColumns).toContain(3); expect(result.hiddenValue).toBe(42); expect(result.dimensions.pane.kind).toBe('frozen'); expect(result.dimensions.pane).not.toHaveProperty('activePane');
    expect(result.filled).toEqual([[1], [2], [3], [4]]); expect(result.fillUndone).toEqual([[1], [2], [null], [null]]);
    expect(result.mergeRejected).toBe(true); expect(result.merged).toEqual([[10, null]]); expect(result.merges).toHaveLength(1); expect(result.mergeUndo).toEqual([[10, 20]]);
    expect(result.copied).toBe(42); expect(result.deleted).toBe('INVALID_ARGUMENT'); expect(result.revived).toBe(true); expect(result.removedAgain).toBe('INVALID_ARGUMENT'); expect(result.revivedAgain).toBe(true); expect(result.reopenedValue).toBe(42);
    expect(result.reopenedSnapshot.rowHeightsPx[2]).toBe(30); expect(result.reopenedSnapshot.pane.kind).toBe('frozen'); expect(result.reopenedSnapshot.pane).not.toHaveProperty('activePane'); expect(result.reopenedReference.formula).toBe('=Values!A1'); expect(result.reopenedReference.calculatedValue).toBe(10);
    expect(rejected).toEqual([]); diagnostics.assertClean();
  } finally { await sdk.evaluate(sdk => sdk.dispose()); }
});

test('O1.1-b/g: real viewer denies a whole matrix and error-value copies leave the target unchanged', async ({ page, browser }) => {
  const diagnostics = installBrowserDiagnostics(page), owner = await ownerSdk(page);
  const context = await browser.newContext({ baseURL: 'http://127.0.0.1:4180' }), viewerPage = await context.newPage(), viewerDiagnostics = installBrowserDiagnostics(viewerPage);
  let viewer: JSHandle<SpreadsheetSdk> | undefined;
  try {
    const setup = await owner.evaluate(async (sdk, input) => {
      await sdk.identity.createUser({ username: input.username, displayName: 'Matrix viewer', password: input.password });
      const user = (await sdk.identity.listUsers()).find(user => user.username === input.username)!;
      const entry = await sdk.workbooks.create({ name: input.username }), workbook = await sdk.workbooks.open(entry.unitId);
      await workbook.worksheets.at(0).ranges.get('A1:B2').setValues([[1, 2], [3, 4]]); await workbook.flush();
      await sdk.workbooks.grantAccess(workbook.id, user.id, 'viewer');
      return { id: workbook.id, username: input.username };
    }, { username: `matrix-viewer-${Date.now()}`, password });
    viewer = await publicSdk(viewerPage); await viewer.evaluate((sdk, input) => sdk.auth.authenticate(input.username, input.password), { username: setup.username, password });
    const result = await viewer.evaluate(async (sdk, id) => {
      const workbook = await sdk.workbooks.open(id), range = workbook.worksheets.at(0).ranges.get('A1:B2'), before = await range.readValues();
      let code = ''; try { await range.setValues([[10, 20], [30, 40]]); } catch (error) { code = (error as { code?: string }).code ?? ''; }
      let moveCode = ''; try { await range.moveTo(workbook.worksheets.at(0).ranges.get('D4:E5')); } catch (error) { moveCode = (error as { code?: string }).code ?? ''; }
      return { before, after: await range.readValues(), code, moveCode, destination: await workbook.worksheets.at(0).ranges.get('D4:E5').readValues() };
    }, setup.id);
    expect(result.before).toEqual([[1, 2], [3, 4]]); expect(result.after).toEqual(result.before); expect(result.code).toBe('FORBIDDEN'); expect(result.moveCode).toBe('FORBIDDEN'); expect(result.destination).toEqual([[null, null], [null, null]]);
    const copy = await owner.evaluate(async (sdk, id) => {
      const source = await sdk.workbooks.open(id), targetEntry = await sdk.workbooks.create({ name: `Error copy ${Date.now()}` }), target = await sdk.workbooks.open(targetEntry.unitId);
      const to = target.worksheets.at(0).ranges.get('A1:B2'); await to.setValues([[5, 6], [7, 8]]); await target.flush();
      await source.worksheets.at(0).cells.get('A1').setFormula('=1/0'); await source.flush();
      let code = ''; try { await source.worksheets.at(0).ranges.get('A1:B2').copyValuesTo(to); } catch (error) { code = (error as { code?: string }).code ?? ''; }
      return { code, after: await to.readValues() };
    }, setup.id);
    expect(copy.code).toBe('UNSUPPORTED_FEATURE'); expect(copy.after).toEqual([[5, 6], [7, 8]]);
    diagnostics.assertClean(); viewerDiagnostics.assertClean();
  } finally { if (viewer) await viewer.evaluate(sdk => sdk.dispose()); await context.close(); await owner.evaluate(sdk => sdk.dispose()); }
});


test('O2.1-a/b/c/d: public names, rich text and protection preserve canonical state and real native output', async ({ page }) => {
  const diagnostics = installBrowserDiagnostics(page), sdk = await ownerSdk(page);
  try {
    const result = await sdk.evaluate(async sdk => {
      const entry = await sdk.workbooks.create({ name: `O2 objects ${Date.now()}` });
      let workbook = await sdk.workbooks.open(entry.unitId), sheet = workbook.worksheets.at(0);
      await sheet.ranges.get('A1:B1').setValues([[2, 3]]); await workbook.flush();
      await workbook.names.define({ name: 'Rate', scope: 'workbook', formula: '=2', hidden: true, comment: 'Global' }); await workbook.flush();
      const local = await workbook.names.define({ name: 'Rate', scope: 'sheet', sheetId: sheet.id, formula: '=3' }); await workbook.flush();
      await workbook.names.define({ name: 'Relative', scope: 'workbook', formula: '=A1', anchor: { sheetId: sheet.id, row: 3, column: 3 }, comment: 'Anchored' }); await workbook.flush();
      await sheet.cells.get('D1').setFormula('=Rate*10'); await sheet.cells.get('D4').setFormula('=Relative'); await sheet.cells.get('E4').setFormula('=Relative'); await workbook.flush();
      const values = await Promise.all(['D1', 'D4', 'E4'].map(async address => (await sheet.cells.get(address).read()).calculatedValue));
      await local.setFormula('=4'); await workbook.flush(); const changed = (await sheet.cells.get('D1').read()).calculatedValue;
      await local.remove(); await workbook.flush(); const removed = (await sheet.cells.get('D1').read()).calculatedValue;
      await workbook.undo(); await workbook.flush(); const restored = (await sheet.cells.get('D1').read()).calculatedValue;
      await sheet.ranges.get('D8:E9').setRichText('=literal', [{ text: '=lit', style: { bold: true } }, { text: 'eral', style: { italic: true, textColor: '#123456' } }]); await workbook.flush();
      const rich = await sheet.ranges.get('D8:E9').read();
      await workbook.undo(); await workbook.flush(); const richUndo = await sheet.ranges.get('D8:E9').readValues();
      await workbook.redo(); await workbook.flush();
      await sheet.cells.get('E8').setStyle({ locked: false }); await workbook.flush();
      await sheet.protection.set({ id: 'sheet-lock', scope: 'sheet', sheetId: sheet.id, locked: true, allow: { selectLocked: true, selectUnlocked: true, formatCells: true } }); await workbook.flush();
      let protectedCode = ''; try { await sheet.cells.get('D8').setValue(99); } catch (error) { protectedCode = (error as { code?: string }).code ?? ''; }
      await sheet.cells.get('E8').setValue(9); await workbook.flush(); const unlocked = (await sheet.cells.get('E8').read()).value;
      await workbook.undo(); await workbook.flush();
      await sheet.protection.remove('sheet-lock'); await workbook.flush(); await workbook.undo(); await workbook.flush();
      await workbook.redo(); await workbook.flush(); const unprotected = sheet.protection.list(); await workbook.undo(); await workbook.flush();
      await workbook.save(); const oldName = workbook.names.byName('Relative', 'workbook'), oldProtection = sheet.protection; workbook.close();
      let retiredName = '', retiredProtection = ''; try { oldName.snapshot(); } catch (error) { retiredName = (error as { code?: string }).code ?? ''; } try { oldProtection.list(); } catch (error) { retiredProtection = (error as { code?: string }).code ?? ''; }
      workbook = await sdk.workbooks.open(entry.unitId); sheet = workbook.worksheets.at(0);
      const name = workbook.names.byName('Relative', 'workbook').snapshot(), global = workbook.names.byName('Rate', 'workbook').snapshot(), reopenedRich = await sheet.ranges.get('D8:E9').read(), protection = sheet.protection.list();
      const output = await sdk.workbooks.exportWorkbook(entry.unitId, { fileName: 'sdk-o21.xlsx' });
      return { id: entry.unitId, sheetId: sheet.id, names: workbook.names.list().map(name => name.snapshot().name), values, changed, removed, restored, rich, richUndo, protectedCode, unlocked, unprotected, retiredName, retiredProtection, name, global, reopenedRich, protection, bytes: Array.from(new Uint8Array(output.buffer)) };
    });
    expect(result.names).toEqual(['Rate', 'Rate', 'Relative']); expect(result.values).toEqual([30, 2, 3]); expect(result.changed).toBe(40); expect(result.removed).toBe(20); expect(result.restored).toBe(40);
    for (const row of result.rich) for (const cell of row) { expect(cell.value).toBe('=literal'); expect(cell.formula).toBeUndefined(); expect(cell.richText).toEqual([{ text: '=lit', style: { bold: true } }, { text: 'eral', style: { italic: true, textColor: '#123456' } }]); }
    expect(result.richUndo).toEqual([[null, null], [null, null]]); expect(result.protectedCode).toBe('FORBIDDEN'); expect(result.unlocked).toBe(9); expect(result.unprotected).toEqual([]);
    expect(result.retiredName).toBe('RUNTIME_DISPOSED'); expect(result.retiredProtection).toBe('RUNTIME_DISPOSED');
    expect(result.name.anchor).toEqual({ sheetId: result.sheetId, row: 3, column: 3 }); expect(result.name.comment).toBe('Anchored'); expect(result.global).toMatchObject({ scope: 'workbook', hidden: true, comment: 'Global' });
    expect(result.reopenedRich.map(row => row.map(cell => cell.value))).toEqual([['=literal', '=literal'], ['=literal', '=literal']]); expect(result.protection).toEqual([{ id: 'sheet-lock', scope: 'sheet', sheetId: result.sheetId, locked: true, allow: { selectLocked: true, selectUnlocked: true, formatCells: true } }]);
    const bytes = Uint8Array.from(result.bytes); await writeFile(path.join(process.env.SDK_UAT_EVIDENCE_DIR!, 'sdk-o21.xlsx'), bytes);
    const imported = await importOoxmlDocument({ fileName: 'sdk-o21.xlsx', buffer: bytes.buffer, options: { compatibilityTarget: 'B' } });
    expect(imported.snapshot.definedNameModels).toBeDefined();
    expect(imported.snapshot.definedNameModels!.find(name => name.name === 'Relative')?.anchor).toEqual(result.name.anchor);
    expect(imported.snapshot.sheets[0]!.cells['7']!['3']!.value).toBe('=literal'); expect(imported.snapshot.sheets[0]!.cells['7']!['3']!.richText?.map(run => run.text).join('')).toBe('=literal'); expect(imported.snapshot.sheets[0]!.protectionRules).toEqual([{ id: `protection-sheet-1`, scope: 'sheet', sheetId: result.sheetId, locked: true, allow: { selectLocked: true, selectUnlocked: true, formatCells: true, insertRows: false, insertColumns: false, deleteRows: false, deleteColumns: false, sort: false, autoFilter: false, editObjects: false } }]);
    expect(imported.snapshot.definedNameModels!.find(name => name.name === 'Relative')?.comment).toBe('Anchored');
    diagnostics.assertClean();
  } finally { await sdk.evaluate(sdk => sdk.dispose()); }
});

test('O2.1-a/b/c/e: viewer and malformed real protection operations are rejected without a revision or partial write', async ({ page, browser }) => {
  const diagnostics = installBrowserDiagnostics(page), owner = await ownerSdk(page);
  const context = await browser.newContext({ baseURL: 'http://127.0.0.1:4180' }), viewerPage = await context.newPage();
  const viewerDiagnostics = installBrowserDiagnostics(viewerPage); let viewer: JSHandle<SpreadsheetSdk> | undefined;
  try {
    const setup = await owner.evaluate(async (sdk, input) => {
      await sdk.identity.createUser({ username: input.username, displayName: 'O2 viewer', password: input.password });
      const user = (await sdk.identity.listUsers()).find(user => user.username === input.username)!;
      const entry = await sdk.workbooks.create({ name: input.username });
      await sdk.workbooks.grantAccess(entry.unitId, user.id, 'viewer');
      const workbook = await sdk.workbooks.open(entry.unitId), sheet = workbook.worksheets.at(0);
      await sheet.ranges.get('D8:E9').setValues([[1, 2], [3, 4]]); await workbook.flush();
      return { id: entry.unitId, sheetId: sheet.id, username: input.username };
    }, { username: `o2-viewer-${Date.now()}`, password });
    const before = await (await page.request.get(`/api/workbooks/${setup.id}/snapshot`)).json();
    viewer = await publicSdk(viewerPage); await viewer.evaluate((sdk, input) => sdk.auth.authenticate(input.username, input.password), { username: setup.username, password });
    const denied = await viewer.evaluate(async (sdk, input) => {
      const workbook = await sdk.workbooks.open(input.id), sheet = workbook.worksheets.byId(input.sheetId), codes: string[] = [];
      for (const action of [() => sheet.ranges.get('D8:E9').setRichText('x', [{ text: 'x' }]), () => workbook.names.define({ name: 'Forged', scope: 'workbook', formula: '=9' }), () => sheet.protection.set({ id: 'forged', scope: 'sheet', locked: true, allow: {} })]) { try { await action(); } catch (error) { codes.push((error as { code?: string }).code ?? ''); } }
      return { codes, values: await sheet.ranges.get('D8:E9').readValues() };
    }, setup);
    expect(denied.codes).toEqual(['FORBIDDEN', 'FORBIDDEN', 'FORBIDDEN']); expect(denied.values).toEqual([[1, 2], [3, 4]]);
    const rule = { id: 'direct', scope: 'sheet', sheetId: setup.sheetId, locked: true, allow: { sort: 'yes' } };
    const operation = { schema: 'OperationEnvelope', clientSessionId: crypto.randomUUID(), operationId: crypto.randomUUID(), unitId: setup.id, clientSequence: 1, baseRevision: before.revision, createdAt: new Date().toISOString(), mutations: [{ id: 'sheet.protect.set', sheetId: setup.sheetId, params: { sheetId: setup.sheetId, rule } }] };
    const ownerCsrf = (await (await page.request.get('/api/auth/session')).json()).csrfToken;
    const invalid = await page.request.post(`/api/workbooks/${setup.id}/operations`, { headers: { 'X-CSRF-TOKEN': ownerCsrf }, data: operation });
    expect(invalid.status()).toBe(400); const invalidBody = await invalid.json(); expect(invalidBody.code).toBe('VALIDATION_ERROR'); expect(invalidBody.message).toMatch(/allow/i);
    const viewerCsrf = (await (await context.request.get('/api/auth/session')).json()).csrfToken;
    const forbidden = await context.request.post(`/api/workbooks/${setup.id}/operations`, { headers: { 'X-CSRF-TOKEN': viewerCsrf }, data: { ...operation, clientSessionId: crypto.randomUUID(), operationId: crypto.randomUUID(), mutations: [{ ...operation.mutations[0]!, params: { sheetId: setup.sheetId, rule: { ...rule, allow: {} } } }] } });
    expect(forbidden.status()).toBe(403); expect((await forbidden.json()).code).toBe('FORBIDDEN');
    const unownedRestore = await page.request.post(`/api/workbooks/${setup.id}/operations`, { headers: { 'X-CSRF-TOKEN': ownerCsrf }, data: { ...operation, clientSessionId: crypto.randomUUID(), operationId: crypto.randomUUID(), mutations: [{ id: 'name.restore', sheetId: setup.sheetId, params: { model: { name: 'Injected', scope: 'workbook', formula: '=9' }, position: 0 } }] } });
    expect(unownedRestore.status()).toBe(409); expect((await unownedRestore.json()).message).toBe('RESTORE_REQUIRES_OWNED_UNDO');
    expect(await (await page.request.get(`/api/workbooks/${setup.id}/snapshot`)).json()).toEqual(before);
    diagnostics.assertClean(); viewerDiagnostics.assertClean();
  } finally { if (viewer) await viewer.evaluate(sdk => sdk.dispose()); await context.close(); await owner.evaluate(sdk => sdk.dispose()); }
});
