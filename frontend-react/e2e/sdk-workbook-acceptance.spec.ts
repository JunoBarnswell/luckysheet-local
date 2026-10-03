import { expect, test, type JSHandle, type Page } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import type { SpreadsheetSdk } from '@react-sheets/sdk';
import { installBrowserDiagnostics } from './support/workbook-fixtures';

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
