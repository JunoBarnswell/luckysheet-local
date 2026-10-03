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
const password = 'Uat-Private-Password-2026';

async function publicSdk(page: Page): Promise<JSHandle<SpreadsheetSdk>> {
  await page.goto('/workbooks');
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
    const forbiddenResponse = readerPage.waitForResponse(response => response.status() === 403 && new URL(response.url()).pathname === `/api/workbooks/${setup.targetId}/external-links/Source.xlsx/inputs`);
    const denied = await calculations(reader, setup.targetId);
    expect(await (await forbiddenResponse).finished()).toBeNull();
    expect(denied.links[0]!.state).toBe('denied'); expect(denied.formulas).toEqual(setup.formulas);
    for (const value of denied.values) expect(value).toMatchObject({ kind: 'error', code: '#BLOCKED!' });
    await owner.evaluate((sdk, input) => sdk.workbooks.grantAccess(input.sourceId, input.readerId, 'viewer'), setup);
    const restored = await calculations(reader, setup.targetId);
    expect(restored.links[0]!.state).toBe('connected'); expect(restored.values).toEqual([50, 25, 2, 10, 40]);
    expect(rejected).toEqual([{ path: `/api/workbooks/${setup.targetId}/external-links/Source.xlsx/inputs`, status: 403 }]);
    expect(readerDiagnostics.pageErrors).toEqual([]); expect(readerDiagnostics.requestFailures).toEqual([]);
    expect(readerDiagnostics.consoleErrors.filter(message => !message.includes('Failed to load resource: the server responded with a status of 403'))).toEqual([]);
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
