import { WorkbookModel } from '@react-sheets/core-model';
import { exportSnapshotToOoxmlBuffer, loadOpcPackageGraph, zipOpcPartsBuffer } from '../packages/exchange-excel-ooxml/src/archive';
import { importOoxmlDocument } from '../packages/exchange-excel-ooxml/src/import';
import { strFromU8, strToU8 } from 'fflate';
import { readFile, mkdir } from 'node:fs/promises';
import { expect, test, type Page, type BrowserContext } from '@playwright/test';
import { installBrowserDiagnostics } from './support/workbook-fixtures';

// Requires the real Java/H2 service with an isolated data directory; no route mocks.
test.describe('SDK product UAT against Java authority', () => {
  test.skip(!process.env.SDK_UAT_ENABLED, 'Run with SDK_UAT_ENABLED=1 against an isolated real Java/H2 service');
  test.describe.configure({ mode: 'serial' });
  test.setTimeout(120_000);
  const runId = Date.now().toString();
  const password = 'Uat-Private-Password-2026';
  const users = new Map<string, string>();
  const userPasswords = new Map<string, string>();
  const roleStates = new Map<string, Awaited<ReturnType<BrowserContext['storageState']>>>();
  let adminState: Awaited<ReturnType<BrowserContext['storageState']>>;
  let unitId = '';
  let name = `SDK UAT ${runId}`;
  const evidence = process.env.SDK_UAT_EVIDENCE_DIR ?? '/tmp/sdk-uat/evidence';

  async function login(page: Page, username = 'uat-admin', userPassword = password) {
    await page.goto('/workbooks');
    await page.getByLabel('用户名', { exact: true }).fill(username);
    await page.getByLabel('密码', { exact: true }).fill(userPassword);
    await page.getByRole('button', { name: '登录', exact: true }).click();
    await expect(page.getByTestId('workbook-hub')).toBeVisible();
  }
  async function ownerPage(context: BrowserContext, page: Page) {
    await context.addCookies(adminState.cookies);
    await page.goto('/workbooks');
    await expect(page.getByTestId('workbook-hub')).toBeVisible();
  }
  const menu = (page: Page) => page.locator('[data-menu-owners]').filter({ has: page.getByText('文件操作', { exact: true }) });
  async function openMenu(page: Page, workbookName = name) {
    await page.getByRole('button', { name: `打开 ${workbookName} 的更多操作`, exact: true }).click();
  }
  async function screenshot(page: Page, file: string) {
    await mkdir(evidence, { recursive: true });
    await page.screenshot({ path: `${evidence}/${file}.png`, fullPage: false });
  }

  test('AUTH-01 ID-01: bootstrap or login administrator and create three ordinary users through SDK UI actions', async ({ page, context, browser }) => {
    const diagnostics = installBrowserDiagnostics(page);
    await page.goto('/workbooks');
    await expect(page.getByRole('heading', { name: /初始化管理员|登录 React Sheets/ })).toBeVisible();
    if (await page.getByRole('heading', { name: '初始化管理员' }).isVisible()) {
      const token = (await readFile(process.env.SDK_UAT_BOOTSTRAP_FILE ?? '/tmp/sdk-uat/data/bootstrap-token', 'utf8')).trim();
      await page.getByLabel('初始化凭据', { exact: true }).fill(token);
      await page.getByLabel('用户名', { exact: true }).fill('uat-admin');
      await page.getByLabel('显示名称', { exact: true }).fill('SDK UAT Admin');
      await page.getByLabel('密码', { exact: true }).fill(password);
      const committed = page.waitForResponse(response => response.url().endsWith('/api/auth/bootstrap') && response.request().method() === 'POST');
      await page.getByRole('button', { name: '创建管理员', exact: true }).click();
      expect((await committed).ok()).toBe(true);
      await expect(page.getByTestId('workbook-hub')).toBeVisible();
    } else await login(page);
    adminState = await context.storageState();
    await page.getByRole('button', { name: '用户管理', exact: true }).click();
    await expect(page).toHaveURL('/admin/users');
    for (const role of ['editor', 'commenter', 'viewer']) {
      const username = `${role}-${runId}`;
      await page.getByLabel('用户名', { exact: true }).fill(username);
      await page.getByLabel('显示名称', { exact: true }).fill(`${role} UAT`);
      await page.getByLabel('初始密码', { exact: true }).fill(password);
      const committed = page.waitForResponse(response => response.url().endsWith('/api/admin/users') && response.request().method() === 'POST');
      await page.getByRole('button', { name: '创建用户', exact: true }).click();
      expect((await committed).ok()).toBe(true);
      const row = page.getByRole('row').filter({ hasText: username });
      await expect(row).toBeVisible();
      const id = (await row.locator('td').nth(1).innerText()).trim();
      expect(id).not.toBe('');
      users.set(role, id);
    }
    const invalidatedContext = await browser.newContext({ baseURL: 'http://127.0.0.1:4180' });
    const invalidatedPage = await invalidatedContext.newPage();
    await login(invalidatedPage, `viewer-${runId}`);
    const viewer = page.getByRole('row').filter({ hasText: `viewer-${runId}` });
    await viewer.getByRole('button', { name: '禁用', exact: true }).click();
    await expect(viewer.getByRole('button', { name: '启用', exact: true })).toBeVisible();
    await invalidatedPage.reload();
    await expect(invalidatedPage.getByRole('heading', { name: '登录 React Sheets' })).toBeVisible();
    await viewer.getByRole('button', { name: '启用', exact: true }).click();
    await expect(viewer.getByRole('button', { name: '禁用', exact: true })).toBeVisible();
    await login(invalidatedPage, `viewer-${runId}`);
    await viewer.getByRole('button', { name: '重置密码', exact: true }).click();
    const replacement = 'Uat-Replacement-Password-2026';
    await page.getByLabel('viewer UAT 的新密码', { exact: false }).fill(replacement);
    await page.getByRole('button', { name: '确认重置', exact: true }).click();
    await expect(page.getByRole('button', { name: '确认重置', exact: true })).toHaveCount(0);
    userPasswords.set('viewer', replacement);
    await invalidatedPage.reload();
    await expect(invalidatedPage.getByRole('heading', { name: '登录 React Sheets' })).toBeVisible();
    await invalidatedContext.close();
    await screenshot(page, 'identity-users');
    diagnostics.assertClean();
  });

  test('HUB-01 RT-01: create, edit, save and reload a server-authoritative workbook', async ({ page, context }) => {
    const diagnostics = installBrowserDiagnostics(page);
    await ownerPage(context, page);
    await page.getByRole('button', { name: '新建工作簿', exact: true }).click();
    const dialog = page.getByTestId('create-workbook-dialog');
    await dialog.getByLabel('工作簿名称').fill(name);
    await dialog.getByLabel('保存位置').selectOption('server');
    await dialog.getByRole('button', { name: '创建工作簿', exact: true }).click();
    await expect(page.getByTestId('designer-shell')).toHaveAttribute('data-workspace-phase', 'ready');
    unitId = decodeURIComponent(new URL(page.url()).pathname.split('/').at(-1)!);
    const nameBox = page.getByTestId('name-box');
    await nameBox.fill('A1'); await nameBox.press('Enter');
    const canvas = page.getByTestId('sheet-canvas');
    await canvas.focus();
    await page.keyboard.type('SDK persisted value');
    await page.keyboard.press('Enter');
    await page.keyboard.press('Control+s');
    await expect.poll(async () => {
      const response = await context.request.get(`/api/workbooks/${encodeURIComponent(unitId)}/snapshot`);
      if (!response.ok()) return '';
      const body = await response.json();
      return JSON.stringify(body.snapshot);
    }).toContain('SDK persisted value');
    await page.reload();
    await expect(page.getByTestId('designer-shell')).toHaveAttribute('data-workspace-phase', 'ready');
    await nameBox.fill('A1'); await nameBox.press('Enter');
    await expect(page.getByTestId('formula-input')).toHaveValue('SDK persisted value');
    await screenshot(page, 'workbook-persisted');
    diagnostics.assertClean();
  });

  test('SIZE-01: SDK dimensions persist multi-column width, atomic unhide and AutoFit', async ({ page, context }) => {
    const diagnostics = installBrowserDiagnostics(page);
    await context.addCookies(adminState.cookies);
    await page.addInitScript(() => localStorage.setItem('react-sheets:locale', 'en-US'));
    await page.goto(`/workbooks/${unitId}`);
    await expect(page.getByTestId('designer-shell')).toHaveAttribute('data-workspace-phase', 'ready');
    const nameBox = page.getByTestId('name-box');
    const select = async () => { await nameBox.fill('A1:B2'); await nameBox.press('Enter'); };
    const format = async (label: string) => {
      await page.getByTestId('ribbon-tab-home').click();
      await page.getByRole('button', { name: 'Cells工具', exact: true }).click();
      await page.getByRole('button', { name: 'Format', exact: true }).click();
      await page.getByRole('button', { name: label, exact: true }).click();
    };
    const width = async (value: string) => {
      await format('Column Width…');
      const dialog = page.getByRole('dialog', { name: 'Column Width', exact: true });
      await dialog.getByLabel('Excel character width').fill(value);
      await dialog.getByRole('button', { name: 'OK', exact: true }).click();
    };
    const sheet = async () => {
      const response = await context.request.get(`/api/workbooks/${unitId}/snapshot`);
      expect(response.ok()).toBe(true);
      return (await response.json()).snapshot.sheets[0];
    };
    await select(); await width('0');
    await expect.poll(async () => (await sheet()).hiddenColumns).toEqual([0, 1]);
    await select(); await width('12');
    await expect.poll(async () => (await sheet()).hiddenColumns).toEqual([]);
    const sized = await sheet();
    expect(sized.columnWidthsPx[0]).toBeGreaterThan(64);
    expect(sized.columnWidthsPx[0]).toBe(sized.columnWidthsPx[1]);
    await page.getByTestId('sheet-canvas').focus(); await page.keyboard.press('Control+z');
    await expect.poll(async () => (await sheet()).hiddenColumns).toEqual([0, 1]);
    await page.keyboard.press('Control+y');
    await expect.poll(async () => (await sheet()).hiddenColumns).toEqual([]);
    await select(); await format('Row Height…');
    const rowDialog = page.getByRole('dialog', { name: 'Row Height', exact: true });
    await rowDialog.getByLabel('Row height in points').fill('20');
    await rowDialog.getByRole('button', { name: 'OK', exact: true }).click();
    await expect.poll(async () => (await sheet()).rowHeightsPx[0]).toBeGreaterThan(26);
    expect((await sheet()).rowHeightsPx[0]).toBe((await sheet()).rowHeightsPx[1]);
    await select(); await format('AutoFit Column Width');
    await expect.poll(async () => (await sheet()).columnWidthsPx[1]).toBe(8);
    expect((await sheet()).columnWidthsPx[0]).toBeGreaterThan(8);
    await page.reload();
    await expect(page.getByTestId('designer-shell')).toHaveAttribute('data-workspace-phase', 'ready');
    await screenshot(page, 'sdk-dimensions');
    diagnostics.assertClean();
  });

  test('DATA-01 DATA-02: SDK sort, filter, subtotal and split persist and undo through Java authority', async ({ page, context }) => {
    const diagnostics = installBrowserDiagnostics(page);
    await ownerPage(context, page);
    await page.addInitScript(() => localStorage.setItem('react-sheets:locale', 'en-US'));
    await page.getByRole('button', { name: '新建工作簿', exact: true }).click();
    const create = page.getByTestId('create-workbook-dialog');
    await create.getByLabel('工作簿名称').fill(`SDK data UAT ${runId}`);
    await create.getByLabel('保存位置').selectOption('server');
    await create.getByRole('button', { name: '创建工作簿', exact: true }).click();
    await expect(page.getByTestId('designer-shell')).toHaveAttribute('data-workspace-phase', 'ready');
    const dataId = decodeURIComponent(new URL(page.url()).pathname.split('/').at(-1)!);
    const nameBox = page.getByTestId('name-box');
    const canvas = page.getByTestId('sheet-canvas');
    const select = async (address: string) => { await nameBox.fill(address); await nameBox.press('Enter'); };
    const enter = async (address: string, value: string) => {
      await select(address); await canvas.focus(); await page.keyboard.type(value); await page.keyboard.press('Enter');
    };
    const sheet = async () => {
      const response = await context.request.get(`/api/workbooks/${dataId}/snapshot`);
      expect(response.ok()).toBe(true);
      return (await response.json()).snapshot.sheets[0];
    };
    // Widen only this browser host so every Data group is directly accessible.
    await page.setViewportSize({ width: 1920, height: 1080 });
    for (const [address, value] of [['A1', 'Group'], ['B1', 'Amount'], ['A2', 'East'], ['B2', '10'], ['A3', 'East'], ['B3', '5'], ['A4', 'West'], ['B4', '7']]) await enter(address!, value!);
    await expect.poll(async () => (await sheet()).cells['3']?.['1']?.value).toBe(7);
    const original = await sheet();
    await select('A1:B4');
    await page.getByTestId('ribbon-tab-data').click();
    await page.getByRole('button', { name: 'Subtotal', exact: true }).click();
    await expect.poll(async () => (await sheet()).cells['6']?.['1']?.formula).toBe('=SUBTOTAL(9,B2:B3)');
    expect((await sheet()).outline.groups).toHaveLength(2);
    await canvas.focus(); await page.keyboard.press('Control+z');
    await expect.poll(async () => (await sheet()).cells).toEqual(original.cells);
    expect((await sheet()).outline).toEqual(original.outline);
    await page.keyboard.press('Control+y');
    await expect.poll(async () => (await sheet()).cells['7']?.['1']?.formula).toBe('=SUBTOTAL(9,B4:B4)');
    await canvas.focus(); await page.keyboard.press('Control+z');
    await expect.poll(async () => (await sheet()).cells).toEqual(original.cells);
    await select('A1:B4');
    await page.getByRole('button', { name: 'Sort Z to A', exact: true }).click();
    await expect.poll(async () => (await sheet()).cells['1']?.['0']?.value).toBe('West');
    await canvas.focus(); await page.keyboard.press('Control+z');
    await expect.poll(async () => (await sheet()).cells).toEqual(original.cells);
    await select('A1:B4');
    await page.getByRole('button', { name: 'Filter Selection', exact: true }).click();
    await expect.poll(async () => Boolean((await sheet()).autoFilter)).toBe(true);
    await page.getByRole('button', { name: 'Clear Filter', exact: true }).click();
    await page.getByRole('button', { name: 'Filter Selection', exact: true }).click();
    await expect.poll(async () => Boolean((await sheet()).autoFilter)).toBe(false);
    for (const [address, value] of [['F1', 'Key'], ['G1', 'Value'], ['F2', 'A'], ['G2', '1'], ['F3', 'A'], ['G3', '1'], ['F4', 'B'], ['G4', '2']]) await enter(address!, value!);
    await expect.poll(async () => (await sheet()).cells['3']?.['6']?.value).toBe(2);
    const beforeDuplicates = await sheet();
    await select('F1:G4');
    await page.getByRole('button', { name: 'Remove Duplicates', exact: true }).click();
    await expect.poll(async () => (await sheet()).cells['2']?.['5']?.value).toBe('B');
    await canvas.focus(); await page.keyboard.press('Control+z');
    await expect.poll(async () => (await sheet()).cells).toEqual(beforeDuplicates.cells);
    await select('A1:B4'); await canvas.focus(); await page.keyboard.press('Control+t');
    const tableDialog = page.getByTestId('create-table-dialog');
    await expect(tableDialog).toBeVisible();
    await tableDialog.getByTestId('create-table-confirm').click();
    await expect.poll(async () => (await sheet()).sheetTables.length).toBe(1);
    await page.getByTestId('ribbon-tab-data').click();
    await page.getByRole('button', { name: 'Filter Selection', exact: true }).click();
    await expect.poll(async () => (await sheet()).sheetTables[0]?.showFilterButton).toBe(false);
    expect((await sheet()).sheetTables[0].autoFilter).toBeUndefined();
    await page.getByTestId('ribbon-tab-data').click();
    await page.getByRole('button', { name: 'Filter Selection', exact: true }).click();
    await expect.poll(async () => (await sheet()).sheetTables[0]?.showFilterButton).toBe(true);
    expect((await sheet()).sheetTables[0].autoFilter).toBeTruthy();
    await page.getByTestId('ribbon-tab-data').click();
    await page.getByRole('button', { name: 'Clear Filter', exact: true }).click();
    expect((await sheet()).sheetTables[0].showFilterButton).toBe(true);
    await enter('D1', 'a,b,c'); await select('D1:D2');
    await expect.poll(async () => (await sheet()).cells['0']?.['3']?.value).toBe('a,b,c');
    const beforeSplit = await sheet();
    await page.getByTestId('ribbon-tab-data').click();
    await page.getByRole('button', { name: 'Text to Columns', exact: true }).click();
    await expect.poll(async () => (await sheet()).cells['0']?.['5']?.value).toBe('c');
    await canvas.focus(); await page.keyboard.press('Control+z');
    await expect.poll(async () => (await sheet()).cells['0']?.['3']?.value).toBe('a,b,c');
    await expect.poll(async () => (await sheet()).cells).toEqual(beforeSplit.cells);
    await page.reload();
    await expect(page.getByTestId('designer-shell')).toHaveAttribute('data-workspace-phase', 'ready');
    await screenshot(page, 'sdk-data-undo');
    diagnostics.assertClean();
  });

  test('DOC-01.b REVIEW-01.a: empty-cell hyperlinks survive real upload edit save and download; invalid native bounds refuse', async ({ page, context }) => {
    const diagnostics = installBrowserDiagnostics(page);
    await ownerPage(context, page);
    const nativeName = `SDK hyperlinks ${runId}`;
    const workbook = new WorkbookModel('native-links-fixture', nativeName);
    const source = workbook.getSheet(workbook.primarySheetId);
    const target = workbook.addSheet('target-native', 'Target', 20, 20);
    workbook.setDefinedName({ name: 'SalesTotal', formula: '=Sheet1!A1', scope: 'workbook' });
    source.cells.set(0, 0, { value: 'Native hyperlink source' });
    source.hyperlinks.set('0:0', { id: 'url', target: { kind: 'url', url: 'https://openai.com/' } });
    source.hyperlinks.set('1:0', { id: 'email', target: { kind: 'email', address: 'team@example.com', subject: 'Review' } });
    source.hyperlinks.set('2:0', { id: 'sheet', target: { kind: 'sheet', sheetId: target.id, address: 'B2' } });
    source.hyperlinks.set('3:0', { id: 'name', target: { kind: 'name', name: 'SalesTotal' } });
    const buffer = exportSnapshotToOoxmlBuffer(workbook.snapshot());
    const parts = loadOpcPackageGraph(buffer).packageGraph.parts;
    const xml = strFromU8(parts['xl/worksheets/sheet1.xml']!);
    expect(xml).toContain('<dimension ref="A1:A1"');
    parts['xl/worksheets/sheet1.xml'] = strToU8(xml.replace('ref="A2"', 'ref="A1048577"'));
    const malformed = zipOpcPartsBuffer(parts);
    let imports = 0;
    page.on('request', request => { if (request.method() === 'POST' && request.url().endsWith('/api/workbook-imports')) imports += 1; });
    await page.getByRole('button', { name: '上传导入', exact: true }).click();
    const dialog = page.getByTestId('import-workbook-dialog');
    await dialog.locator('input[type="file"]').setInputFiles({ name: 'invalid-links.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: Buffer.from(malformed) });
    await dialog.getByLabel('导入到', { exact: true }).selectOption('server');
    await dialog.getByRole('button', { name: '开始导入', exact: true }).click();
    await expect(page.getByText(/NATIVE_DOCUMENT_INVALID: Imported hyperlink anchor/)).toBeVisible();
    expect(imports).toBe(0);
    await expect(dialog).toBeVisible();
    await dialog.locator('input[type="file"]').setInputFiles({ name: `${nativeName}.xlsx`, mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: Buffer.from(buffer) });
    await dialog.getByRole('button', { name: '开始导入', exact: true }).click();
    await expect(page.getByTestId('designer-shell')).toHaveAttribute('data-workspace-phase', 'ready');
    const importedId = page.url().split('/workbooks/')[1]!.split('?')[0]!;
    const readSnapshot = async () => {
      const response = await context.request.get(`/api/workbooks/${importedId}/snapshot`);
      expect(response.ok()).toBe(true); return (await response.json()).snapshot;
    };
    const imported = await readSnapshot();
    expect(imported.sheets[0].rowCount).toBe(4);
    expect(imported.sheets[1].rowCount).toBe(2);
    expect(imported.sheets[0].hyperlinks).toHaveLength(4);
    expect(imported.sheets[0].cells['1']).toBeUndefined();
    const nameBox = page.getByTestId('name-box');
    await nameBox.fill('A1'); await nameBox.press('Enter');
    const formula = page.getByTestId('formula-input');
    await formula.fill('Native hyperlinks edited'); await formula.press('Enter');
    await page.getByTestId('sheet-canvas').focus();
    await page.keyboard.press('Control+s');
    await expect.poll(async () => (await readSnapshot()).sheets[0].cells['0']?.['0']?.value).toBe('Native hyperlinks edited');
    await page.reload();
    await expect(page.getByTestId('designer-shell')).toHaveAttribute('data-workspace-phase', 'ready');
    await screenshot(page, 'native-hyperlinks-edited');
    await page.goto('/workbooks'); await expect(page.getByTestId('workbook-hub')).toBeVisible();
    await openMenu(page, nativeName);
    const downloading = page.waitForEvent('download');
    await menu(page).getByRole('button', { name: '导出副本', exact: true }).click();
    const download = await downloading;
    expect(await download.failure()).toBeNull();
    const downloadedPath = `${evidence}/native-hyperlinks-edited.xlsx`;
    await download.saveAs(downloadedPath);
    const bytes = await readFile(downloadedPath);
    const reimported = await importOoxmlDocument({ fileName: download.suggestedFilename(), buffer: Uint8Array.from(bytes).buffer, options: { compatibilityTarget: 'B' } });
    expect(reimported.snapshot.sheets[0]!.cells['0']!['0']!.value).toBe('Native hyperlinks edited');
    expect(reimported.snapshot.sheets[0]!.hyperlinks).toEqual(imported.sheets[0].hyperlinks);
    expect(imports).toBe(1);
    diagnostics.assertClean();
  });

  test('ROLE-02 HUB-02: owner menu, rename, favorite and share roles use real SDK catalog actions', async ({ page, context }) => {
    const diagnostics = installBrowserDiagnostics(page);
    await ownerPage(context, page);
    await openMenu(page);
    for (const label of ['重命名', '移动到', '共享', '移到回收站']) await expect(menu(page).getByRole('button', { name: label, exact: true })).toBeVisible();
    await menu(page).getByRole('button', { name: '重命名', exact: true }).click();
    const rename = page.getByTestId('rename-workbook-dialog');
    const newName = `${name} renamed`;
    await rename.getByLabel('名称', { exact: true }).fill(newName);
    await rename.getByRole('button', { name: '保存', exact: true }).click();
    await expect(rename).not.toBeVisible();
    name = newName;
    await openMenu(page);
    await page.getByRole('button', { name: '添加星标', exact: true }).click();
    await expect(page.getByRole('button', { name: `取消 ${name} 的星标`, exact: true })).toBeVisible();
    await openMenu(page);
    await expect(page.getByRole('button', { name: '取消星标', exact: true })).toBeVisible();
    await page.keyboard.press('Escape');
    for (const role of ['editor', 'commenter', 'viewer']) {
      await openMenu(page);
      await menu(page).getByRole('button', { name: '共享', exact: true }).click();
      const share = page.getByTestId('share-workbook-dialog');
      await share.getByLabel('成员邮箱或账号').fill(users.get(role)!);
      await share.getByLabel('访问权限').selectOption(role);
      const committed = page.waitForResponse(response => response.url().includes(`/api/workbooks/${unitId}/acl`) && response.request().method() === 'PUT');
      await share.getByRole('button', { name: '发送邀请', exact: true }).click();
      expect((await committed).ok()).toBe(true);
      await expect(share).not.toBeVisible();
    }
    await screenshot(page, 'owner-catalog');
    diagnostics.assertClean();
  });

  for (const role of ['editor', 'commenter', 'viewer']) {
    test(`ROLE-02 ID-02: ${role} capability menu and server authorization`, async ({ page, context }) => {
      const diagnostics = installBrowserDiagnostics(page);
      await login(page, `${role}-${runId}`, userPasswords.get(role) ?? password);
      await expect(page.getByRole('button', { name: '用户管理', exact: true })).toHaveCount(0);
      await openMenu(page);
      for (const label of ['共享', '移到回收站', '恢复', '永久删除']) await expect(menu(page).getByRole('button', { name: label, exact: true })).toHaveCount(0);
      for (const label of ['重命名', '移动到']) {
        if (role === 'editor') await expect(menu(page).getByRole('button', { name: label, exact: true })).toBeVisible();
        else await expect(menu(page).getByRole('button', { name: label, exact: true })).toHaveCount(0);
      }
      const sessionResponse = await context.request.get('/api/auth/session');
      const csrf = (await sessionResponse.json()).csrfToken;
      for (const suffix of ['', '/purge']) {
        const denied = await context.request.delete(`/api/workbooks/${unitId}${suffix}`, { headers: { 'X-CSRF-TOKEN': csrf } });
        expect(denied.status()).toBe(403);
        expect((await denied.json()).code).toBe('FORBIDDEN');
      }
      roleStates.set(role, await context.storageState());
      await screenshot(page, `role-${role}`);
      await page.keyboard.press('Escape');
      await page.goto('/admin/users');
      await expect(page.getByRole('alert')).toHaveText('需要管理员权限。');
      const rejected = await context.request.get('/api/admin/users');
      expect(rejected.status()).toBe(403);
      diagnostics.assertClean();
    });
  }

  test('HUB-03: owner trash, restore and purge retain lifecycle restrictions', async ({ page, context, browser }) => {
    const diagnostics = installBrowserDiagnostics(page);
    await ownerPage(context, page);
    await openMenu(page);
    await page.getByRole('button', { name: '移到回收站', exact: true }).click();
    await page.getByTestId('delete-workbook-dialog').getByRole('button', { name: '移到回收站', exact: true }).click();
    await expect(page.getByRole('button', { name: `打开 ${name} 的更多操作`, exact: true })).toHaveCount(0);
    await page.getByRole('button', { name: '回收站', exact: true }).click();
    await openMenu(page);
    await expect(page.getByRole('button', { name: '恢复', exact: true })).toBeVisible();
    await expect(page.getByRole('button', { name: '永久删除', exact: true })).toBeVisible();
    await expect(menu(page).getByRole('button', { name: '共享', exact: true })).toHaveCount(0);
    for (const role of ['editor', 'commenter', 'viewer']) {
      const nonOwner = await browser.newContext({ baseURL: 'http://127.0.0.1:4180', storageState: roleStates.get(role)! });
      const sessionResponse = await nonOwner.request.get('/api/auth/session');
      const csrf = (await sessionResponse.json()).csrfToken;
      const denied = await nonOwner.request.delete(`/api/workbooks/${unitId}/purge`, { headers: { 'X-CSRF-TOKEN': csrf } });
      expect(denied.status()).toBe(403);
      expect((await denied.json()).code).toBe('FORBIDDEN');
      await nonOwner.close();
    }
    await page.getByRole('button', { name: '恢复', exact: true }).click();
    await expect(page.getByRole('button', { name: `打开 ${name} 的更多操作`, exact: true })).toHaveCount(0);
    await page.getByRole('button', { name: '最近', exact: true }).click();
    await openMenu(page);
    await page.getByRole('button', { name: '移到回收站', exact: true }).click();
    await page.getByTestId('delete-workbook-dialog').getByRole('button', { name: '移到回收站', exact: true }).click();
    await expect(page.getByRole('button', { name: `打开 ${name} 的更多操作`, exact: true })).toHaveCount(0);
    await page.getByRole('button', { name: '回收站', exact: true }).click();
    await openMenu(page);
    await page.getByRole('button', { name: '永久删除', exact: true }).click();
    await page.getByTestId('delete-workbook-dialog').getByRole('button', { name: '永久删除', exact: true }).click();
    await expect(page.getByRole('button', { name: `打开 ${name} 的更多操作`, exact: true })).toHaveCount(0);
    await screenshot(page, 'trash-purged');
    diagnostics.assertClean();
  });
  test('AUTH-02: invalid login does not create identity; successful login and logout follow server session', async ({ page, context }) => {
    const consoleErrors: { text: string; url: string }[] = [];
    const pageErrors: string[] = [];
    page.on('console', message => { if (message.type() === 'error') consoleErrors.push({ text: message.text(), url: message.location().url }); });
    page.on('pageerror', error => pageErrors.push(error.message));
    await page.goto('/workbooks');
    await page.getByLabel('用户名', { exact: true }).fill('uat-admin');
    await page.getByLabel('密码', { exact: true }).fill('incorrect-password');
    const rejected = page.waitForResponse(response => response.url().endsWith('/api/auth/login'));
    await page.getByRole('button', { name: '登录', exact: true }).click();
    expect((await rejected).status()).toBe(401);
    await expect(page.getByRole('alert')).toContainText('401');
    await expect(page.getByTestId('workbook-hub')).toHaveCount(0);
    await page.getByLabel('密码', { exact: true }).fill(password);
    await page.getByRole('button', { name: '登录', exact: true }).click();
    await expect(page.getByTestId('workbook-hub')).toBeVisible();
    await page.getByRole('button', { name: '退出登录', exact: true }).click();
    await expect(page.getByRole('heading', { name: '登录 React Sheets' })).toBeVisible();
    const response = await context.request.get('/api/auth/session');
    expect((await response.json()).authenticated).toBe(false);
    await page.reload();
    await expect(page.getByRole('heading', { name: '登录 React Sheets' })).toBeVisible();
    expect(pageErrors).toEqual([]);
    expect(consoleErrors.filter(error => !(error.url.endsWith('/api/auth/login') && error.text.includes('401')))).toEqual([]);
    await screenshot(page, 'auth-logged-out');
  });

});
