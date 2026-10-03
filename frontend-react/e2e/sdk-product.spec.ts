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
  let adminState: Awaited<ReturnType<BrowserContext['storageState']>>;
  let unitId = '';
  const name = `SDK UAT ${runId}`;
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

  test('ROLE-02 HUB-02: owner menu, rename, favorite and share roles use real SDK catalog actions', async ({ page, context }) => {
    const diagnostics = installBrowserDiagnostics(page);
    await ownerPage(context, page);
    await openMenu(page);
    for (const label of ['重命名', '移动到', '共享', '移到回收站']) await expect(page.getByRole('button', { name: label, exact: true })).toBeVisible();
    await page.getByRole('button', { name: '添加星标', exact: true }).click();
    await openMenu(page);
    await expect(page.getByRole('button', { name: '取消星标', exact: true })).toBeVisible();
    await page.keyboard.press('Escape');
    for (const role of ['editor', 'commenter', 'viewer']) {
      await openMenu(page);
      await page.getByRole('button', { name: '共享', exact: true }).click();
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
      for (const label of ['共享', '移到回收站', '恢复', '永久删除']) await expect(page.getByRole('button', { name: label, exact: true })).toHaveCount(0);
      for (const label of ['重命名', '移动到']) {
        if (role === 'editor') await expect(page.getByRole('button', { name: label, exact: true })).toBeVisible();
        else await expect(page.getByRole('button', { name: label, exact: true })).toHaveCount(0);
      }
      await screenshot(page, `role-${role}`);
      await page.keyboard.press('Escape');
      await page.goto('/admin/users');
      await expect(page.getByRole('alert')).toHaveText('需要管理员权限。');
      const rejected = await context.request.get('/api/admin/users');
      expect(rejected.status()).toBe(403);
      diagnostics.assertClean();
    });
  }

  test('HUB-03: owner trash, restore and purge retain lifecycle restrictions', async ({ page, context }) => {
    const diagnostics = installBrowserDiagnostics(page);
    await ownerPage(context, page);
    await openMenu(page);
    await page.getByRole('button', { name: '移到回收站', exact: true }).click();
    await page.getByTestId('delete-workbook-dialog').getByRole('button', { name: '移到回收站', exact: true }).click();
    await page.getByRole('button', { name: '回收站', exact: true }).click();
    await openMenu(page);
    await expect(page.getByRole('button', { name: '恢复', exact: true })).toBeVisible();
    await expect(page.getByRole('button', { name: '永久删除', exact: true })).toBeVisible();
    await expect(page.getByRole('button', { name: '共享', exact: true })).toHaveCount(0);
    await page.getByRole('button', { name: '恢复', exact: true }).click();
    await page.getByRole('button', { name: '最近', exact: true }).click();
    await openMenu(page);
    await page.getByRole('button', { name: '移到回收站', exact: true }).click();
    await page.getByTestId('delete-workbook-dialog').getByRole('button', { name: '移到回收站', exact: true }).click();
    await page.getByRole('button', { name: '回收站', exact: true }).click();
    await openMenu(page);
    await page.getByRole('button', { name: '永久删除', exact: true }).click();
    await page.getByRole('button', { name: '永久删除', exact: true }).click();
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
