import { expect, test } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { randomBytes } from 'node:crypto';

// Run against a fresh, disposable backend; never consume a deployed bootstrap token.
test('security remediation: server persistence, inert replacements and AutoSum', async ({ page, context }, testInfo) => {
  test.skip(!process.env.SECURITY_ACCEPTANCE_BOOTSTRAP_FILE, 'Requires a disposable backend bootstrap file');
  const session = await (await context.request.get('/api/auth/session')).json();
  expect(session.bootstrapRequired).toBe(true);
  const bootstrap = await context.request.post('/api/auth/bootstrap', {
    headers: { 'X-CSRF-TOKEN': session.csrfToken },
    data: {
      token: readFileSync(process.env.SECURITY_ACCEPTANCE_BOOTSTRAP_FILE!, 'utf8'),
      username: 'acceptance-admin', password: randomBytes(24).toString('base64url'), displayName: 'Acceptance',
    },
  });
  expect(bootstrap.ok()).toBe(true);
  const diagnostics = { consoleErrors: [] as string[], pageErrors: [] as string[], requestFailures: [] as string[], httpErrors: [] as string[] };
  page.on('console', message => { if (message.type() === 'error') diagnostics.consoleErrors.push(message.text()); });
  page.on('pageerror', error => diagnostics.pageErrors.push(error.message));
  page.on('requestfailed', request => diagnostics.requestFailures.push(`${request.method()} ${request.url()}`));
  page.on('response', response => { if (response.status() >= 400) diagnostics.httpErrors.push(`${response.status()} ${response.url()}`); });
  await page.goto('/workbooks');
  await expect(page.getByTestId('workbook-hub')).toBeVisible();
  await page.getByRole('button', { name: '新建工作簿' }).click();
  const dialog = page.getByTestId('create-workbook-dialog');
  await dialog.getByLabel('工作簿名称').fill('Security remediation acceptance');
  await dialog.getByLabel('保存位置').selectOption('server');
  await dialog.getByRole('button', { name: '创建工作簿' }).click();
  await expect(page.getByTestId('designer-shell')).toHaveAttribute('data-workspace-phase', 'ready', { timeout: 30_000 });
  const canvas = page.getByTestId('sheet-canvas');
  const formula = page.getByTestId('formula-input');
  const select = async (address: string) => {
    await page.getByTestId('name-box').fill(address);
    await page.getByTestId('name-box').press('Enter');
    await canvas.focus();
  };
  await select('A1');
  await page.keyboard.type('normal business input');
  await page.keyboard.press('Enter');
  await select('A1');
  await expect(formula).toHaveValue('normal business input');
  await canvas.press('Control+Z');
  await expect(formula).toHaveValue('');
  await canvas.press('Control+Y');
  await expect(formula).toHaveValue('normal business input');
  await canvas.press('F2');
  const editor = page.getByLabel('Cell editor');
  await expect(editor).toBeFocused();
  await editor.fill('A measured draft wider than a cell');
  expect((await editor.boundingBox())!.width).toBeGreaterThan(64);
  await editor.press('Escape');
  await expect(editor).toHaveCount(0);

  await canvas.press('Control+h');
  const replace = page.getByTestId('find-replace-dialog');
  await expect(replace).toBeVisible();
  await page.getByTestId('find-input').fill('normal business input');
  await page.getByTestId('replace-input').fill('=1+2');
  await page.getByTestId('find-replace-all').click();
  await page.keyboard.press('Escape');
  await expect(replace).toHaveCount(0);
  await select('A1');
  await expect(formula).toHaveValue('=1+2');

  for (const [address, value] of [['D1', '1'], ['D2', '2']] as const) {
    await select(address);
    await page.keyboard.type(value);
    await page.keyboard.press('Enter');
  }
  await select('D1');
  await page.getByTestId('ribbon-tab-formulas').click();
  await page.locator('[data-ribbon-command="autoSum"]').click();
  await select('D3');
  await expect(formula).toHaveValue('=SUM(D1:D2)');
  const unitId = new URL(page.url()).pathname.split('/').at(-1)!;
  await expect.poll(async () => {
    const response = await context.request.get(`/api/workbooks/${encodeURIComponent(unitId)}/snapshot`);
    expect(response.ok()).toBe(true);
    const sheet = (await response.json()).snapshot.sheets[0];
    return { replacement: sheet.cells['0']?.['0'], sum: sheet.cells['2']?.['3']?.formula,
      sources: [sheet.cells['0']?.['3']?.value, sheet.cells['1']?.['3']?.value] };
  }).toMatchObject({ replacement: { value: '=1+2' }, sum: '=SUM(D1:D2)', sources: [1, 2] });
  const snapshot = await (await context.request.get(`/api/workbooks/${encodeURIComponent(unitId)}/snapshot`)).json();
  expect(snapshot.snapshot.sheets[0].cells['0']['0'].formula).toBeUndefined();
  await page.reload();
  await expect(page.getByTestId('designer-shell')).toHaveAttribute('data-workspace-phase', 'ready', { timeout: 30_000 });
  await select('A1');
  await expect(formula).toHaveValue('=1+2');
  await select('D3');
  await expect(formula).toHaveValue('=SUM(D1:D2)');
  await page.screenshot({ path: testInfo.outputPath('acceptance.png'), fullPage: true });
  await testInfo.attach('diagnostics', { body: JSON.stringify(diagnostics, null, 2), contentType: 'application/json' });
  expect(diagnostics).toEqual({ consoleErrors: [], pageErrors: [], requestFailures: [], httpErrors: [] });
});
