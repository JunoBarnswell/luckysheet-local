import { existsSync } from 'node:fs';
import { expect, test } from '@playwright/test';
import { installBrowserDiagnostics } from './support/workbook-fixtures';

const fixturePath = process.env.OCR_XLSX_FIXTURE ?? 'C:\\Users\\kuo13\\Downloads\\OCR结果.xlsx';

test.describe('Pivot worker runtime', () => {
  test.use({ viewport: { width: 1770, height: 1041 }, deviceScaleFactor: 1 });

  test('imports the real OCR workbook and creates its Pivot Field List without blocking', async ({ page }, testInfo) => {
    test.setTimeout(60_000);
    test.skip(!existsSync(fixturePath), 'OCR fixture is not installed on this host');
    const diagnostics = installBrowserDiagnostics(page);
    await page.addInitScript(() => window.localStorage.setItem('react-sheets:locale', 'zh-CN'));
    await page.goto('/workbooks?dialog=import');
    const chooser = page.waitForEvent('filechooser');
    await page.getByRole('button', { name: '选择文件', exact: true }).click();
    await (await chooser).setFiles(fixturePath);
    const importStartedAt = performance.now();
    await page.getByRole('button', { name: '开始导入', exact: true }).click();
    await expect(page).toHaveURL(/\/workbooks\/[^/]+(?:\?.*)?$/, { timeout: 10_000 });
    await expect(page.getByTestId('designer-shell')).toHaveAttribute('data-workspace-phase', 'ready', { timeout: 10_000 });
    const importMs = performance.now() - importStartedAt;

    const nameBox = page.getByTestId('name-box');
    await nameBox.fill('A1:W4059');
    await nameBox.press('Enter');
    await page.getByTestId('ribbon-tab-insert').click();
    await page.getByRole('button', { name: '透视表', exact: true }).click();
    const createStartedAt = performance.now();
    await page.getByTestId('create-pivot-confirm').click();
    await expect(page.getByText('选择要添加到报表的字段', { exact: true })).toBeVisible({ timeout: 5_000 });
    const createMs = performance.now() - createStartedAt;

    await expect(page.getByTestId('pivot-field-list').getByRole('button', { name: /^字段菜单:/ })).toHaveCount(23);
    await expect(page.getByRole('region', { name: '筛选 field area' })).toBeVisible();
    await expect(page.getByRole('region', { name: '列 field area' })).toBeVisible();
    await expect(page.getByRole('region', { name: '行 field area' })).toBeVisible();
    await expect(page.getByRole('region', { name: '值 field area' })).toBeVisible();
    await expect(page.getByRole('alert')).toHaveCount(0);
    expect(importMs).toBeLessThan(2_000);
    expect(createMs).toBeLessThan(1_000);
    testInfo.annotations.push({ type: 'performance', description: `import=${Math.round(importMs)}ms create=${Math.round(createMs)}ms` });
    diagnostics.assertClean();
  });
});
