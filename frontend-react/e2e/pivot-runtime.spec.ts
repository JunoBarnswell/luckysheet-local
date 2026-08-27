import { existsSync } from 'node:fs';
import { expect, test } from '@playwright/test';
import { installBrowserDiagnostics } from './support/workbook-fixtures';

const fixturePath = process.env.OCR_XLSX_FIXTURE ?? 'C:\\Users\\kuo13\\Downloads\\OCR结果.xlsx';

test.describe('Pivot worker runtime', () => {
  test.use({ viewport: { width: 1770, height: 1041 }, deviceScaleFactor: 1 });

  test('imports the real OCR workbook and applies a deferred Pivot layout without blocking', async ({ page }, testInfo) => {
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

    await page.getByRole('button', { name: '字段菜单: 页码', exact: true }).click();
    await page.getByRole('button', { name: '行', exact: true }).last().click();
    await page.getByRole('button', { name: '字段菜单: BOM1', exact: true }).click();
    await page.getByRole('button', { name: '值', exact: true }).last().click();
    const applyButton = page.getByRole('button', { name: '应用布局', exact: true });
    await expect(applyButton).toBeEnabled();
    const applyStartedAt = performance.now();
    await applyButton.click();
    await expect(page.getByRole('button', { name: '应用布局', exact: true })).toBeDisabled({ timeout: 5_000 });
    const applyMs = performance.now() - applyStartedAt;

    await expect(page.getByRole('region', { name: '行 field area' }).getByText('页码', { exact: true })).toBeVisible();
    await expect(page.getByRole('region', { name: '值 field area' }).getByText('BOM1', { exact: true })).toBeVisible();
    await expect(page.getByRole('alert')).toHaveCount(0);
    expect(importMs).toBeLessThan(2_000);
    expect(createMs).toBeLessThan(1_000);
    expect(applyMs).toBeLessThan(1_000);
    testInfo.annotations.push({ type: 'performance', description: `import=${Math.round(importMs)}ms create=${Math.round(createMs)}ms apply=${Math.round(applyMs)}ms` });
    diagnostics.assertClean();
  });
});
