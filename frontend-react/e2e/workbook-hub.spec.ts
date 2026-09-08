import { expect, test } from '@playwright/test';
import { openConnectedWorkbook } from './support/workbook-fixtures';

test.describe('workbook hub', () => {
  test('renders the file-center shell and never exposes inactive workbook commands', async ({ page }) => {
    await page.goto('/');
    await expect(page).toHaveURL('/workbooks');
    await expect(page.getByTestId('workbook-hub')).toBeVisible();
    await expect(page.getByRole('heading', { name: '早上好' })).toBeVisible();
    await expect(page.getByRole('button', { name: '信息' })).toBeDisabled();
    await expect(page.getByRole('button', { name: '保存' })).toBeDisabled();
    await expect(page.getByRole('button', { name: '导出', exact: true })).toBeDisabled();
    await expect(page.getByRole('button', { name: '关闭' })).toBeDisabled();
    await expect(page.getByRole('button', { name: '空白工作簿', exact: true })).toBeVisible();
    await expect(page.getByRole('button', { name: '打开 / 导入原生文档', exact: true })).toBeVisible();
  });

  test('creates a cloud workbook then preserves its session through Backstage', async ({ page }) => {
    await openConnectedWorkbook(page, 'zh-CN', 'Backstage UAT');
    await page.getByRole('button', { name: 'Open workbook menu', exact: true }).click();
    await page.getByText('File / 工作簿', { exact: true }).click();
    await expect(page.getByTestId('workbook-backstage')).toBeVisible();
    await expect(page.getByTestId('workbook-backstage').getByText('文件名：Backstage UAT', { exact: true })).toBeVisible();
    await page.getByRole('button', { name: '返回编辑器', exact: true }).click();
  await expect(page.getByTestId('designer-shell')).toHaveAttribute('data-workspace-phase', 'ready');
  });

  test('does not turn an unknown unauthenticated workbook into a local session', async ({ page }) => {
    await page.goto('/workbooks/not-a-cloud-workbook');
    await expect(page.getByTestId('designer-shell')).toHaveCount(0);
    await expect(page.getByRole('heading', { name: '无法打开工作簿' })).toBeVisible();
    await expect(page.getByText('Authoritative workbook service is unavailable: not-a-cloud-workbook', { exact: true })).toBeVisible();
  });
});

test.describe('workbook hub desktop reference anchors', () => {
  test.use({ viewport: { width: 1672, height: 941 }, deviceScaleFactor: 1 });

  test('keeps the supplied 1672×941 shell geometry', async ({ page }) => {
    await page.goto('/workbooks');
    await expect(page.getByTestId('workbook-hub')).toBeVisible();
    const header = await page.locator('header').boundingBox();
    const sidebar = await page.getByRole('navigation', { name: '工作簿导航' }).boundingBox();
    const title = await page.getByRole('heading', { name: '早上好' }).boundingBox();
    const firstCard = await page.getByRole('button', { name: '空白工作簿', exact: true }).boundingBox();
    expect(header).toMatchObject({ x: 0, y: 0, height: 58 });
    expect(sidebar).toMatchObject({ x: 0, y: 58, width: 168 });
    expect(title).toMatchObject({ x: 218, y: 88 });
    expect(firstCard).toMatchObject({ x: 236, y: 192, height: 182 });
    expect(await page.evaluate(() => document.documentElement.scrollWidth)).toBeLessThanOrEqual(1672);
  });
});
