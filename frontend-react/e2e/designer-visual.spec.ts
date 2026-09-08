import { expect, test } from '@playwright/test';
import { DESIGNER_GEOMETRY } from '@react-sheets/ui-system';
import { openConnectedWorkbook } from './support/workbook-fixtures';

const viewports = [
  { width: 1280, height: 720 },
  { width: 1366, height: 768 },
  { width: 1440, height: 900 },
  { width: 1917, height: 900 },
  { width: 1920, height: 1080 },
] as const;

async function openDemo(page: import('@playwright/test').Page, viewport: { width: number; height: number }) {
  await page.setViewportSize(viewport);
  await openConnectedWorkbook(page, 'zh-CN', `Designer Visual ${viewport.width}x${viewport.height} ${Date.now()}`);
  await page.getByTestId('name-box').fill('B1');
  await page.getByTestId('name-box').press('Enter');
}

for (const viewport of viewports) {
  test(`Designer Shell visual contract ${viewport.width}x${viewport.height}`, async ({ page }) => {
    await openDemo(page, viewport);
    await expect(page.getByTestId('name-box')).toHaveValue('B1');
    const geometry = await page.evaluate(() => {
      const rect = (selector: string) => {
        const node = document.querySelector(selector);
        if (!node) return null;
        const box = node.getBoundingClientRect();
        return { x: Math.round(box.x), y: Math.round(box.y), width: Math.round(box.width), height: Math.round(box.height) };
      };
      return {
        shell: rect('[data-testid="designer-shell"]'),
        ribbon: rect('[data-testid="designer-ribbon"]'),
        formula: rect('[data-testid="designer-formula-bar"]'),
        workspace: rect('[data-testid="designer-workspace"]'),
        tabs: rect('[data-testid="designer-sheet-tabs"]'),
        status: rect('[data-testid="designer-status-bar"]'),
        scrollWidth: document.documentElement.scrollWidth,
        scrollHeight: document.documentElement.scrollHeight,
      };
    });
    expect(geometry.shell).toMatchObject({ x: 0, y: 0, width: viewport.width, height: viewport.height });
    const ribbonY = DESIGNER_GEOMETRY.documentBarHeight;
    const formulaY = ribbonY + DESIGNER_GEOMETRY.ribbonHeight;
    const workspaceY = formulaY + DESIGNER_GEOMETRY.formulaBarHeight;
    const tabsY = viewport.height - DESIGNER_GEOMETRY.statusBarHeight - DESIGNER_GEOMETRY.sheetTabsHeight;
    expect(geometry.ribbon).toMatchObject({ x: 0, y: ribbonY, width: viewport.width, height: DESIGNER_GEOMETRY.ribbonHeight });
    expect(geometry.formula).toMatchObject({ x: 0, y: formulaY, width: viewport.width, height: DESIGNER_GEOMETRY.formulaBarHeight });
    expect(geometry.workspace).toMatchObject({
      x: 0,
      y: workspaceY,
      width: viewport.width,
      height: viewport.height - DESIGNER_GEOMETRY.statusBarHeight - workspaceY,
    });
    expect(geometry.tabs).toMatchObject({ x: 0, y: tabsY, width: viewport.width, height: DESIGNER_GEOMETRY.sheetTabsHeight });
    expect(geometry.status).toMatchObject({ x: 0, y: viewport.height - DESIGNER_GEOMETRY.statusBarHeight, width: viewport.width, height: DESIGNER_GEOMETRY.statusBarHeight });
    expect(geometry.scrollWidth).toBeLessThanOrEqual(viewport.width);
    expect(geometry.scrollHeight).toBeLessThanOrEqual(viewport.height);
    await page.getByRole('tab', { name: '视图' }).click();
    await page.getByRole('button', { name: '命令面板' }).click();
    await expect(page.getByTestId('command-palette')).toBeVisible();
    await expect(page.getByRole('textbox', { name: 'Command search' })).toBeFocused();
    await expect(page.getByTestId('designer-shell')).toHaveAttribute('data-workspace-phase', 'ready');
  });
}
