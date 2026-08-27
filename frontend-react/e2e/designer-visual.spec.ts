import { expect, test } from '@playwright/test';

const viewports = [
  { width: 1280, height: 720 },
  { width: 1366, height: 768 },
  { width: 1440, height: 900 },
  { width: 1917, height: 900 },
  { width: 1920, height: 1080 },
] as const;

async function openDemo(page: import('@playwright/test').Page, viewport: { width: number; height: number }) {
  await page.setViewportSize(viewport);
  await page.addInitScript(() => window.localStorage.setItem('react-sheets:locale', 'zh-CN'));
  await page.goto('/');
  await expect(page.getByTestId('workbook-hub')).toBeVisible();
  await page.getByRole('button', { name: 'Designer Demo' }).click();
  const dialog = page.getByTestId('create-workbook-dialog');
  await expect(dialog).toBeVisible();
  await dialog.getByLabel('工作簿名称').fill(`Designer Demo Visual ${viewport.width}x${viewport.height} ${Date.now()}`);
  await dialog.getByLabel('保存位置').selectOption('local');
  await dialog.getByRole('button', { name: '创建工作簿' }).click();
  await expect(page).toHaveURL(/\/workbooks\/[^/]+(?:\?.*)?$/);
  await expect(page.getByTestId('designer-shell')).toHaveAttribute('data-workspace-phase', 'ready', { timeout: 30_000 });
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
    expect(geometry.ribbon).toMatchObject({ x: 0, y: 0, width: viewport.width, height: 195 });
    expect(geometry.formula).toMatchObject({ x: 0, y: 195, width: viewport.width, height: 48 });
    expect(geometry.workspace).toMatchObject({ x: 0, y: 243, width: viewport.width, height: viewport.height - 265 });
    expect(geometry.tabs).toMatchObject({ x: 0, y: viewport.height - 51, width: viewport.width, height: 29 });
    expect(geometry.status).toMatchObject({ x: 0, y: viewport.height - 22, width: viewport.width, height: 22 });
    expect(geometry.scrollWidth).toBeLessThanOrEqual(viewport.width);
    expect(geometry.scrollHeight).toBeLessThanOrEqual(viewport.height);
    if (viewport.width === 1917) {
      const homeAnchors = await page.evaluate(() => {
        const bounds = (node: Element | null) => {
          if (!node) return null;
          const box = node.getBoundingClientRect();
          return { x: Math.round(box.x), width: Math.round(box.width), height: Math.round(box.height) };
        };
        const tabIds = ['home', 'insert', 'pageLayout', 'formulas', 'data', 'view', 'review', 'settings'];
        const groups = document.querySelector('[data-testid="home-ribbon-groups"]');
        return {
          file: bounds(document.querySelector('[aria-label="Open workbook menu"]')),
          tabs: tabIds.map((id) => bounds(document.querySelector(`[data-testid="ribbon-tab-${id}"]`))),
          separators: groups ? Array.from(groups.children).filter((child) => child.tagName === 'HR').map((child) => Math.round(child.getBoundingClientRect().x)) : [],
          nameBox: bounds(document.querySelector('[data-testid="name-box"]')),
          formulaInput: bounds(document.querySelector('[data-testid="formula-input"]')),
        };
      });
      expect(homeAnchors.file).toMatchObject({ x: 0, width: 61 });
      expect(homeAnchors.tabs.map((tab) => tab?.width)).toEqual([62, 68, 91, 72, 57, 66, 66, 66]);
      expect(homeAnchors.separators).toEqual([60, 182, 405, 659, 813, 1112, 1228]);
      expect(homeAnchors.nameBox).toMatchObject({ x: 5, width: 165, height: 37 });
      expect(homeAnchors.formulaInput).toMatchObject({ x: 315, width: 1593, height: 37 });
      await page.screenshot({ path: 'test-results/grape-city-ribbon-1917-home.png' });
      await page.getByRole('tab', { name: '插入' }).click();
      const insertAnchors = await page.evaluate(() => {
        const root = document.querySelector('[data-testid="insert-ribbon-groups"]');
        return root ? Array.from(root.children).filter((child) => child.tagName === 'HR').map((child) => Math.round(child.getBoundingClientRect().x)) : [];
      });
      expect(insertAnchors).toEqual([135, 271, 408, 499, 675, 766]);
      await page.screenshot({ path: 'test-results/grape-city-ribbon-1917-insert.png' });
    }
    await page.getByRole('tab', { name: '视图' }).click();
    await page.getByRole('button', { name: '命令面板' }).click();
    await expect(page.getByTestId('command-palette')).toBeVisible();
    await expect(page.getByRole('textbox', { name: 'Command search' })).toBeFocused();
    await expect(page.getByTestId('designer-shell')).toHaveAttribute('data-workspace-phase', 'ready');
  });
}
