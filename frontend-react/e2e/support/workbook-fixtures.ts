import { expect, type Page } from '@playwright/test';

export type FixtureLocale = 'zh-CN' | 'en-US';

export interface BrowserDiagnostics {
  readonly consoleErrors: string[];
  readonly pageErrors: string[];
  readonly requestFailures: string[];
  assertClean(): void;
}

export function installBrowserDiagnostics(page: Page): BrowserDiagnostics {
  const consoleErrors: string[] = [];
  const pageErrors: string[] = [];
  const requestFailures: string[] = [];
  page.on('console', (message) => {
    if (message.type() === 'error') consoleErrors.push(message.text());
  });
  page.on('pageerror', (error) => pageErrors.push(error.message));
  page.on('requestfailed', (request) => requestFailures.push(`${request.method()} ${request.url()}: ${request.failure()?.errorText ?? 'unknown failure'}`));
  return {
    consoleErrors,
    pageErrors,
    requestFailures,
    assertClean() {
      expect(consoleErrors, 'browser console errors').toEqual([]);
      expect(pageErrors, 'browser page errors').toEqual([]);
      expect(requestFailures, 'browser request failures').toEqual([]);
    },
  };
}

export async function openLocalWorkbook(page: Page, locale: FixtureLocale, name: string): Promise<void> {
  await page.addInitScript((value) => window.localStorage.setItem('react-sheets:locale', value), locale);
  await page.goto('/workbooks');
  await expect(page.getByTestId('workbook-hub')).toBeVisible();
  await page.getByRole('button', { name: '新建工作簿' }).click();
  const dialog = page.getByTestId('create-workbook-dialog');
  await expect(dialog).toBeVisible();
  await dialog.getByLabel('工作簿名称').fill(name);
  await dialog.getByLabel('保存位置').selectOption('local');
  await dialog.getByRole('button', { name: '创建工作簿' }).click();
  await expect(page).toHaveURL(/\/workbooks\/[^/]+(?:\?.*)?$/);
  await expect(page.getByTestId('designer-shell')).toHaveAttribute('data-workspace-phase', 'ready', { timeout: 30_000 });
}

export async function selectRibbonTab(page: Page, tab: 'home' | 'insert' | 'pageLayout' | 'formulas' | 'data'): Promise<void> {
  await page.getByTestId(`ribbon-tab-${tab}`).click();
  await expect(page.getByTestId(tab === 'home' ? 'home-ribbon-groups' : tab === 'insert' ? 'insert-ribbon-groups' : `ribbon-layout-${tab}`)).toBeVisible();
}

export async function assertSurfaceVisible(page: Page, surfaceId: string): Promise<void> {
  await expect(page.locator(`[data-ribbon-surface="${surfaceId}"]`).first()).toBeVisible();
}

export async function focusCanvas(page: Page): Promise<ReturnType<Page['getByTestId']>> {
  const canvas = page.getByTestId('sheet-canvas');
  await expect(canvas).toBeVisible();
  await canvas.focus();
  return canvas;
}
