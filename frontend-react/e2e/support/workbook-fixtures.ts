import { expect, test, type Page } from '@playwright/test';
import { appendFileSync, mkdirSync, readFileSync } from 'node:fs';
import path from 'node:path';
import { RIBBON_TAB_SURFACES, type RibbonSurfaceDefinition } from '@react-sheets/spreadsheet-app';

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

interface ConnectedAuthFixture {
  readonly authority: string;
  readonly clientId: string;
  readonly userJson: string;
}

const connectedWorkbookLedgerPath = path.resolve(import.meta.dirname, '../../test-results/connected-workbooks.ndjson');

function recordConnectedWorkbook(unitId: string): void {
  mkdirSync(path.dirname(connectedWorkbookLedgerPath), { recursive: true });
  appendFileSync(connectedWorkbookLedgerPath, `${JSON.stringify({ unitId })}\n`, 'utf8');
}

function connectedAuthFixture(): ConnectedAuthFixture | null {
  const authority = process.env.E2E_OIDC_AUTHORITY?.trim();
  const clientId = process.env.E2E_OIDC_CLIENT_ID?.trim();
  const userFile = process.env.E2E_OIDC_USER_FILE?.trim();
  if (!authority || !clientId || !userFile) return null;
  const userJson = readFileSync(userFile, 'utf8');
  const user = JSON.parse(userJson) as { access_token?: unknown; expires_at?: unknown; profile?: { sub?: unknown } };
  if (typeof user.access_token !== 'string' || typeof user.expires_at !== 'number' || user.expires_at <= Date.now() / 1000 || typeof user.profile?.sub !== 'string') {
    throw new Error('E2E_OIDC_USER_FILE must contain a valid, unexpired oidc-client-ts User payload');
  }
  return { authority, clientId, userJson };
}

export async function openConnectedWorkbook(page: Page, locale: FixtureLocale, name: string): Promise<void> {
  const auth = connectedAuthFixture();
  if (!auth && process.env.CI) {
    throw new Error('Connected browser acceptance requires E2E_OIDC_AUTHORITY, E2E_OIDC_CLIENT_ID, and E2E_OIDC_USER_FILE in CI');
  }
  test.skip(!auth, 'Requires a configured backend and a real OIDC user via E2E_OIDC_AUTHORITY, E2E_OIDC_CLIENT_ID, and E2E_OIDC_USER_FILE');
  if (!auth) return;
  await page.addInitScript((value) => window.localStorage.setItem('react-sheets:locale', value), locale);
  await page.addInitScript(({ key, userJson }) => window.sessionStorage.setItem(key, userJson), {
    key: `oidc.user:${auth.authority}:${auth.clientId}`,
    userJson: auth.userJson,
  });
  await page.goto('/workbooks');
  await expect(page.getByTestId('workbook-hub')).toBeVisible();
  await page.getByRole('button', { name: '新建工作簿' }).click();
  const dialog = page.getByTestId('create-workbook-dialog');
  await expect(dialog).toBeVisible();
  await dialog.getByLabel('工作簿名称').fill(name);
  await expect(dialog.getByLabel('保存位置')).toBeEnabled();
  await dialog.getByRole('button', { name: '创建工作簿' }).click();
  await expect(page).toHaveURL(/\/workbooks\/[^/]+(?:\?.*)?$/);
  const unitId = decodeURIComponent(new URL(page.url()).pathname.split('/').filter(Boolean).at(-1) ?? '');
  if (!unitId) throw new Error(`Connected workbook URL does not contain a unit id: ${page.url()}`);
  recordConnectedWorkbook(unitId);
  await expect(page.getByTestId('designer-shell')).toHaveAttribute('data-workspace-phase', 'ready', { timeout: 30_000 });
}

export async function selectRibbonTab(page: Page, tab: 'home' | 'insert' | 'pageLayout' | 'formulas' | 'data'): Promise<void> {
  await page.getByTestId(`ribbon-tab-${tab}`).click();
  await expect(page.getByTestId(tab === 'home' ? 'home-ribbon-groups' : tab === 'insert' ? 'insert-ribbon-groups' : `ribbon-layout-${tab}`)).toBeVisible();
}

export async function assertSurfaceVisible(page: Page, surfaceId: string): Promise<void> {
  await expect(page.locator(`[data-ribbon-surface="${surfaceId}"]`).first()).toBeVisible();
}

export function ribbonSurface(surfaceId: string): RibbonSurfaceDefinition {
  const surface = RIBBON_TAB_SURFACES.find((candidate) => candidate.id === surfaceId);
  if (!surface) throw new Error(`Unknown ribbon surface: ${surfaceId}`);
  return surface;
}

export async function openRibbonGroup(page: Page, groupId: string): Promise<void> {
  const group = page.locator(`[data-ribbon-group="${groupId}"]`).first();
  await expect(group).toBeVisible();
  const layout = group.locator('xpath=ancestor::*[@data-ribbon-breakpoint][1]');
  if (await layout.getAttribute('data-ribbon-breakpoint') === 'wide') return;
  await group.getByRole('button').first().click();
}

/** Reveals one catalog surface through its real responsive group/menu path. */
export async function revealRibbonSurface(page: Page, surface: RibbonSurfaceDefinition): Promise<ReturnType<Page['locator']>> {
  await page.keyboard.press('Escape');
  const target = page.locator(`[data-ribbon-surface="${surface.id}"]`).first();
  if (!await target.isVisible()) await openRibbonGroup(page, surface.group);
  if (surface.menuId && !await target.isVisible()) {
    let menu = page.locator(`[data-ribbon-menu="${surface.menuId}"]`).first();
    if (!await menu.isVisible()) {
      await page.keyboard.press('Escape');
      const owner = RIBBON_TAB_SURFACES.find((candidate) => candidate.id === surface.menuId);
      if (!owner) throw new Error(`Unknown ribbon menu owner: ${surface.menuId}`);
      await openRibbonGroup(page, owner.group);
      menu = page.locator(`[data-ribbon-menu="${surface.menuId}"]`).first();
    }
    await expect(menu, `menu ${surface.menuId} must expose ${surface.id}`).toBeVisible();
    await menu.click();
  }
  await expect(target).toBeVisible();
  return target;
}

export async function revealRibbonSurfaceById(page: Page, surfaceId: string): Promise<ReturnType<Page['locator']>> {
  return revealRibbonSurface(page, ribbonSurface(surfaceId));
}

export async function waitForServerSaved(page: Page): Promise<void> {
  await expect(page.getByTestId('document-save-state')).toHaveAttribute('data-save-state', 'saved', { timeout: 30_000 });
}

export async function focusCanvas(page: Page): Promise<ReturnType<Page['getByTestId']>> {
  const canvas = page.getByTestId('sheet-canvas');
  await expect(canvas).toBeVisible();
  await canvas.focus();
  return canvas;
}
