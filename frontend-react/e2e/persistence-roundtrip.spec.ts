import { expect, test } from '@playwright/test';
import { PERSISTENCE_CASES } from './acceptance-matrix';
import { focusCanvas, installBrowserDiagnostics, openConnectedWorkbook, waitForServerSaved } from './support/workbook-fixtures';

test.describe('Persistence acceptance', () => {
  test('retains a server-acknowledged cell edit after reload', async ({ page }) => {
    const apiRequests: string[] = [];
    page.on('request', (request) => {
      if (new URL(request.url()).pathname.startsWith('/api/')) apiRequests.push(request.url());
    });
    const diagnostics = installBrowserDiagnostics(page);
    await openConnectedWorkbook(page, 'zh-CN', 'Persistence round-trip');
    const canvas = await focusCanvas(page);
    await page.keyboard.type('persisted-acceptance-value');
    await page.keyboard.press('Enter');
    await canvas.press('ArrowUp');
    await expect(page.getByTestId('formula-input')).toHaveValue('persisted-acceptance-value');
    await waitForServerSaved(page);
    await page.reload();
    await expect(page.getByTestId('designer-shell')).toHaveAttribute('data-workspace-phase', 'ready', { timeout: 30_000 });
    await page.getByTestId('name-box').fill('A1');
    await page.getByTestId('name-box').press('Enter');
    await expect(page.getByTestId('formula-input')).toHaveValue('persisted-acceptance-value');
    expect(apiRequests.length).toBeGreaterThan(0);
    expect(PERSISTENCE_CASES.some((entry) => entry.id === 'G-SERVER-COLLABORATION-ROUNDTRIP')).toBe(true);
    diagnostics.assertClean();
  });
});
