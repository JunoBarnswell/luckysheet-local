import { expect, test } from '@playwright/test';
import { PERSISTENCE_CASES } from './acceptance-matrix';
import { focusCanvas, installBrowserDiagnostics, openLocalWorkbook } from './support/workbook-fixtures';

test.describe('Persistence acceptance', () => {
  test('preserves a local authored value through reload without remote traffic', async ({ page }) => {
    const apiRequests: string[] = [];
    let websocketCount = 0;
    page.on('request', (request) => {
      if (new URL(request.url()).pathname.startsWith('/api/')) apiRequests.push(request.url());
    });
    page.on('websocket', (socket) => {
      if (new URL(socket.url()).pathname === '/ws') websocketCount += 1;
    });
    const diagnostics = installBrowserDiagnostics(page);
    await openLocalWorkbook(page, 'zh-CN', 'Persistence round-trip');
    const canvas = await focusCanvas(page);
    await page.keyboard.type('persisted-acceptance-value');
    await page.keyboard.press('Enter');
    await canvas.press('ArrowUp');
    await expect(page.getByTestId('formula-input')).toHaveValue('persisted-acceptance-value');
    await page.reload();
    await expect(page.getByTestId('designer-shell')).toHaveAttribute('data-workspace-phase', 'ready');
    await expect(page.getByTestId('formula-input')).toHaveValue('');
    await page.getByTestId('name-box').fill('A1');
    await page.getByTestId('name-box').press('Enter');
    await expect(page.getByTestId('formula-input')).toHaveValue('persisted-acceptance-value');
    expect(apiRequests).toEqual([]);
    expect(websocketCount).toBe(0);
    expect(PERSISTENCE_CASES.some((entry) => entry.id === 'G-LOCAL-SAVE-RELOAD')).toBe(true);
    diagnostics.assertClean();
  });
});
