import { expect, test } from '@playwright/test';
import { PERSISTENCE_CASES } from './acceptance-matrix';
import { focusCanvas, installBrowserDiagnostics, openLocalWorkbook } from './support/workbook-fixtures';

test.describe('Persistence acceptance', () => {
  test('clears page-session local data on reload without remote traffic', async ({ page }) => {
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
    await expect(page.getByRole('heading', { name: '内存会话已重置' })).toBeVisible();
    await expect(page.getByText('本地工作簿只存在于当前页面的内存会话中；刷新或关闭页面后无法恢复。请返回工作簿中心重新创建或导入。')).toBeVisible();
    expect(apiRequests).toEqual([]);
    expect(websocketCount).toBe(0);
    expect(PERSISTENCE_CASES.some((entry) => entry.id === 'G-LOCAL-SAVE-RELOAD')).toBe(true);
    diagnostics.assertClean();
  });
});
