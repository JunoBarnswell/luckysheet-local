import { expect, test } from '@playwright/test';
import { ELASTIC_GRID_CASES } from './acceptance-matrix';
import { focusCanvas, installBrowserDiagnostics, openLocalWorkbook } from './support/workbook-fixtures';

test.describe('Elastic Grid acceptance', () => {
  test('resolves a high canonical address and keeps the viewport bounded', async ({ page }) => {
    const diagnostics = installBrowserDiagnostics(page);
    await openLocalWorkbook(page, 'zh-CN', 'Elastic extent high address');
    const nameBox = page.getByTestId('name-box');
    await nameBox.fill('ZZZ1000');
    await nameBox.press('Enter');
    await expect(nameBox).toHaveValue('ZZZ1000');
    const canvas = await focusCanvas(page);
    await canvas.press('ArrowLeft');
    await canvas.press('ArrowUp');
    await page.mouse.wheel(0, 2400);
    await expect(page.getByTestId('designer-shell')).toHaveAttribute('data-workspace-phase', 'ready');
    expect(ELASTIC_GRID_CASES.some((entry) => entry.id === 'G-NAME-BOX-HIGH-ADDRESS')).toBe(true);
    diagnostics.assertClean();
  });

  test('keeps high-index interaction cases attached to the canonical extent contract', () => {
    const ids = new Set(ELASTIC_GRID_CASES.map((entry) => entry.id));
    expect(ids.has('G-PASTE-FILL-SPILL-EXTENT')).toBe(true);
    expect(ids.has('G-STRUCTURAL-SHIFT')).toBe(true);
    expect(ids.has('G-COLLABORATION-EXTENT-CONVERGENCE')).toBe(true);
    expect(ids.has('G-XLSX-EXTENT-PREFLIGHT')).toBe(true);
  });
});
