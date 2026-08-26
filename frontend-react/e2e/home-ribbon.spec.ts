import { expect, test } from '@playwright/test';
import { ACCEPTANCE_LOCALES, ACCEPTANCE_VIEWPORTS, HOME_BEHAVIOR_CASES, HOME_SURFACE_CASES } from './acceptance-matrix';
import { assertSurfaceVisible, focusCanvas, installBrowserDiagnostics, openLocalWorkbook, selectRibbonTab } from './support/workbook-fixtures';

for (const locale of ACCEPTANCE_LOCALES) {
  for (const viewport of ACCEPTANCE_VIEWPORTS) {
    test.describe(`Home ribbon acceptance ${locale} ${viewport.width}x${viewport.height}`, () => {
      test.use({ viewport, deviceScaleFactor: 1 });

      test('renders every Home surface from the canonical catalog', async ({ page }) => {
        const diagnostics = installBrowserDiagnostics(page);
        await openLocalWorkbook(page, locale, `Home matrix ${locale} ${viewport.width}`);
        await selectRibbonTab(page, 'home');
        for (const entry of HOME_SURFACE_CASES) {
          if (entry.surface.menuId) {
            await assertSurfaceVisible(page, entry.surface.menuId);
            await page.locator(`[data-ribbon-surface="${entry.surface.menuId}"]`).first().click();
            await assertSurfaceVisible(page, entry.surface.id);
            await page.keyboard.press('Escape');
          } else {
            await assertSurfaceVisible(page, entry.surface.id);
          }
        }
        diagnostics.assertClean();
      });

      test('keeps the Home authored smoke on one selection and history path', async ({ page }) => {
        const diagnostics = installBrowserDiagnostics(page);
        await openLocalWorkbook(page, locale, `Home transaction ${locale} ${viewport.width}`);
        const canvas = await focusCanvas(page);
        await page.keyboard.type('home-matrix-value');
        await page.keyboard.press('Enter');
        await canvas.press('ArrowUp');
        await expect(page.getByTestId('formula-input')).toHaveValue('home-matrix-value');
        await canvas.press('Control+Z');
        await expect(page.getByTestId('formula-input')).toHaveValue('');
        await canvas.press('Control+Y');
        await expect(page.getByTestId('formula-input')).toHaveValue('home-matrix-value');
        expect(HOME_BEHAVIOR_CASES.every((entry) => entry.layers.includes('contract'))).toBe(true);
        diagnostics.assertClean();
      });
    });
  }
}
