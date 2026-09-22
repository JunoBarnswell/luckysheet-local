import { expect, test } from '@playwright/test';
import { ACCEPTANCE_LOCALES, ACCEPTANCE_VIEWPORTS, INSERT_SURFACE_CASES, INSERT_VARIANT_GROUPS } from './acceptance-matrix';
import { assertSurfaceVisible, installBrowserDiagnostics, openLocalWorkbook, selectRibbonTab } from './support/workbook-fixtures';

for (const locale of ACCEPTANCE_LOCALES) {
  for (const viewport of ACCEPTANCE_VIEWPORTS) {
    test.describe(`Insert ribbon acceptance ${locale} ${viewport.width}x${viewport.height}`, () => {
      test.use({ viewport, deviceScaleFactor: 1 });

      test('renders every Insert surface from the canonical catalog', async ({ page }) => {
        const diagnostics = installBrowserDiagnostics(page);
        await openLocalWorkbook(page, locale, `Insert matrix ${locale} ${viewport.width}`);
        await selectRibbonTab(page, 'insert');
        for (const entry of INSERT_SURFACE_CASES) await assertSurfaceVisible(page, entry.surface.id);
        diagnostics.assertClean();
      });

      test('exposes every typed gallery variant through its catalog root', async ({ page }) => {
        const diagnostics = installBrowserDiagnostics(page);
        await openLocalWorkbook(page, locale, `Insert variants ${locale} ${viewport.width}`);
        await selectRibbonTab(page, 'insert');
        for (const group of INSERT_VARIANT_GROUPS) {
          await assertSurfaceVisible(page, group.rootSurfaceId);
          await page.locator(`[data-ribbon-surface="${group.rootSurfaceId}"]`).first().click();
          for (const variant of group.variants) {
            await expect(page.locator(`[data-ribbon-variant="${variant.id}"]`).first()).toBeVisible();
          }
          await page.keyboard.press('Escape');
        }
        expect(INSERT_VARIANT_GROUPS.every((group) => group.variants.length > 0)).toBe(true);
        diagnostics.assertClean();
      });
    });
  }
}
