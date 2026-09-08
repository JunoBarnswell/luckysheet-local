import { expect, test } from '@playwright/test';
import { INSERT_CHART_FAMILIES } from '../apps/web/src/components/insert-ribbon-catalog';
import { ACCEPTANCE_LOCALES, ACCEPTANCE_VIEWPORTS, INSERT_SURFACE_CASES, INSERT_VARIANT_GROUPS } from './acceptance-matrix';
import { installBrowserDiagnostics, openConnectedWorkbook, revealRibbonSurface, revealRibbonSurfaceById, selectRibbonTab } from './support/workbook-fixtures';

for (const locale of ACCEPTANCE_LOCALES) {
  for (const viewport of ACCEPTANCE_VIEWPORTS) {
    test.describe(`Insert ribbon acceptance ${locale} ${viewport.width}x${viewport.height}`, () => {
      test.use({ viewport, deviceScaleFactor: 1 });

      test('renders every Insert surface from the canonical catalog', async ({ page }) => {
        const diagnostics = installBrowserDiagnostics(page);
        await openConnectedWorkbook(page, locale, `Insert matrix ${locale} ${viewport.width}`);
        await selectRibbonTab(page, 'insert');
        for (const entry of INSERT_SURFACE_CASES) await revealRibbonSurface(page, entry.surface);
        await page.keyboard.press('Escape');
        diagnostics.assertClean();
      });

      test('exposes every typed gallery variant through its catalog root', async ({ page }) => {
        const diagnostics = installBrowserDiagnostics(page);
        await openConnectedWorkbook(page, locale, `Insert variants ${locale} ${viewport.width}`);
        await selectRibbonTab(page, 'insert');
        for (const group of INSERT_VARIANT_GROUPS) {
          await revealRibbonSurfaceById(page, group.rootSurfaceId);
          const familyMenus = page.locator('[data-ribbon-gallery-family]');
          if (group.rootSurfaceId === 'charts.gallery' && await familyMenus.count() > 0) {
            for (const family of INSERT_CHART_FAMILIES) {
              await page.locator(`[data-ribbon-gallery-family="${family.id}"]:visible`).click();
              for (const variant of family.variants) await expect(page.locator(`[data-ribbon-variant="${variant.id}"]:visible`).first()).toBeVisible();
              await page.keyboard.press('Escape');
            }
          } else {
            const firstVariant = page.locator(`[data-ribbon-variant="${group.variants[0]!.id}"]:visible`).first();
            if (!await firstVariant.isVisible()) {
              const menu = page.locator(`[data-ribbon-menu="${group.rootSurfaceId}"]`).first();
              await expect(menu).toBeVisible();
              await menu.click();
            }
            for (const variant of group.variants) await expect(page.locator(`[data-ribbon-variant="${variant.id}"]:visible`).first()).toBeVisible();
          }
          await page.keyboard.press('Escape');
        }
        expect(INSERT_VARIANT_GROUPS.every((group) => group.variants.length > 0)).toBe(true);
        diagnostics.assertClean();
      });
    });
  }
}
