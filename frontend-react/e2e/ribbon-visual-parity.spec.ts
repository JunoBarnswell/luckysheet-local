import { expect, test } from '@playwright/test';
import { ACCEPTANCE_LOCALES, ACCEPTANCE_RIBBON_TABS, ACCEPTANCE_VIEWPORTS, RIBBON_VISUAL_GOLDEN_CASES } from './acceptance-matrix';
import { installBrowserDiagnostics, openLocalWorkbook, selectRibbonTab } from './support/workbook-fixtures';

for (const locale of ACCEPTANCE_LOCALES) {
  for (const viewport of ACCEPTANCE_VIEWPORTS) {
    test.describe(`Ribbon visual golden ${locale} ${viewport.width}x${viewport.height}`, () => {
      test.use({ viewport, deviceScaleFactor: 1 });

      test('matches the Designer shell and every primary ribbon tab visual golden', async ({ page }) => {
        const diagnostics = installBrowserDiagnostics(page);
        await openLocalWorkbook(page, locale, `Ribbon golden ${locale} ${viewport.width}`);
        const shellGolden = RIBBON_VISUAL_GOLDEN_CASES.find((entry) => entry.locale === locale && entry.viewport.width === viewport.width && entry.tab === 'home');
        if (!shellGolden) throw new Error(`Missing shell visual golden contract for ${locale} ${viewport.width}x${viewport.height}`);
        await expect(page.getByTestId('designer-shell')).toHaveScreenshot(shellGolden.shellScreenshot, { animations: 'disabled', caret: 'hide' });
        for (const tab of ACCEPTANCE_RIBBON_TABS) {
          const golden = RIBBON_VISUAL_GOLDEN_CASES.find((entry) => entry.locale === locale && entry.viewport.width === viewport.width && entry.tab === tab);
          if (!golden) throw new Error(`Missing visual golden contract for ${tab}/${locale}/${viewport.width}x${viewport.height}`);
          await selectRibbonTab(page, tab);
          const ribbon = page.getByTestId(tab === 'home' ? 'home-ribbon-groups' : tab === 'insert' ? 'insert-ribbon-groups' : `ribbon-layout-${tab}`);
          await expect(ribbon).toHaveAttribute('data-ribbon-layout', tab);
          await expect(ribbon).toHaveScreenshot(golden.screenshot, { animations: 'disabled', caret: 'hide' });
        }
        diagnostics.assertClean();
      });
    });
  }
}
