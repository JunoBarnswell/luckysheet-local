import { expect, test } from '@playwright/test';
import { ACCEPTANCE_LOCALES, ACCEPTANCE_VIEWPORTS } from './acceptance-matrix';
import { installBrowserDiagnostics, openLocalWorkbook, selectRibbonTab } from './support/workbook-fixtures';

for (const locale of ACCEPTANCE_LOCALES) {
  for (const viewport of ACCEPTANCE_VIEWPORTS) {
    test.describe(`Ribbon visual golden ${locale} ${viewport.width}x${viewport.height}`, () => {
      test.use({ viewport, deviceScaleFactor: 1 });

      test('matches the Home and Insert structural visual goldens', async ({ page }) => {
        const diagnostics = installBrowserDiagnostics(page);
        await openLocalWorkbook(page, locale, `Ribbon golden ${locale} ${viewport.width}`);
        await selectRibbonTab(page, 'home');
        const home = page.getByTestId('home-ribbon-groups');
        await expect(home).toHaveScreenshot(`home-${locale}-${viewport.width}x${viewport.height}.png`, { animations: 'disabled', caret: 'hide' });
        await selectRibbonTab(page, 'insert');
        const insert = page.getByTestId('insert-ribbon-groups');
        await expect(insert).toHaveScreenshot(`insert-${locale}-${viewport.width}x${viewport.height}.png`, { animations: 'disabled', caret: 'hide' });
        diagnostics.assertClean();
      });
    });
  }
}
