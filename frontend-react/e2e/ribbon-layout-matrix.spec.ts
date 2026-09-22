import { expect, test } from '@playwright/test';
import { ACCEPTANCE_LOCALES, ACCEPTANCE_VIEWPORTS, DATA_CASES, FORMULAS_CASES, PAGE_LAYOUT_CASES } from './acceptance-matrix';
import { installBrowserDiagnostics, openLocalWorkbook, selectRibbonTab } from './support/workbook-fixtures';

const primaryTabCases = [
  ['pageLayout', PAGE_LAYOUT_CASES],
  ['formulas', FORMULAS_CASES],
  ['data', DATA_CASES],
] as const;

for (const locale of ACCEPTANCE_LOCALES) {
  for (const viewport of ACCEPTANCE_VIEWPORTS) {
    test.describe(`Primary ribbon layout matrix ${locale} ${viewport.width}x${viewport.height}`, () => {
      test.use({ viewport, deviceScaleFactor: 1 });

      for (const [tab, cases] of primaryTabCases) {
        test(`exposes ${tab} layout nodes with canonical command identity`, async ({ page }) => {
          const diagnostics = installBrowserDiagnostics(page);
          await openLocalWorkbook(page, locale, `${tab} matrix ${locale} ${viewport.width}`);
          await selectRibbonTab(page, tab);
          const layout = page.getByTestId(`ribbon-layout-${tab}`);
          await expect(layout).toHaveAttribute('data-ribbon-layout', tab);
          for (const entry of cases) {
            const command = page.locator(`[data-ribbon-command="${entry.commandIds[0]}"][data-ribbon-layout-node="${entry.nodeId}"]`).first();
            await expect(command, `${entry.id} must retain its layout-node and command identity`).toBeAttached();
          }
          diagnostics.assertClean();
        });
      }
    });
  }
}
