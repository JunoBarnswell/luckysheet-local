import assert from 'node:assert/strict';
import test from 'node:test';
import { WorkbookModel } from '@react-sheets/core-model';
import { normalizeAutoFilterModel } from '@react-sheets/sheet-features';
import { createWorkbookRowVisibilityResolver } from './formula-visibility';

test('row visibility invalidation recomputes filter-hidden rows after inputs change', () => {
  const workbook = new WorkbookModel('visibility-cache', 'Visibility');
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.cells.set(0, 0, { value: 'Value' });
  sheet.cells.set(1, 0, { value: 'Alpha' });
  sheet.cells.set(2, 0, { value: 'Beta' });
  sheet.autoFilter = normalizeAutoFilterModel({
    sheetId: sheet.id,
    range: { sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 0 },
    columns: {
      0: {
        column: 0,
        showButton: true,
        hiddenButton: false,
        criterion: { kind: 'custom', join: 'and', conditions: [{ operator: 'equals', value: 'Alpha' }] },
      },
    },
  });
  const visibility = createWorkbookRowVisibilityResolver(workbook, '1900', () => undefined);

  assert.equal(visibility.resolve(sheet.id, 2).filterHidden, true);
  sheet.cells.set(2, 0, { value: 'Alpha' });
  assert.equal(visibility.resolve(sheet.id, 2).filterHidden, true);
  visibility.invalidate();
  assert.equal(visibility.resolve(sheet.id, 2).filterHidden, false);
});
