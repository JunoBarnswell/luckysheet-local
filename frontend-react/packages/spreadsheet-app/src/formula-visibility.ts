import {
  computeFilterHiddenRows,
  computeOutlineHiddenRows,
} from '@react-sheets/sheet-features';
import {
  resolveFilterCellValue,
  type WorkbookModel,
  type WorksheetModel,
} from '@react-sheets/core-model';
import type {
  FormulaVisibilitySnapshot,
  RowVisibility,
  RowVisibilityResolver,
} from '@react-sheets/formula-engine';

interface VisibilityCache {
  readonly revision: number;
  readonly rows: Map<number, RowVisibility>;
}

export interface WorkbookRowVisibilityResolver extends RowVisibilityResolver {
  invalidate(): void;
}

/**
 * Workbook-owned visibility projection shared by canvas/filter/formula paths.
 * The resolver caches only the derived row flags; CellMatrix and filter models
 * remain the sole sources of values and visibility inputs.
 */
export function createWorkbookRowVisibilityResolver(
  workbook: WorkbookModel,
  dateSystem: '1900' | '1904',
  readFormulaValue: (sheet: WorksheetModel, row: number, column: number) => unknown,
): WorkbookRowVisibilityResolver {
  let revision = 0;
  const caches = new Map<string, VisibilityCache>();

  const rebuild = (sheet: WorksheetModel): VisibilityCache => {
    const readFilterCell = (row: number, column: number) => resolveFilterCellValue(
      sheet.cells.get(row, column),
      readFormulaValue(sheet, row, column),
      dateSystem,
    );
    const filterHidden = computeFilterHiddenRows(sheet, readFilterCell, dateSystem);
    const outlineHidden = computeOutlineHiddenRows(sheet);
    const rows = new Map<number, RowVisibility>();
    for (const row of sheet.hiddenRows) rows.set(row, {
      manualHidden: true,
      filterHidden: filterHidden.has(row),
      outlineHidden: outlineHidden.has(row),
    });
    for (const row of filterHidden) rows.set(row, {
      manualHidden: sheet.hiddenRows.has(row),
      filterHidden: true,
      outlineHidden: outlineHidden.has(row),
    });
    for (const row of outlineHidden) rows.set(row, {
      manualHidden: sheet.hiddenRows.has(row),
      filterHidden: filterHidden.has(row),
      outlineHidden: true,
    });
    const cache = { revision, rows };
    caches.set(sheet.id, cache);
    return cache;
  };

  const cacheFor = (sheet: WorksheetModel): VisibilityCache => {
    const cached = caches.get(sheet.id);
    return cached?.revision === revision ? cached : rebuild(sheet);
  };

  return {
    resolve: (sheetId, row) => cacheFor(workbook.getSheet(sheetId)).rows.get(row) ?? {
      manualHidden: false,
      filterHidden: false,
      outlineHidden: false,
    },
    invalidate: () => {
      revision += 1;
      caches.clear();
    },
    snapshot: (): FormulaVisibilitySnapshot => {
      const rows: FormulaVisibilitySnapshot['rows'][number][] = [];
      for (const sheet of workbook.getSheets()) {
        for (const [row, visibility] of cacheFor(sheet).rows) {
          rows.push({ sheetId: sheet.id, row, ...visibility });
        }
      }
      return { revision, rows };
    },
  };
}
