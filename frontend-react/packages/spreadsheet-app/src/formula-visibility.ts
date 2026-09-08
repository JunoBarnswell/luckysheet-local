import type { WorkbookModel, WorksheetModel } from '@react-sheets/core-model';
import { executeKernelFilter, resolvedVisibilityFromKernel, resolveAutoFilters, type ResolvedVisibility } from '@react-sheets/sheet-features';
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

export type KernelVisibilityResolver = (sheet: WorksheetModel) => ResolvedVisibility;

export function createKernelWorkbookVisibilityResolver(
  workbook: WorkbookModel,
  revision: () => number,
): KernelVisibilityResolver {
  return (sheet) => {
    const filters = resolveAutoFilters(sheet);
    const owners = filters.map(({ owner, autoFilter }) => ({
      id: owner.kind === 'worksheet' ? `${sheet.id}:worksheet` : `${sheet.id}:table:${owner.tableId}`,
      range: structuredClone(autoFilter.range),
      // Criteria cross the boundary in their canonical typed form. Rust owns
      // conjunctions, dynamic/date/color/icon/top10-percent semantics.
      columns: Object.values(autoFilter.columns).filter((column) => column.criterion).map((column) => ({ column: column.column, predicate: column.criterion })),
    }));
    const range = { sheetId: sheet.id, startRow: 0, endRow: Math.max(0, sheet.rowCount - 1), startColumn: 0, endColumn: Math.max(0, sheet.columnCount - 1) };
    const result = executeKernelFilter(workbook.unitId, { revision: revision(), range, owners, limit: sheet.rowCount });
    return resolvedVisibilityFromKernel(sheet, result, range.startRow);
  };
}

/**
 * Workbook-owned visibility projection shared by canvas/filter/formula paths.
 * The resolver caches only the derived row flags; CellMatrix and filter models
 * remain the sole sources of values and visibility inputs.
 */
export function createWorkbookRowVisibilityResolver(
  workbook: WorkbookModel,
  resolveKernelVisibility: KernelVisibilityResolver,
): WorkbookRowVisibilityResolver {
  let revision = 0;
  const caches = new Map<string, VisibilityCache>();

  const rebuild = (sheet: WorksheetModel): VisibilityCache => {
    const kernelVisibility = resolveKernelVisibility(sheet);
    const rows = new Map<number, RowVisibility>();
    for (const [row, flags] of kernelVisibility.rows) rows.set(row, { ...flags });
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
