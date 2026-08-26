import type { CellAddress } from './ast';
import type { RangeDependency } from './range-index';
import type { FormulaEvaluationContext } from './evaluator';
import type { FormulaValue } from './values';

export type ReferenceFormulaKind = 'ordinary' | 'subtotal' | 'aggregate';

export interface RowVisibility {
  readonly manualHidden: boolean;
  readonly filterHidden: boolean;
  readonly outlineHidden: boolean;
}

export const VISIBLE_ROW: RowVisibility = Object.freeze({
  manualHidden: false,
  filterHidden: false,
  outlineHidden: false,
});

export interface RowVisibilityResolver {
  resolve(sheetId: string, row: number): RowVisibility;
  snapshot?(): FormulaVisibilitySnapshot;
}

export interface FormulaVisibilitySnapshot {
  readonly revision: number;
  readonly rows: readonly FormulaVisibilityRowSnapshot[];
}

export interface FormulaVisibilityRowSnapshot extends RowVisibility {
  readonly sheetId: string;
  readonly row: number;
}

export interface ReferenceCell {
  readonly address: CellAddress;
  readonly value: FormulaValue;
  readonly visibility: RowVisibility;
  readonly formulaKind: ReferenceFormulaKind;
}

/**
 * The sole range-to-cell projection for provenance/visibility-aware functions.
 * It is intentionally lazy so large references do not become a second matrix.
 */
export function createReferenceCursor(
  range: RangeDependency,
  context: Pick<FormulaEvaluationContext, 'readCell' | 'rowVisibility' | 'readFormulaKind'>,
): Iterable<ReferenceCell> {
  return (function* cursor(): Generator<ReferenceCell> {
    for (let row = range.start.row; row <= range.end.row; row += 1) {
      for (let column = range.start.column; column <= range.end.column; column += 1) {
        const address = { sheetId: range.start.sheetId, row, column };
        yield {
          address,
          value: context.readCell(address),
          visibility: context.rowVisibility?.resolve(address.sheetId, row) ?? VISIBLE_ROW,
          formulaKind: context.readFormulaKind?.(address) ?? 'ordinary',
        };
      }
    }
  })();
}

export function createSnapshotVisibilityResolver(snapshot: FormulaVisibilitySnapshot): RowVisibilityResolver {
  const rows = new Map<string, RowVisibility>();
  for (const entry of snapshot.rows) rows.set(`${entry.sheetId}\u0000${entry.row}`, {
    manualHidden: entry.manualHidden,
    filterHidden: entry.filterHidden,
    outlineHidden: entry.outlineHidden,
  });
  return {
    resolve: (sheetId, row) => rows.get(`${sheetId}\u0000${row}`) ?? VISIBLE_ROW,
    snapshot: () => structuredClone(snapshot),
  };
}

export function assertFormulaVisibilitySnapshot(value: unknown): asserts value is FormulaVisibilitySnapshot {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) throw new Error('Formula visibility snapshot must be an object');
  const candidate = value as Record<string, unknown>;
  if (!Number.isSafeInteger(candidate.revision) || Number(candidate.revision) < 0 || !Array.isArray(candidate.rows)) {
    throw new Error('Formula visibility snapshot has invalid revision or rows');
  }
  for (const row of candidate.rows) {
    if (typeof row !== 'object' || row === null || Array.isArray(row)) throw new Error('Formula visibility row is invalid');
    const entry = row as Record<string, unknown>;
    if (typeof entry.sheetId !== 'string' || entry.sheetId.length === 0
      || !Number.isSafeInteger(entry.row) || Number(entry.row) < 0
      || typeof entry.manualHidden !== 'boolean'
      || typeof entry.filterHidden !== 'boolean'
      || typeof entry.outlineHidden !== 'boolean') {
      throw new Error('Formula visibility row is invalid');
    }
  }
}
