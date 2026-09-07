import type { CellAddress } from './ast';
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
