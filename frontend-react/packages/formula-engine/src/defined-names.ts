import type { CellAddress } from './ast';
/** Formula-engine representation of the workbook's canonical scoped names. */
export interface FormulaDefinedName {
  readonly name: string;
  readonly formula: string;
  readonly scope: 'workbook' | 'sheet';
  readonly sheetId?: string;
  readonly anchor?: CellAddress;
}
