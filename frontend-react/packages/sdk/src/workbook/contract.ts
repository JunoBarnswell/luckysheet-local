import type { CellValue, FormulaValue } from '@react-sheets/core-model';
export interface CellSnapshot {
  readonly value: CellValue;
  readonly formula?: string;
  readonly calculatedValue: FormulaValue;
  readonly formulaHidden: boolean;
}
export interface ExternalLinkSnapshot {
  readonly token: string;
  readonly state: string;
  readonly sourceRevision: number;
  readonly error?: Readonly<{ code: string; message: string }>;
}
