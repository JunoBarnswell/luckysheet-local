import type { CellStyle, CellValue, FormulaValue, RichTextRun, SheetKind, WorksheetPane, ProtectionRule } from '@react-sheets/core-model';
export type { DefinedNameModel, DefinedNameScope, ProtectionRule, ProtectionAllow, CellStyle, RichTextRun, WorksheetPane, BorderPlacement, BorderLine } from '@react-sheets/core-model';
export type { ClearFamily, FillDirection, FillMode, FillSeriesOptions } from '@react-sheets/sheet-features';
export type CellInput = { readonly kind: 'value'; readonly value: CellValue } | { readonly kind: 'formula'; readonly formula: string };
export interface WorksheetSnapshot {
  readonly id: string;
  readonly name: string;
  readonly kind: SheetKind;
  readonly rowCount: number;
  readonly columnCount: number;
  readonly hidden: boolean;
  readonly pane: Readonly<WorksheetPane>;
  readonly protectionRules: readonly Readonly<ProtectionRule>[];
  readonly defaultRowHeightPx: number;
  readonly defaultColumnWidthPx: number;
  readonly rowHeightsPx: Readonly<Record<number, number>>;
  readonly columnWidthsPx: Readonly<Record<number, number>>;
  readonly hiddenRows: readonly number[];
  readonly hiddenColumns: readonly number[];
  readonly merges: readonly Readonly<import('@react-sheets/core-model').MergeSpan>[];
}
export interface WorksheetCreateOptions { readonly name: string; readonly rowCount?: number; readonly columnCount?: number; }
export interface RangeStyleOptions { readonly numberFormat?: string; readonly replaceStyle?: boolean; readonly clearNumberFormat?: boolean; }
export interface CellSnapshot {
  readonly value: CellValue;
  readonly formula?: string;
  readonly calculatedValue: FormulaValue;
  readonly formulaHidden: boolean;
  readonly style?: Readonly<CellStyle>;
  readonly numberFormat?: string;
  readonly richText?: readonly Readonly<RichTextRun>[];
}
export interface ExternalLinkSnapshot {
  readonly token: string;
  readonly state: string;
  readonly sourceRevision: number;
  readonly error?: Readonly<{ code: string; message: string }>;
}
