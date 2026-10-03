import type { SdkError } from '../error';

export type DataActionResult = { readonly status: 'applied' } | { readonly status: 'rejected'; readonly error: SdkError };
export type SubtotalFunction = 'AVERAGE' | 'COUNT' | 'COUNTA' | 'MAX' | 'MIN' | 'PRODUCT' | 'STDEV' | 'STDEVP' | 'SUM' | 'VAR' | 'VARP';
export interface DataRangeAddress {
  readonly sheetId: string;
  readonly address: string;
}
export interface SubtotalOptions {
  readonly range?: DataRangeAddress;
  readonly groupColumn?: number;
  readonly valueColumn?: number;
  readonly functionName?: SubtotalFunction;
  /** Always excludes filter-hidden rows; this also excludes manual/outline-hidden rows. */
  readonly excludeHiddenRows?: boolean;
}
export interface DataActions {
  quickSort(ascending: boolean): Promise<DataActionResult>;
  toggleFilter(): Promise<DataActionResult>;
  clearFilter(): Promise<DataActionResult>;
  textToColumns(input?: { readonly delimiter?: string; readonly maxColumns?: number }): Promise<DataActionResult>;
  removeDuplicates(input?: { readonly columns?: readonly number[]; readonly hasHeader?: boolean }): Promise<DataActionResult>;
  subtotal(input?: SubtotalOptions): Promise<DataActionResult>;
}
