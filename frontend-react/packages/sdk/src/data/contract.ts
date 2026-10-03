import type { SdkError } from '../error';

export type DataActionResult = { readonly status: 'applied' } | { readonly status: 'rejected'; readonly error: SdkError };
export interface DataActions {
  quickSort(ascending: boolean): Promise<DataActionResult>;
  toggleFilter(): Promise<DataActionResult>;
  clearFilter(): Promise<DataActionResult>;
  textToColumns(input?: { readonly delimiter?: string; readonly maxColumns?: number }): Promise<DataActionResult>;
  removeDuplicates(input?: { readonly columns?: readonly number[]; readonly hasHeader?: boolean }): Promise<DataActionResult>;
  subtotal(input?: { readonly groupColumn?: number; readonly valueColumn?: number; readonly functionName?: 'SUM' | 'COUNT' | 'AVERAGE' }): Promise<DataActionResult>;
}
