/** Canonical JSON contract from kernel/host-protocol.md. */
export const KERNEL_PROTOCOL_VERSION = 1 as const;
export const KERNEL_WORKBOOK_MANIFEST_VERSION = 11 as const;
export const KERNEL_PAGE_ROWS = 1024 as const;
export const KERNEL_PAGE_COLUMNS = 32 as const;

export type KernelScalar = string | number | boolean | null | KernelFormulaError;
export interface KernelFormulaError { readonly kind: 'error'; readonly code: string; readonly message: string; }
export interface KernelCell { readonly value: KernelScalar; readonly formula?: string; readonly [authoredMetadata: string]: unknown; }
export interface KernelCellAddress { readonly sheetId: string; readonly row: number; readonly column: number; }
export interface KernelRangeRef { readonly sheetId: string; readonly startRow: number; readonly endRow: number; readonly startColumn: number; readonly endColumn: number; }
export interface KernelSheetManifest { readonly sheetId: string; readonly name: string; readonly rowCount: number; readonly columnCount: number; readonly metadata: Record<string, unknown>; }
export interface KernelPageDescriptor { readonly sheetId: string; readonly pageRow: number; readonly pageColumn: number; readonly revision: number; readonly checksum: string; readonly byteLength: number; readonly cellCount: number; readonly occupiedRange: KernelRangeRef | null; }
export interface WorkbookManifest { readonly schema: 'WorkbookManifest'; readonly version: typeof KERNEL_WORKBOOK_MANIFEST_VERSION; readonly unitId: string; readonly name: string; readonly revision: number; readonly sheets: KernelSheetManifest[]; readonly pages: KernelPageDescriptor[]; readonly metadata: Record<string, unknown>; }
export interface KernelPagePayload extends KernelPageDescriptor { readonly payloadBase64: string; }
export interface KernelPageKey { readonly sheetId: string; readonly pageRow: number; readonly pageColumn: number; }
export interface KernelRevisionPin { readonly unitId: string; readonly revision: number; }
export interface KernelChangeSet { readonly operationId: string; readonly baseRevision: number; readonly revision: number; readonly manifest: WorkbookManifest; readonly pages: KernelPagePayload[]; readonly removedPages: KernelPageKey[]; readonly affectedRanges: KernelRangeRef[]; }
export interface KernelAnalyticsRequest { readonly kind: 'filter' | 'query' | 'pivot'; readonly params: unknown; }
export interface KernelAnalyticsResponse { readonly kind: KernelAnalyticsRequest['kind']; readonly revision: number; readonly [resultField: string]: unknown; }

export type KernelOperation = 'init' | 'open' | 'create' | 'manifest' | 'sheet.stats' | 'cell.get' | 'range.get' | 'page.get' | 'page.load' | 'command' | 'close' | 'formula.evaluate' | 'formula.recalculate' | 'analytics.execute' | 'geometry.computePaneMap' | 'geometry.hitTest' | 'geometry.cellRect' | 'geometry.headerRect' | 'document.import' | 'document.export';
export interface KernelInitResponse { readonly protocolVersion: typeof KERNEL_PROTOCOL_VERSION; readonly manifestVersion: typeof KERNEL_WORKBOOK_MANIFEST_VERSION; readonly operations: KernelOperation[]; }
export interface KernelOpenRequest { readonly manifest: WorkbookManifest; readonly pages?: KernelPagePayload[]; }
export interface KernelCreateRequest { readonly unitId: string; readonly name: string; readonly sheets: KernelSheetManifest[]; }
export type KernelCreateResponse = WorkbookManifest;
export interface KernelManifestRequest { readonly unitId: string; readonly revision: number; }
export interface KernelCellGetRequest extends KernelRevisionPin { readonly address: KernelCellAddress; }
export interface KernelCellGetResponse { readonly revision: number; readonly cell: KernelCell | null; }
export interface KernelRangeGetRequest extends KernelRevisionPin { readonly range: KernelRangeRef; }
export interface KernelRangeGetResponse { readonly revision: number; readonly cells: Array<{ address: KernelCellAddress; cell: KernelCell }>; }
export interface KernelPageGetRequest extends KernelRevisionPin { readonly sheetId: string; readonly pageRow: number; readonly pageColumn: number; }
export interface KernelPageLoadRequest extends KernelRevisionPin { readonly page: KernelPagePayload; }
export interface KernelCommandRequest { readonly unitId: string; readonly baseRevision: number; readonly operationId: string; readonly commandId: string; readonly params: unknown; }
export interface KernelCloseResponse { readonly closed: true; }
export interface KernelFormulaEvaluateRequest extends KernelRevisionPin { readonly address: KernelCellAddress; readonly formula: string; }
export interface KernelFormulaEvaluateResponse { readonly revision: number; readonly value: KernelScalar | KernelScalar[][]; }
export interface KernelFormulaRecalculateRequest extends KernelRevisionPin {}
export interface KernelFormulaRecalculateResponse { readonly revision: number; readonly generation: number; readonly recalculatedCount: number; readonly pendingRecalculation: boolean; }
export interface KernelAnalyticsExecuteRequest extends KernelRevisionPin { readonly request: KernelAnalyticsRequest; }
export interface KernelError { readonly code: string; readonly message: string; readonly object?: string; readonly recovery: string; }

export interface KernelOperationMap {
  init: { request: Record<string, never>; response: KernelInitResponse };
  open: { request: KernelOpenRequest; response: { unitId: string; revision: number; pageCount: number } };
  create: { request: KernelCreateRequest; response: KernelCreateResponse };
  manifest: { request: KernelManifestRequest; response: WorkbookManifest };
  'sheet.stats': { request: KernelRevisionPin & { sheetId: string }; response: { sheetId: string; cellCount: number; occupiedRange: KernelRangeRef | null } };
  'cell.get': { request: KernelCellGetRequest; response: KernelCellGetResponse };
  'range.get': { request: KernelRangeGetRequest; response: KernelRangeGetResponse };
  'page.get': { request: KernelPageGetRequest; response: KernelPagePayload };
  'page.load': { request: KernelPageLoadRequest; response: { revision: number; loaded: KernelPageDescriptor } };
  command: { request: KernelCommandRequest; response: KernelChangeSet };
  close: { request: { unitId: string }; response: KernelCloseResponse };
  'formula.evaluate': { request: KernelFormulaEvaluateRequest; response: KernelFormulaEvaluateResponse };
  'formula.recalculate': { request: KernelFormulaRecalculateRequest; response: KernelFormulaRecalculateResponse };
  'analytics.execute': { request: KernelAnalyticsExecuteRequest; response: KernelAnalyticsResponse };
  'geometry.computePaneMap': { request: unknown; response: unknown };
  'geometry.hitTest': { request: unknown; response: unknown };
  'geometry.cellRect': { request: unknown; response: unknown };
  'geometry.headerRect': { request: unknown; response: unknown };
  'document.import': { request: unknown; response: unknown };
  'document.export': { request: unknown; response: unknown };
}
export type KernelRequest<Op extends KernelOperation> = KernelOperationMap[Op]['request'];
export type KernelResponse<Op extends KernelOperation> = KernelOperationMap[Op]['response'];
/** kernelInvoke is synchronous after initialization and throws KernelError. */
export interface KernelClient { initializeKernel(): Promise<void>; kernelInvoke<Op extends KernelOperation>(operation: Op, params: KernelRequest<Op>): KernelResponse<Op>; }
