export type { WorkbookRole } from '@react-sheets/protocol';
export { workbookCapabilities, type WorkbookCapabilities } from './identity/workbook-capabilities';
export { createSpreadsheetSdk, type SpreadsheetSdk } from './sdk';
export { SdkError, type SdkErrorCode } from './error';
export type { AuthSession, AuthSnapshot, AuthPhase, AuthOptions, OidcConfiguration } from './auth/contract';
export type { IdentityActions, LocalUser } from './identity/contract';

export { useSdkServices, useWorkbook } from './react/runtime';
export type { CatalogEntry } from './workbooks/domain';

export type { DimensionsActions, DimensionResult } from './dimensions/contract';
export { MAX_EXCEL_ROW_HEIGHT_POINTS } from './dimensions/domain';
