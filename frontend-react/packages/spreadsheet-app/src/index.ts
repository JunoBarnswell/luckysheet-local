export {
  SpreadsheetApplication,
  getInitialAppPhase,
  resolveUnitId,
  resolveActorId,
  resolveShareToken,
  type SpreadsheetApplicationOptions,
  type UiSnapshot,
} from './application';
export { useSpreadsheetApp, type UseSpreadsheetAppResult } from './react';
export {
  registerSpreadsheetFeatures,
  getFeatureRegistry,
  type SpreadsheetFeatureManifest,
} from './feature-registry';
export { registerEditingFeatures, buildSelectionSnapshot, type SetSelectionParams } from './features/editing/index';
export { registerDrawingFeature, DrawingRuntime } from './features/drawing/index';
export { EditSession } from './edit-session';
export { SelectionService, createInitialSelection, type SelectionState, type SelectionSnapshot } from './selection-service';
export { PermissionService, type ShareRole, type ActorContext, type PermissionCheckInput, type PermissionResult, type PermissionCapabilities } from './permission-service';
export { canExecuteCommand, buildPermissionCapabilities, type PermissionAction } from './features/permission';
export { buildCollaborationSnapshot, type CollaborationSnapshot } from './collaboration';
export { buildRestoreParams, revisionToHistoryMeta } from './features/history';
export {
  exchangeExportXlsx,
  exchangeImportXlsx,
  summarizeCompatibilityReport,
} from './features/xlsx';
export {
  buildPersistenceMeta,
  LocalWorkspaceStore,
  getLocalWorkspaceStore,
  WorkspacePersistence,
  buildWorkspaceRecord,
  verifyWorkspaceRecord,
  verifyPendingOperationJournal,
  type PersistenceSnapshotMeta,
  type WorkspaceRecord,
  type WorkspaceRecordInput,
  type PendingOperationJournal,
  type LocalWorkspaceSummary,
  type IndexedDbWorkspaceStoreOptions,
} from './features/persistence';
export {
  buildPrintSnapshot,
  summarizePrintSnapshot,
  printLayoutToPageSetup,
  pageSetupToPrintLayout,
  resolvePrintArea,
  buildPrintLayoutModel,
  toPrintPageSnapshots,
  type PrintSnapshot,
  type PrintPageSnapshot,
  type PrintPreviewCommandParams,
  type PrintAreaSetCommandParams,
  type PrintLayout,
  type PrintDocument,
  type PageSetup,
  type PrintArea,
  type PrintPageBreak,
} from './features/print';
export {
  buildQueryResultSnapshot,
  summarizeQueryResult,
  executeQueryDefinition,
  resolveLoadTarget,
  createInlineJsonQuery,
  queryResultToRangeValues,
  type QueryResultSnapshot,
  type QuerySessionEntry,
  type QueryLoadCommandPayload,
} from './features/query';
export type { QueryDefinition, LoadTarget, QueryStep } from './features/query/query-steps';
export type { ConnectorKind, QueryResult, DataConnector } from './features/query';
export {
  runAutomationScript,
  summarizeScriptResult,
  SAMPLE_AUTOMATION_SCRIPT,
  type AutomationSnapshot,
} from './features/automation';
export type { ScriptRunResult } from './features/automation';
export type { PlatformCapability, CapabilityDescriptor } from './features/extended';
export type { GoalSeekParams, GoalSeekResult } from './features/extended/what-if';
export {
  runGoalSeek,
  summarizeGoalSeekResult,
  evaluateCapability,
  type ExtendedSnapshot,
} from './features/extended';
export {
  HistoryPreviewSession,
  type HistoryEntryMeta,
  type HistoryPreviewProjection,
  type RestoreCommandParams,
  type ServerRestoreMutationParams,
} from './features/history';
export { CollaborationSession } from './collaboration';
export { computePivotResult, getPivotFieldCatalog, getPivotRevisionKey } from './features/pivot/engine';
export { cellAddress, columnLabel, parseAddress } from './address';
export type { CanvasSheetSnapshot, CanvasCellSnapshot, PreviewRowSnapshot } from './ui-snapshot';
export type { AppPhase, RibbonTabId, SidebarPanelId, SaveState, PeerCursor, UiSessionIntent } from './types';
