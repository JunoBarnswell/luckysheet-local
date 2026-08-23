export {
  SpreadsheetApplication,
  getInitialAppPhase,
  resolveUnitId,
  resolveActorId,
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
export { canExecuteCommand, buildPermissionCapabilities, type PermissionAction } from './permission-bridge';
export { buildCollaborationSnapshot, type CollaborationSnapshot } from './collaboration-bridge';
export { buildRestoreParams, revisionToHistoryMeta } from './history-bridge';
export {
  exchangeExportXlsx,
  exchangeImportXlsx,
  summarizeCompatibilityReport,
} from './xlsx-bridge';
export {
  buildPersistenceMeta,
  buildLocalDraftRecord,
  LocalDraftStore,
  type PersistenceSnapshotMeta,
  type LocalDraftRecord,
} from './persistence-bridge';
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
} from './print-bridge';
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
} from './query-bridge';
export type { QueryDefinition, LoadTarget, QueryStep } from './features/query/query-steps';
export type { ConnectorKind, QueryResult, DataConnector } from './features/query';
export {
  runAutomationScript,
  summarizeScriptResult,
  SAMPLE_AUTOMATION_SCRIPT,
  type AutomationSnapshot,
} from './automation-bridge';
export type { ScriptRunResult } from './features/automation';
export type { PlatformCapability, CapabilityDescriptor } from './features/extended';
export type { GoalSeekParams, GoalSeekResult } from './features/extended/what-if';
export {
  runGoalSeek,
  summarizeGoalSeekResult,
  evaluateCapability,
  type ExtendedSnapshot,
} from './extended-bridge';
export { HistoryPreviewSession, type HistoryEntryMeta } from './features/history';
export { CollaborationSession } from './collaboration';
export { executeUiCommand, isUiCommand, type UiCommandId } from './execute-command';
export { cellAddress, columnLabel, parseAddress } from './address';
export type { CanvasSheetSnapshot, CanvasCellSnapshot, PreviewRowSnapshot } from './ui-snapshot';
export type { AppPhase, RibbonTabId, SidebarPanelId, SaveState, PeerCursor } from './types';
