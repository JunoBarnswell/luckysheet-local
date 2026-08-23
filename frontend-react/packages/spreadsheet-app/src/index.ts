export {
  WorkbookSession,
  getInitialSessionPhase,
  resolveUnitId,
  resolveActorId,
  resolveShareToken,
  type WorkbookSessionOptions,
  type UiSnapshot,
} from './workbook-session';
export { useWorkbookSession, type UseWorkbookSessionResult } from './workbook-session-react';
export {
  registerSpreadsheetFeatures,
  getFeatureRegistry,
  type SpreadsheetFeatureManifest,
} from './feature-registry';
export * from './ui-command-catalog';
export * from './features/formula-audit';
export { registerEditingFeatures, buildSelectionSnapshot, type SetSelectionParams } from './features/editing/index';
export { registerDrawingFeature, DrawingRuntime } from './features/drawing/index';
export { EditSession } from './edit-session';
export { SelectionService, createInitialSelection, type SelectionState, type SelectionSnapshot } from './selection-service';
export { resolveContextHit, type ContextHitInput, type ContextTargetKind, type ResolvedContextHit } from './context';
export {
  ShortcutRegistry,
  createSpreadsheetShortcutRegistry,
  type ShortcutBinding,
  type ShortcutContext,
  type ShortcutEventLike,
  type ShortcutScope,
} from './input/shortcut-registry';
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
  LocalDataBlockStore,
  DataBlockSynchronizer,
  LocalXlsxArtifactStore,
  buildWorkspaceRecord,
  verifyWorkspaceRecord,
  verifyPendingOperationJournal,
  type PersistenceSnapshotMeta,
  type WorkspaceRecord,
  type WorkspaceRecordInput,
  type PendingOperationJournal,
  type LocalWorkspaceSummary,
  type IndexedDbWorkspaceStoreOptions,
  type DataBlockRecord,
  type DataBlockSyncOptions,
  type XlsxArtifactRecord,
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
  type PrintPageSetupCommandParams,
  type PrintTitlesSetCommandParams,
  type PrintScaleSetCommandParams,
  type PrintToggleCommandParams,
  type PrintLayout,
  type PrintDocument,
  type PageSetup,
  type PrintArea,
  type PrintPageBreak,
  type PrintTitleSpan,
} from './features/print';
export {
  FormulaAuditController,
  evaluateFormulaStep,
  getFormulaDependents,
  getFormulaPrecedents,
  projectFormulaCells,
  removeFormulaAuditArrows,
  scanFormulaErrors,
  registerFormulaAuditCommands,
  type FormulaAuditArrow,
  type FormulaAuditControllerOptions,
  type FormulaAuditDirection,
  type FormulaAuditEmptyParams,
  type FormulaAuditError,
  type FormulaAuditErrorScanParams,
  type FormulaAuditEvaluationProjection,
  type FormulaAuditEvaluationStep,
  type FormulaAuditFormulaProjection,
  type FormulaAuditProjection,
  type FormulaAuditAddressParams,
  type FormulaAuditShowFormulasParams,
  type FormulaErrorScanOptions,
} from './features/formula-audit';
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
export type { GoalSeekParams, GoalSeekResult } from './features/extended/what-if';
export {
  runGoalSeek,
  summarizeGoalSeekResult,
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
