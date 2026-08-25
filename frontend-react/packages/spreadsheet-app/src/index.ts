export {
  WorkbookSession,
  getInitialSessionPhase,
  resolveUnitId,
  resolveActorId,
  resolveShareToken,
  type WorkbookSessionOptions,
  type DispatchOutcome,
  type ClipboardExecutionOutcome,
  type DispatchErrorCode,
  CommandDispatchError,
  type DefinedNameCommandInput,
  type UiSnapshot,
} from './workbook-session';
export { writeSystemClipboard, type BrowserClipboardPort, type SystemClipboardWriteOutcome } from './clipboard-browser';
export { useWorkbookSession, createWorkbookSessionFactory, type UseWorkbookSessionResult, type WorkbookSessionFactory } from './workbook-session-react';
export {
  registerSpreadsheetFeatures,
  getFeatureRegistry,
  type SpreadsheetFeatureManifest,
} from './feature-registry';
export * from './ui-command-catalog';
export * from './features/formula-audit';
export * from './features/pivot-controls';
export { buildPivotChartData, type PivotChartCategory, type PivotChartData, type PivotChartSeries } from './features/chart/data';
export * from './features/data-source';
export { registerEditingFeatures, buildSelectionSnapshot, type SetSelectionParams } from './features/editing/index';
export { registerDrawingFeature, DrawingRuntime } from './features/drawing/index';
export { EditSession } from './edit-session';
export { SelectionService, createInitialSelection, type SelectionState, type SelectionSnapshot } from './selection-service';
export type { HomeRibbonState, HomeSelectionValue, HomeStyleKey, SheetDialogState, MergeOperation } from './types';
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
export * from './features/workbook-catalog';
export {
  buildPersistenceMeta,
  LocalWorkspaceStore,
  getLocalWorkspaceStore,
  WorkspacePersistence,
  LocalDataBlockStore,
  DataBlockSynchronizer,
  LocalNativePackageStore,
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
  type NativePackageRecord,
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
  type FormulaCalculationModeParams,
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
export { computePivotResult, getPivotFieldCatalog, getPivotRevisionKey, pivotResultMatchesLayoutAndFilter, pivotResultMatchesRevision } from './features/pivot/engine';
export { buildGanttProjection, type GanttProjection, type GanttTaskProjection } from './features/gantt/projection';
export { buildReportProjection, type ReportCellProjection, type ReportProjection } from './features/report/projection';
export { cellAddress, columnLabel, parseAddress } from './address';
export type { CanvasSheetSnapshot, CanvasCellSnapshot, PreviewRowSnapshot } from './ui-snapshot';
export type { AppPhase, FocusState, FocusTarget, InputMode, RibbonTabId, SidebarPanelId, SaveState, PeerCursor, UiSessionIntent } from './types';
