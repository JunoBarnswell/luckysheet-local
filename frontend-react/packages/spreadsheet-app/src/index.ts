export {
  WorkbookSession,
  getInitialSessionPhase,
  resolveUnitId,
  resolveActorId,
  resolveShareToken,
  type WorkbookSessionOptions,
  type PivotCreateOutcome,
  type PivotUpdateOutcome,
  type PivotCreateTaskState,
  type PivotTaskState,
  type DispatchOutcome,
  type ClipboardExecutionOutcome,
  type DispatchErrorCode,
  type FindDialogParams,
  CommandDispatchError,
  type DefinedNameCommandInput,
  type UiSnapshot,
  type SheetTabSnapshot,
} from './workbook-session';
export { InsertCoordinator, type InsertIdentity, type InsertResult, type InsertRequest, type DrawingInsertRequest, type InsertMutationRequest } from './insert-coordinator';
export { writeSystemClipboard, type BrowserClipboardPort, type SystemClipboardWriteOutcome } from './clipboard-browser';
export { useCellEdit, useWorkbookSession, createWorkbookSessionFactory, type UseWorkbookSessionResult, type WorkbookSessionFactory } from './workbook-session-react';
export {
  registerSpreadsheetFeatures,
  getFeatureRegistry,
  type SpreadsheetFeatureManifest,
} from './feature-registry';
export * from './ui-command-catalog';
export * from './features/formula-audit';
export * from './features/pivot-controls';
export { buildPivotChartData, resolveStructuredChartBindings, type PivotChartCategory, type PivotChartData, type PivotChartSeries, type StructuredChartData, type StructuredChartSeries, type StructuredChartSheet } from './features/chart/data';
export { recommendCharts, type ChartRecommendation } from './features/chart/recommendation';
export { recommendPivotTables, type PivotTableRecommendation } from './features/pivot/recommendation';
export * from './features/data-source';
export { registerEditingFeatures, buildSelectionSnapshot, type SetSelectionParams } from './features/editing/index';
export { registerDrawingFeature, DrawingRuntime } from './features/drawing/index';
export * from './cell-edit';
export { SelectionService, createInitialSelection, type SelectionState, type SelectionSnapshot } from './selection-service';
export type { SelectionArea, SelectionKind, SelectionMode } from './selection-service';
export { reduceSelectionInteraction, selectionFromGesture, moveSelection, selectionArea, type SelectionInteractionEvent, type SelectionGesture, type SelectionBounds } from './selection-interaction-machine';
export { containsRange, expandSelectionRangeForMerges, intersectsRange, nextVisibleCell, resolveSelectionTarget, type ResolvedSelectionTarget, type SelectionTargetSurface } from './selection-target-resolver';
export type { HomeRibbonState, HomeSelectionValue, HomeStyleAggregate, HomeStyleFieldState, HomeStyleKey, SheetDialogState, MergeOperation, FindDialogMode } from './types';
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
  MemoryWorkspaceStore,
  WorkspacePersistence,
  LocalDataBlockStore,
  DataBlockSynchronizer,
  LocalNativePackageStore,
  LocalAssetStore,
  RemoteAssetStore,
  migrateLegacyImageAssets,
  buildWorkspaceRecord,
  verifyWorkspaceRecord,
  verifyPendingOperationJournal,
  WorkspaceStorageError,
  isWorkspaceStorageError,
  WorkspaceMemoryCoordinator,
  type WorkspacePersistenceState,
  type WorkspacePersistenceMode,
  type PersistenceSnapshotMeta,
  type WorkspaceRecord,
  type WorkspaceRecordInput,
  type PendingOperationJournal,
  type LocalWorkspaceSummary,
  type WorkspacePersistenceOptions,
  type DataBlockRecord,
  type DataBlockSyncOptions,
  type NativePackageRecord,
  type AssetStore,
  type AssetPutInput,
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
  prepareQueryLoadPayload,
  type QueryResultSnapshot,
  type QuerySessionEntry,
  type QueryLoadCommandPayload,
} from './features/query';
export type { QueryDefinition, LoadTarget, QueryStep } from './features/query/query-steps';
export type { ConnectorKind, QueryResult, DataConnector } from './features/query';
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
export { buildPivotGroupedFilterMembers, computePivotResult, evaluatePivotTask, findPivotProjectionCellAt, getPivotFieldCatalog, getPivotRevisionKey, pivotResultMatchesLayoutAndFilter, pivotResultMatchesRevision, preparePivotTaskInput, type PivotGroupedFilterMember, type PivotTaskControl, type PivotTaskEvaluationInput } from './features/pivot/engine';
export { BrowserPivotTaskPort, InlinePivotTaskPort, createBrowserPivotTaskPort, type PivotTaskPort } from './features/pivot/task-port';
export { createPivotCalculateRequest, createPivotSourceRegisterRequest, createPivotSourceReleaseRequest, type PivotTaskError, type PivotTaskErrorCode, type PivotTaskResult } from './features/pivot/task-protocol';
export { createPivotSourceIndex, estimatePivotSourceIndexBytes, type PivotSourceIndex } from './features/pivot/source-index';
export { buildGanttProjection, type GanttProjection, type GanttTaskProjection } from './features/gantt/projection';
export { buildReportProjection, type ReportCellProjection, type ReportProjection } from './features/report/projection';
export { cellAddress, columnLabel, parseAddress } from './address';
export type { CanvasSheetSnapshot, CanvasCellSnapshot } from './ui-snapshot';
export type { AppPhase, FocusState, FocusTarget, InputMode, RibbonTabId, SidebarPanelId, SaveState, PeerCursor, UiSessionIntent } from './types';
