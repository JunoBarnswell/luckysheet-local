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
  activateSpreadsheetFeatures,
  advanceSpreadsheetFeatures,
  createSpreadsheetFeatureRuntime,
  compileFeatureSurfaceSchema,
  getFeatureRegistry,
  getExcelParityReport,
  SpreadsheetFeatureRuntime,
  type SpreadsheetFeatureManifest,
  type SpreadsheetFeatureSurface,
  type FeatureLifecyclePhase,
  type FeatureLifecycleContext,
  type FeatureLifecycleHooks,
  type FeatureRuntimeActivation,
  type CompiledFeatureSurfaceEntry,
  type CompiledFeatureSurfaceSchema,
} from './feature-registry';
export * from './ui-command-catalog';
export {
  EXCEL_PARITY_MANIFEST,
  EXCEL_SHORTCUT_MANIFEST,
  assertExcelParityGate,
  buildExcelParityReport,
  createExcelFeatureRegistry,
  validateExcelFeatureRegistry,
  type ExcelFeatureRegistry,
  type ExcelParityClass,
  type ExcelParityItem,
  type ExcelParityReport,
  type ExcelParityScope,
  type ExcelParityStatus,
} from './excel-parity';
export * from './features/formula-audit';
export * from './features/pivot-controls';
export {
  buildPivotChartData,
  chartSourceRevision,
  ChartDataCache,
  chartNumericValue,
  resolveChartData,
  resolveChartDataFromSources,
  resolveStructuredChartBindings,
  resolveSparklineSeries,
  type ChartBindingModel,
  type ChartDataStatus,
  type ChartDataSourceKind,
  type PivotChartCategory,
  type PivotChartData,
  type PivotChartSeries,
  type ResolvedChartData,
  type ResolvedChartSeries,
  type ResolvedSparklineSeries,
  type StructuredChartData,
  type StructuredChartSeries,
  type StructuredChartSheet,
} from './features/chart/data';
export {
  buildChartLayout,
  ChartLayoutCache,
  type ChartLayout,
  type ChartLayoutBar,
  type ChartLayoutPoint,
  type ChartLayoutSeries,
  type ChartLayoutTrendline,
  type ChartAxisLayout,
  type ChartPieSliceLayout,
  type ChartHistogramBinLayout,
  type ChartBoxLayout,
  type ChartWaterfallBarLayout,
} from './features/chart/layout';
export { recommendCharts, type ChartRecommendation } from './features/chart/recommendation';
export { recommendPivotTables, type PivotTableRecommendation } from './features/pivot/recommendation';
export { resolveSparklineData } from './features/sparkline/helpers';
export * from './features/data-source';
export { registerEditingFeatures, buildSelectionSnapshot, planRangeDrag, isRangeBorderPoint, rangeDragMode, type SetSelectionParams, type RangeDragMode, type RangeDragPlan } from './features/editing/index';
export { registerDrawingFeature, DrawingRuntime } from './features/drawing/index';
export * from './cell-edit';
export { SelectionService, createInitialSelection, type SelectionState, type SelectionSnapshot, type SelectionInteractionMode } from './selection-service';
export type { SelectionArea, SelectionKind, SelectionMode } from './selection-service';
export { reduceSelectionInteraction, selectionFromGesture, moveSelection, selectionArea, type SelectionInteractionEvent, type SelectionGesture, type SelectionBounds } from './selection-interaction-machine';
export { applyHeaderSelection, headerContextMenuCatalog, headerRange, headerTargetSelected, selectedHeaderIndices, type DimensionSelectionOptions, type HeaderBounds, type HeaderContextAction, type HeaderContextMenuDescriptor, type HeaderIntent, type HeaderTarget } from './header-interaction-domain';
export { containsRange, expandSelectionRangeForMerges, intersectsRange, nextVisibleCell, resolveSelectionTarget, type ResolvedSelectionTarget, type SelectionTargetSurface } from './selection-target-resolver';
export type { HomeRibbonState, HomeSelectionValue, HomeStyleAggregate, HomeStyleFieldState, HomeStyleKey, SheetDialogState, MergeOperation, FindDialogMode, ChartElementSelection, LocalObjectDialogKind } from './types';
export { resolveContextHit, type ContextHitInput, type ContextTargetKind, type ResolvedContextHit } from './context';
export {
  ShortcutRegistry,
  canonicalKeyGesture,
  createSpreadsheetShortcutRegistry,
  type ShortcutBinding,
  type ShortcutContext,
  type ShortcutEventLike,
  type ShortcutChord,
  type ShortcutSequenceBinding,
  type ShortcutSequenceState,
  type ResolvedShortcutSequence,
  type ShortcutScope,
} from './input/shortcut-registry';
export {
  EXCEL_KEY_TIP_BINDINGS,
  INITIAL_KEY_TIP_STATE,
  keyTipCandidates,
  keyTipTransition,
  type KeyTipBinding,
  type KeyTipState,
  type KeyTipTransition,
} from './input/key-tip-state';
export { canExecuteCommand, buildPermissionCapabilities, type PermissionAction } from './features/permission';
export { buildCollaborationSnapshot, type CollaborationSnapshot } from './collaboration';
export { buildRestoreParams, revisionToHistoryMeta } from './features/history';
export {
  createNativeDocumentTransaction,
  NativeDocumentTransactionRegistry,
  summarizeCompatibilityReport,
  type NativeDocumentExchangeResult,
  type NativeDocumentTransaction,
} from './features/native-document';
export * from './features/workbook-catalog';
export {
  buildPersistenceMeta,
  LocalWorkspaceStore,
  MemoryWorkspaceStore,
  WorkspacePersistence,
  LocalDataBlockStore,
  DataBlockSynchronizer,
  LocalNativeDocumentStore,
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
  type NativeDocumentRecord,
  type AssetStore,
  type AssetPutInput,
} from './features/persistence';
export {
  buildPrintSnapshot,
  buildPrintProjection,
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
  type PrintCellReader,
  type PrintProjectionOptions,
  type PrintProjection,
  type PrintProjectionCell,
  type PrintProjectionDrawing,
  type PrintChartProjection,
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
export type { ConnectorKind, ConnectorManifest, ConnectorInputField, QueryResult, DataConnector } from './features/query';
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
