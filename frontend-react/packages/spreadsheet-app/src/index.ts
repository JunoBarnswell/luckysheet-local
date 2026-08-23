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
export { PermissionService, type ShareRole, type ActorContext, type PermissionCheckInput, type PermissionResult } from './permission-service';
export { executeUiCommand, isUiCommand, type UiCommandId } from './execute-command';
export { cellAddress, columnLabel, parseAddress } from './address';
export type { CanvasSheetSnapshot, CanvasCellSnapshot, PreviewRowSnapshot } from './ui-snapshot';
export type { AppPhase, RibbonTabId, SidebarPanelId, SaveState, PeerCursor } from './types';
