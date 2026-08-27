import type { CellData, RangeRef, RichTextRun, RichTextRunStyle, SheetId } from '@react-sheets/core-model';
import type { SelectionSnapshot } from '../selection-service';

export type CellEditorStatus = 'ready' | 'enter' | 'edit' | 'point';
export type CellEditSurface = 'grid' | 'formula-bar' | 'formula-panel';
export type CellEditSource = 'direct-typing' | 'double-click' | 'f2' | 'formula-bar' | 'function-insert' | 'cell-control';
export type BuiltinCellEditAdapterKind = 'text' | 'number' | 'datetime' | 'validation-list' | 'combo-box' | 'checkbox' | 'mask' | 'formula' | 'rich-text';
export type CellEditAdapterKind = BuiltinCellEditAdapterKind | `custom:${string}`;
export interface CellEditorSurfaceDescriptor { kind: 'text' | 'rich-text' | 'list' | 'checkbox' | 'custom'; inputMode?: 'text' | 'decimal' | 'numeric'; multiline: boolean; }

export interface CellEditAddress {
  sheetId: SheetId;
  row: number;
  column: number;
}

/**
 * `display` is the PaneMap/render address. `canonical` is the model write
 * address. Keeping both in one immutable target prevents table-sheet and
 * merged-cell projection coordinates from leaking into commit semantics.
 */
export interface CanonicalCellEditTarget {
  display: CellEditAddress;
  canonical: CellEditAddress;
  mergedRange?: RangeRef;
}

export interface CellEditCaret {
  start: number;
  end: number;
}

export interface CellEditComposition {
  active: boolean;
  text: string;
}

export type CellEditDraft =
  | { kind: 'plain'; text: string }
  | { kind: 'rich-text'; text: string; runs: RichTextRun[] };

export interface FormulaReferenceSelection {
  id: string;
  sheetId: SheetId;
  range: RangeRef;
  tokenSpan: CellEditCaret;
  colorIndex: number;
  operation: 'insert' | 'replace' | 'move' | 'resize';
}

export interface FormulaAutocompleteCandidate {
  id: string;
  kind: 'function' | 'defined-name' | 'table' | 'table-column' | 'argument';
  label: string;
  insertionText: string;
  detail?: string;
}
export interface CellEditorListItem { label: string; text: string; }

export type CellEditOverlayState =
  | { kind: 'none' }
  | { kind: 'input-message'; title?: string; message: string }
  | { kind: 'function-hint'; functionName: string; argumentIndex: number }
  | { kind: 'value-autocomplete'; prefix: string; candidate: string }
  | { kind: 'autocomplete'; candidates: readonly FormulaAutocompleteCandidate[]; activeIndex: number; revision: number; replacementSpan: CellEditCaret }
  | { kind: 'editor-list'; items: readonly CellEditorListItem[]; activeIndex: number }
  | { kind: 'validation-confirmation'; title?: string; message: string; alertStyle: 'warning' | 'information' };

export type CellEditValidationState =
  | { kind: 'idle' }
  | { kind: 'blocking-error'; code: string; message: string }
  | { kind: 'confirmation-required'; title?: string; message: string; alertStyle: 'warning' | 'information' };

export interface CellEditSession {
  target: CanonicalCellEditTarget;
  source: CellEditSource;
  status: Exclude<CellEditorStatus, 'ready'>;
  surface: CellEditSurface;
  adapterKind: CellEditAdapterKind;
  editorSurface: CellEditorSurfaceDescriptor;
  draft: CellEditDraft;
  baselineDraft: CellEditDraft;
  caret: CellEditCaret;
  composition: CellEditComposition;
  overtype: boolean;
  referenceSelections: readonly FormulaReferenceSelection[];
  activeReferenceId: string | null;
  originalSelection: SelectionSnapshot;
  originalCell: CellData | null;
  validation: CellEditValidationState;
  overlay: CellEditOverlayState;
  baseCellFingerprint: string;
  dirty: boolean;
  pendingCommit: CellEditCommitRequest | null;
  enterMove: CellEditMoveAfter;
  groupedSheetIds: readonly SheetId[];
}

export interface CellEditSnapshot {
  revision: number;
  status: CellEditorStatus;
  session: CellEditSession | null;
}

export interface CellEditEntryContext {
  target: CanonicalCellEditTarget;
  source: CellEditSource;
  surface: CellEditSurface;
  adapterKind: CellEditAdapterKind;
  editorSurface: CellEditorSurfaceDescriptor;
  initialDraft: CellEditDraft;
  beforeInitialDraft?: CellEditDraft;
  caret: CellEditCaret;
  originalSelection: SelectionSnapshot;
  originalCell: CellData | null;
  baseCellFingerprint: string;
  referenceSelections?: readonly FormulaReferenceSelection[];
  activeReferenceId?: string | null;
  inputMessage?: { title?: string; message: string };
  enterMove: CellEditMoveAfter;
  groupedSheetIds: readonly SheetId[];
}

export interface CellEditBeginRequest {
  type: 'begin.request';
  source: CellEditSource;
  surface?: CellEditSurface;
  initialText?: string;
  caret?: CellEditCaret;
}

export type CellEditMoveAfter = 'down' | 'up' | 'left' | 'right' | 'none';

export interface CellEditCommitRequest {
  target: CanonicalCellEditTarget;
  draft: CellEditDraft;
  adapterKind: CellEditAdapterKind;
  moveAfter: CellEditMoveAfter;
  toSelection: boolean;
  validationConfirmation: boolean;
  baseCellFingerprint: string;
  originalSelection: SelectionSnapshot;
  groupedSheetIds: readonly SheetId[];
}

export interface CanonicalKeyGesture {
  key: string;
  code: string;
  alt: boolean;
  ctrl: boolean;
  meta: boolean;
  shift: boolean;
  repeat: boolean;
  composing: boolean;
}

export type CellEditIntent =
  | { type: 'begin'; entry: CellEditEntryContext }
  | { type: 'text.replace'; text: string; caret: CellEditCaret }
  | { type: 'text.insert'; text: string }
  | { type: 'text.delete-backward' }
  | { type: 'text.delete-forward' }
  | { type: 'caret.set'; caret: CellEditCaret }
  | { type: 'composition.start' }
  | { type: 'composition.update'; text: string }
  | { type: 'composition.end'; text: string; caret: CellEditCaret }
  | { type: 'surface.focus'; surface: CellEditSurface }
  | { type: 'keyboard'; gesture: CanonicalKeyGesture }
  | { type: 'status.toggle' }
  | { type: 'reference.insert'; referenceText: string; selection: FormulaReferenceSelection }
  | { type: 'reference.begin' }
  | { type: 'reference.set'; selections: readonly FormulaReferenceSelection[]; activeReferenceId: string | null }
  | { type: 'reference.toggle-absolute' }
  | { type: 'reference.gesture.begin' }
  | { type: 'reference.gesture.end' }
  | { type: 'reference.gesture.cancel' }
  | { type: 'autocomplete.open'; candidates: readonly FormulaAutocompleteCandidate[]; revision: number; replacementSpan: CellEditCaret }
  | { type: 'autocomplete.close' }
  | { type: 'autocomplete.move'; delta: number }
  | { type: 'autocomplete.accept' }
  | { type: 'function-hint.open'; functionName: string; argumentIndex: number }
  | { type: 'function-hint.close' }
  | { type: 'value-autocomplete.apply'; prefix: string; candidate: string }
  | { type: 'value-autocomplete.close' }
  | { type: 'editor-list.open'; items: readonly CellEditorListItem[] }
  | { type: 'editor-list.move'; delta: number }
  | { type: 'editor-list.accept'; index?: number }
  | { type: 'editor-list.close' }
  | { type: 'rich-text.format'; style: Partial<RichTextRunStyle> }
  | { type: 'draft.undo' }
  | { type: 'draft.redo' }
  | { type: 'commit'; moveAfter?: CellEditMoveAfter; toSelection?: boolean }
  | { type: 'commit.succeeded' }
  | { type: 'commit.failed'; error: CellEditFailure }
  | { type: 'validation.confirm' }
  | { type: 'validation.reject' }
  | { type: 'cancel' };

export type CellEditUserIntent =
  | CellEditBeginRequest
  | Exclude<CellEditIntent, { type: 'begin' | 'commit.succeeded' | 'commit.failed' }>;

export type CellEditEffect =
  | { type: 'commit'; request: CellEditCommitRequest }
  | { type: 'cancel'; originalSelection: SelectionSnapshot }
  | { type: 'focus'; surface: CellEditSurface }
  | { type: 'reference.move'; referenceId: string | null; rowDelta: number; columnDelta: number; extend: boolean; jump: boolean }
  | { type: 'overlay.toggle-request' }
  | { type: 'insert-current'; value: 'date' | 'time' }
  | { type: 'reference.switch-sheet'; direction: 'previous' | 'next' }
  | { type: 'defined-name.request' }
  | { type: 'lifecycle'; event: CellEditLifecycleEvent };

export interface CellEditDispatchResult {
  handled: boolean;
  preventDefault: boolean;
  status: CellEditorStatus;
  effects: readonly CellEditEffect[];
  failure?: CellEditFailure;
}

export interface CellEditController {
  subscribe(listener: () => void): () => void;
  getSnapshot(): CellEditSnapshot;
  dispatch(intent: CellEditUserIntent): CellEditDispatchResult;
}

export type CellEditLifecycleEvent =
  | { type: 'EditStarted'; target: CanonicalCellEditTarget; status: Exclude<CellEditorStatus, 'ready'> }
  | { type: 'EditChanged'; target: CanonicalCellEditTarget; revision: number }
  | { type: 'EditorStatusChanged'; target: CanonicalCellEditTarget; status: CellEditorStatus }
  | { type: 'EditSurfaceChanged'; target: CanonicalCellEditTarget; surface: CellEditSurface }
  | { type: 'EditEnded'; target: CanonicalCellEditTarget; committed: boolean }
  | { type: 'ValidationError'; target: CanonicalCellEditTarget; code: string; message: string };

export interface CellEditFailure {
  code: CellEditErrorCode;
  message: string;
  target?: CellEditAddress;
  recovery: string;
  alertStyle?: 'stop' | 'warning' | 'information';
  title?: string;
}

export type CellEditErrorCode =
  | 'CELL_EDIT_NOT_ACTIVE'
  | 'CELL_EDIT_ALREADY_ACTIVE'
  | 'CELL_EDIT_PERMISSION_DENIED'
  | 'CELL_EDIT_PROTECTED'
  | 'CELL_EDIT_FORMULA_HIDDEN'
  | 'CELL_EDIT_SPILL_CHILD'
  | 'CELL_EDIT_UNSUPPORTED_TARGET'
  | 'CELL_EDIT_REVISION_CONFLICT'
  | 'CELL_EDIT_VALIDATION_BLOCKED'
  | 'CELL_EDIT_CONFIRMATION_REQUIRED'
  | 'CELL_EDIT_INVALID_FORMULA'
  | 'CELL_EDIT_ADAPTER_NOT_FOUND'
  | 'CELL_EDIT_COMMIT_REJECTED';
