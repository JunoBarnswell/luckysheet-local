import type { CellData, RichTextRun } from '@react-sheets/core-model';
import type { CellInputInterpretationContext } from '@react-sheets/sheet-features';
import type { CellEditDraft, CellEditIntent, CanonicalCellEditTarget, CanonicalKeyGesture, CellEditorListItem, CellEditorRuntimeKind, CellEditorSurfaceDescriptor, CellEditSource } from './contracts';
import { CellEditError } from './error';

export interface CellEditorContext {
  target: CanonicalCellEditTarget;
  source: CellEditSource;
  cell: CellData | null;
  inputContext: CellInputInterpretationContext;
  validationValues?: readonly string[];
  initialText?: string;
}

export type CellEditorEntryDecision =
  | { allowed: true }
  | { allowed: false; code: 'CELL_EDIT_UNSUPPORTED_TARGET'; message: string; recovery: string };

export type CellEditorValidationResult =
  | { valid: true }
  | { valid: false; code: 'CELL_EDIT_COMMIT_REJECTED' | 'CELL_EDIT_INVALID_FORMULA'; message: string; recovery: string };

export type CellEditCommitPayload =
  | { kind: 'raw-text'; text: string }
  | { kind: 'typed-value'; value: CellData['value'] }
  | { kind: 'rich-text'; text: string; runs: RichTextRun[] };

export type CellControlHit = { kind: 'toggle' | 'begin-edit' };

export interface CellEditorBehavior<TDraft extends CellEditDraft = CellEditDraft> {
  readonly kind: CellEditorRuntimeKind;
  readonly surface: CellEditorSurfaceDescriptor;
  readonly valueAutocomplete?: boolean;
  canEnter(context: CellEditorContext): CellEditorEntryDecision;
  createDraft(context: CellEditorContext): TDraft;
  reduce(intent: CellEditIntent, draft: TDraft, context: CellEditorContext): TDraft;
  ownsKey(gesture: CanonicalKeyGesture, context: CellEditorContext): boolean;
  validate(draft: TDraft, context: CellEditorContext): CellEditorValidationResult;
  toCommitPayload(draft: TDraft, context: CellEditorContext): CellEditCommitPayload;
  hitTestControl?(point: { x: number; y: number }, rect: { width: number; height: number }, context: CellEditorContext): CellControlHit | null;
  controlActionForKey?(gesture: CanonicalKeyGesture, context: CellEditorContext): CellControlHit | null;
  listItems?(context: CellEditorContext): readonly CellEditorListItem[];
}

export class CellEditorRegistry {
  private readonly editors = new Map<CellEditorRuntimeKind, CellEditorBehavior>();

  register<TDraft extends CellEditDraft>(behavior: CellEditorBehavior<TDraft>): void {
    if (this.editors.has(behavior.kind)) throw new Error(`Cell editor behavior is already registered: ${behavior.kind}`);
    this.editors.set(behavior.kind, behavior as CellEditorBehavior);
  }

  get(kind: CellEditorRuntimeKind): CellEditorBehavior {
    const behavior = this.editors.get(kind);
    if (!behavior) {
      throw new CellEditError({
        code: 'CELL_EDIT_EDITOR_NOT_FOUND',
        message: `Cell editor behavior is not registered: ${kind}`,
        recovery: 'Register the workbook-owned editor kind before entering edit mode.',
      });
    }
    return behavior;
  }

  resolve(context: CellEditorContext): CellEditorBehavior {
    const configured = context.cell?.editor?.kind === 'custom' ? `custom:${context.cell.editor.editorId}` as const : context.cell?.editor?.kind;
    if (context.initialText?.startsWith('=') && configured !== 'text') return this.get('formula');
    if (!configured && context.validationValues && context.validationValues.length > 0) return this.get('validation-list');
    const kind: CellEditorRuntimeKind = configured
      ?? (context.cell?.formula ? 'formula' : context.cell?.richText ? 'rich-text' : 'text');
    return this.get(kind);
  }

  listKinds(): readonly CellEditorRuntimeKind[] {
    return [...this.editors.keys()];
  }
}
