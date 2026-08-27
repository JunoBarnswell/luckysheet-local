import type { CellData, CellEditorConfig, RichTextRun } from '@react-sheets/core-model';
import type { CellInputInterpretationContext } from '@react-sheets/sheet-features';
import type { CellEditAdapterKind, CellEditDraft, CellEditIntent, CanonicalCellEditTarget, CanonicalKeyGesture, CellEditorSurfaceDescriptor } from './contracts';
import { CellEditError } from './error';

export interface CellEditorContext {
  target: CanonicalCellEditTarget;
  cell: CellData | null;
  config?: CellEditorConfig;
  inputContext: CellInputInterpretationContext;
  validationValues?: readonly string[];
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

export interface CellEditorAdapter<TDraft extends CellEditDraft = CellEditDraft> {
  readonly kind: CellEditAdapterKind;
  readonly surface: CellEditorSurfaceDescriptor;
  canEnter(context: CellEditorContext): CellEditorEntryDecision;
  createDraft(context: CellEditorContext): TDraft;
  reduce(intent: CellEditIntent, draft: TDraft, context: CellEditorContext): TDraft;
  ownsKey(gesture: CanonicalKeyGesture, context: CellEditorContext): boolean;
  validate(draft: TDraft, context: CellEditorContext): CellEditorValidationResult;
  toCommitPayload(draft: TDraft, context: CellEditorContext): CellEditCommitPayload;
}

export class CellEditorAdapterRegistry {
  private readonly adapters = new Map<CellEditAdapterKind, CellEditorAdapter>();

  register<TDraft extends CellEditDraft>(adapter: CellEditorAdapter<TDraft>): void {
    if (this.adapters.has(adapter.kind)) throw new Error(`Cell editor adapter is already registered: ${adapter.kind}`);
    this.adapters.set(adapter.kind, adapter as CellEditorAdapter);
  }

  get(kind: CellEditAdapterKind): CellEditorAdapter {
    const adapter = this.adapters.get(kind);
    if (!adapter) {
      throw new CellEditError({
        code: 'CELL_EDIT_ADAPTER_NOT_FOUND',
        message: `Cell editor adapter is not registered: ${kind}`,
        recovery: 'Register the workbook-owned editor kind before entering edit mode.',
      });
    }
    return adapter;
  }

  resolve(context: CellEditorContext): CellEditorAdapter {
    const configured = context.config?.kind === 'custom' ? `custom:${context.config.adapterId}` as const : context.config?.kind;
    const kind: CellEditAdapterKind = configured
      ?? (context.cell?.formula ? 'formula' : context.cell?.richText ? 'rich-text' : 'text');
    return this.get(kind);
  }

  listKinds(): readonly CellEditAdapterKind[] {
    return [...this.adapters.keys()];
  }
}
