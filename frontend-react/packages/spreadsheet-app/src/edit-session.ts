import type { CellData, Row, Column, SheetId } from '@react-sheets/core-model';
import { formatFormula, parseFormula, type CellReferenceNode, type FormulaAst, type FormulaReferenceNode, type RangeReferenceNode } from '@react-sheets/formula-engine';
import type { SelectionSnapshot } from './selection-service';

export type EditSource = 'directTyping' | 'f2' | 'formulaBar' | 'functionInsert';

export interface EditCaret { start: number; end: number; }
export interface EditComposition { active: boolean; text: string; }

export interface EditSessionState {
  sheetId: SheetId;
  row: Row;
  column: Column;
  originalValue: CellData | null;
  originalFormula?: string;
  originalSelection: SelectionSnapshot;
  currentDraft: string;
  caret: EditCaret;
  composition: EditComposition;
  source: EditSource;
  referenceMode: boolean;
  baseRevision?: number;
  isDirty: boolean;
}

function clampCaret(value: number, length: number): number {
  return Math.max(0, Math.min(length, Number.isSafeInteger(value) ? value : length));
}

interface FormulaReferenceToken {
  start: number;
  end: number;
  node: FormulaReferenceNode;
}

function containsCaret(span: { start: number; end: number }, caret: EditCaret): boolean {
  const selectionStart = Math.min(caret.start, caret.end);
  const selectionEnd = Math.max(caret.start, caret.end);
  return selectionStart >= span.start && selectionEnd <= span.end;
}

function referenceTokenAt(node: FormulaAst, caret: EditCaret): FormulaReferenceToken | null {
  if (!containsCaret(node.span, caret)) return null;
  switch (node.type) {
    case 'cell-reference':
    case 'whole-column-reference':
    case 'whole-row-reference':
    case 'table-reference':
    case 'invalid-reference':
    case 'range-reference':
    case 'sheet-range-reference':
    case 'external-reference':
      return { start: node.span.start, end: node.span.end, node };
    case 'spill-reference':
      return referenceTokenAt(node.operand, caret);
    case 'reference-union':
      for (const reference of node.references) {
        const token = referenceTokenAt(reference, caret);
        if (token) return token;
      }
      return null;
    case 'reference-intersection':
      return referenceTokenAt(node.left, caret) ?? referenceTokenAt(node.right, caret);
    case 'unary-expression':
      return referenceTokenAt(node.operand, caret);
    case 'binary-expression':
      return referenceTokenAt(node.left, caret) ?? referenceTokenAt(node.right, caret);
    case 'function-call':
      for (const argument of node.arguments) {
        const token = referenceTokenAt(argument, caret);
        if (token) return token;
      }
      return null;
    default:
      return null;
  }
}

function formulaReferenceAt(text: string, caret: EditCaret): FormulaReferenceToken | null {
  try {
    return referenceTokenAt(parseFormula(text), caret);
  } catch {
    // An incomplete formula is still editable. The caller will use the
    // canonical caret as the insertion point, while F4 remains fail-closed
    // until the formula has a parseable reference token.
    return null;
  }
}

function nextReferenceState(reference: CellReferenceNode): { absoluteColumn: boolean; absoluteRow: boolean } {
  const state = (reference.reference.absoluteColumn ? 1 : 0) + (reference.reference.absoluteRow ? 2 : 0);
  const nextState = state === 0 ? 1 : state === 1 ? 2 : state === 2 ? 3 : 0;
  return {
    absoluteColumn: nextState === 1 || nextState === 3,
    absoluteRow: nextState === 1 || nextState === 2,
  };
}

function toggleReferenceNode(node: FormulaReferenceNode): FormulaReferenceNode | null {
  switch (node.type) {
    case 'cell-reference': {
      const next = nextReferenceState(node);
      return { ...node, reference: { ...node.reference, ...next } };
    }
    case 'range-reference':
      return {
        ...node,
        start: toggleReferenceNode(node.start) as CellReferenceNode,
        end: toggleReferenceNode(node.end) as CellReferenceNode,
      };
    case 'sheet-range-reference': {
      const reference = toggleReferenceNode(node.reference);
      return reference ? { ...node, reference: reference as typeof node.reference } : null;
    }
    case 'external-reference': {
      const reference = toggleReferenceNode(node.reference);
      return reference ? { ...node, reference: reference as typeof node.reference } : null;
    }
    default:
      return null;
  }
}

export class EditSession {
  private session: EditSessionState | null = null;

  get active(): EditSessionState | null { return this.session; }
  get editingCell(): { row: number; column: number } | null {
    return this.session ? { row: this.session.row, column: this.session.column } : null;
  }
  get composing(): boolean { return this.session?.composition.active ?? false; }

  begin(params: {
    sheetId: SheetId;
    row: Row;
    column: Column;
    cell: CellData | undefined;
    selection: SelectionSnapshot;
    initialText?: string;
    source?: EditSource;
    baseRevision?: number;
    caret?: EditCaret;
  }): void {
    const raw = params.initialText ?? params.cell?.formula ?? (params.cell?.value == null ? '' : String(params.cell.value));
    const caret = params.caret ?? { start: raw.length, end: raw.length };
    this.session = {
      sheetId: params.sheetId,
      row: params.row,
      column: params.column,
      originalValue: params.cell ? structuredClone(params.cell) : null,
      originalFormula: params.cell?.formula,
      originalSelection: structuredClone(params.selection),
      currentDraft: raw,
      caret: { start: clampCaret(caret.start, raw.length), end: clampCaret(caret.end, raw.length) },
      composition: { active: false, text: '' },
      source: params.source ?? (params.initialText === undefined ? 'f2' : 'directTyping'),
      referenceMode: false,
      ...(params.baseRevision === undefined ? {} : { baseRevision: params.baseRevision }),
      isDirty: params.initialText !== undefined,
    };
  }

  setDraft(value: string, caret?: EditCaret): void {
    if (!this.session) return;
    const nextCaret = caret ?? { start: value.length, end: value.length };
    this.session.currentDraft = value;
    this.session.caret = { start: clampCaret(nextCaret.start, value.length), end: clampCaret(nextCaret.end, value.length) };
    this.session.isDirty = value !== (this.session.originalFormula ?? String(this.session.originalValue?.value ?? ''));
  }

  setCaret(caret: EditCaret): void {
    if (!this.session) return;
    this.session.caret = { start: clampCaret(caret.start, this.session.currentDraft.length), end: clampCaret(caret.end, this.session.currentDraft.length) };
  }

  compositionStart(): void { if (this.session) this.session.composition = { active: true, text: '' }; }
  compositionUpdate(text: string): void { if (this.session) this.session.composition = { active: true, text }; }
  compositionEnd(): void { if (this.session) this.session.composition = { active: false, text: '' }; }
  enterReferenceMode(): void { if (this.session) this.session.referenceMode = true; }

  insertRef(refText: string): void {
    if (!this.session || !refText) return;
    const draft = this.session.currentDraft;
    const token = formulaReferenceAt(draft, this.session.caret);
    const start = token?.start ?? Math.min(this.session.caret.start, this.session.caret.end);
    const end = token?.end ?? Math.max(this.session.caret.start, this.session.caret.end);
    const next = `${draft.slice(0, start)}${refText}${draft.slice(end)}`;
    this.setDraft(next, { start: start + refText.length, end: start + refText.length });
    this.session.referenceMode = true;
  }

  toggleAbsoluteReference(): void {
    if (!this.session) return;
    const token = formulaReferenceAt(this.session.currentDraft, this.session.caret);
    if (!token) return;
    const rewritten = toggleReferenceNode(token.node);
    if (!rewritten) return;
    const next = formatFormula(rewritten).slice(1);
    this.setDraft(`${this.session.currentDraft.slice(0, token.start)}${next}${this.session.currentDraft.slice(token.end)}`, { start: token.start, end: token.start + next.length });
    this.session.referenceMode = true;
  }

  cancel(): string {
    if (!this.session) return '';
    const restore = this.session.originalFormula ?? (this.session.originalValue?.value == null ? '' : String(this.session.originalValue.value));
    this.session = null;
    return restore;
  }

  apply(): { row: number; column: number; draft: string } | null {
    if (!this.session) return null;
    const result = { row: this.session.row, column: this.session.column, draft: this.session.currentDraft };
    this.session = null;
    return result;
  }
}
