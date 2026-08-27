import {
  formatFormula,
  parseFormula,
  type CellReferenceNode,
  type FormulaAst,
  type FormulaReferenceNode,
} from '@react-sheets/formula-engine';
import type { CellEditCaret } from './contracts';

interface FormulaReferenceToken {
  start: number;
  end: number;
  node: FormulaReferenceNode;
}

export interface FormulaReferenceRewrite {
  text: string;
  caret: CellEditCaret;
  tokenSpan: CellEditCaret;
}

export interface ParsedFormulaReference {
  tokenSpan: CellEditCaret;
  sheetName?: string;
  startRow?: number;
  endRow?: number;
  startColumn?: number;
  endColumn?: number;
  wholeRow?: boolean;
  wholeColumn?: boolean;
  table?: { name: string; specifier?: 'all' | 'headers' | 'data' | 'totals'; columnName?: string; columnEndName?: string; thisRow: boolean };
  definedName?: string;
}

function containsCaret(span: { start: number; end: number }, caret: CellEditCaret): boolean {
  const start = Math.min(caret.start, caret.end);
  const end = Math.max(caret.start, caret.end);
  return start >= span.start && end <= span.end;
}

function referenceTokenAt(node: FormulaAst, caret: CellEditCaret): FormulaReferenceToken | null {
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

function formulaReferenceAt(text: string, caret: CellEditCaret): FormulaReferenceToken | null {
  if (!text.startsWith('=')) return null;
  try {
    return referenceTokenAt(parseFormula(text), caret);
  } catch {
    // Incomplete formulas remain editable. Reference-only actions fail closed
    // until the caret is inside a parseable reference token.
    return null;
  }
}

function descriptorOf(node: FormulaAst, inheritedSheet?: string): ParsedFormulaReference | null {
  switch (node.type) {
    case 'cell-reference':
      return {
        tokenSpan: { start: node.span.start, end: node.span.end },
        sheetName: node.reference.sheetId ?? inheritedSheet,
        startRow: node.reference.row,
        endRow: node.reference.row,
        startColumn: node.reference.column,
        endColumn: node.reference.column,
      };
    case 'range-reference':
      return {
        tokenSpan: { start: node.span.start, end: node.span.end },
        sheetName: node.start.reference.sheetId ?? node.end.reference.sheetId ?? inheritedSheet,
        startRow: Math.min(node.start.reference.row, node.end.reference.row),
        endRow: Math.max(node.start.reference.row, node.end.reference.row),
        startColumn: Math.min(node.start.reference.column, node.end.reference.column),
        endColumn: Math.max(node.start.reference.column, node.end.reference.column),
      };
    case 'whole-row-reference':
      return { tokenSpan: { start: node.span.start, end: node.span.end }, sheetName: node.sheetId ?? inheritedSheet, startRow: node.startRow, endRow: node.endRow, wholeRow: true };
    case 'whole-column-reference':
      return { tokenSpan: { start: node.span.start, end: node.span.end }, sheetName: node.sheetId ?? inheritedSheet, startColumn: node.startColumn, endColumn: node.endColumn, wholeColumn: true };
    case 'sheet-range-reference':
      return descriptorOf(node.reference, node.qualifier.startSheetId);
    case 'table-reference':
      return { tokenSpan: { start: node.span.start, end: node.span.end }, table: { name: node.tableName, ...(node.specifier ? { specifier: node.specifier } : {}), ...(node.columnName ? { columnName: node.columnName } : {}), ...(node.columnEndName ? { columnEndName: node.columnEndName } : {}), thisRow: node.thisRow } };
    case 'name-reference':
      return { tokenSpan: { start: node.span.start, end: node.span.end }, definedName: node.name };
    default:
      return null;
  }
}

function collectReferenceDescriptors(node: FormulaAst, output: ParsedFormulaReference[]): void {
  const direct = descriptorOf(node);
  if (direct) {
    output.push(direct);
    return;
  }
  switch (node.type) {
    case 'spill-reference':
      collectReferenceDescriptors(node.operand, output);
      return;
    case 'reference-union':
      for (const reference of node.references) collectReferenceDescriptors(reference, output);
      return;
    case 'reference-intersection':
      collectReferenceDescriptors(node.left, output);
      collectReferenceDescriptors(node.right, output);
      return;
    case 'unary-expression':
      collectReferenceDescriptors(node.operand, output);
      return;
    case 'binary-expression':
      collectReferenceDescriptors(node.left, output);
      collectReferenceDescriptors(node.right, output);
      return;
    case 'function-call':
      for (const argument of node.arguments) collectReferenceDescriptors(argument, output);
      return;
    default:
      return;
  }
}

export function parseFormulaReferences(text: string): readonly ParsedFormulaReference[] {
  if (!text.startsWith('=')) return [];
  try {
    const references: ParsedFormulaReference[] = [];
    collectReferenceDescriptors(parseFormula(text), references);
    return references;
  } catch {
    return [];
  }
}

function nextReferenceState(reference: CellReferenceNode): Pick<CellReferenceNode['reference'], 'absoluteColumn' | 'absoluteRow'> {
  const { absoluteColumn, absoluteRow } = reference.reference;
  if (!absoluteColumn && !absoluteRow) return { absoluteColumn: true, absoluteRow: true };
  if (absoluteColumn && absoluteRow) return { absoluteColumn: false, absoluteRow: true };
  if (!absoluteColumn && absoluteRow) return { absoluteColumn: true, absoluteRow: false };
  return { absoluteColumn: false, absoluteRow: false };
}

function toggleReferenceNode(node: FormulaReferenceNode): FormulaReferenceNode | null {
  switch (node.type) {
    case 'cell-reference':
      return { ...node, reference: { ...node.reference, ...nextReferenceState(node) } };
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

export function rewriteAbsoluteReferenceAtCaret(text: string, caret: CellEditCaret): FormulaReferenceRewrite | null {
  const token = formulaReferenceAt(text, caret);
  if (!token) return null;
  const rewritten = toggleReferenceNode(token.node);
  if (!rewritten) return null;
  const replacement = formatFormula(rewritten).slice(1);
  const nextText = `${text.slice(0, token.start)}${replacement}${text.slice(token.end)}`;
  const tokenSpan = { start: token.start, end: token.start + replacement.length };
  return { text: nextText, caret: tokenSpan, tokenSpan };
}

export function replaceReferenceAtCaret(text: string, caret: CellEditCaret, referenceText: string): FormulaReferenceRewrite {
  const token = formulaReferenceAt(text, caret);
  const start = token?.start ?? Math.min(caret.start, caret.end);
  const end = token?.end ?? Math.max(caret.start, caret.end);
  const nextText = `${text.slice(0, start)}${referenceText}${text.slice(end)}`;
  const nextCaret = start + referenceText.length;
  return {
    text: nextText,
    caret: { start: nextCaret, end: nextCaret },
    tokenSpan: { start, end: nextCaret },
  };
}
