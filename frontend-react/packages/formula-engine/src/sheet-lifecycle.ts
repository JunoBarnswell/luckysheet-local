import type { FormulaAst } from './ast';
import type { FormulaSheetIdentity } from './sheet-reference';
import { sameFormulaSheetName } from './sheet-reference';
import { parseFormula } from './parser';
import { formatFormula } from './ast-format';

export interface SheetLifecycleChange {
  readonly before: readonly FormulaSheetIdentity[];
  readonly after: readonly FormulaSheetIdentity[];
  readonly sheetId: string;
  readonly kind: 'delete' | 'move';
}

/** Endpoint identity follows the surviving ordered interval; foreign qualifiers are opaque. */
export function rewriteSheetLifecycleFormula(formula: string, change: SheetLifecycleChange): string {
  const hasEquals = formula.trim().startsWith('=');
  const ast = parseFormula(hasEquals ? formula : `=${formula}`);
  const target = change.before.find(sheet => sheet.id === change.sheetId);
  if (!target) throw new Error('SHEET_LIFECYCLE_IDENTITY: missing source sheet');
  const matches = (name?: string): boolean => name !== undefined && (name === target.id || sameFormulaSheetName(name, target.name));
  const invalid = (node: FormulaAst): FormulaAst => ({ type: 'invalid-reference', code: '#REF!', span: node.span });
  const visit = (node: FormulaAst): FormulaAst => {
    switch (node.type) {
      case 'external-reference': return node;
      case 'sheet-range-reference': {
        const indexOf = (name: string): number => change.before.findIndex(s => s.id === name || sameFormulaSheetName(s.name, name));
        const start = indexOf(node.qualifier.startSheetId), end = indexOf(node.qualifier.endSheetId);
        if (start < 0 || end < 0) return node;
        const low = Math.min(start, end), high = Math.max(start, end);
        const oldInterval = change.before.slice(low, high + 1);
        let survivors = oldInterval.filter(s => change.after.some(next => next.id === s.id));
        if (change.kind === 'move' && (matches(node.qualifier.startSheetId) || matches(node.qualifier.endSheetId))) {
          const moved = change.after.findIndex(s => s.id === target.id);
          const other = change.before[start === indexOf(target.id) ? end : start]!;
          const otherIndex = change.after.findIndex(s => s.id === other.id);
          const wasFirst = change.before[low]!.id === target.id;
          if ((wasFirst && moved > otherIndex) || (!wasFirst && moved < otherIndex)) survivors = survivors.filter(s => s.id !== target.id);
        }
        if (survivors.length === 0) return invalid(node);
        const first = survivors[0]!, last = survivors[survivors.length - 1]!;
        const qualifier = { startSheetId: (start <= end ? first : last).name, endSheetId: (start <= end ? last : first).name };
        if (sameFormulaSheetName(qualifier.startSheetId, node.qualifier.startSheetId) && sameFormulaSheetName(qualifier.endSheetId, node.qualifier.endSheetId)) return node;
        return { ...node, qualifier };
      }
      case 'cell-reference': return change.kind === 'delete' && matches(node.reference.sheetId) ? invalid(node) : node;
      case 'range-reference': return change.kind === 'delete' && (matches(node.start.reference.sheetId) || matches(node.end.reference.sheetId)) ? invalid(node) : node;
      case 'whole-row-reference': case 'whole-column-reference': return change.kind === 'delete' && matches(node.sheetId) ? invalid(node) : node;
      case 'reference-union': return { ...node, references: node.references.map(n => visit(n) as typeof n) };
      case 'reference-intersection': return { ...node, left: visit(node.left) as typeof node.left, right: visit(node.right) as typeof node.right };
      case 'unary-expression': case 'spill-reference': return { ...node, operand: visit(node.operand) };
      case 'binary-expression': return { ...node, left: visit(node.left), right: visit(node.right) };
      case 'function-call': return { ...node, ...(node.callee ? { callee: visit(node.callee) } : {}), arguments: node.arguments.map(visit) };
      default: return node;
    }
  };
  const next = visit(ast);
  if (JSON.stringify(next) === JSON.stringify(ast)) return formula;
  const formatted = formatFormula(next);
  return hasEquals ? formatted : formatted.replace(/^=/, '');
}
