import type { FormulaAst } from './ast';
import { getFunctionDescriptor } from './functions';

export function collectNameReferences(ast: FormulaAst): string[] {
  const names: string[] = [];
  const seen = new Set<string>();
  visit(ast);
  return names;

  function visit(node: FormulaAst): void {
    switch (node.type) {
      case 'name-reference': {
        const upper = node.name.toUpperCase();
        if (!seen.has(upper)) {
          seen.add(upper);
          names.push(upper);
        }
        return;
      }
      case 'unary-expression':
        visit(node.operand);
        return;
      case 'spill-reference':
        visit(node.operand);
        return;
      case 'reference-union':
        for (const reference of node.references) visit(reference);
        return;
      case 'reference-intersection':
        visit(node.left);
        visit(node.right);
        return;
      case 'sheet-range-reference':
      case 'external-reference':
        visit(node.reference);
        return;
      case 'binary-expression':
        visit(node.left);
        visit(node.right);
        return;
      case 'function-call':
        for (const argument of node.arguments) visit(argument);
        return;
      default:
        return;
    }
  }
}

export function formulaUsesVolatile(ast: FormulaAst): boolean {
  let volatile = false;
  visit(ast);
  return volatile;

  function visit(node: FormulaAst): void {
    if (volatile) return;
    switch (node.type) {
      case 'function-call': {
        const descriptor = getFunctionDescriptor(node.name);
        if (descriptor?.volatile) volatile = true;
        for (const argument of node.arguments) visit(argument);
        return;
      }
      case 'unary-expression':
        visit(node.operand);
        return;
      case 'spill-reference':
        visit(node.operand);
        return;
      case 'reference-union':
        for (const reference of node.references) visit(reference);
        return;
      case 'reference-intersection':
        visit(node.left);
        visit(node.right);
        return;
      case 'sheet-range-reference':
      case 'external-reference':
        visit(node.reference);
        return;
      case 'binary-expression':
        visit(node.left);
        visit(node.right);
        return;
      default:
        return;
    }
  }
}
