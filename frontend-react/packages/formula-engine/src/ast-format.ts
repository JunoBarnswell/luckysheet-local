import type {
  FormulaAst,
  ParsedCellReference,
} from './ast';
import { columnToLabel } from './address';

function formatReference(reference: ParsedCellReference): string {
  const column = (reference.absoluteColumn ? '$' : '') + columnToLabel(reference.column);
  const row = (reference.absoluteRow ? '$' : '') + String(reference.row + 1);
  return column + row;
}

function formatNode(node: FormulaAst): string {
  switch (node.type) {
    case 'number-literal':
      return String(node.value);
    case 'string-literal':
      return '"' + node.value.replace(/"/g, '""') + '"';
    case 'boolean-literal':
      return node.value ? 'TRUE' : 'FALSE';
    case 'cell-reference': {
      const prefix = node.reference.sheetId !== undefined ? node.reference.sheetId + '!' : '';
      return prefix + formatReference(node.reference);
    }
    case 'range-reference':
      return formatNode(node.start) + ':' + formatNode(node.end);
    case 'name-reference':
      return node.name;
    case 'unary-expression':
      if (node.operator === '%') return formatNode(node.operand) + '%';
      return node.operator + formatNode(node.operand);
    case 'binary-expression':
      return formatNode(node.left) + node.operator + formatNode(node.right);
    case 'function-call':
      return node.name.toUpperCase() + '(' + node.arguments.map(formatNode).join(',') + ')';
  }
}

/** 将 AST 序列化为规范公式文本(含前导 =) */
export function formatFormula(ast: FormulaAst): string {
  return '=' + formatNode(ast);
}
