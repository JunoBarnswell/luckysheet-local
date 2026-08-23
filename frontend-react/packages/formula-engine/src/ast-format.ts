import type {
  FormulaAst,
  ParsedCellReference,
  TableReferenceSpecifier,
} from './ast';
import { columnToLabel } from './address';

function formatTableSpecifier(specifier: TableReferenceSpecifier): string {
  switch (specifier) {
    case 'all': return '#All';
    case 'headers': return '#Headers';
    case 'data': return '#Data';
    case 'totals': return '#Totals';
  }
}

function formatSheetId(sheetId: string): string {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheetId)
    ? sheetId
    : `'${sheetId.replaceAll("'", "''")}'`;
}

function formatReference(reference: ParsedCellReference): string {
  const column = (reference.absoluteColumn ? '$' : '') + columnToLabel(reference.column);
  const row = (reference.absoluteRow ? '$' : '') + String(reference.row + 1);
  return column + row;
}

type BinaryNode = Extract<FormulaAst, { type: 'binary-expression' }>;

function precedence(node: FormulaAst): number {
  switch (node.type) {
    case 'binary-expression':
      switch (node.operator) {
        case '=':
        case '<>':
        case '<':
        case '<=':
        case '>':
        case '>=':
          return 1;
        case '&': return 2;
        case '+':
        case '-': return 3;
        case '*':
        case '/': return 4;
        case '^': return 5;
      }
    case 'unary-expression': return 6;
    default: return 7;
  }
}

function formatBinaryChild(node: FormulaAst, parent: BinaryNode, side: 'left' | 'right'): string {
  const child = formatNode(node, precedence(parent));
  // The parser retains explicit grouping. The precedence pass handles ASTs
  // created by callers as well, including right operands of non-associative
  // operators where equal precedence changes the result.
  if (
    node.type === 'binary-expression'
    && precedence(node) === precedence(parent)
    && side === 'right'
    && (parent.operator === '-' || parent.operator === '/' || parent.operator === '^' || parent.operator === '&')
  ) {
    return `(${child})`;
  }
  return child;
}

function formatNode(node: FormulaAst, parentPrecedence = 0): string {
  let content: string;
  switch (node.type) {
    case 'number-literal':
      content = String(node.value);
      break;
    case 'string-literal':
      content = '"' + node.value.replace(/"/g, '""') + '"';
      break;
    case 'boolean-literal':
      content = node.value ? 'TRUE' : 'FALSE';
      break;
    case 'cell-reference': {
      const prefix = node.reference.sheetId !== undefined ? formatSheetId(node.reference.sheetId) + '!' : '';
      content = prefix + formatReference(node.reference);
      break;
    }
    case 'invalid-reference':
      content = node.code;
      break;
    case 'range-reference':
      content = formatNode(node.start) + ':' + formatNode(node.end);
      break;
    case 'name-reference':
      content = node.name;
      break;
    case 'table-reference':
      if (node.specifier && node.columnName) {
        content = `${node.tableName}[[${formatTableSpecifier(node.specifier)}],[${node.columnName}]]`;
      } else if (node.specifier) {
        content = `${node.tableName}[${formatTableSpecifier(node.specifier)}]`;
      } else {
        content = `${node.tableName}[${node.thisRow ? '@' : ''}${node.columnName ?? ''}]`;
      }
      break;
    case 'unary-expression':
      content = node.operator === '%' ? `${formatNode(node.operand, precedence(node))}%` : node.operator + formatNode(node.operand, precedence(node));
      break;
    case 'binary-expression':
      content = `${formatBinaryChild(node.left, node, 'left')}${node.operator}${formatBinaryChild(node.right, node, 'right')}`;
      break;
    case 'function-call':
      content = node.name.toUpperCase() + '(' + node.arguments.map((argument) => formatNode(argument)).join(',') + ')';
      break;
  }

  const needsParentheses = node.parenthesized === true || precedence(node) < parentPrecedence;
  return needsParentheses ? `(${content})` : content;
}

/** Serialize an AST to canonical formula text (including the leading `=`). */
export function formatFormula(ast: FormulaAst): string {
  return '=' + formatNode(ast);
}
