import type { FormulaAst, ParsedCellReference } from './ast';

export interface StructuralShift {
  axis: 'row' | 'column';
  at: number;
  count: number;
  op: 'insert' | 'delete';
}

/** 结构变更后引用重映射:插入下移/右移,删除上移/左移,被删区域引用变 #REF! 语义(此处钳制到边界) */
export function remapReference(ref: ParsedCellReference, shift: StructuralShift): ParsedCellReference {
  if (shift.op === 'insert') {
    const before = shift.axis === 'row' ? ref.row : ref.column;
    if (before >= shift.at) {
      return shift.axis === 'row'
        ? { ...ref, row: ref.row + shift.count }
        : { ...ref, column: ref.column + shift.count };
    }
    return ref;
  }
  const position = shift.axis === 'row' ? ref.row : ref.column;
  const end = shift.at + shift.count - 1;
  if (position > end) {
    return shift.axis === 'row'
      ? { ...ref, row: ref.row - shift.count }
      : { ...ref, column: ref.column - shift.count };
  }
  if (position >= shift.at) {
    // 引用落在删除区:钳制到删除起点
    return shift.axis === 'row' ? { ...ref, row: shift.at } : { ...ref, column: shift.at };
  }
  return ref;
}

export function remapAst(node: FormulaAst, shift: StructuralShift): FormulaAst {
  switch (node.type) {
    case 'number-literal':
    case 'string-literal':
    case 'boolean-literal':
    case 'name-reference':
    case 'table-reference':
      return node;
    case 'cell-reference':
      return { ...node, reference: remapReference(node.reference, shift) };
    case 'range-reference':
      return {
        ...node,
        start: { ...node.start, reference: remapReference(node.start.reference, shift) },
        end: { ...node.end, reference: remapReference(node.end.reference, shift) },
      };
    case 'unary-expression':
      return { ...node, operand: remapAst(node.operand, shift) };
    case 'binary-expression':
      return { ...node, left: remapAst(node.left, shift), right: remapAst(node.right, shift) };
    case 'function-call':
      return { ...node, arguments: node.arguments.map((argument) => remapAst(argument, shift)) };
  }
}
