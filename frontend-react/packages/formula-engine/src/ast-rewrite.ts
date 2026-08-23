import type { FormulaAst, ParsedCellReference } from './ast';

export interface StructuralShift {
  axis: 'row' | 'column';
  at: number;
  count: number;
  op: 'insert' | 'delete';
}

/**
 * A structural reference transform can invalidate a reference.  `undefined`
 * is deliberately distinct from a surviving coordinate: callers must render
 * it as a real `#REF!` AST node rather than clamping it to an arbitrary cell.
 */
export type FormulaReferenceMapper = (reference: ParsedCellReference) => ParsedCellReference | undefined;

/**
 * Apply a reference mapper to every cell reference in an AST.
 *
 * This is deliberately the only tree-walking primitive used by structural
 * transforms, paste/autofill and sheet rename. Callers can restrict the
 * mapper to a target worksheet or a selected region without textual search.
 */
export function mapAstReferences(node: FormulaAst, mapper: FormulaReferenceMapper): FormulaAst {
  switch (node.type) {
    case 'number-literal':
    case 'string-literal':
    case 'boolean-literal':
    case 'name-reference':
    case 'table-reference':
    case 'invalid-reference':
      return node;
    case 'cell-reference': {
      const mapped = mapper(node.reference);
      return mapped === undefined
        ? { type: 'invalid-reference', code: '#REF!', span: node.span, parenthesized: node.parenthesized }
        : { ...node, reference: mapped };
    }
    case 'range-reference': {
      const mappedStart = mapper(node.start.reference);
      const mappedEnd = mapper(node.end.reference);
      // A range endpoint which no longer exists invalidates the range.  This
      // avoids the old clamp behaviour (`A2:A2`) which silently changes the
      // meaning of formulas after row/column deletion.
      if (mappedStart === undefined || mappedEnd === undefined) {
        return { type: 'invalid-reference', code: '#REF!', span: node.span, parenthesized: node.parenthesized };
      }
      return {
        ...node,
        start: { ...node.start, reference: mappedStart },
        end: { ...node.end, reference: mappedEnd },
      };
    }
    case 'unary-expression':
      return { ...node, operand: mapAstReferences(node.operand, mapper) };
    case 'binary-expression':
      return {
        ...node,
        left: mapAstReferences(node.left, mapper),
        right: mapAstReferences(node.right, mapper),
      };
    case 'function-call':
      return { ...node, arguments: node.arguments.map((argument) => mapAstReferences(argument, mapper)) };
  }
}

/**
 * Remap references after inserting/deleting a whole row or column.
 * Absolute markers affect copy/fill operations, not workbook structure
 * changes, so both absolute and relative references move with the structure.
 */
export function remapReference(ref: ParsedCellReference, shift: StructuralShift): ParsedCellReference | undefined {
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
    // A reference into the deleted region is invalid.  Returning undefined is
    // handled by mapAstReferences and produces a first-class #REF! node.
    return undefined;
  }
  return ref;
}

export function remapAst(
  node: FormulaAst,
  shift: StructuralShift,
  shouldRemap: (reference: ParsedCellReference) => boolean = () => true,
): FormulaAst {
  return mapAstReferences(node, (reference) =>
    shouldRemap(reference) ? remapReference(reference, shift) : reference,
  );
}

/** Shift relative references when a formula is copied to another cell. */
export function offsetReference(
  reference: ParsedCellReference,
  rowOffset: number,
  columnOffset: number,
): ParsedCellReference {
  return {
    ...reference,
    row: reference.absoluteRow ? reference.row : Math.max(0, reference.row + rowOffset),
    column: reference.absoluteColumn ? reference.column : Math.max(0, reference.column + columnOffset),
  };
}

export function offsetAst(node: FormulaAst, rowOffset: number, columnOffset: number): FormulaAst {
  return mapAstReferences(node, (reference) => offsetReference(reference, rowOffset, columnOffset));
}

/** Rename qualified worksheet references without touching string literals. */
export function renameAstSheetReferences(
  node: FormulaAst,
  oldName: string,
  newName: string,
): FormulaAst {
  const normalizedOld = oldName.trim().toLocaleLowerCase();
  return mapAstReferences(node, (reference) => {
    if (reference.sheetId?.trim().toLocaleLowerCase() !== normalizedOld) return reference;
    return { ...reference, sheetId: newName };
  });
}
