import type { FormulaAst, ParsedCellReference } from './ast';

const MAX_ROW_INDEX = 1_048_575;
const MAX_COLUMN_INDEX = 16_383;

export interface StructuralShift {
  axis: 'row' | 'column';
  at: number;
  count: number;
  op: 'insert' | 'delete';
}

export interface StructuralReferenceContext {
  readonly shift: StructuralShift;
  readonly cellShift?: CellShiftReferenceTransform;
  readonly ownerSheetId: string;
  readonly targetSheetId: string;
  readonly targetSheetName?: string;
  readonly sheetOrder?: readonly { readonly id: string; readonly name: string }[];
}

export interface CellShiftReferenceTransform {
  readonly axis: 'row' | 'column';
  readonly selection: {
    readonly startRow: number;
    readonly endRow: number;
    readonly startColumn: number;
    readonly endColumn: number;
  };
  readonly direction: 1 | -1;
}

export interface MoveRangeReferenceTransform {
  readonly selection: {
    readonly sheetId: string;
    readonly startRow: number;
    readonly endRow: number;
    readonly startColumn: number;
    readonly endColumn: number;
  };
  readonly rowDelta: number;
  readonly columnDelta: number;
  readonly ownerSheetId: string;
  readonly targetSheetId: string;
  readonly targetSheetName?: string;
  readonly sheetOrder?: readonly { readonly id: string; readonly name: string }[];
}

/**
 * Transform an inclusive worksheet interval as one reference. Structural
 * deletion removes the deleted coordinates from the interval and joins the
 * surviving sides; it invalidates the reference only when no coordinate
 * survives. This differs from mapping each endpoint independently.
 */
export function transformReferenceInterval(
  start: number,
  end: number,
  shift: StructuralShift,
): { readonly start: number; readonly end: number } | undefined {
  const low = Math.min(start, end);
  const high = Math.max(start, end);
  let interval: { readonly start: number; readonly end: number } | undefined;
  if (shift.op === 'insert') {
    if (shift.at <= low) interval = { start: low + shift.count, end: high + shift.count };
    else if (shift.at <= high) interval = { start: low, end: high + shift.count };
    else interval = { start: low, end: high };
  } else {
    const deletedEnd = shift.at + shift.count - 1;
    if (high < shift.at) interval = { start: low, end: high };
    else if (low > deletedEnd) interval = { start: low - shift.count, end: high - shift.count };
    else {
      const nextStart = low < shift.at ? low : shift.at;
      const nextEnd = high > deletedEnd ? high - shift.count : shift.at - 1;
      if (nextStart <= nextEnd) interval = { start: nextStart, end: nextEnd };
    }
  }
  const maximum = shift.axis === 'row' ? MAX_ROW_INDEX : MAX_COLUMN_INDEX;
  return interval && interval.start >= 0 && interval.end <= maximum ? interval : undefined;
}

function sameSheet(left: string | undefined, right: string): boolean {
  return left?.trim().toLocaleLowerCase() === right.trim().toLocaleLowerCase();
}

function referenceTargetsSheet(sheetId: string | undefined, context: StructuralReferenceContext): boolean {
  if (sheetId === undefined) return sameSheet(context.ownerSheetId, context.targetSheetId);
  return sameSheet(sheetId, context.targetSheetId)
    || (context.targetSheetName !== undefined && sameSheet(sheetId, context.targetSheetName));
}

function mapStructuralCellReference(
  reference: ParsedCellReference,
  context: StructuralReferenceContext,
): ParsedCellReference | undefined {
  if (!referenceTargetsSheet(reference.sheetId, context)) return reference;
  if (context.cellShift) return mapCellShiftReference(reference, context.cellShift);
  return remapReference(reference, context.shift);
}

function mapCellShiftReference(
  reference: ParsedCellReference,
  transform: CellShiftReferenceTransform,
): ParsedCellReference | undefined {
  const { selection, axis, direction } = transform;
  if (axis === 'row') {
    if (reference.column < selection.startColumn || reference.column > selection.endColumn
      || reference.row < selection.startRow) return reference;
    if (direction < 0 && reference.row <= selection.endRow) return undefined;
    const count = selection.endRow - selection.startRow + 1;
    const row = reference.row + direction * count;
    if (row > MAX_ROW_INDEX) throw new Error('UNSUPPORTED_FEATURE: cell shift moves a reference outside worksheet row bounds');
    return { ...reference, row };
  }
  if (reference.row < selection.startRow || reference.row > selection.endRow
    || reference.column < selection.startColumn) return reference;
  if (direction < 0 && reference.column <= selection.endColumn) return undefined;
  const count = selection.endColumn - selection.startColumn + 1;
  const column = reference.column + direction * count;
  if (column > MAX_COLUMN_INDEX) throw new Error('UNSUPPORTED_FEATURE: cell shift moves a reference outside worksheet column bounds');
  return { ...reference, column };
}

interface ReferenceRectangle {
  readonly startRow: number;
  readonly endRow: number;
  readonly startColumn: number;
  readonly endColumn: number;
}

function splitInterval(start: number, end: number, boundaries: readonly number[]): Array<[number, number]> {
  const cuts = [...new Set([start, ...boundaries.filter((value) => value > start && value <= end), end + 1])]
    .sort((left, right) => left - right);
  return cuts.slice(0, -1).map((value, index) => [value, cuts[index + 1]! - 1]);
}

function mergeReferenceRectangles(rectangles: ReferenceRectangle[]): ReferenceRectangle[] {
  const result = [...rectangles];
  let merged = true;
  while (merged) {
    merged = false;
    for (let leftIndex = 0; leftIndex < result.length && !merged; leftIndex += 1) {
      for (let rightIndex = leftIndex + 1; rightIndex < result.length; rightIndex += 1) {
        const left = result[leftIndex]!;
        const right = result[rightIndex]!;
        const sameRows = left.startRow === right.startRow && left.endRow === right.endRow;
        const sameColumns = left.startColumn === right.startColumn && left.endColumn === right.endColumn;
        if (sameRows && (left.endColumn + 1 === right.startColumn || right.endColumn + 1 === left.startColumn)) {
          result[leftIndex] = {
            ...left,
            startColumn: Math.min(left.startColumn, right.startColumn),
            endColumn: Math.max(left.endColumn, right.endColumn),
          };
        } else if (sameColumns && (left.endRow + 1 === right.startRow || right.endRow + 1 === left.startRow)) {
          result[leftIndex] = {
            ...left,
            startRow: Math.min(left.startRow, right.startRow),
            endRow: Math.max(left.endRow, right.endRow),
          };
        } else continue;
        result.splice(rightIndex, 1);
        merged = true;
        break;
      }
    }
  }
  return result;
}

function transformCellShiftRange(
  start: ParsedCellReference,
  end: ParsedCellReference,
  transform: CellShiftReferenceTransform,
): ReferenceRectangle | undefined {
  const lowRow = Math.min(start.row, end.row);
  const highRow = Math.max(start.row, end.row);
  const lowColumn = Math.min(start.column, end.column);
  const highColumn = Math.max(start.column, end.column);
  const rowCuts = transform.axis === 'row'
    ? [transform.selection.startRow, ...(transform.direction < 0 ? [transform.selection.endRow + 1] : [])]
    : [transform.selection.startRow, transform.selection.endRow + 1];
  const columnCuts = transform.axis === 'column'
    ? [transform.selection.startColumn, ...(transform.direction < 0 ? [transform.selection.endColumn + 1] : [])]
    : [transform.selection.startColumn, transform.selection.endColumn + 1];
  const rectangles: ReferenceRectangle[] = [];
  for (const [startRow, endRow] of splitInterval(lowRow, highRow, rowCuts)) {
    for (const [startColumn, endColumn] of splitInterval(lowColumn, highColumn, columnCuts)) {
      const mappedStart = mapCellShiftReference({
        sheetId: start.sheetId ?? end.sheetId,
        row: startRow,
        column: startColumn,
        absoluteRow: false,
        absoluteColumn: false,
      }, transform);
      if (!mappedStart) continue;
      const mappedEnd = mapCellShiftReference({
        sheetId: start.sheetId ?? end.sheetId,
        row: endRow,
        column: endColumn,
        absoluteRow: false,
        absoluteColumn: false,
      }, transform);
      if (!mappedEnd) continue;
      rectangles.push({
        startRow: Math.min(mappedStart.row, mappedEnd.row),
        endRow: Math.max(mappedStart.row, mappedEnd.row),
        startColumn: Math.min(mappedStart.column, mappedEnd.column),
        endColumn: Math.max(mappedStart.column, mappedEnd.column),
      });
    }
  }

  const merged = mergeReferenceRectangles(rectangles);
  if (merged.length > 1) {
    throw new Error('UNSUPPORTED_FEATURE: cell shift makes this formula range non-contiguous');
  }
  return merged[0];
}

function rangeTargetsSheet(
  start: ParsedCellReference,
  end: ParsedCellReference,
  context: StructuralReferenceContext,
): boolean {
  if (start.sheetId !== undefined && end.sheetId !== undefined && !sameSheet(start.sheetId, end.sheetId)) {
    throw new Error('UNSUPPORTED_STRUCTURAL_REFERENCE: range endpoints target different worksheets');
  }
  return referenceTargetsSheet(start.sheetId ?? end.sheetId, context);
}

function transformStructuralRange(
  start: ParsedCellReference,
  end: ParsedCellReference,
  context: StructuralReferenceContext,
): { readonly start: ParsedCellReference; readonly end: ParsedCellReference } | undefined {
  if (!rangeTargetsSheet(start, end, context)) return { start, end };
  if (context.cellShift) {
    const rectangle = transformCellShiftRange(start, end, context.cellShift);
    if (!rectangle) return undefined;
    const reverseRows = start.row > end.row;
    const reverseColumns = start.column > end.column;
    return {
      start: {
        ...start,
        row: reverseRows ? rectangle.endRow : rectangle.startRow,
        column: reverseColumns ? rectangle.endColumn : rectangle.startColumn,
      },
      end: {
        ...end,
        row: reverseRows ? rectangle.startRow : rectangle.endRow,
        column: reverseColumns ? rectangle.startColumn : rectangle.endColumn,
      },
    };
  }

  const startCoordinate = context.shift.axis === 'row' ? start.row : start.column;
  const endCoordinate = context.shift.axis === 'row' ? end.row : end.column;
  const interval = transformReferenceInterval(startCoordinate, endCoordinate, context.shift);
  if (!interval) return undefined;

  const reversed = startCoordinate > endCoordinate;
  const mappedStartCoordinate = reversed ? interval.end : interval.start;
  const mappedEndCoordinate = reversed ? interval.start : interval.end;
  const mapCoordinate = (reference: ParsedCellReference, coordinate: number): ParsedCellReference =>
    context.shift.axis === 'row' ? { ...reference, row: coordinate } : { ...reference, column: coordinate };
  return {
    start: mapCoordinate(start, mappedStartCoordinate),
    end: mapCoordinate(end, mappedEndCoordinate),
  };
}

/**
 * Structural reference transform for formula ASTs. `$` markers are retained
 * but do not pin coordinates during insert/delete operations.
 */
export function mapAstStructuralReferences(
  node: FormulaAst,
  context: StructuralReferenceContext,
): FormulaAst {
  const invalid = (): FormulaAst => ({
    type: 'invalid-reference',
    code: '#REF!',
    span: node.span,
    parenthesized: 'parenthesized' in node ? node.parenthesized : undefined,
  });
  switch (node.type) {
    case 'number-literal':
    case 'string-literal':
    case 'boolean-literal':
    case 'name-reference':
    case 'table-reference':
    case 'invalid-reference':
    case 'external-reference':
      return node;
    case 'cell-reference': {
      const reference = mapStructuralCellReference(node.reference, context);
      return reference === undefined ? invalid() : { ...node, reference };
    }
    case 'range-reference': {
      const mapped = transformStructuralRange(node.start.reference, node.end.reference, context);
      if (!mapped) return invalid();
      return {
        ...node,
        start: { ...node.start, reference: mapped.start },
        end: { ...node.end, reference: mapped.end },
      };
    }
    case 'whole-row-reference': {
      if (context.cellShift) return node;
      if (context.shift.axis !== 'row' || !referenceTargetsSheet(node.sheetId, context)) return node;
      const interval = transformReferenceInterval(node.startRow, node.endRow, context.shift);
      return interval ? { ...node, startRow: interval.start, endRow: interval.end } : invalid();
    }
    case 'whole-column-reference': {
      if (context.cellShift) return node;
      if (context.shift.axis !== 'column' || !referenceTargetsSheet(node.sheetId, context)) return node;
      const interval = transformReferenceInterval(node.startColumn, node.endColumn, context.shift);
      return interval ? { ...node, startColumn: interval.start, endColumn: interval.end } : invalid();
    }
    case 'spill-reference':
      return { ...node, operand: mapAstStructuralReferences(node.operand, context) };
    case 'reference-union':
      return { ...node, references: node.references.map((reference) => mapAstStructuralReferences(reference, context) as typeof reference) };
    case 'reference-intersection':
      return {
        ...node,
        left: mapAstStructuralReferences(node.left, context) as typeof node.left,
        right: mapAstStructuralReferences(node.right, context) as typeof node.right,
      };
    case 'sheet-range-reference': {
      const findSheetIndex = (reference: string): number => context.sheetOrder?.findIndex((sheet) =>
        sameSheet(sheet.id, reference) || sameSheet(sheet.name, reference)) ?? -1;
      const start = findSheetIndex(node.qualifier.startSheetId);
      const end = findSheetIndex(node.qualifier.endSheetId);
      if (start < 0 || end < 0) throw new Error('UNSUPPORTED_STRUCTURAL_REFERENCE: 3D reference sheet boundary is unresolved');
      const target = context.sheetOrder?.findIndex((sheet) => sameSheet(sheet.id, context.targetSheetId)) ?? -1;
      if (target >= Math.min(start, end) && target <= Math.max(start, end)) {
        throw new Error('UNSUPPORTED_STRUCTURAL_REFERENCE: structural edits cannot rewrite one sheet inside a 3D reference');
      }
      return node;
    }
    case 'unary-expression':
      return { ...node, operand: mapAstStructuralReferences(node.operand, context) };
    case 'binary-expression':
      return {
        ...node,
        left: mapAstStructuralReferences(node.left, context),
        right: mapAstStructuralReferences(node.right, context),
      };
    case 'function-call':
      return { ...node, arguments: node.arguments.map((argument) => mapAstStructuralReferences(argument, context)) };
  }
}

/**
 * Rewrites references to a moved cell block without flattening partially
 * intersecting ranges into an invalid A1 rectangle. Such ranges can become
 * multi-area references and are rejected until that reference family is
 * represented canonically.
 */
export function mapAstMovedReferences(node: FormulaAst, context: MoveRangeReferenceTransform): FormulaAst {
  const targetsSheet = (sheetId: string | undefined): boolean => sheetId === undefined
    ? sameSheet(context.ownerSheetId, context.targetSheetId)
    : sameSheet(sheetId, context.targetSheetId)
      || (context.targetSheetName !== undefined && sameSheet(sheetId, context.targetSheetName));
  const mapCell = (reference: ParsedCellReference): ParsedCellReference => {
    if (!targetsSheet(reference.sheetId)) return reference;
    const { selection } = context;
    if (reference.row < selection.startRow || reference.row > selection.endRow
      || reference.column < selection.startColumn || reference.column > selection.endColumn) return reference;
    return { ...reference, row: reference.row + context.rowDelta, column: reference.column + context.columnDelta };
  };
  const mapRange = (start: ParsedCellReference, end: ParsedCellReference): { start: ParsedCellReference; end: ParsedCellReference } => {
    if (start.sheetId !== undefined && end.sheetId !== undefined && !sameSheet(start.sheetId, end.sheetId)) {
      throw new Error('UNSUPPORTED_STRUCTURAL_REFERENCE: move range endpoints target different worksheets');
    }
    if (!targetsSheet(start.sheetId ?? end.sheetId)) return { start, end };
    const lowRow = Math.min(start.row, end.row);
    const highRow = Math.max(start.row, end.row);
    const lowColumn = Math.min(start.column, end.column);
    const highColumn = Math.max(start.column, end.column);
    const selection = context.selection;
    const intersects = lowRow <= selection.endRow && highRow >= selection.startRow
      && lowColumn <= selection.endColumn && highColumn >= selection.startColumn;
    if (!intersects) return { start, end };
    const contained = lowRow >= selection.startRow && highRow <= selection.endRow
      && lowColumn >= selection.startColumn && highColumn <= selection.endColumn;
    if (!contained) throw new Error('UNSUPPORTED_FEATURE: moving this range would make a formula reference non-contiguous');
    return { start: mapCell(start), end: mapCell(end) };
  };
  const map3d = (startSheetId: string, endSheetId: string): void => {
    const sheetIndex = (reference: string): number => context.sheetOrder?.findIndex((sheet) =>
      sameSheet(sheet.id, reference) || sameSheet(sheet.name, reference)) ?? -1;
    const start = sheetIndex(startSheetId);
    const end = sheetIndex(endSheetId);
    const target = sheetIndex(context.targetSheetId);
    if (start < 0 || end < 0 || target < 0) {
      throw new Error('UNSUPPORTED_STRUCTURAL_REFERENCE: 3D reference boundary is unresolved');
    }
    if (target >= Math.min(start, end) && target <= Math.max(start, end)) {
      throw new Error('UNSUPPORTED_STRUCTURAL_REFERENCE: moving cells cannot rewrite one sheet inside a 3D reference');
    }
  };
  const assertWholeAxisMoveIsRepresentable = (
    referenceStart: number,
    referenceEnd: number,
    sourceStart: number,
    sourceEnd: number,
    delta: number,
    label: 'row' | 'column',
  ): void => {
    if (delta === 0) return;
    const targetStart = sourceStart + delta;
    const targetEnd = sourceEnd + delta;
    const intersects = (start: number, end: number): boolean => start <= referenceEnd && end >= referenceStart;
    const contains = (start: number, end: number): boolean => start >= referenceStart && end <= referenceEnd;
    const sourceIntersects = intersects(sourceStart, sourceEnd);
    const targetIntersects = intersects(targetStart, targetEnd);
    const sourceAndTargetCovered = contains(sourceStart, sourceEnd) && contains(targetStart, targetEnd);
    if (!sourceAndTargetCovered && (sourceIntersects || targetIntersects)) {
      throw new Error(`UNSUPPORTED_STRUCTURAL_REFERENCE: moving cells would make a whole-${label} reference non-contiguous`);
    }
  };
  switch (node.type) {
    case 'number-literal':
    case 'string-literal':
    case 'boolean-literal':
    case 'name-reference':
    case 'table-reference':
    case 'invalid-reference':
    case 'external-reference':
      return node;
    case 'whole-row-reference':
      if (targetsSheet(node.sheetId)) {
        assertWholeAxisMoveIsRepresentable(
          Math.min(node.startRow, node.endRow), Math.max(node.startRow, node.endRow),
          context.selection.startRow, context.selection.endRow, context.rowDelta, 'row',
        );
      }
      return node;
    case 'whole-column-reference':
      if (targetsSheet(node.sheetId)) {
        assertWholeAxisMoveIsRepresentable(
          Math.min(node.startColumn, node.endColumn), Math.max(node.startColumn, node.endColumn),
          context.selection.startColumn, context.selection.endColumn, context.columnDelta, 'column',
        );
      }
      return node;
    case 'cell-reference':
      return { ...node, reference: mapCell(node.reference) };
    case 'range-reference': {
      const mapped = mapRange(node.start.reference, node.end.reference);
      return {
        ...node,
        start: { ...node.start, reference: mapped.start },
        end: { ...node.end, reference: mapped.end },
      };
    }
    case 'spill-reference':
      return { ...node, operand: mapAstMovedReferences(node.operand, context) };
    case 'reference-union':
      return { ...node, references: node.references.map((reference) => mapAstMovedReferences(reference, context) as typeof reference) };
    case 'reference-intersection':
      return {
        ...node,
        left: mapAstMovedReferences(node.left, context) as typeof node.left,
        right: mapAstMovedReferences(node.right, context) as typeof node.right,
      };
    case 'sheet-range-reference':
      map3d(node.qualifier.startSheetId, node.qualifier.endSheetId);
      return node;
    case 'unary-expression':
      return { ...node, operand: mapAstMovedReferences(node.operand, context) };
    case 'binary-expression':
      return {
        ...node,
        left: mapAstMovedReferences(node.left, context),
        right: mapAstMovedReferences(node.right, context),
      };
    case 'function-call':
      return { ...node, arguments: node.arguments.map((argument) => mapAstMovedReferences(argument, context)) };
  }
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
    case 'spill-reference':
      return { ...node, operand: mapAstReferences(node.operand, mapper) };
    case 'whole-column-reference':
    case 'whole-row-reference':
      return node;
    case 'reference-union':
      return { ...node, references: node.references.map((reference) => mapAstReferences(reference, mapper) as typeof reference) };
    case 'reference-intersection':
      return {
        ...node,
        left: mapAstReferences(node.left, mapper) as typeof node.left,
        right: mapAstReferences(node.right, mapper) as typeof node.right,
      };
    case 'sheet-range-reference':
      return { ...node, reference: mapAstReferences(node.reference, mapper) as typeof node.reference };
    case 'external-reference':
      return { ...node, reference: mapAstReferences(node.reference, mapper) as typeof node.reference };
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
function remapReference(ref: ParsedCellReference, shift: StructuralShift): ParsedCellReference | undefined {
  if (shift.op === 'insert') {
    const before = shift.axis === 'row' ? ref.row : ref.column;
    if (before >= shift.at) {
      const coordinate = before + shift.count;
      const maximum = shift.axis === 'row' ? MAX_ROW_INDEX : MAX_COLUMN_INDEX;
      if (coordinate > maximum) return undefined;
      return shift.axis === 'row' ? { ...ref, row: coordinate } : { ...ref, column: coordinate };
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

/** Shift relative references when a formula is copied to another cell. */
export function offsetReference(
  reference: ParsedCellReference,
  rowOffset: number,
  columnOffset: number,
): ParsedCellReference | undefined {
  if (!Number.isSafeInteger(rowOffset) || !Number.isSafeInteger(columnOffset)) {
    throw new Error('Formula reference offset must use safe integer coordinates');
  }
  const row = reference.absoluteRow ? reference.row : reference.row + rowOffset;
  const column = reference.absoluteColumn ? reference.column : reference.column + columnOffset;
  if (!Number.isSafeInteger(row) || row < 0 || row > MAX_ROW_INDEX
    || !Number.isSafeInteger(column) || column < 0 || column > MAX_COLUMN_INDEX) return undefined;
  return {
    ...reference,
    row,
    column,
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
  const mapped = mapAstReferences(node, (reference) => {
    if (reference.sheetId?.trim().toLocaleLowerCase() !== normalizedOld) return reference;
    return { ...reference, sheetId: newName };
  });
  return renameQualifiedSheets(mapped, normalizedOld, newName);
}

function renameQualifiedSheets(node: FormulaAst, normalizedOld: string, newName: string): FormulaAst {
  switch (node.type) {
    case 'whole-column-reference':
    case 'whole-row-reference':
      return node.sheetId?.trim().toLocaleLowerCase() === normalizedOld ? { ...node, sheetId: newName } : node;
    case 'sheet-range-reference':
      return {
        ...node,
        qualifier: {
          startSheetId: node.qualifier.startSheetId.trim().toLocaleLowerCase() === normalizedOld ? newName : node.qualifier.startSheetId,
          endSheetId: node.qualifier.endSheetId.trim().toLocaleLowerCase() === normalizedOld ? newName : node.qualifier.endSheetId,
        },
        reference: renameQualifiedSheets(node.reference, normalizedOld, newName) as typeof node.reference,
      };
    case 'external-reference':
      return {
        ...node,
        qualifier: {
          ...node.qualifier,
          sheetId: node.qualifier.sheetId?.trim().toLocaleLowerCase() === normalizedOld ? newName : node.qualifier.sheetId,
        },
        reference: renameQualifiedSheets(node.reference, normalizedOld, newName) as typeof node.reference,
      };
    case 'reference-union':
      return { ...node, references: node.references.map((reference) => renameQualifiedSheets(reference, normalizedOld, newName) as typeof reference) };
    case 'reference-intersection':
      return { ...node, left: renameQualifiedSheets(node.left, normalizedOld, newName) as typeof node.left, right: renameQualifiedSheets(node.right, normalizedOld, newName) as typeof node.right };
    case 'spill-reference':
      return { ...node, operand: renameQualifiedSheets(node.operand, normalizedOld, newName) };
    case 'range-reference':
      return { ...node, start: renameQualifiedSheets(node.start, normalizedOld, newName) as typeof node.start, end: renameQualifiedSheets(node.end, normalizedOld, newName) as typeof node.end };
    default:
      return node;
  }
}
