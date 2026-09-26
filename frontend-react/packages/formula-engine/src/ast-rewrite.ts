import type { FormulaAst, ParsedCellReference, TableReferenceNode } from './ast';
import { formulaSheetReferenceIndex, sameFormulaSheetName } from './sheet-reference';
import { ReferenceTransformDomain, MAX_COLUMN_INDEX, MAX_ROW_INDEX } from './reference-transform-domain';
import type { StructuralShift } from './reference-transform-domain';
import { lexFormula } from './lexer';
import { parseFormula } from './parser';

export { MAX_COLUMN_INDEX, MAX_ROW_INDEX } from './reference-transform-domain';
export type { StructuralShift } from './reference-transform-domain';

export function mapAstTableReferences(
  node: FormulaAst,
  mapper: (reference: TableReferenceNode) => TableReferenceNode,
): FormulaAst {
  switch (node.type) {
    case 'table-reference':
      return mapper(node);
    case 'spill-reference':
      return { ...node, operand: mapAstTableReferences(node.operand, mapper) as typeof node.operand };
    case 'reference-union':
      return { ...node, references: node.references.map((reference) => mapAstTableReferences(reference, mapper) as typeof reference) };
    case 'reference-intersection':
      return {
        ...node,
        left: mapAstTableReferences(node.left, mapper) as typeof node.left,
        right: mapAstTableReferences(node.right, mapper) as typeof node.right,
      };
    case 'sheet-range-reference':
    case 'external-reference':
      return { ...node, reference: mapAstTableReferences(node.reference, mapper) as typeof node.reference };
    case 'unary-expression':
      return { ...node, operand: mapAstTableReferences(node.operand, mapper) };
    case 'binary-expression':
      return {
        ...node,
        left: mapAstTableReferences(node.left, mapper),
        right: mapAstTableReferences(node.right, mapper),
      };
    case 'function-call':
      return { ...node, arguments: node.arguments.map((argument) => mapAstTableReferences(argument, mapper)) };
    default:
      return node;
  }
}

/** Rewrite only the table-name token of matching structured references. */
export function rewriteFormulaTableReferences(source: string, tableName: string, replacement: string): string {
  const normalizedName = tableName.trim().toUpperCase();
  if (!normalizedName || tableName === replacement) return source;

  let ast: FormulaAst;
  try {
    ast = parseFormula(source);
  } catch (error) {
    if (hasStructuredTableCandidate(source, normalizedName)) {
      throw new Error(`UNSUPPORTED_STRUCTURAL_REFERENCE: formula contains an unparseable reference to table ${tableName}`, { cause: error });
    }
    return source;
  }

  const spans: Array<{ start: number; end: number; token: string }> = [];
  mapAstTableReferences(ast, (reference) => {
    if (reference.tableName.trim().toUpperCase() === normalizedName) {
      spans.push({ start: reference.span.start, end: reference.span.start + reference.tableName.length, token: reference.tableName });
    }
    return reference;
  });
  if (spans.length === 0) return source;
  spans.sort((left, right) => right.start - left.start);
  let rewritten = source;
  for (const span of spans) {
    if (rewritten.slice(span.start, span.end).toUpperCase() !== span.token.toUpperCase()) {
      throw new Error(`STRUCTURAL_PATCH_INVARIANT: table reference span no longer matches ${span.token}`);
    }
    rewritten = `${rewritten.slice(0, span.start)}${replacement}${rewritten.slice(span.end)}`;
  }
  return rewritten;
}

function hasStructuredTableCandidate(source: string, normalizedName: string): boolean {
  try {
    const tokens = lexFormula(source);
    return tokens.some((token, index) => token.kind === 'identifier'
      && token.lexeme.trim().toUpperCase() === normalizedName
      && !isExternalWorkbookTableName(tokens, index)
      && tokens[index + 1]?.kind === 'left-bracket');
  } catch {
    let index = 0;
    while (index < source.length) {
      if (source[index] === '"') {
        const quoteStart = index;
        index += 1;
        let terminated = false;
        while (index < source.length) {
          if (source[index] !== '"') {
            index += 1;
            continue;
          }
          if (source[index + 1] === '"') {
            index += 2;
            continue;
          }
          index += 1;
          terminated = true;
          break;
        }
        if (!terminated && source.slice(quoteStart + 1).toUpperCase().includes(`${normalizedName}[`)) return true;
        continue;
      }
      const start = source[index]!;
      if (!/[A-Za-z_]/.test(start) || index > 0 && /[\p{L}\p{N}_.]/u.test(source[index - 1]!)) {
        index += 1;
        continue;
      }
      let end = index + 1;
      while (end < source.length && /[A-Za-z0-9_.]/.test(source[end]!)) end += 1;
      if (source.slice(index, end).toUpperCase() === normalizedName && source[end] === '[') return true;
      index = end;
    }
    return false;
  }
}

function isExternalWorkbookTableName(tokens: ReturnType<typeof lexFormula>, index: number): boolean {
  return tokens[index - 1]?.kind === 'right-bracket'
    && tokens[index - 2]?.kind === 'identifier'
    && tokens[index - 3]?.kind === 'left-bracket';
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

function sameSheet(left: string | undefined, right: string): boolean {
  return sameFormulaSheetName(left, right);
}

function referenceTargetsSheet(sheetId: string | undefined, context: StructuralReferenceContext): boolean {
  return referenceTargetsWorksheet(
    sheetId,
    context.ownerSheetId,
    context.targetSheetId,
    context.targetSheetName,
    context.sheetOrder,
  );
}

function referenceTargetsWorksheet(
  reference: string | undefined,
  ownerSheetId: string,
  targetSheetId: string,
  targetSheetName: string | undefined,
  sheetOrder: readonly { readonly id: string; readonly name: string }[] | undefined,
): boolean {
  if (reference === undefined) return ownerSheetId === targetSheetId;
  if (sheetOrder) {
    const index = sheetReferenceIndex(reference, sheetOrder);
    return index >= 0 && sheetOrder[index]?.id === targetSheetId;
  }
  return reference === targetSheetId
    || (targetSheetName !== undefined && sameSheet(reference, targetSheetName));
}

function sheetReferenceIndex(
  reference: string,
  sheetOrder: readonly { readonly id: string; readonly name: string }[] | undefined,
): number {
  return formulaSheetReferenceIndex(reference, sheetOrder);
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
  const mapped = ReferenceTransformDomain.mapCellShiftPoint(
    reference.row,
    reference.column,
    transform.selection,
    transform.axis,
    transform.direction === 1 ? 'insert' : 'delete',
  );
  if (mapped.kind === 'deleted') return undefined;
  if (mapped.kind === 'out-of-bounds') {
    throw new Error(`UNSUPPORTED_FEATURE: cell shift moves a reference outside worksheet ${transform.axis} bounds`);
  }
  if (mapped.row === reference.row && mapped.column === reference.column) return reference;
  return { ...reference, row: mapped.row, column: mapped.column };
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
  const startTargets = referenceTargetsSheet(start.sheetId, context);
  const endTargets = referenceTargetsSheet(end.sheetId, context);
  if (!startTargets && !endTargets) return false;
  if (startTargets !== endTargets) {
    throw new Error('UNSUPPORTED_STRUCTURAL_REFERENCE: structural transform cannot rewrite a partially qualified range');
  }
  return true;
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
  const interval = ReferenceTransformDomain.mapInterval(startCoordinate, endCoordinate, context.shift);
  if (interval.kind !== 'mapped') return undefined;

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
      const interval = ReferenceTransformDomain.mapInterval(node.startRow, node.endRow, context.shift);
      return interval.kind === 'mapped' ? { ...node, startRow: interval.start, endRow: interval.end } : invalid();
    }
    case 'whole-column-reference': {
      if (context.cellShift) return node;
      if (context.shift.axis !== 'column' || !referenceTargetsSheet(node.sheetId, context)) return node;
      const interval = ReferenceTransformDomain.mapInterval(node.startColumn, node.endColumn, context.shift);
      return interval.kind === 'mapped' ? { ...node, startColumn: interval.start, endColumn: interval.end } : invalid();
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
      const findSheetIndex = (reference: string): number => sheetReferenceIndex(reference, context.sheetOrder);
      const start = findSheetIndex(node.qualifier.startSheetId);
      const end = findSheetIndex(node.qualifier.endSheetId);
      if (start < 0 || end < 0) throw new Error('UNSUPPORTED_STRUCTURAL_REFERENCE: 3D reference sheet boundary is unresolved');
      const target = context.sheetOrder?.findIndex((sheet) => sheet.id === context.targetSheetId) ?? -1;
      if (target < 0) throw new Error('UNSUPPORTED_STRUCTURAL_REFERENCE: target worksheet identity is unresolved');
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
  const targetsSheet = (sheetId: string | undefined): boolean => referenceTargetsWorksheet(
    sheetId,
    context.ownerSheetId,
    context.targetSheetId,
    context.targetSheetName,
    context.sheetOrder,
  );
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
    const startTargets = targetsSheet(start.sheetId);
    const endTargets = targetsSheet(end.sheetId);
    if (!startTargets && !endTargets) return { start, end };
    if (startTargets !== endTargets) {
      throw new Error('UNSUPPORTED_STRUCTURAL_REFERENCE: moved range has a partially qualified formula reference');
    }
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
    const sheetIndex = (reference: string): number => sheetReferenceIndex(reference, context.sheetOrder);
    const start = sheetIndex(startSheetId);
    const end = sheetIndex(endSheetId);
    const target = context.sheetOrder?.findIndex((sheet) => sheet.id === context.targetSheetId) ?? -1;
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
  const position = shift.axis === 'row' ? ref.row : ref.column;
  const mapped = ReferenceTransformDomain.mapPoint(
    position,
    shift.at,
    shift.count,
    shift.op === 'insert' ? 1 : -1,
    shift.axis === 'row' ? MAX_ROW_INDEX : MAX_COLUMN_INDEX,
  );
  if (mapped.kind !== 'mapped') return undefined;
  if (mapped.position === position) return ref;
  return shift.axis === 'row' ? { ...ref, row: mapped.position } : { ...ref, column: mapped.position };
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

/** Shift relative whole-row and whole-column endpoints during copy/fill. */
export function offsetWholeAxisReferences(node: FormulaAst, rowOffset: number, columnOffset: number): FormulaAst {
  if (!Number.isSafeInteger(rowOffset) || !Number.isSafeInteger(columnOffset)) {
    throw new Error('Formula reference offset must use safe integer coordinates');
  }
  switch (node.type) {
    case 'whole-column-reference': {
      const startColumn = node.absoluteStartColumn ? node.startColumn : node.startColumn + columnOffset;
      const endColumn = node.absoluteEndColumn ? node.endColumn : node.endColumn + columnOffset;
      if (!Number.isSafeInteger(startColumn) || startColumn < 0 || startColumn > MAX_COLUMN_INDEX
        || !Number.isSafeInteger(endColumn) || endColumn < 0 || endColumn > MAX_COLUMN_INDEX) {
        return { type: 'invalid-reference', code: '#REF!', span: node.span, parenthesized: node.parenthesized };
      }
      return { ...node, startColumn, endColumn };
    }
    case 'whole-row-reference': {
      const startRow = node.absoluteStartRow ? node.startRow : node.startRow + rowOffset;
      const endRow = node.absoluteEndRow ? node.endRow : node.endRow + rowOffset;
      if (!Number.isSafeInteger(startRow) || startRow < 0 || startRow > MAX_ROW_INDEX
        || !Number.isSafeInteger(endRow) || endRow < 0 || endRow > MAX_ROW_INDEX) {
        return { type: 'invalid-reference', code: '#REF!', span: node.span, parenthesized: node.parenthesized };
      }
      return { ...node, startRow, endRow };
    }
    case 'spill-reference':
      return { ...node, operand: offsetWholeAxisReferences(node.operand, rowOffset, columnOffset) };
    case 'reference-union':
      return { ...node, references: node.references.map((reference) => offsetWholeAxisReferences(reference, rowOffset, columnOffset) as typeof reference) };
    case 'reference-intersection':
      return {
        ...node,
        left: offsetWholeAxisReferences(node.left, rowOffset, columnOffset) as typeof node.left,
        right: offsetWholeAxisReferences(node.right, rowOffset, columnOffset) as typeof node.right,
      };
    case 'sheet-range-reference':
    case 'external-reference':
      return { ...node, reference: offsetWholeAxisReferences(node.reference, rowOffset, columnOffset) as typeof node.reference };
    case 'unary-expression':
      return { ...node, operand: offsetWholeAxisReferences(node.operand, rowOffset, columnOffset) };
    case 'binary-expression':
      return {
        ...node,
        left: offsetWholeAxisReferences(node.left, rowOffset, columnOffset),
        right: offsetWholeAxisReferences(node.right, rowOffset, columnOffset),
      };
    case 'function-call':
      return { ...node, arguments: node.arguments.map((argument) => offsetWholeAxisReferences(argument, rowOffset, columnOffset)) };
    case 'number-literal':
    case 'string-literal':
    case 'boolean-literal':
    case 'name-reference':
    case 'table-reference':
    case 'cell-reference':
    case 'range-reference':
    case 'invalid-reference':
      return node;
  }
}

export function offsetAst(node: FormulaAst, rowOffset: number, columnOffset: number): FormulaAst {
  return offsetWholeAxisReferences(
    mapAstReferences(node, (reference) => offsetReference(reference, rowOffset, columnOffset)),
    rowOffset,
    columnOffset,
  );
}

/** Rename qualified worksheet references without touching string literals. */
export function renameAstSheetReferences(
  node: FormulaAst,
  oldName: string,
  newName: string,
): FormulaAst {
  const normalizedOld = oldName.trim().toLowerCase();
  const mapped = mapAstReferences(node, (reference) => {
    if (reference.sheetId?.trim().toLowerCase() !== normalizedOld) return reference;
    return { ...reference, sheetId: newName };
  });
  return renameQualifiedSheets(mapped, normalizedOld, newName);
}

function renameQualifiedSheets(node: FormulaAst, normalizedOld: string, newName: string): FormulaAst {
  switch (node.type) {
    case 'whole-column-reference':
    case 'whole-row-reference':
      return node.sheetId?.trim().toLowerCase() === normalizedOld ? { ...node, sheetId: newName } : node;
    case 'sheet-range-reference':
      return {
        ...node,
        qualifier: {
          startSheetId: node.qualifier.startSheetId.trim().toLowerCase() === normalizedOld ? newName : node.qualifier.startSheetId,
          endSheetId: node.qualifier.endSheetId.trim().toLowerCase() === normalizedOld ? newName : node.qualifier.endSheetId,
        },
        reference: renameQualifiedSheets(node.reference, normalizedOld, newName) as typeof node.reference,
      };
    case 'external-reference':
      return {
        ...node,
        qualifier: {
          ...node.qualifier,
          sheetId: node.qualifier.sheetId?.trim().toLowerCase() === normalizedOld ? newName : node.qualifier.sheetId,
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
