import type { RangeRef } from '@react-sheets/core-model';
import { formatFormula, mapAstStructuralReferences, parseFormula, ReferenceTransformDomain, MAX_COLUMN_INDEX, MAX_ROW_INDEX } from '@react-sheets/formula-engine';
import type { FormulaAst } from '@react-sheets/formula-engine';
import { cellAddress, parseAddress } from '../address';
import type { ClassifiedMutation, CollaborationOperationKind } from './operation-types';

export interface StructuralDelta {
  kind: 'insert-rows' | 'delete-rows' | 'insert-columns' | 'delete-columns';
  sheetId: string;
  at: number;
  count: number;
}

export interface RebaseResult {
  rebased: ClassifiedMutation;
  transformed: boolean;
}

export interface StructuralRebaseContext {
  readonly sheetOrder: readonly { readonly id: string; readonly name: string }[];
}

type StructuralAxis = 'row' | 'column';
function deltaAxis(delta: StructuralDelta): StructuralAxis {
  return delta.kind.endsWith('rows') ? 'row' : 'column';
}

function deltaOperation(delta: StructuralDelta): 'insert' | 'delete' {
  return delta.kind.startsWith('insert-') ? 'insert' : 'delete';
}

function rebaseConflict(message: string): never {
  throw new Error(`STRUCTURAL_REBASE_CONFLICT: ${message}`);
}

function shiftPoint(value: number, delta: StructuralDelta): number {
  const maximum = deltaAxis(delta) === 'row' ? MAX_ROW_INDEX : MAX_COLUMN_INDEX;
  if (!Number.isSafeInteger(value) || value < 0 || value > maximum) rebaseConflict('coordinate is outside worksheet bounds');
  const mapped = ReferenceTransformDomain.mapPoint(value, delta.at, delta.count, deltaOperation(delta) === 'insert' ? 1 : -1, maximum);
  if (mapped.kind === 'deleted') rebaseConflict(`coordinate ${value} was deleted by the committed operation`);
  if (mapped.kind === 'out-of-bounds') rebaseConflict('coordinate exceeds worksheet bounds');
  return mapped.position;
}

function shiftRange(range: RangeRef, delta: StructuralDelta): RangeRef {
  if (range.sheetId !== delta.sheetId) return range;
  const axis = deltaAxis(delta);
  const startKey = axis === 'row' ? 'startRow' : 'startColumn';
  const endKey = axis === 'row' ? 'endRow' : 'endColumn';
  const start = range[startKey];
  const end = range[endKey];
  if (!Number.isSafeInteger(start + delta.count) || !Number.isSafeInteger(end + delta.count)) {
    rebaseConflict('range exceeds the safe integer range after transformation');
  }
  if (deltaOperation(delta) === 'insert' && delta.at > start && delta.at <= end) {
    rebaseConflict('insertion splits the pending range into non-contiguous coordinates');
  }
  const interval = ReferenceTransformDomain.mapInterval(range[startKey], range[endKey], {
    axis,
    at: delta.at,
    count: delta.count,
    op: deltaOperation(delta),
  });
  if (interval.kind === 'deleted') rebaseConflict('the committed deletion removed the entire pending range');
  if (interval.kind === 'out-of-bounds') rebaseConflict('range exceeds worksheet bounds after transformation');
  const maximum = axis === 'row' ? MAX_ROW_INDEX : MAX_COLUMN_INDEX;
  if (interval.start < 0 || interval.end > maximum) rebaseConflict('range exceeds worksheet bounds after transformation');
  return { ...range, [startKey]: interval.start, [endKey]: interval.end };
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object' && !Array.isArray(value);
}

function isRange(value: Record<string, unknown>): value is Record<string, unknown> & RangeRef {
  return typeof value.sheetId === 'string'
    && Number.isSafeInteger(value.startRow) && Number.isSafeInteger(value.endRow)
    && Number.isSafeInteger(value.startColumn) && Number.isSafeInteger(value.endColumn);
}

const UNSUPPORTED_STRUCTURAL_KINDS = new Set<CollaborationOperationKind>([
  'move-range', 'sort', 'table-resize', 'sheet-identity', 'pivot-config',
]);

function transformParams(
  value: unknown,
  delta: StructuralDelta,
  ownerSheetId: string,
  context: StructuralRebaseContext,
  field = '',
): unknown {
  if (typeof value === 'string' && isFormulaField(field, value) && value.trim() !== '') {
    return transformFormulaValue(value, delta, ownerSheetId, context);
  }
  if (Array.isArray(value)) {
    return value.map((entry) => transformParams(entry, delta, ownerSheetId, context, field));
  }
  if (!isRecord(value)) return value;

  const sheetId = typeof value.sheetId === 'string' ? value.sheetId : ownerSheetId;
  const formulaOwnerSheetId = isRecord(value.formulaAnchor) && typeof value.formulaAnchor.sheetId === 'string'
    ? value.formulaAnchor.sheetId
    : sheetId;
  if (isRange(value)) return shiftRange(value, delta);

  const next: Record<string, unknown> = {};
  for (const [key, entry] of Object.entries(value)) {
    if (typeof entry === 'string' && (isFormulaField(key, entry) || isRuleFormulaField(key, entry, value))) {
      next[key] = transformFormulaValue(entry, delta, formulaOwnerSheetId, context);
    } else {
      next[key] = transformParams(entry, delta, sheetId, context, key);
    }
  }
  return next;
}

function assertStructuralEditDoesNotIntersectRange(range: RangeRef, delta: StructuralDelta, label: string): void {
  if (range.sheetId !== delta.sheetId || deltaOperation(delta) !== 'delete') return;
  const start = deltaAxis(delta) === 'row' ? range.startRow : range.startColumn;
  const end = deltaAxis(delta) === 'row' ? range.endRow : range.endColumn;
  const deletedEnd = delta.at + delta.count - 1;
  if (start <= deletedEnd && end >= delta.at) {
    rebaseConflict(`committed deletion intersects pending ${label}`);
  }
}

function shiftCellCoordinate(value: unknown, sheetId: string, delta: StructuralDelta, label: string): Record<string, unknown> {
  if (!isRecord(value)) rebaseConflict(`${label} is not a coordinate record`);
  const ownerSheetId = typeof value.sheetId === 'string' ? value.sheetId : sheetId;
  if (ownerSheetId !== delta.sheetId) return value;
  const key = deltaAxis(delta) === 'row' ? 'row' : 'column';
  const coordinate = value[key];
  if (typeof coordinate !== 'number') rebaseConflict(`${label} has no ${key} coordinate`);
  return { ...value, [key]: shiftPoint(coordinate, delta) };
}

function transformCellFormulaOwners(
  value: unknown,
  ownerSheetId: string,
  delta: StructuralDelta,
  context: StructuralRebaseContext,
  label: string,
): Record<string, unknown> {
  if (!isRecord(value)) rebaseConflict(`${label} is not a cell snapshot`);
  const next = { ...value };
  if (value.formula !== undefined) {
    if (typeof value.formula !== 'string') rebaseConflict(`${label} formula is invalid`);
    next.formula = transformFormulaValue(value.formula, delta, ownerSheetId, context);
  }
  if (value.formulaMetadata !== undefined) {
    if (!isRecord(value.formulaMetadata)) rebaseConflict(`${label} formula metadata is invalid`);
    if (value.formulaMetadata.kind !== 'normal' || value.formulaMetadata.range !== undefined
      || value.formulaMetadata.preservedOnly === true) {
      rebaseConflict(`${label} contains formula-group metadata without a canonical rebase transform`);
    }
    if (value.formulaMetadata.sourceFormula !== undefined) {
      if (typeof value.formulaMetadata.sourceFormula !== 'string') rebaseConflict(`${label} source formula is invalid`);
      next.formulaMetadata = {
        ...value.formulaMetadata,
        sourceFormula: transformFormulaValue(value.formulaMetadata.sourceFormula, delta, ownerSheetId, context),
      };
    }
  }
  if (isRecord(value.presentation) && value.presentation.kind === 'barcode'
    && isRecord(value.presentation.source) && value.presentation.source.kind === 'formula') {
    const formula = value.presentation.source.formula;
    if (typeof formula !== 'string') rebaseConflict(`${label} barcode formula is invalid`);
    next.presentation = {
      ...value.presentation,
      source: { ...value.presentation.source, formula: transformFormulaValue(formula, delta, ownerSheetId, context) },
    };
  }
  return next;
}

function transformCellMutationFormulaOwners(
  mutationId: string,
  originalParams: unknown,
  transformedParams: unknown,
  delta: StructuralDelta,
  ownerSheetId: string,
  context: StructuralRebaseContext,
): unknown {
  if (!isRecord(originalParams) || !isRecord(transformedParams)) {
    rebaseConflict(`pending ${mutationId} parameters are invalid`);
  }
  const field = mutationId === 'cell.set' ? 'value' : 'previous';
  const sheetId = typeof originalParams.sheetId === 'string' ? originalParams.sheetId : ownerSheetId;
  const next = shiftCellCoordinate(transformedParams, sheetId, delta, `pending ${mutationId} address`);
  if (originalParams[field] !== undefined) {
    next[field] = transformCellFormulaOwners(originalParams[field], sheetId, delta, context, `pending ${mutationId} cell`);
  }
  if (mutationId === 'cell.set' && originalParams.writeAuthority !== undefined) {
    const originalAuthority = originalParams.writeAuthority;
    const transformedAuthority = transformedParams.writeAuthority;
    if (!isRecord(originalAuthority) || !isRecord(transformedAuthority)
      || !isRecord(originalAuthority.target) || !isRecord(transformedAuthority.target)
      || !isRecord(originalAuthority.candidate)) {
      rebaseConflict('pending cell.set write authority is invalid');
    }
    const writeAuthority = { ...transformedAuthority };
    writeAuthority.target = shiftCellCoordinate(
      transformedAuthority.target,
      sheetId,
      delta,
      'pending cell.set write authority target',
    );
    writeAuthority.candidate = transformCellFormulaOwners(
      originalAuthority.candidate,
      sheetId,
      delta,
      context,
      'pending cell.set write authority candidate',
    );
    next.writeAuthority = writeAuthority;
  }
  return next;
}

function transformRangeSetCoordinates(
  originalParams: unknown,
  transformedParams: unknown,
  delta: StructuralDelta,
  ownerSheetId: string,
  context: StructuralRebaseContext,
): unknown {
  if (!isRecord(originalParams) || !isRecord(transformedParams) || !Array.isArray(originalParams.values)) {
    rebaseConflict('pending range.set parameters are invalid');
  }
  const sheetId = typeof originalParams.sheetId === 'string' ? originalParams.sheetId : ownerSheetId;
  if (sheetId !== delta.sheetId) return transformedParams;
  const row = originalParams.startRow;
  const column = originalParams.startColumn;
  if (typeof row !== 'number' || typeof column !== 'number') rebaseConflict('pending range.set origin is invalid');
  const originCoordinate = deltaAxis(delta) === 'row' ? row : column;
  const rowCount = originalParams.values.length;
  const columnCount = originalParams.values.reduce((max, values) => {
    if (!Array.isArray(values)) rebaseConflict('pending range.set contains an invalid value row');
    return Math.max(max, values.length);
  }, 0);
  const extent = deltaAxis(delta) === 'row' ? rowCount : columnCount;
  if (extent > 0) {
    const end = originCoordinate + extent - 1;
    if (!Number.isSafeInteger(end)) rebaseConflict('pending range.set exceeds the safe integer range');
    if (deltaOperation(delta) === 'delete' && originCoordinate <= delta.at + delta.count - 1 && end >= delta.at) {
      rebaseConflict('committed deletion intersects pending range.set writes');
    }
    if (deltaOperation(delta) === 'insert' && delta.at > originCoordinate && delta.at <= end) {
      rebaseConflict('committed insertion splits pending range.set writes');
    }
  }
  const originKey = deltaAxis(delta) === 'row' ? 'startRow' : 'startColumn';
  const values = originalParams.values.map((rowValues) => {
    if (!Array.isArray(rowValues)) rebaseConflict('pending range.set contains an invalid value row');
    return rowValues.map((cell) => cell === undefined || cell === null
      ? cell
      : transformCellFormulaOwners(cell, sheetId, delta, context, 'pending range.set cell'));
  });
  return { ...transformedParams, [originKey]: shiftPoint(originCoordinate, delta), values };
}

function transformFillCoordinates(
  mutationId: string,
  originalParams: unknown,
  transformedParams: unknown,
  delta: StructuralDelta,
  ownerSheetId: string,
  context: StructuralRebaseContext,
): unknown {
  if (!isRecord(originalParams) || !isRecord(transformedParams)) {
    rebaseConflict(`pending ${mutationId} parameters are invalid`);
  }
  const originalWrites = originalParams.writes;
  const transformedWrites = transformedParams.writes;
  const sourceRange = originalParams.sourceRange;
  const targetRange = originalParams.targetRange;
  if (!Array.isArray(originalWrites) || !Array.isArray(transformedWrites)
    || !isRecord(sourceRange) || !isRange(sourceRange)
    || !isRecord(targetRange) || !isRange(targetRange)) {
    rebaseConflict(`pending ${mutationId} parameters are invalid`);
  }
  assertStructuralEditDoesNotIntersectRange(sourceRange, delta, `${mutationId} source range`);
  assertStructuralEditDoesNotIntersectRange(targetRange, delta, `${mutationId} target range`);
  const sheetId = typeof originalParams.sheetId === 'string' ? originalParams.sheetId : ownerSheetId;
  const writes = transformedWrites.map((write, index) => {
    const shifted = shiftCellCoordinate(write, sheetId, delta, `${mutationId} write`);
    const source = originalWrites[index];
    if (!isRecord(source)) rebaseConflict(`pending ${mutationId} write is invalid`);
    for (const key of ['before', 'after'] as const) {
      if (source[key] !== undefined) shifted[key] = transformCellFormulaOwners(source[key], sheetId, delta, context, `${mutationId} ${key} cell`);
    }
    return shifted;
  });
  return { ...transformedParams, writes };
}

function transformFindReplacementCoordinates(
  originalParams: unknown,
  transformedParams: unknown,
  delta: StructuralDelta,
  ownerSheetId: string,
  context: StructuralRebaseContext,
): unknown {
  if (!isRecord(originalParams) || !isRecord(transformedParams) || !Array.isArray(transformedParams.patches)) {
    rebaseConflict('pending find.replaced parameters are invalid');
  }
  const originalPatches = originalParams.patches;
  const transformedPatches = transformedParams.patches;
  if (!Array.isArray(originalPatches) || !Array.isArray(transformedPatches)) {
    rebaseConflict('pending find.replaced source patches are invalid');
  }
  const patches = transformedPatches.map((patch, index) => {
    if (!isRecord(patch) || !isRecord(patch.match)) rebaseConflict('pending find.replaced patch has no match address');
    const sourcePatch = originalPatches[index];
    if (!isRecord(sourcePatch) || !isRecord(sourcePatch.match) || typeof sourcePatch.match.sheetId !== 'string') {
      rebaseConflict('pending find.replaced source patch has no worksheet identity');
    }
    const match = shiftCellCoordinate(patch.match, ownerSheetId, delta, 'pending find.replaced match');
    if (match.sheetId === delta.sheetId) {
      if (typeof match.row !== 'number' || typeof match.column !== 'number' || typeof match.target !== 'string') {
        rebaseConflict('pending find.replaced match identity is invalid');
      }
      match.key = `${match.sheetId}!${match.row}:${match.column}:${match.target}:${typeof match.sourceId === 'string' ? match.sourceId : ''}`;
    }
    const nextPatch: Record<string, unknown> = { ...patch, match };
    for (const key of ['previous', 'next'] as const) {
      if (sourcePatch[key] !== undefined) {
        nextPatch[key] = transformCellFormulaOwners(sourcePatch[key], sourcePatch.match.sheetId, delta, context, `pending find.replaced ${key} cell`);
      }
    }
    return nextPatch;
  });
  return { ...transformedParams, patches };
}

function transformCommentAddCoordinates(originalParams: unknown, transformedParams: unknown, delta: StructuralDelta, ownerSheetId: string): unknown {
  if (!isRecord(originalParams) || !isRecord(transformedParams) || !isRecord(originalParams.thread) || !isRecord(transformedParams.thread)) {
    rebaseConflict('pending comment.add parameters are invalid');
  }
  const sheetId = typeof originalParams.sheetId === 'string' ? originalParams.sheetId : ownerSheetId;
  const envelope = shiftCellCoordinate(transformedParams, sheetId, delta, 'pending comment.add address');
  const thread = shiftCellCoordinate(transformedParams.thread, sheetId, delta, 'pending comment.add thread');
  return { ...envelope, thread };
}

function transformCommentUpdateCoordinates(originalParams: unknown, transformedParams: unknown, delta: StructuralDelta, ownerSheetId: string): unknown {
  if (!isRecord(originalParams) || !isRecord(transformedParams)) rebaseConflict('pending comment.update parameters are invalid');
  const sheetId = typeof originalParams.sheetId === 'string' ? originalParams.sheetId : ownerSheetId;
  return shiftCellCoordinate(transformedParams, sheetId, delta, 'pending comment.update address');
}

function transformReviewCellCoordinates(
  mutationId: string,
  originalParams: unknown,
  transformedParams: unknown,
  delta: StructuralDelta,
  ownerSheetId: string,
): unknown {
  if (!isRecord(originalParams) || !isRecord(transformedParams)) {
    rebaseConflict(`pending ${mutationId} parameters are invalid`);
  }
  const sheetId = typeof originalParams.sheetId === 'string' ? originalParams.sheetId : ownerSheetId;
  const next = shiftCellCoordinate(transformedParams, sheetId, delta, `pending ${mutationId} address`);
  if (mutationId === 'hyperlink.set' && originalParams.hyperlink !== undefined) {
    next.hyperlink = transformHyperlinkReference(originalParams.hyperlink, delta, 'pending hyperlink.set');
  }
  return next;
}

function transformVisibilityCoordinates(
  mutationId: string,
  originalParams: unknown,
  transformedParams: unknown,
  delta: StructuralDelta,
  ownerSheetId: string,
): unknown {
  if (!isRecord(originalParams) || !isRecord(transformedParams)) {
    rebaseConflict(`pending ${mutationId} states are invalid`);
  }
  const originalStates = originalParams.states;
  const transformedStates = transformedParams.states;
  if (!Array.isArray(originalStates) || !Array.isArray(transformedStates)
    || originalStates.length !== transformedStates.length) {
    rebaseConflict(`pending ${mutationId} states are invalid`);
  }
  const sheetId = typeof originalParams.sheetId === 'string' ? originalParams.sheetId : ownerSheetId;
  const axis = mutationId === 'rows.visibility' ? 'row' : 'column';
  if (sheetId !== delta.sheetId || axis !== deltaAxis(delta)) return transformedParams;
  const states = transformedStates.map((state, index) => {
    const originalState = originalStates[index];
    if (!isRecord(state) || !isRecord(originalState)) {
      rebaseConflict(`pending ${mutationId} contains an invalid ${axis} state`);
    }
    const coordinate = originalState[axis];
    if (typeof coordinate !== 'number') rebaseConflict(`pending ${mutationId} contains an invalid ${axis} state`);
    return { ...state, [axis]: shiftPoint(coordinate, delta) };
  });
  return { ...transformedParams, states };
}

function transformPasteTargetCoordinates(originalParams: unknown, transformedParams: unknown, delta: StructuralDelta, ownerSheetId: string): unknown {
  if (!isRecord(originalParams) || !isRecord(transformedParams) || !isRecord(originalParams.targetOrigin)) {
    rebaseConflict('pending range.paste parameters are invalid');
  }
  const sheetId = typeof originalParams.sheetId === 'string' ? originalParams.sheetId : ownerSheetId;
  if (originalParams.sourceExtent !== undefined || originalParams.clipboard !== undefined) {
    if (!isRecord(originalParams.sourceExtent) || !isRecord(originalParams.spec)
      || typeof originalParams.spec.transpose !== 'boolean') {
      rebaseConflict('pending range.paste has no canonical target extent');
    }
    const sourceRows = originalParams.sourceExtent.rows;
    const sourceColumns = originalParams.sourceExtent.columns;
    const rowCount = originalParams.spec.transpose ? sourceColumns : sourceRows;
    const columnCount = originalParams.spec.transpose ? sourceRows : sourceColumns;
    const row = originalParams.targetOrigin.row;
    const column = originalParams.targetOrigin.column;
    if (typeof sourceRows !== 'number' || typeof sourceColumns !== 'number'
      || typeof rowCount !== 'number' || typeof columnCount !== 'number'
      || typeof row !== 'number' || typeof column !== 'number'
      || ![sourceRows, sourceColumns, rowCount, columnCount, row, column].every(Number.isSafeInteger)
      || rowCount <= 0 || columnCount <= 0) {
      rebaseConflict('pending range.paste target extent is invalid');
    }
    const targetRange: RangeRef = {
      sheetId,
      startRow: row,
      endRow: row + rowCount - 1,
      startColumn: column,
      endColumn: column + columnCount - 1,
    };
    if (targetRange.endRow > MAX_ROW_INDEX || targetRange.endColumn > MAX_COLUMN_INDEX) {
      rebaseConflict('pending range.paste target exceeds worksheet bounds');
    }
    assertStructuralEditDoesNotIntersectRange(targetRange, delta, 'range.paste target');
  }
  const targetOrigin = shiftCellCoordinate(transformedParams.targetOrigin, sheetId, delta, 'pending range.paste target');
  return { ...transformedParams, targetOrigin };
}

function transformPendingStructuralCoordinates(
  pendingKind: CollaborationOperationKind,
  originalParams: unknown,
  transformedParams: unknown,
  delta: StructuralDelta,
  ownerSheetId: string,
): unknown {
  if (!STRUCTURAL_KINDS.has(pendingKind) || !isStructuralKind(pendingKind, deltaAxis(delta))) return transformedParams;
  if (!isRecord(originalParams) || !isRecord(transformedParams)) rebaseConflict('pending structural mutation parameters are invalid');
  const sheetId = typeof originalParams.sheetId === 'string' ? originalParams.sheetId : ownerSheetId;
  if (sheetId !== delta.sheetId) return transformedParams;
  const coordinateKey = typeof originalParams.at === 'number' ? 'at' : deltaAxis(delta) === 'row' ? 'row' : 'column';
  const coordinate = originalParams[coordinateKey];
  if (typeof coordinate !== 'number') rebaseConflict('pending structural mutation index is invalid');
  return { ...transformedParams, [coordinateKey]: shiftPoint(coordinate, delta) };
}

function transformKnownMutationCoordinates(
  mutationId: string,
  pendingKind: CollaborationOperationKind,
  originalParams: unknown,
  transformedParams: unknown,
  delta: StructuralDelta,
  ownerSheetId: string,
  context: StructuralRebaseContext,
): unknown {
  const structuralParams = transformPendingStructuralCoordinates(
    pendingKind, originalParams, transformedParams, delta, ownerSheetId,
  );
  if (structuralParams !== transformedParams) return structuralParams;
  if (mutationId === 'name.set') {
    return transformDefinedNameSetCoordinates(originalParams, delta, context);
  }
  if (mutationId === 'name.remove') return transformedParams;
  if (mutationId === 'range.set') {
    return transformRangeSetCoordinates(originalParams, transformedParams, delta, ownerSheetId, context);
  }
  if (mutationId === 'fill.applied' || mutationId === 'fill.restored') {
    return transformFillCoordinates(mutationId, originalParams, transformedParams, delta, ownerSheetId, context);
  }
  if (mutationId === 'find.replaced') {
    return transformFindReplacementCoordinates(originalParams, transformedParams, delta, ownerSheetId, context);
  }
  if (mutationId === 'cell.set' || mutationId === 'cell.restore') {
    return transformCellMutationFormulaOwners(mutationId, originalParams, transformedParams, delta, ownerSheetId, context);
  }
  if (mutationId === 'comment.add') {
    return transformCommentAddCoordinates(originalParams, transformedParams, delta, ownerSheetId);
  }
  if (mutationId === 'comment.update') {
    return transformCommentUpdateCoordinates(originalParams, transformedParams, delta, ownerSheetId);
  }
  if (mutationId === 'note.set' || mutationId === 'note.remove' || mutationId === 'note.visibility'
    || mutationId === 'hyperlink.set' || mutationId === 'hyperlink.remove') {
    return transformReviewCellCoordinates(mutationId, originalParams, transformedParams, delta, ownerSheetId);
  }
  if (mutationId === 'rows.visibility' || mutationId === 'columns.visibility') {
    return transformVisibilityCoordinates(mutationId, originalParams, transformedParams, delta, ownerSheetId);
  }
  if (mutationId === 'range.paste') {
    return transformPasteTargetCoordinates(originalParams, transformedParams, delta, ownerSheetId);
  }
  return transformedParams;
}

function hasUnqualifiedFormulaReference(node: FormulaAst): boolean {
  switch (node.type) {
    case 'cell-reference': return node.reference.sheetId === undefined;
    case 'range-reference':
      return node.start.reference.sheetId === undefined && node.end.reference.sheetId === undefined;
    case 'whole-column-reference':
    case 'whole-row-reference': return node.sheetId === undefined;
    case 'spill-reference': return hasUnqualifiedFormulaReference(node.operand);
    case 'reference-union': return node.references.some(hasUnqualifiedFormulaReference);
    case 'reference-intersection':
      return hasUnqualifiedFormulaReference(node.left) || hasUnqualifiedFormulaReference(node.right);
    case 'unary-expression': return hasUnqualifiedFormulaReference(node.operand);
    case 'binary-expression':
      return hasUnqualifiedFormulaReference(node.left) || hasUnqualifiedFormulaReference(node.right);
    case 'function-call': return node.arguments.some(hasUnqualifiedFormulaReference);
    case 'number-literal':
    case 'string-literal':
    case 'boolean-literal':
    case 'name-reference':
    case 'table-reference':
    case 'invalid-reference':
    case 'sheet-range-reference':
    case 'external-reference':
      return false;
  }
}

function transformDefinedNameSetCoordinates(
  params: unknown,
  delta: StructuralDelta,
  context: StructuralRebaseContext,
): unknown {
  if (!isRecord(params) || !isRecord(params.model)) rebaseConflict('pending name.set has no defined-name model');
  const model = params.model;
  if (typeof model.formula !== 'string' || model.formula.trim() === '') {
    rebaseConflict('pending name.set has no canonical formula');
  }
  const formulaValue = model.formula;
  if (model.scope !== 'workbook' && model.scope !== 'sheet') {
    rebaseConflict('pending name.set has an unknown defined-name scope');
  }

  let ownerSheetId: string | undefined;
  const originalAnchor = model.anchor;
  if (originalAnchor !== undefined) {
    if (!isRecord(originalAnchor) || typeof originalAnchor.sheetId !== 'string' || originalAnchor.sheetId.trim() === ''
      || !Number.isSafeInteger(originalAnchor.row) || !Number.isSafeInteger(originalAnchor.column)
      || (originalAnchor.row as number) < 0 || (originalAnchor.row as number) > MAX_ROW_INDEX
      || (originalAnchor.column as number) < 0 || (originalAnchor.column as number) > MAX_COLUMN_INDEX) {
      rebaseConflict('pending name.set has an invalid formula anchor');
    }
    ownerSheetId = originalAnchor.sheetId;
  } else if (model.scope === 'sheet') {
    if (typeof model.sheetId !== 'string' || model.sheetId.trim() === '') {
      rebaseConflict('pending sheet-scoped name.set has no worksheet identity');
    }
    ownerSheetId = model.sheetId;
  }

  let formulaOwnerSheetId = ownerSheetId;
  if (formulaOwnerSheetId === undefined) {
    try {
      const source = formulaValue.trimStart().startsWith('=') ? formulaValue : `=${formulaValue}`;
      if (hasUnqualifiedFormulaReference(parseFormula(source))) {
        rebaseConflict('workbook-scoped defined name has an unqualified reference but no formula anchor');
      }
    } catch (error) {
      if (error instanceof Error && error.message.startsWith('STRUCTURAL_REBASE_CONFLICT:')) throw error;
      rebaseConflict(`pending defined-name formula cannot be parsed: ${error instanceof Error ? error.message : String(error)}`);
    }
    // The empty identity keeps explicitly qualified references transformable
    // without assigning unresolved relative references to the primary sheet.
    formulaOwnerSheetId = '';
  }

  const formula = transformFormulaValue(formulaValue, delta, formulaOwnerSheetId, context);
  const anchor = originalAnchor === undefined
    ? undefined
    : shiftCellCoordinate(originalAnchor, ownerSheetId!, delta, 'pending name.set formula anchor');
  return {
    ...params,
    model: {
      ...model,
      formula,
      ...(anchor === undefined ? {} : { anchor }),
    },
  };
}

function isFormulaField(field: string, formula: string): boolean {
  const normalized = field.toLowerCase();
  return normalized === 'formula'
    || ((normalized === 'formula1' || normalized === 'formula2') && formula.trimStart().startsWith('='));
}

function isRuleFormulaField(field: string, formula: string, rule: Record<string, unknown>): boolean {
  const startsWithEquals = formula.trimStart().startsWith('=');
  if ((field === 'value1' || field === 'value2') && startsWithEquals) return true;
  if (field === 'value1' && rule.operator === 'formula') return true;
  if (field === 'formula1' && (rule.operator === 'formula' || rule.type === 'custom')) return true;
  return field === 'formula2' && rule.type === 'custom';
}

function transformFormulaValue(
  formula: string,
  delta: StructuralDelta,
  ownerSheetId: string,
  context: StructuralRebaseContext,
): string {
  const targetSheet = context.sheetOrder.find((sheet) => sheet.id === delta.sheetId);
  if (!targetSheet) rebaseConflict(`worksheet identity is missing for formula references on ${delta.sheetId}`);
  try {
    const hasPrefix = formula.trimStart().startsWith('=');
    const source = hasPrefix ? formula : `=${formula}`;
    const transformed = formatFormula(mapAstStructuralReferences(parseFormula(source), {
      shift: { axis: deltaAxis(delta), at: delta.at, count: delta.count, op: deltaOperation(delta) },
      ownerSheetId,
      targetSheetId: delta.sheetId,
      targetSheetName: targetSheet.name,
      sheetOrder: context.sheetOrder,
    }));
    return hasPrefix ? transformed : transformed.replace(/^=/, '');
  } catch (error) {
    rebaseConflict(`pending formula cannot be structurally transformed: ${error instanceof Error ? error.message : String(error)}`);
  }
}

function shiftSnapshotCellKey(value: unknown, sheetId: string, delta: StructuralDelta): string {
  if (typeof value !== 'string') rebaseConflict('paste snapshot contains a non-string cell metadata key');
  const match = value.match(/^(\d+):(\d+)$/);
  if (!match) rebaseConflict('paste snapshot contains an invalid cell metadata key');
  if (sheetId !== delta.sheetId) return value;
  const row = Number(match[1]);
  const column = Number(match[2]);
  return `${deltaAxis(delta) === 'row' ? shiftPoint(row, delta) : row}:${deltaAxis(delta) === 'column' ? shiftPoint(column, delta) : column}`;
}

function shiftSnapshotAddress(value: unknown, sheetId: string, delta: StructuralDelta): void {
  if (!isRecord(value)) rebaseConflict('paste snapshot contains an invalid cell coordinate');
  const ownerSheetId = typeof value.sheetId === 'string' ? value.sheetId : sheetId;
  if (ownerSheetId !== delta.sheetId) return;
  const coordinateKey = deltaAxis(delta) === 'row' ? 'row' : 'column';
  const coordinate = value[coordinateKey];
  if (typeof coordinate !== 'number') rebaseConflict(`paste snapshot ${coordinateKey} coordinate is invalid`);
  value[coordinateKey] = shiftPoint(coordinate, delta);
}

function transformHyperlinkReference(value: unknown, delta: StructuralDelta, label: string): unknown {
  if (!isRecord(value) || !isRecord(value.target)) rebaseConflict(`${label} contains an invalid hyperlink`);
  const target = value.target;
  if (target.kind !== 'sheet' || target.sheetId !== delta.sheetId) return value;
  let row: number;
  let column: number;
  if (typeof target.address === 'string' && target.row === undefined && target.column === undefined) {
    const parsed = parseAddress(target.address);
    if (!parsed) rebaseConflict(`${label} contains an invalid worksheet hyperlink address`);
    row = parsed.row;
    column = parsed.column;
  } else if (target.address === undefined && typeof target.row === 'number' && typeof target.column === 'number') {
    row = target.row;
    column = target.column;
  } else {
    rebaseConflict(`${label} contains a non-canonical worksheet hyperlink target`);
  }
  if (deltaAxis(delta) === 'row') row = shiftPoint(row, delta);
  else column = shiftPoint(column, delta);
  const nextTarget = target.address === undefined
    ? { ...target, row, column }
    : { ...target, address: cellAddress(row, column) };
  return { ...value, target: nextTarget };
}

function shiftPasteSnapshot(
  originalValue: unknown,
  transformedValue: unknown,
  sheetId: string,
  delta: StructuralDelta,
  context: StructuralRebaseContext,
): unknown {
  if (!isRecord(originalValue) || !isRecord(transformedValue)) rebaseConflict('pending range.paste has an invalid snapshot');
  const originalSnapshot = originalValue;
  const snapshot = transformedValue;
  const originalCells = originalSnapshot.cells;
  const cells = snapshot.cells;
  if (!Array.isArray(originalCells) || !Array.isArray(cells) || originalCells.length !== cells.length) {
    rebaseConflict('pending range.paste snapshot cells are invalid');
  }
  for (let index = 0; index < cells.length; index += 1) {
    const originalCell = originalCells[index];
    const cell = cells[index];
    if (!isRecord(originalCell) || !isRecord(cell)) rebaseConflict('pending range.paste snapshot cell is invalid');
    shiftSnapshotAddress(cell, sheetId, delta);
    if (originalCell.value !== undefined) {
      const formulaOwnerSheetId = typeof originalCell.sheetId === 'string' ? originalCell.sheetId : sheetId;
      cell.value = transformCellFormulaOwners(
        originalCell.value,
        formulaOwnerSheetId,
        delta,
        context,
        'pending range.paste snapshot cell value',
      );
    }
  }
  for (const field of ['notes', 'hyperlinks'] as const) {
    const entries = snapshot[field];
    if (entries === undefined) continue;
    if (!Array.isArray(entries)) rebaseConflict(`pending range.paste snapshot ${field} are invalid`);
    for (const entry of entries) {
      if (!isRecord(entry)) rebaseConflict(`pending range.paste snapshot ${field} entry is invalid`);
      entry.key = shiftSnapshotCellKey(entry.key, sheetId, delta);
      if (field === 'hyperlinks' && entry.value !== undefined) {
        entry.value = transformHyperlinkReference(entry.value, delta, 'paste snapshot');
      }
    }
  }
  if (snapshot.commentCells !== undefined) {
    if (!Array.isArray(snapshot.commentCells)) rebaseConflict('pending range.paste snapshot commentCells are invalid');
    snapshot.commentCells = snapshot.commentCells.map((key) => shiftSnapshotCellKey(key, sheetId, delta));
  }
  if (snapshot.comments !== undefined) {
    if (!Array.isArray(snapshot.comments)) rebaseConflict('pending range.paste snapshot comments are invalid');
    for (const comment of snapshot.comments) shiftSnapshotAddress(comment, sheetId, delta);
  }
  for (const field of ['validations', 'conditionalFormats'] as const) {
    const rules = snapshot[field];
    if (rules === undefined) continue;
    if (!Array.isArray(rules)) rebaseConflict(`pending range.paste snapshot ${field} are invalid`);
    for (const rule of rules) {
      if (!isRecord(rule)) rebaseConflict(`pending range.paste snapshot ${field} rule is invalid`);
      if (rule.formulaAnchor !== undefined) {
        rule.formulaAnchor = shiftCellCoordinate(rule.formulaAnchor, sheetId, delta, `pending range.paste ${field} formula anchor`);
      }
    }
  }
  if (snapshot.columnWidths !== undefined) {
    if (!Array.isArray(snapshot.columnWidths)) rebaseConflict('pending range.paste snapshot columnWidths are invalid');
    if (deltaAxis(delta) === 'column' && sheetId === delta.sheetId) {
      for (const entry of snapshot.columnWidths) {
        if (!isRecord(entry) || typeof entry.column !== 'number') rebaseConflict('pending range.paste snapshot column width is invalid');
        entry.column = shiftPoint(entry.column, delta);
      }
    }
  }
  return snapshot;
}

function transformPasteSnapshots(
  params: unknown,
  transformedParams: unknown,
  ownerSheetId: string,
  delta: StructuralDelta,
  context: StructuralRebaseContext,
): unknown {
  if (!isRecord(params) || !isRecord(transformedParams)) rebaseConflict('pending range.paste parameters are invalid');
  const result = transformedParams;
  if (Object.prototype.hasOwnProperty.call(params, 'snapshot')) {
    result.snapshot = shiftPasteSnapshot(params.snapshot, result.snapshot, ownerSheetId, delta, context);
  }
  if (Object.prototype.hasOwnProperty.call(params, 'clipboard')) {
    result.clipboard = structuredClone(params.clipboard);
  }
  return result;
}

function isStructuralKind(kind: CollaborationOperationKind, axis: StructuralAxis): boolean {
  return kind === `${kind.startsWith('insert-') ? 'insert' : 'delete'}-${axis}s`;
}

function extractStructuralDelta(mutation: ClassifiedMutation): StructuralDelta | undefined {
  const p = mutation.params as { at?: number; count?: number; row?: number; column?: number } | null;
  if (!p) return undefined;
  const at = p.at ?? p.row ?? p.column;
  const count = p.count ?? 1;
  if (at == null) return undefined;

  if (!Number.isSafeInteger(at) || at < 0 || !Number.isSafeInteger(count) || count < 1) {
    rebaseConflict(`committed ${mutation.mutationId} has invalid structural bounds`);
  }
  if (!Number.isSafeInteger(at + count)) rebaseConflict(`committed ${mutation.mutationId} exceeds the safe integer range`);
  const maximum = mutation.kind.endsWith('rows') ? MAX_ROW_INDEX : MAX_COLUMN_INDEX;
  if (at > maximum || at + count > maximum + 1) {
    rebaseConflict(`committed ${mutation.mutationId} exceeds worksheet bounds`);
  }
  switch (mutation.kind) {
    case 'insert-rows': return { kind: 'insert-rows', sheetId: mutation.sheetId, at, count };
    case 'delete-rows': return { kind: 'delete-rows', sheetId: mutation.sheetId, at, count };
    case 'insert-columns': return { kind: 'insert-columns', sheetId: mutation.sheetId, at, count };
    case 'delete-columns': return { kind: 'delete-columns', sheetId: mutation.sheetId, at, count };
    default: return undefined;
  }
}

const STRUCTURAL_KINDS = new Set<CollaborationOperationKind>([
  'insert-rows', 'delete-rows', 'insert-columns', 'delete-columns',
]);

/** 将 pending 操作按已提交的结构变更 rebase — 例: A 插第 5 行，B 改 A10 → A11 */
export function rebaseMutation(
  pending: ClassifiedMutation,
  committed: ClassifiedMutation,
  context: StructuralRebaseContext = { sheetOrder: [] },
): RebaseResult {
  if (committed.kind === 'unknown') {
    rebaseConflict(`committed ${committed.mutationId} has no registered structural transform`);
  }
  const delta = extractStructuralDelta(committed);
  if (!delta) {
    if (STRUCTURAL_KINDS.has(committed.kind)) rebaseConflict(`committed ${committed.mutationId} has no structural bounds`);
    if (UNSUPPORTED_STRUCTURAL_KINDS.has(committed.kind)) {
      rebaseConflict(`committed ${committed.mutationId} has no canonical structural patch`);
    }
    return { rebased: pending, transformed: false };
  }
  if (!STRUCTURAL_KINDS.has(committed.kind)) {
    return { rebased: pending, transformed: false };
  }
  if (pending.kind === 'unknown') {
    rebaseConflict(`Cannot rebase unknown mutation ${pending.mutationId} across ${committed.mutationId}`);
  }
  if (UNSUPPORTED_STRUCTURAL_KINDS.has(pending.kind)) {
    rebaseConflict(`pending ${pending.mutationId} has no canonical structural patch for ${committed.mutationId}`);
  }

  const rebasedRanges = pending.affectedRanges.map((range) => {
    return shiftRange(range, delta);
  });

  const paramsForTransform = isRecord(pending.params) ? { ...pending.params } : pending.params;
  if (pending.mutationId === 'range.paste' && isRecord(paramsForTransform)) delete paramsForTransform.clipboard;
  const transformedParams = pending.mutationId === 'name.set'
    ? paramsForTransform
    : transformParams(paramsForTransform, delta, pending.sheetId, context);
  const mappedParams = transformKnownMutationCoordinates(pending.mutationId, pending.kind, pending.params, transformedParams, delta, pending.sheetId, context);
  const rebasedParams = pending.mutationId === 'range.paste'
    ? transformPasteSnapshots(pending.params, mappedParams, pending.sheetId, delta, context)
    : mappedParams;

  return {
    rebased: { ...pending, affectedRanges: rebasedRanges, params: rebasedParams },
    transformed: true,
  };
}

/** 按 revision 顺序依次 rebase 一批 pending 操作 */
export function rebaseAgainstHistory(
  pending: ClassifiedMutation,
  committedHistory: ClassifiedMutation[],
  context: StructuralRebaseContext = { sheetOrder: [] },
): RebaseResult {
  let current = pending;
  let transformed = false;
  for (const committed of committedHistory) {
    const result = rebaseMutation(current, committed, context);
    current = result.rebased;
    transformed = transformed || result.transformed;
  }
  return { rebased: current, transformed };
}
