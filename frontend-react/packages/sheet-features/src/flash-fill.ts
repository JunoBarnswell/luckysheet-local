import type { CellData, RangeRef, WorksheetModel } from '@react-sheets/core-model';

export interface FlashFillParams {
  sheetId: string;
  sourceRange: RangeRef;
  targetRange: RangeRef;
}

export type FlashFillOperation =
  | { kind: 'identity' }
  | { kind: 'case'; mode: 'upper' | 'lower' | 'capitalize' }
  | { kind: 'trim' }
  | { kind: 'substring'; start: number; length: number }
  | { kind: 'token'; delimiter: string; index: number }
  | { kind: 'prefix'; value: string }
  | { kind: 'suffix'; value: string }
  | { kind: 'numeric-delta'; delta: number };

export interface FlashFillWrite {
  readonly row: number;
  readonly column: number;
  readonly before?: CellData;
  readonly after?: CellData;
}

export interface FlashFillPlan {
  readonly sheetId: string;
  readonly sourceRange: RangeRef;
  readonly targetRange: RangeRef;
  readonly operation: FlashFillOperation;
  readonly writes: readonly FlashFillWrite[];
}

const MAX_FLASH_FILL_CELLS = 100_000;
const TOKEN_DELIMITERS = [' ', ',', ';', '-', '_', '/', '\\', '@', '.'] as const;

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null;
}

function isFiniteInt(value: unknown): value is number {
  return Number.isSafeInteger(value) && Number(value) >= 0;
}

function isRange(value: unknown): value is RangeRef {
  if (!isRecord(value)) return false;
  return typeof value.sheetId === 'string'
    && isFiniteInt(value.startRow) && isFiniteInt(value.endRow)
    && isFiniteInt(value.startColumn) && isFiniteInt(value.endColumn)
    && value.endRow >= value.startRow && value.endColumn >= value.startColumn;
}

function normalizeRange(range: RangeRef, sheetId: string, name: string): RangeRef {
  if (!isRange(range) || range.sheetId !== sheetId) throw new Error(`Flash Fill ${name} must target the command sheet`);
  return {
    sheetId,
    startRow: Math.min(range.startRow, range.endRow),
    endRow: Math.max(range.startRow, range.endRow),
    startColumn: Math.min(range.startColumn, range.endColumn),
    endColumn: Math.max(range.startColumn, range.endColumn),
  };
}

function rangeHeight(range: RangeRef): number {
  return range.endRow - range.startRow + 1;
}

function rangeWidth(range: RangeRef): number {
  return range.endColumn - range.startColumn + 1;
}

function rangesEqual(left: RangeRef, right: RangeRef): boolean {
  return left.sheetId === right.sheetId
    && left.startRow === right.startRow && left.endRow === right.endRow
    && left.startColumn === right.startColumn && left.endColumn === right.endColumn;
}

function assertBounds(sheet: WorksheetModel, range: RangeRef, name: string): void {
  if (range.startRow < 0 || range.startColumn < 0 || range.endRow >= sheet.rowCount || range.endColumn >= sheet.columnCount) {
    throw new Error(`Flash Fill ${name} is outside worksheet bounds`);
  }
  if (rangeHeight(range) * rangeWidth(range) > MAX_FLASH_FILL_CELLS) throw new Error(`Flash Fill ${name} is too large`);
}

function cellAt(sheet: WorksheetModel, row: number, column: number): CellData | undefined {
  const cell = sheet.cells.get(row, column);
  return cell === undefined ? undefined : structuredClone(cell);
}

function isBlank(cell: CellData | undefined): boolean {
  return cell === undefined || (cell.value === null && cell.formula === undefined && cell.formulaValue === undefined);
}

function scalarText(cell: CellData | undefined, name: string): string | number | undefined {
  if (cell === undefined) return undefined;
  if (cell.formula !== undefined || cell.formulaValue !== undefined) throw new Error(`Flash Fill ${name} cannot use formula cells as examples`);
  if (cell.value === null || cell.value === undefined) return undefined;
  if (typeof cell.value !== 'string' && typeof cell.value !== 'number' && typeof cell.value !== 'boolean') {
    throw new Error(`Flash Fill ${name} contains an unsupported value`);
  }
  return typeof cell.value === 'boolean' ? String(cell.value) : cell.value;
}

interface Sample {
  readonly source: string | number;
  readonly target: string | number;
}

function textValue(value: string | number): string {
  return String(value);
}

function sameNumber(left: number, right: number): boolean {
  const tolerance = Number.EPSILON * Math.max(1, Math.abs(left), Math.abs(right)) * 32;
  return Math.abs(left - right) <= tolerance;
}

function capitalize(value: string): string {
  return value.replace(/(^|[\\s_-])([a-z])/gi, (_match, prefix: string, character: string) => `${prefix}${character.toLocaleUpperCase()}`);
}

function applyOperation(operation: FlashFillOperation, source: string | number): string | number {
  if (operation.kind === 'numeric-delta') return Number(source) + operation.delta;
  const value = textValue(source);
  switch (operation.kind) {
    case 'identity': return value;
    case 'case': return operation.mode === 'upper' ? value.toLocaleUpperCase() : operation.mode === 'lower' ? value.toLocaleLowerCase() : capitalize(value);
    case 'trim': return value.trim();
    case 'substring': return Array.from(value).slice(operation.start, operation.start + operation.length).join('');
    case 'token': return value.split(operation.delimiter)[operation.index] ?? '';
    case 'prefix': return `${operation.value}${value}`;
    case 'suffix': return `${value}${operation.value}`;
  }
}

function matchesTarget(operation: FlashFillOperation, sample: Sample): boolean {
  const expected = applyOperation(operation, sample.source);
  if (typeof expected === 'number' && typeof sample.target === 'number') return sameNumber(expected, sample.target);
  return String(expected) === String(sample.target);
}

function substringCandidates(source: string, target: string): FlashFillOperation[] {
  const candidates: FlashFillOperation[] = [];
  const targetChars = Array.from(target);
  if (targetChars.length === 0) return candidates;
  let index = source.indexOf(target);
  while (index >= 0) {
    const start = Array.from(source.slice(0, index)).length;
    candidates.push({ kind: 'substring', start, length: targetChars.length });
    index = source.indexOf(target, index + 1);
  }
  return candidates;
}

function inferOperation(samples: readonly Sample[]): FlashFillOperation {
  if (samples.length < 2) throw new Error('Flash Fill needs at least two non-blank examples to infer a pattern');
  const first = samples[0]!;
  const candidates: FlashFillOperation[] = [
    { kind: 'identity' },
    { kind: 'case', mode: 'upper' },
    { kind: 'case', mode: 'lower' },
    { kind: 'case', mode: 'capitalize' },
    { kind: 'trim' },
  ];
  if (typeof first.source === 'number' && typeof first.target === 'number') {
    const delta = first.target - first.source;
    candidates.push({ kind: 'numeric-delta', delta });
  }
  const firstSourceText = textValue(first.source);
  const firstTargetText = textValue(first.target);
  for (const candidate of substringCandidates(firstSourceText, firstTargetText)) candidates.push(candidate);
  if (firstTargetText.endsWith(firstSourceText)) candidates.push({ kind: 'prefix', value: firstTargetText.slice(0, -firstSourceText.length) });
  if (firstTargetText.startsWith(firstSourceText)) candidates.push({ kind: 'suffix', value: firstTargetText.slice(firstSourceText.length) });
  for (const delimiter of TOKEN_DELIMITERS) {
    const tokens = firstSourceText.split(delimiter);
    tokens.forEach((token, index) => {
      if (token === firstTargetText) candidates.push({ kind: 'token', delimiter, index });
    });
  }
  const unique = new Map<string, FlashFillOperation>();
  for (const candidate of candidates) unique.set(JSON.stringify(candidate), candidate);
  for (const candidate of unique.values()) if (samples.every((sample) => matchesTarget(candidate, sample))) return candidate;
  throw new Error('Flash Fill could not infer one deterministic pattern from the supplied examples');
}

function outputCell(existing: CellData | undefined, value: string | number): CellData {
  const next = existing === undefined ? { value: value as CellData['value'] } : structuredClone(existing);
  next.value = value as CellData['value'];
  delete next.formula;
  delete next.formulaValue;
  delete next.formulaMetadata;
  delete next.displayValue;
  delete next.richText;
  return next;
}

function validateGeometry(sheet: WorksheetModel, params: FlashFillParams): { sourceRange: RangeRef; targetRange: RangeRef } {
  if (params.sheetId !== sheet.id) throw new Error('Flash Fill sheet mismatch');
  const sourceRange = normalizeRange(params.sourceRange, params.sheetId, 'source range');
  const targetRange = normalizeRange(params.targetRange, params.sheetId, 'target range');
  assertBounds(sheet, sourceRange, 'source range');
  assertBounds(sheet, targetRange, 'target range');
  if (rangeHeight(sourceRange) !== rangeHeight(targetRange) || rangeWidth(sourceRange) !== rangeWidth(targetRange)) {
    throw new Error('Flash Fill source and target ranges must have identical dimensions');
  }
  if (rangesEqual(sourceRange, targetRange)) throw new Error('Flash Fill source and target ranges must be different');
  if (sourceRange.startRow <= targetRange.endRow && sourceRange.endRow >= targetRange.startRow
    && sourceRange.startColumn <= targetRange.endColumn && sourceRange.endColumn >= targetRange.startColumn) {
    throw new Error('Flash Fill source and target ranges cannot overlap');
  }
  return { sourceRange, targetRange };
}

/** Build the one side-effect-free Flash Fill plan used by command and replay. */
export function planFlashFill(sheet: WorksheetModel, params: FlashFillParams): FlashFillPlan {
  const { sourceRange, targetRange } = validateGeometry(sheet, params);
  const samples: Sample[] = [];
  for (let rowOffset = 0; rowOffset < rangeHeight(sourceRange); rowOffset += 1) {
    for (let columnOffset = 0; columnOffset < rangeWidth(sourceRange); columnOffset += 1) {
      const source = cellAt(sheet, sourceRange.startRow + rowOffset, sourceRange.startColumn + columnOffset);
      const target = cellAt(sheet, targetRange.startRow + rowOffset, targetRange.startColumn + columnOffset);
      const sourceValue = scalarText(source, 'source range');
      const targetValue = scalarText(target, 'target range');
      if (targetValue !== undefined) {
        if (sourceValue === undefined) throw new Error('Flash Fill example has a blank source value');
        samples.push({ source: sourceValue, target: targetValue });
      }
    }
  }
  const operation = inferOperation(samples);
  const writes: FlashFillWrite[] = [];
  for (let rowOffset = 0; rowOffset < rangeHeight(targetRange); rowOffset += 1) {
    for (let columnOffset = 0; columnOffset < rangeWidth(targetRange); columnOffset += 1) {
      const row = targetRange.startRow + rowOffset;
      const column = targetRange.startColumn + columnOffset;
      const current = cellAt(sheet, row, column);
      if (!isBlank(current)) continue;
      const source = cellAt(sheet, sourceRange.startRow + rowOffset, sourceRange.startColumn + columnOffset);
      const sourceValue = scalarText(source, 'source range');
      if (sourceValue === undefined) throw new Error('Flash Fill target contains a blank source value');
      const after = outputCell(current, applyOperation(operation, sourceValue));
      writes.push({ row, column, before: current, after });
    }
  }
  return { sheetId: sheet.id, sourceRange, targetRange, operation, writes };
}

export function isFlashFillOperation(value: unknown): value is FlashFillOperation {
  if (!isRecord(value) || typeof value.kind !== 'string') return false;
  if (value.kind === 'identity' || value.kind === 'trim') return true;
  if (value.kind === 'case') return value.mode === 'upper' || value.mode === 'lower' || value.mode === 'capitalize';
  if (value.kind === 'substring') return isFiniteInt(value.start) && isFiniteInt(value.length);
  if (value.kind === 'token') return typeof value.delimiter === 'string' && value.delimiter.length > 0 && isFiniteInt(value.index);
  if (value.kind === 'prefix' || value.kind === 'suffix') return typeof value.value === 'string';
  return value.kind === 'numeric-delta' && typeof value.delta === 'number' && Number.isFinite(value.delta);
}

export function isFlashFillParams(value: unknown): value is FlashFillParams {
  return isRecord(value) && typeof value.sheetId === 'string' && isRange(value.sourceRange) && isRange(value.targetRange);
}

export function isFlashFillWrite(value: unknown): value is FlashFillWrite {
  const validCell = (cell: unknown): cell is CellData | undefined => cell === undefined || (isRecord(cell) && 'value' in cell);
  return isRecord(value) && isFiniteInt(value.row) && isFiniteInt(value.column) && validCell(value.before) && validCell(value.after);
}

export function isFlashFillPlan(value: unknown): value is FlashFillPlan {
  return isRecord(value) && typeof value.sheetId === 'string' && isRange(value.sourceRange) && isRange(value.targetRange)
    && isFlashFillOperation(value.operation) && Array.isArray(value.writes) && value.writes.every(isFlashFillWrite);
}
