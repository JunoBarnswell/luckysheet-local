import type {
  PivotFieldDataType,
  PivotScalar,
  PivotSourceRowPath,
} from '@react-sheets/core-model';
import {
  createPivotMemberKey,
  isPivotError,
  pivotTimelineInstant,
  pivotMemberKey,
} from '@react-sheets/core-model';

export const PIVOT_SOURCE_INDEX_SCHEMA = 'PivotSourceIndex' as const;
export const PIVOT_SOURCE_INDEX_VERSION = 1 as const;
const MAX_UINT32 = 0xffffffff;

export interface PivotSourceFieldInput {
  fieldId: string;
  name: string;
  ordinal: number;
  dataType?: PivotFieldDataType;
}

export interface PivotSourceColumnValues {
  field: PivotSourceFieldInput;
  values: readonly PivotScalar[];
}

export interface PivotDictionaryColumn {
  kind: 'dictionary';
  dictionary: PivotScalar[];
  /** Zero is blank; every non-zero code addresses dictionary[code - 1]. */
  codes: Uint32Array;
}

export interface PivotNumberColumn {
  kind: 'number';
  values: Float64Array;
  validity: Uint8Array;
}

export interface PivotBooleanColumn {
  kind: 'boolean';
  values: Uint8Array;
  validity: Uint8Array;
}

export type PivotSourceColumn = PivotDictionaryColumn | PivotNumberColumn | PivotBooleanColumn;

/**
 * The only worksheet/data-source derivative consumed by Pivot calculation.
 * Field order owns column order; row values never duplicate field-id strings.
 */
export interface PivotSourceIndex {
  schema: typeof PIVOT_SOURCE_INDEX_SCHEMA;
  version: typeof PIVOT_SOURCE_INDEX_VERSION;
  rowCount: number;
  fields: PivotSourceFieldInput[];
  columns: PivotSourceColumn[];
  rowPathPool: PivotSourceRowPath[];
  rowPathOffsets: Uint32Array;
}

export type PivotSourceIndexBuildInput = {
  columns: readonly PivotSourceColumnValues[];
  rowPaths: readonly (readonly PivotSourceRowPath[])[];
  rowCount?: never;
  rowPathAt?: never;
} | {
  columns: readonly PivotSourceColumnValues[];
  rowCount: number;
  rowPathAt: (row: number) => readonly PivotSourceRowPath[];
  rowPaths?: never;
};

export function inferPivotSourceFieldType(values: readonly PivotScalar[]): PivotFieldDataType {
  const present = values.filter((value) => value != null && value !== '');
  if (!present.length && values.some((value) => typeof value === 'string') && values.every((value) => value == null || typeof value === 'string')) return 'text';
  if (!present.length) return 'mixed';
  if (present.every(isPivotError)) return 'error';
  if (present.some(isPivotError)) return 'mixed';
  if (present.every((value) => typeof value === 'boolean')) return 'boolean';
  if (present.every((value) => typeof value === 'number' && Number.isFinite(value))) return 'number';
  const dateLike = present.every((value) => typeof value === 'string'
    && /^\d{4}-\d{2}-\d{2}(?:[T ].*)?$/.test(value)
    && !Number.isNaN(Date.parse(value)));
  if (dateLike) return 'date';
  if (present.every((value) => typeof value === 'string')) return 'text';
  return 'mixed';
}

export function createPivotSourceIndex(input: PivotSourceIndexBuildInput): PivotSourceIndex {
  const rowCount = input.rowPaths === undefined ? input.rowCount : input.rowPaths.length;
  if (!Number.isSafeInteger(rowCount) || rowCount < 0) throw new Error('Pivot source row count is invalid');
  if (rowCount >= MAX_UINT32) throw new Error('Pivot source row count exceeds the transferable index limit');
  const fieldIds = new Set<string>();
  const fields = input.columns.map(({ field }, ordinal) => {
    if (!field.fieldId.trim()) throw new Error(`Pivot source field ${String(ordinal)} has no stable fieldId`);
    if (fieldIds.has(field.fieldId)) throw new Error(`Pivot source fieldId is duplicated: ${field.fieldId}`);
    if (field.ordinal !== ordinal) throw new Error(`Pivot source field ${field.fieldId} has a non-contiguous ordinal`);
    fieldIds.add(field.fieldId);
    return { ...field };
  });
  const columns = input.columns.map(({ field, values }, ordinal) => {
    if (values.length !== rowCount) {
      throw new Error(`Pivot source field ${field.fieldId} has ${String(values.length)} rows; expected ${String(rowCount)}`);
    }
    const inferred = inferPivotSourceFieldType(values);
    const dataType = field.dataType ?? inferred;
    validateDeclaredType(field.fieldId, dataType, values);
    fields[ordinal]!.dataType = dataType;
    return encodeColumn(dataType, values);
  });
  const rowPathPool: PivotSourceRowPath[] = [];
  const rowPathOffsets = new Uint32Array(rowCount + 1);
  for (let row = 0; row < rowCount; row += 1) {
    rowPathOffsets[row] = rowPathPool.length;
    const paths = input.rowPaths === undefined ? input.rowPathAt(row) : input.rowPaths[row]!;
    for (const path of paths) rowPathPool.push(structuredClone(path));
  }
  rowPathOffsets[rowCount] = rowPathPool.length;
  return {
    schema: PIVOT_SOURCE_INDEX_SCHEMA,
    version: PIVOT_SOURCE_INDEX_VERSION,
    rowCount,
    fields,
    columns,
    rowPathPool,
    rowPathOffsets,
  };
}

export function assertPivotSourceIndex(index: PivotSourceIndex): void {
  if (index.schema !== PIVOT_SOURCE_INDEX_SCHEMA || index.version !== PIVOT_SOURCE_INDEX_VERSION) {
    throw new Error('Pivot source index protocol is invalid');
  }
  if (!Number.isSafeInteger(index.rowCount) || index.rowCount < 0) throw new Error('Pivot source index rowCount is invalid');
  if (index.rowCount >= MAX_UINT32) throw new Error('Pivot source index rowCount exceeds the transferable index limit');
  if (!Array.isArray(index.fields) || !Array.isArray(index.columns) || !Array.isArray(index.rowPathPool)) throw new Error('Pivot source index collections are invalid');
  if (!(index.rowPathOffsets instanceof Uint32Array)) throw new Error('Pivot source index row-path offsets are invalid');
  if (index.fields.length !== index.columns.length) throw new Error('Pivot source index field/column count mismatch');
  if (index.rowPathOffsets.length !== index.rowCount + 1) throw new Error('Pivot source index row-path offsets are invalid');
  if (index.rowPathOffsets[0] !== 0) throw new Error('Pivot source index row-path offsets must start at zero');
  if (index.rowPathOffsets[index.rowCount] !== index.rowPathPool.length) throw new Error('Pivot source index row-path pool is invalid');
  if (index.rowPathPool.length > MAX_UINT32) throw new Error('Pivot source index row-path pool exceeds the transferable limit');
  for (let row = 0; row < index.rowCount; row += 1) {
    const start = index.rowPathOffsets[row]!;
    const end = index.rowPathOffsets[row + 1]!;
    if (start > end || end > index.rowPathPool.length) throw new Error(`Pivot source index row-path offsets are invalid at row ${String(row)}`);
  }
  const fieldTypes = new Set<PivotFieldDataType>(['number', 'boolean', 'text', 'date', 'mixed', 'error']);
  const fieldIds = new Set<string>();
  index.fields.forEach((field, ordinal) => {
    if (typeof field.fieldId !== 'string' || field.fieldId.trim() === '' || fieldIds.has(field.fieldId)) throw new Error(`Pivot source field ${field.fieldId} is invalid`);
    fieldIds.add(field.fieldId);
    if (typeof field.name !== 'string' || field.dataType === undefined || !fieldTypes.has(field.dataType)) throw new Error(`Pivot source field ${field.fieldId} metadata is invalid`);
    if (field.ordinal !== ordinal) throw new Error(`Pivot source field ${field.fieldId} has a non-contiguous ordinal`);
    const column = index.columns[ordinal]!;
    if (!column || !['dictionary', 'number', 'boolean'].includes(column.kind)) throw new Error(`Pivot source column ${field.fieldId} kind is invalid`);
    const length = column.kind === 'dictionary' ? column.codes.length : column.values.length;
    if (length !== index.rowCount) throw new Error(`Pivot source column ${field.fieldId} row count mismatch`);
    if (column.kind === 'dictionary') {
      if (!(column.codes instanceof Uint32Array) || !Array.isArray(column.dictionary)) throw new Error(`Pivot source dictionary column ${field.fieldId} is invalid`);
      for (const code of column.codes) if (code > column.dictionary.length) throw new Error(`Pivot source dictionary code is out of bounds for ${field.fieldId}`);
    } else {
      if (!(column.values instanceof (column.kind === 'number' ? Float64Array : Uint8Array))
        || !(column.validity instanceof Uint8Array)
        || column.validity.length !== index.rowCount) throw new Error(`Pivot source typed column ${field.fieldId} is invalid`);
      for (let row = 0; row < index.rowCount; row += 1) {
        const valid = column.validity[row]!;
        if (valid !== 0 && valid !== 1) throw new Error(`Pivot source validity bitmap is invalid for ${field.fieldId}`);
        if (column.kind === 'number' && valid === 1 && !Number.isFinite(column.values[row]!)) throw new Error(`Pivot source number column contains a non-finite value: ${field.fieldId}`);
        if (column.kind === 'boolean' && column.values[row]! > 1) throw new Error(`Pivot source boolean column contains an invalid value: ${field.fieldId}`);
      }
    }
  });
}

export function pivotSourceValueAt(index: PivotSourceIndex, fieldOrdinal: number, row: number): PivotScalar {
  if (row < 0 || row >= index.rowCount) throw new Error(`Pivot source row is out of bounds: ${String(row)}`);
  const column = index.columns[fieldOrdinal];
  if (!column) throw new Error(`Pivot source field ordinal is out of bounds: ${String(fieldOrdinal)}`);
  if (column.kind === 'number') return column.validity[row] === 0 ? null : column.values[row]!;
  if (column.kind === 'boolean') return column.validity[row] === 0 ? null : column.values[row] === 1;
  const code = column.codes[row]!;
  return code === 0 ? null : column.dictionary[code - 1] ?? null;
}

export function pivotSourceColumnValues(index: PivotSourceIndex, fieldOrdinal: number): PivotScalar[] {
  return Array.from({ length: index.rowCount }, (_, row) => pivotSourceValueAt(index, fieldOrdinal, row));
}

export function pivotSourceRowPaths(index: PivotSourceIndex, row: number): PivotSourceRowPath[] {
  if (row < 0 || row >= index.rowCount) throw new Error(`Pivot source row is out of bounds: ${String(row)}`);
  const start = index.rowPathOffsets[row]!;
  const end = index.rowPathOffsets[row + 1]!;
  return index.rowPathPool.slice(start, end);
}

export function pivotSourceIndexTransferables(index: PivotSourceIndex): ArrayBuffer[] {
  const buffers = new Set<ArrayBuffer>([index.rowPathOffsets.buffer as ArrayBuffer]);
  for (const column of index.columns) {
    if (column.kind === 'dictionary') buffers.add(column.codes.buffer as ArrayBuffer);
    else {
      buffers.add(column.values.buffer as ArrayBuffer);
      buffers.add(column.validity.buffer as ArrayBuffer);
    }
  }
  return [...buffers];
}

export function estimatePivotSourceIndexBytes(index: PivotSourceIndex): number {
  let bytes = index.rowPathOffsets.byteLength;
  for (const field of index.fields) bytes += (field.fieldId.length + field.name.length) * 2 + 16;
  for (const column of index.columns) {
    if (column.kind === 'dictionary') {
      bytes += column.codes.byteLength;
      bytes += column.dictionary.reduce<number>((sum, value) => sum + scalarBytes(value), 0);
    } else bytes += column.values.byteLength + column.validity.byteLength;
  }
  bytes += index.rowPathPool.reduce((sum, path) => sum + 40
    + (path.sheetId?.length ?? 0) * 2
    + (path.sourceId?.length ?? 0) * 2
    + (path.recordId?.length ?? 0) * 2, 0);
  return bytes;
}

function encodeColumn(dataType: PivotFieldDataType, values: readonly PivotScalar[]): PivotSourceColumn {
  if (dataType === 'number') {
    const encoded = new Float64Array(values.length);
    const validity = new Uint8Array(values.length);
    values.forEach((value, index) => {
      if (value == null || value === '') return;
      encoded[index] = value as number;
      validity[index] = 1;
    });
    return { kind: 'number', values: encoded, validity };
  }
  if (dataType === 'boolean') {
    const encoded = new Uint8Array(values.length);
    const validity = new Uint8Array(values.length);
    values.forEach((value, index) => {
      if (value == null || value === '') return;
      encoded[index] = value === true ? 1 : 0;
      validity[index] = 1;
    });
    return { kind: 'boolean', values: encoded, validity };
  }
  const dictionary: PivotScalar[] = [];
  const dictionaryIndex = new Map<string, number>();
  const codes = new Uint32Array(values.length);
  values.forEach((value, index) => {
    if (value == null || value === '') return;
    const key = pivotMemberKey(createPivotMemberKey(value));
    let code = dictionaryIndex.get(key);
    if (code === undefined) {
      dictionary.push(structuredClone(value));
      code = dictionary.length;
      dictionaryIndex.set(key, code);
    }
    codes[index] = code;
  });
  return { kind: 'dictionary', dictionary, codes };
}

function validateDeclaredType(fieldId: string, dataType: PivotFieldDataType, values: readonly PivotScalar[]): void {
  const invalid = values.find((value) => {
    if (value == null || value === '') return false;
    // Typed numeric/boolean/date columns cannot carry a formula error: the
    // transferable representation has no error slot and would otherwise turn
    // the object into NaN/false. Mixed and error columns retain the canonical
    // error value in their dictionary representation.
    if (isPivotError(value)) return dataType !== 'mixed' && dataType !== 'error';
    if (dataType === 'number') return typeof value !== 'number' || !Number.isFinite(value);
    if (dataType === 'boolean') return typeof value !== 'boolean';
    if (dataType === 'text') return typeof value !== 'string';
    if (dataType === 'date') return !((typeof value === 'number' && Number.isFinite(value))
      || (typeof value === 'string' && pivotTimelineInstant(value) !== undefined));
    if (dataType === 'error') return !isPivotError(value);
    return false;
  });
  if (invalid !== undefined) throw new Error(`Pivot source field ${fieldId} contains a value incompatible with ${dataType}`);
}

function scalarBytes(value: PivotScalar): number {
  if (typeof value === 'string') return value.length * 2 + 8;
  if (typeof value === 'number') return 8;
  if (typeof value === 'boolean') return 1;
  if (isPivotError(value)) return (value.code.length + (value.message?.length ?? 0)) * 2 + 16;
  return 1;
}
