import {
  DEFAULT_DATA_BLOCK_ROW_COUNT,
  type DataSourceFieldType,
  type TableScalar,
} from '@react-sheets/core-model';

/** The wire encoding used by DataBlockRef. */
export const COLUMNAR_BLOCK_ENCODING = 'columnar-v1' as const;
export const COLUMNAR_BLOCK_MAGIC = 'LDBK' as const;
export const COLUMNAR_BLOCK_FORMAT_VERSION = 1 as const;
export const COLUMNAR_BLOCK_CHECKSUM_BYTES = 32;

const MAGIC_BYTES = Uint8Array.from([...COLUMNAR_BLOCK_MAGIC].map((character) => character.charCodeAt(0)));
const FIXED_HEADER_BYTES = 40;
const FIELD_DESCRIPTOR_BYTES = 32;
const FIELD_SECTION_HEADER_BYTES = 24;
const MAX_UINT32 = 0xffff_ffff;

const FieldTypeCode = {
  text: 0,
  number: 1,
  boolean: 2,
  date: 3,
  mixed: 4,
} as const;
type FieldTypeCode = (typeof FieldTypeCode)[keyof typeof FieldTypeCode];

/** The schema stored in a block and checked before its rows are consumed. */
export interface ColumnarBlockField {
  id: string;
  name: string;
  ordinal: number;
  type: DataSourceFieldType;
}

export interface EncodeColumnarBlockInput {
  fields: readonly ColumnarBlockField[];
  rows: readonly (readonly TableScalar[])[];
}

export interface DecodeColumnarBlockOptions {
  expectedRowCount?: number;
  expectedFields?: readonly ColumnarBlockField[];
  expectedChecksum?: string;
}

export interface DecodedColumnarBlock {
  schema: 'ColumnarBlock';
  formatVersion: typeof COLUMNAR_BLOCK_FORMAT_VERSION;
  encoding: typeof COLUMNAR_BLOCK_ENCODING;
  rowCount: number;
  fields: ColumnarBlockField[];
  rows: TableScalar[][];
  checksum: string;
}

/**
 * Sparse edits deliberately live beside a columnar block. They are not
 * encoded into the block and never replace its row cells. Coordinates are
 * zero-based and relative to the block.
 */
export const SPARSE_CELL_OVERLAY_SCHEMA = 'SparseCellOverlay' as const;

export interface SparseCellOverlayCell {
  row: number;
  column: number;
  value: TableScalar;
}

export interface SparseCellOverlay {
  schema: typeof SPARSE_CELL_OVERLAY_SCHEMA;
  revision: number;
  cells: readonly SparseCellOverlayCell[];
}

export interface SparseCellOverlayBounds {
  rowCount: number;
  columnCount: number;
}

function fail(message: string): never {
  throw new Error(`Columnar data block ${message}`);
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function assertUInt32(value: number, label: string): void {
  if (!Number.isSafeInteger(value) || value < 0 || value > MAX_UINT32) {
    fail(`${label} must be an unsigned 32-bit integer`);
  }
}

function assertLength(value: number, label: string): void {
  assertUInt32(value, label);
}

function checkedAdd(left: number, right: number, label: string): number {
  const value = left + right;
  if (!Number.isSafeInteger(value) || value > MAX_UINT32) fail(`${label} exceeds the binary block limit`);
  return value;
}

function checkedMultiply(left: number, right: number, label: string): number {
  const value = left * right;
  if (!Number.isSafeInteger(value) || value > MAX_UINT32) fail(`${label} exceeds the binary block limit`);
  return value;
}

function typeCode(type: DataSourceFieldType): FieldTypeCode {
  switch (type) {
    case 'text': return FieldTypeCode.text;
    case 'number': return FieldTypeCode.number;
    case 'boolean': return FieldTypeCode.boolean;
    case 'date': return FieldTypeCode.date;
    case 'mixed': return FieldTypeCode.mixed;
    default: return fail(`does not support field type ${String(type)}`);
  }
}

function typeFromCode(code: number): DataSourceFieldType {
  switch (code) {
    case FieldTypeCode.text: return 'text';
    case FieldTypeCode.number: return 'number';
    case FieldTypeCode.boolean: return 'boolean';
    case FieldTypeCode.date: return 'date';
    case FieldTypeCode.mixed: return 'mixed';
    default: return fail(`contains unknown field type code ${String(code)}`);
  }
}

function assertScalar(value: unknown, label: string): asserts value is TableScalar {
  if (value !== null && typeof value !== 'string' && typeof value !== 'boolean'
    && (typeof value !== 'number' || !Number.isFinite(value))) {
    fail(`${label} must be a finite number, boolean, string, or null`);
  }
}

function validateFields(fields: readonly ColumnarBlockField[]): ColumnarBlockField[] {
  if (!Array.isArray(fields) || fields.length === 0) fail('requires at least one field');
  if (fields.length > MAX_UINT32) fail('contains too many fields');
  const ids = new Set<string>();
  return fields.map((field, index) => {
    if (!isRecord(field)) fail(`field ${String(index)} is not an object`);
    if (typeof field.id !== 'string' || field.id.trim().length === 0) fail(`field ${String(index)} has no id`);
    if (typeof field.name !== 'string' || field.name.trim().length === 0) fail(`field ${field.id} has no name`);
    const ordinal = field.ordinal;
    if (!Number.isSafeInteger(ordinal) || ordinal !== index) {
      fail(`field ${field.id} must have ordinal ${String(index)}`);
    }
    if (ids.has(field.id)) fail(`contains duplicate field id ${field.id}`);
    ids.add(field.id);
    typeCode(field.type as DataSourceFieldType);
    return {
      id: field.id,
      name: field.name,
      ordinal: Number(ordinal),
      type: field.type as DataSourceFieldType,
    };
  });
}

function validateValueForField(value: unknown, field: ColumnarBlockField, row: number): void {
  assertScalar(value, `row ${String(row)}, field ${field.id}`);
  if (value === null || field.type === 'mixed') return;
  const valid = field.type === 'text'
    ? typeof value === 'string'
    : field.type === 'boolean'
      ? typeof value === 'boolean'
      : typeof value === 'number';
  if (!valid) fail(`row ${String(row)}, field ${field.id} does not match ${field.type}`);
}

function validateRows(rows: readonly (readonly TableScalar[])[], fields: readonly ColumnarBlockField[]): TableScalar[][] {
  if (!Array.isArray(rows) || rows.length === 0) fail('requires at least one row');
  if (rows.length > DEFAULT_DATA_BLOCK_ROW_COUNT) {
    fail(`row count cannot exceed ${String(DEFAULT_DATA_BLOCK_ROW_COUNT)}`);
  }
  return rows.map((row, rowIndex) => {
    if (!Array.isArray(row) || row.length !== fields.length) {
      fail(`row ${String(rowIndex)} does not match the field count`);
    }
    row.forEach((value, fieldIndex) => validateValueForField(value, fields[fieldIndex]!, rowIndex));
    return [...row];
  });
}

function validityByteLength(rowCount: number): number {
  return Math.ceil(rowCount / 8);
}

function setValidity(bitmap: Uint8Array, row: number): void {
  bitmap[row >> 3] = (bitmap[row >> 3] ?? 0) | (1 << (row & 7));
}

function isValid(bitmap: Uint8Array, row: number): boolean {
  return ((bitmap[row >> 3] ?? 0) & (1 << (row & 7))) !== 0;
}

function concatBytes(parts: readonly Uint8Array[]): Uint8Array {
  let length = 0;
  for (const part of parts) length = checkedAdd(length, part.byteLength, 'byte payload');
  const result = new Uint8Array(length);
  let offset = 0;
  for (const part of parts) {
    result.set(part, offset);
    offset += part.byteLength;
  }
  return result;
}

function encodeFloat64Array(values: Float64Array): Uint8Array {
  const bytes = new Uint8Array(values.length * Float64Array.BYTES_PER_ELEMENT);
  const view = new DataView(bytes.buffer);
  for (let index = 0; index < values.length; index += 1) {
    view.setFloat64(index * Float64Array.BYTES_PER_ELEMENT, values[index]!, true);
  }
  return bytes;
}

function encodeInt32Array(values: Int32Array): Uint8Array {
  const bytes = new Uint8Array(values.length * Int32Array.BYTES_PER_ELEMENT);
  const view = new DataView(bytes.buffer);
  for (let index = 0; index < values.length; index += 1) {
    view.setInt32(index * Int32Array.BYTES_PER_ELEMENT, values[index]!, true);
  }
  return bytes;
}

function decodeFloat64Array(bytes: Uint8Array, rowCount: number): Float64Array {
  const expected = checkedMultiply(rowCount, Float64Array.BYTES_PER_ELEMENT, 'number column');
  if (bytes.byteLength !== expected) fail('has an invalid Float64Array column length');
  const values = new Float64Array(rowCount);
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  for (let index = 0; index < rowCount; index += 1) {
    values[index] = view.getFloat64(index * Float64Array.BYTES_PER_ELEMENT, true);
  }
  return values;
}

function decodeInt32Array(bytes: Uint8Array, rowCount: number): Int32Array {
  const expected = checkedMultiply(rowCount, Int32Array.BYTES_PER_ELEMENT, 'text index column');
  if (bytes.byteLength !== expected) fail('has an invalid Int32Array column length');
  const values = new Int32Array(rowCount);
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  for (let index = 0; index < rowCount; index += 1) {
    values[index] = view.getInt32(index * Int32Array.BYTES_PER_ELEMENT, true);
  }
  return values;
}

function buildDictionary(values: readonly (string | null)[]): { indexes: Int32Array; dictionary: Uint8Array } {
  const indexes = new Int32Array(values.length);
  indexes.fill(-1);
  const dictionaryValues: string[] = [];
  const dictionaryIndexes = new Map<string, number>();
  for (let index = 0; index < values.length; index += 1) {
    const value = values[index];
    if (value === null) continue;
    let dictionaryIndex = dictionaryIndexes.get(value);
    if (dictionaryIndex === undefined) {
      dictionaryIndex = dictionaryValues.length;
      dictionaryIndexes.set(value, dictionaryIndex);
      dictionaryValues.push(value);
    }
    indexes[index] = dictionaryIndex;
  }

  const encoder = new TextEncoder();
  const encodedValues = dictionaryValues.map((value) => encoder.encode(value));
  let stringBytes = 0;
  for (const value of encodedValues) stringBytes = checkedAdd(stringBytes, value.byteLength, 'text dictionary');
  const offsetsBytes = checkedMultiply(dictionaryValues.length + 1, 4, 'text dictionary offsets');
  const total = checkedAdd(4, checkedAdd(offsetsBytes, stringBytes, 'text dictionary'), 'text dictionary');
  const dictionary = new Uint8Array(total);
  const view = new DataView(dictionary.buffer);
  view.setUint32(0, dictionaryValues.length, true);
  let offset = 0;
  let writeOffset = 4 + offsetsBytes;
  for (const value of encodedValues) {
    dictionary.set(value, writeOffset);
    writeOffset += value.byteLength;
    offset += value.byteLength;
  }
  offset = 0;
  view.setUint32(4, 0, true);
  for (let index = 0; index < encodedValues.length; index += 1) {
    offset += encodedValues[index]!.byteLength;
    view.setUint32(4 + (index + 1) * 4, offset, true);
  }
  return { indexes, dictionary };
}

function decodeDictionary(bytes: Uint8Array): string[] {
  if (bytes.byteLength < 8) fail('has an incomplete text dictionary');
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const count = view.getUint32(0, true);
  const offsetTableBytes = checkedMultiply(count + 1, 4, 'text dictionary offsets');
  const dataStart = checkedAdd(4, offsetTableBytes, 'text dictionary');
  if (dataStart > bytes.byteLength) fail('has an out-of-range text dictionary table');
  const decoder = new TextDecoder('utf-8', { fatal: true });
  const values: string[] = [];
  const seen = new Set<string>();
  let previous = view.getUint32(4, true);
  if (previous !== 0) fail('text dictionary must start at offset zero');
  for (let index = 0; index < count; index += 1) {
    const start = view.getUint32(4 + index * 4, true);
    const end = view.getUint32(4 + (index + 1) * 4, true);
    if (start !== previous || end < start || dataStart + end > bytes.byteLength) {
      fail('has invalid text dictionary offsets');
    }
    let value: string;
    try {
      value = decoder.decode(bytes.subarray(dataStart + start, dataStart + end));
    } catch {
      fail('contains invalid UTF-8 text');
    }
    if (seen.has(value)) fail('contains duplicate text dictionary values');
    seen.add(value);
    values.push(value);
    previous = end;
  }
  const finalOffset = view.getUint32(4 + count * 4, true);
  if (finalOffset !== previous || dataStart + finalOffset !== bytes.byteLength) {
    fail('has a trailing or truncated text dictionary payload');
  }
  return values;
}

function createFieldSection(field: ColumnarBlockField, rows: readonly TableScalar[]): Uint8Array {
  const rowCount = rows.length;
  const validity = new Uint8Array(validityByteLength(rowCount));
  let primary: Uint8Array;
  let dictionary: Uint8Array<ArrayBufferLike> = new Uint8Array(0);
  let tags = new Uint8Array(0);

  if (field.type === 'text') {
    const textValues = rows.map((value) => value === null ? null : String(value));
    const encoded = buildDictionary(textValues);
    primary = encodeInt32Array(encoded.indexes);
    dictionary = encoded.dictionary;
  } else if (field.type === 'number' || field.type === 'date') {
    const values = new Float64Array(rowCount);
    rows.forEach((value, index) => {
      if (value !== null) {
        setValidity(validity, index);
        values[index] = value as number;
      }
    });
    primary = encodeFloat64Array(values);
  } else if (field.type === 'boolean') {
    const values = new Uint8Array(rowCount);
    rows.forEach((value, index) => {
      if (value !== null) {
        setValidity(validity, index);
        values[index] = value === true ? 1 : 0;
      }
    });
    primary = values;
  } else {
    const numbers = new Float64Array(rowCount);
    const booleans = new Uint8Array(rowCount);
    const textValues: (string | null)[] = rows.map(() => null);
    const typeTags = new Uint8Array(rowCount);
    rows.forEach((value, index) => {
      if (value === null) return;
      setValidity(validity, index);
      if (typeof value === 'number') {
        typeTags[index] = 1;
        numbers[index] = value;
      } else if (typeof value === 'boolean') {
        typeTags[index] = 2;
        booleans[index] = value ? 1 : 0;
      } else {
        typeTags[index] = 3;
        textValues[index] = value;
      }
    });
    const encoded = buildDictionary(textValues);
    primary = concatBytes([encodeFloat64Array(numbers), booleans, encodeInt32Array(encoded.indexes)]);
    dictionary = encoded.dictionary;
    tags = typeTags;
  }

  if (field.type === 'text') {
    rows.forEach((value, index) => {
      if (value !== null) setValidity(validity, index);
    });
  }

  const total = FIELD_SECTION_HEADER_BYTES;
  const sectionLength = checkedAdd(
    checkedAdd(checkedAdd(checkedAdd(total, validity.byteLength, 'field section'), primary.byteLength, 'field section'), dictionary.byteLength, 'field section'),
    tags.byteLength,
    'field section',
  );
  const section = new Uint8Array(sectionLength);
  const view = new DataView(section.buffer);
  view.setUint8(0, typeCode(field.type));
  view.setUint8(1, field.type === 'mixed' ? 1 : 0);
  view.setUint32(4, rowCount, true);
  view.setUint32(8, validity.byteLength, true);
  view.setUint32(12, primary.byteLength, true);
  view.setUint32(16, dictionary.byteLength, true);
  view.setUint32(20, tags.byteLength, true);
  let offset = FIELD_SECTION_HEADER_BYTES;
  section.set(validity, offset);
  offset += validity.byteLength;
  section.set(primary, offset);
  offset += primary.byteLength;
  section.set(dictionary, offset);
  offset += dictionary.byteLength;
  section.set(tags, offset);
  return section;
}

interface EncodedField {
  field: ColumnarBlockField;
  id: Uint8Array;
  name: Uint8Array;
  section: Uint8Array;
  idOffset: number;
  nameOffset: number;
  sectionOffset: number;
}

function toOwnedBytes(input: ArrayBuffer | ArrayBufferView): Uint8Array {
  if (input instanceof ArrayBuffer) return new Uint8Array(input.slice(0));
  const view = new Uint8Array(input.buffer, input.byteOffset, input.byteLength);
  const result = new Uint8Array(view.byteLength);
  result.set(view);
  return result;
}

function bytesToHex(bytes: Uint8Array): string {
  return [...bytes].map((value) => value.toString(16).padStart(2, '0')).join('');
}

function hexToBytes(value: string): Uint8Array {
  const normalized = value.toLowerCase();
  if (!/^[0-9a-f]+$/.test(normalized) || normalized.length % 2 !== 0) fail('contains an invalid hexadecimal checksum');
  const bytes = new Uint8Array(normalized.length / 2);
  for (let index = 0; index < bytes.length; index += 1) {
    bytes[index] = Number.parseInt(normalized.slice(index * 2, index * 2 + 2), 16);
  }
  return bytes;
}

async function sha256Hex(bytes: Uint8Array): Promise<string> {
  const subtle = globalThis.crypto?.subtle;
  if (!subtle) fail('requires Web Crypto SHA-256 support');
  const digest = await subtle.digest('SHA-256', bytes.slice().buffer as ArrayBuffer);
  return bytesToHex(new Uint8Array(digest));
}

function compareExpectedChecksum(value: string | undefined, actual: string): void {
  if (value === undefined) return;
  if (!/^[0-9a-fA-F]{64}$/.test(value)) fail('expected checksum must be a SHA-256 hexadecimal string');
  if (value.toLowerCase() !== actual) fail('checksum does not match the expected block checksum');
}

/**
 * Encode one local data block. The result is a single binary ArrayBuffer with
 * a self-describing schema, typed columns, and a trailing SHA-256 checksum.
 */
export async function encodeColumnarBlock(input: EncodeColumnarBlockInput): Promise<ArrayBuffer> {
  if (!isRecord(input)) fail('encode input must be an object');
  const fields = validateFields(input.fields as readonly ColumnarBlockField[]);
  const rows = validateRows(input.rows as readonly (readonly TableScalar[])[], fields);
  const encoder = new TextEncoder();
  const encodedFields: EncodedField[] = [];
  const stringParts: Uint8Array[] = [];
  let stringOffset = 0;
  let payloadOffset = 0;
  for (let index = 0; index < fields.length; index += 1) {
    const field = fields[index]!;
    const id = encoder.encode(field.id);
    const name = encoder.encode(field.name);
    assertLength(id.byteLength, `field ${field.id} id length`);
    assertLength(name.byteLength, `field ${field.id} name length`);
    const section = createFieldSection(field, rows.map((row) => row[index]!));
    encodedFields.push({ field, id, name, section, idOffset: stringOffset, nameOffset: stringOffset + id.byteLength, sectionOffset: payloadOffset });
    stringParts.push(id, name);
    stringOffset = checkedAdd(stringOffset, checkedAdd(id.byteLength, name.byteLength, 'field schema'), 'field schema');
    payloadOffset = checkedAdd(payloadOffset, section.byteLength, 'column payload');
  }
  const stringTable = concatBytes(stringParts);
  const descriptorBytes = checkedMultiply(fields.length, FIELD_DESCRIPTOR_BYTES, 'field descriptors');
  const headerBytes = checkedAdd(FIXED_HEADER_BYTES, checkedAdd(descriptorBytes, stringTable.byteLength, 'block header'), 'block header');
  const payloadBytes = payloadOffset;
  const totalBytes = checkedAdd(checkedAdd(headerBytes, payloadBytes, 'block size'), COLUMNAR_BLOCK_CHECKSUM_BYTES, 'block size');
  if (totalBytes > 0x7fff_ffff) fail('is too large for a JavaScript ArrayBuffer');

  const bytes = new Uint8Array(new ArrayBuffer(totalBytes));
  const view = new DataView(bytes.buffer);
  bytes.set(MAGIC_BYTES, 0);
  view.setUint8(4, COLUMNAR_BLOCK_FORMAT_VERSION);
  view.setUint8(5, 0);
  view.setUint32(8, headerBytes, true);
  view.setUint32(12, rows.length, true);
  view.setUint32(16, fields.length, true);
  view.setUint32(20, descriptorBytes, true);
  view.setUint32(24, stringTable.byteLength, true);
  view.setUint32(28, payloadBytes, true);
  view.setUint32(32, COLUMNAR_BLOCK_CHECKSUM_BYTES, true);
  let descriptorOffset = FIXED_HEADER_BYTES;
  for (const encoded of encodedFields) {
    view.setUint32(descriptorOffset, encoded.idOffset, true);
    view.setUint32(descriptorOffset + 4, encoded.id.byteLength, true);
    view.setUint32(descriptorOffset + 8, encoded.nameOffset, true);
    view.setUint32(descriptorOffset + 12, encoded.name.byteLength, true);
    view.setUint32(descriptorOffset + 16, encoded.field.ordinal, true);
    view.setUint8(descriptorOffset + 20, typeCode(encoded.field.type));
    view.setUint8(descriptorOffset + 21, 0);
    view.setUint32(descriptorOffset + 24, encoded.sectionOffset, true);
    view.setUint32(descriptorOffset + 28, encoded.section.byteLength, true);
    descriptorOffset += FIELD_DESCRIPTOR_BYTES;
  }
  bytes.set(stringTable, FIXED_HEADER_BYTES + descriptorBytes);
  let payloadWriteOffset = headerBytes;
  for (const encoded of encodedFields) {
    bytes.set(encoded.section, payloadWriteOffset);
    payloadWriteOffset += encoded.section.byteLength;
  }
  const checksum = hexToBytes(await sha256Hex(bytes.subarray(0, totalBytes - COLUMNAR_BLOCK_CHECKSUM_BYTES)));
  bytes.set(checksum, totalBytes - COLUMNAR_BLOCK_CHECKSUM_BYTES);
  return bytes.buffer;
}

function rangeWithin(start: number, length: number, total: number, label: string): void {
  if (!Number.isSafeInteger(start) || !Number.isSafeInteger(length) || start < 0 || length < 0 || start > total || length > total - start) {
    fail(`${label} is outside the block`);
  }
}

function decodeUtf8(bytes: Uint8Array, start: number, length: number, label: string): string {
  rangeWithin(start, length, bytes.byteLength, label);
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(bytes.subarray(start, start + length));
  } catch {
    fail(`${label} is not valid UTF-8`);
  }
}

function validateBitmap(bitmap: Uint8Array, rowCount: number): void {
  const unusedBits = bitmap.byteLength * 8 - rowCount;
  if (unusedBits > 0 && (bitmap[bitmap.length - 1]! & (0xff << (8 - unusedBits))) !== 0) {
    fail('has non-zero validity bits beyond the row count');
  }
}

function validateExpectedFields(actual: readonly ColumnarBlockField[], expected: readonly ColumnarBlockField[] | undefined): void {
  if (expected === undefined) return;
  const normalized = validateFields(expected);
  if (normalized.length !== actual.length) fail('field schema does not match the expected schema');
  for (let index = 0; index < actual.length; index += 1) {
    const left = actual[index]!;
    const right = normalized[index]!;
    if (left.id !== right.id || left.name !== right.name || left.ordinal !== right.ordinal || left.type !== right.type) {
      fail('field schema does not match the expected schema');
    }
  }
}

function decodeFieldValues(
  field: ColumnarBlockField,
  section: Uint8Array,
  rowCount: number,
): TableScalar[] {
  if (section.byteLength < FIELD_SECTION_HEADER_BYTES) fail(`field ${field.id} has an incomplete column section`);
  const view = new DataView(section.buffer, section.byteOffset, section.byteLength);
  const sectionType = typeFromCode(view.getUint8(0));
  if (view.getUint16(2, true) !== 0 || sectionType !== field.type || view.getUint8(1) !== (field.type === 'mixed' ? 1 : 0)) {
    fail(`field ${field.id} has a schema/type mismatch`);
  }
  if (view.getUint32(4, true) !== rowCount) fail(`field ${field.id} has a row count mismatch`);
  const validityLength = view.getUint32(8, true);
  const primaryLength = view.getUint32(12, true);
  const dictionaryLength = view.getUint32(16, true);
  const tagLength = view.getUint32(20, true);
  const expectedValidityLength = validityByteLength(rowCount);
  if (validityLength !== expectedValidityLength) fail(`field ${field.id} has an invalid validity bitmap length`);
  const expectedPrimaryLength = field.type === 'number' || field.type === 'date'
    ? checkedMultiply(rowCount, Float64Array.BYTES_PER_ELEMENT, `field ${field.id}`)
    : field.type === 'boolean'
      ? rowCount
      : field.type === 'text'
        ? checkedMultiply(rowCount, Int32Array.BYTES_PER_ELEMENT, `field ${field.id}`)
        : checkedMultiply(rowCount, 13, `field ${field.id}`);
  const expectedTagLength = field.type === 'mixed' ? rowCount : 0;
  if (primaryLength !== expectedPrimaryLength || tagLength !== expectedTagLength) {
    fail(`field ${field.id} has an invalid typed column length`);
  }
  if (field.type !== 'text' && field.type !== 'mixed' && dictionaryLength !== 0) {
    fail(`field ${field.id} has an unexpected text dictionary`);
  }
  const expectedLength = checkedAdd(
    checkedAdd(checkedAdd(checkedAdd(FIELD_SECTION_HEADER_BYTES, validityLength, 'field section'), primaryLength, 'field section'), dictionaryLength, 'field section'),
    tagLength,
    'field section',
  );
  if (expectedLength !== section.byteLength) fail(`field ${field.id} has a truncated or trailing section`);
  let offset = FIELD_SECTION_HEADER_BYTES;
  const validity = section.subarray(offset, offset + validityLength);
  offset += validityLength;
  const primary = section.subarray(offset, offset + primaryLength);
  offset += primaryLength;
  const dictionaryBytes = section.subarray(offset, offset + dictionaryLength);
  offset += dictionaryLength;
  const tags = section.subarray(offset, offset + tagLength);
  validateBitmap(validity, rowCount);
  const dictionary = field.type === 'text' || field.type === 'mixed' ? decodeDictionary(dictionaryBytes) : [];
  const values = new Array<TableScalar>(rowCount).fill(null);

  if (field.type === 'number' || field.type === 'date') {
    const numbers = decodeFloat64Array(primary, rowCount);
    for (let index = 0; index < rowCount; index += 1) {
      if (!Number.isFinite(numbers[index]!)) fail(`field ${field.id} contains a non-finite number`);
      if (isValid(validity, index)) values[index] = numbers[index]!;
      else if (numbers[index] !== 0) fail(`field ${field.id} has data for a null number value`);
    }
    return values;
  }
  if (field.type === 'boolean') {
    for (let index = 0; index < rowCount; index += 1) {
      const value = primary[index]!;
      if (value > 1) fail(`field ${field.id} contains an invalid boolean value`);
      if (isValid(validity, index)) values[index] = value === 1;
      else if (value !== 0) fail(`field ${field.id} has data for a null boolean value`);
    }
    return values;
  }
  if (field.type === 'text') {
    const indexes = decodeInt32Array(primary, rowCount);
    for (let index = 0; index < rowCount; index += 1) {
      const dictionaryIndex = indexes[index]!;
      if (!isValid(validity, index)) {
        if (dictionaryIndex !== -1) fail(`field ${field.id} has an index for a null text value`);
        continue;
      }
      if (dictionaryIndex < 0 || dictionaryIndex >= dictionary.length) fail(`field ${field.id} contains an invalid text dictionary index`);
      values[index] = dictionary[dictionaryIndex]!;
    }
    return values;
  }

  const numberBytes = rowCount * Float64Array.BYTES_PER_ELEMENT;
  const booleanBytes = primary.subarray(numberBytes, numberBytes + rowCount);
  const indexes = decodeInt32Array(primary.subarray(numberBytes + rowCount), rowCount);
  const numbers = decodeFloat64Array(primary.subarray(0, numberBytes), rowCount);
  for (let index = 0; index < rowCount; index += 1) {
    const tag = tags[index]!;
    if (!Number.isFinite(numbers[index]!)) fail(`field ${field.id} contains a non-finite number`);
    if (booleanBytes[index]! > 1) fail(`field ${field.id} contains an invalid boolean value`);
    const dictionaryIndex = indexes[index]!;
    if (tag !== 3 && dictionaryIndex !== -1) fail(`field ${field.id} has a text index for a non-text value`);
    if (!isValid(validity, index)) {
      if (tag !== 0) fail(`field ${field.id} has a type tag for a null value`);
      continue;
    }
    if (tag === 1) {
      values[index] = numbers[index]!;
    } else if (tag === 2) {
      values[index] = booleanBytes[index] === 1;
    } else if (tag === 3) {
      if (dictionaryIndex < 0 || dictionaryIndex >= dictionary.length) fail(`field ${field.id} contains an invalid text dictionary index`);
      values[index] = dictionary[dictionaryIndex]!;
    } else {
      fail(`field ${field.id} contains an invalid mixed type tag`);
    }
  }
  return values;
}

/** Compute the SHA-256 checksum used by DataBlockRef over the complete block. */
export async function computeColumnarBlockChecksum(input: ArrayBuffer | ArrayBufferView): Promise<string> {
  const bytes = toOwnedBytes(input);
  if (bytes.byteLength < FIXED_HEADER_BYTES + COLUMNAR_BLOCK_CHECKSUM_BYTES) fail('is too short to contain a complete block');
  return sha256Hex(bytes);
}

/**
 * Decode and validate a complete block. The embedded checksum, row count,
 * field schema, typed column lengths, validity bitmap, and dictionary indexes
 * are all checked before rows are returned.
 */
export async function decodeColumnarBlock(
  input: ArrayBuffer | ArrayBufferView,
  options: DecodeColumnarBlockOptions = {},
): Promise<DecodedColumnarBlock> {
  const bytes = toOwnedBytes(input);
  if (bytes.byteLength < FIXED_HEADER_BYTES + COLUMNAR_BLOCK_CHECKSUM_BYTES) fail('is truncated');
  const view = new DataView(bytes.buffer);
  for (let index = 0; index < MAGIC_BYTES.length; index += 1) {
    if (bytes[index] !== MAGIC_BYTES[index]) fail('has an invalid magic header');
  }
  if (view.getUint8(4) !== COLUMNAR_BLOCK_FORMAT_VERSION) fail('uses an unsupported format version');
  if (view.getUint8(5) !== 0 || view.getUint16(6, true) !== 0 || view.getUint32(36, true) !== 0) fail('has unsupported header flags');
  const headerBytes = view.getUint32(8, true);
  const rowCount = view.getUint32(12, true);
  const fieldCount = view.getUint32(16, true);
  const descriptorBytes = view.getUint32(20, true);
  const stringTableBytes = view.getUint32(24, true);
  const payloadBytes = view.getUint32(28, true);
  const checksumBytes = view.getUint32(32, true);
  if (rowCount === 0 || rowCount > DEFAULT_DATA_BLOCK_ROW_COUNT) fail('has an invalid row count');
  if (fieldCount === 0) fail('has no fields');
  if (descriptorBytes !== checkedMultiply(fieldCount, FIELD_DESCRIPTOR_BYTES, 'field descriptors')) fail('has an invalid descriptor length');
  if (checksumBytes !== COLUMNAR_BLOCK_CHECKSUM_BYTES) fail('has an invalid checksum length');
  if (headerBytes !== FIXED_HEADER_BYTES + descriptorBytes + stringTableBytes) fail('has an invalid header length');
  const expectedTotal = checkedAdd(checkedAdd(headerBytes, payloadBytes, 'block size'), checksumBytes, 'block size');
  if (expectedTotal !== bytes.byteLength) fail('has trailing or missing bytes');
  const checksumOffset = bytes.byteLength - checksumBytes;
  const embeddedChecksum = await sha256Hex(bytes.subarray(0, checksumOffset));
  const storedChecksum = bytesToHex(bytes.subarray(checksumOffset));
  if (embeddedChecksum !== storedChecksum) fail('embedded checksum does not match the block bytes');
  const actualChecksum = await sha256Hex(bytes);
  compareExpectedChecksum(options.expectedChecksum, actualChecksum);

  const stringTableStart = FIXED_HEADER_BYTES + descriptorBytes;
  const payloadStart = headerBytes;
  rangeWithin(stringTableStart, stringTableBytes, bytes.byteLength, 'field schema');
  rangeWithin(payloadStart, payloadBytes, checksumOffset, 'column payload');
  const fields: ColumnarBlockField[] = [];
  const sections: { field: ColumnarBlockField; offset: number; length: number }[] = [];
  const ids = new Set<string>();
  for (let index = 0; index < fieldCount; index += 1) {
    const descriptorOffset = FIXED_HEADER_BYTES + index * FIELD_DESCRIPTOR_BYTES;
    const idOffset = view.getUint32(descriptorOffset, true);
    const idLength = view.getUint32(descriptorOffset + 4, true);
    const nameOffset = view.getUint32(descriptorOffset + 8, true);
    const nameLength = view.getUint32(descriptorOffset + 12, true);
    const ordinal = view.getUint32(descriptorOffset + 16, true);
    const fieldType = typeFromCode(view.getUint8(descriptorOffset + 20));
    if (view.getUint8(descriptorOffset + 21) !== 0 || view.getUint16(descriptorOffset + 22, true) !== 0) fail('has unsupported field descriptor flags');
    rangeWithin(idOffset, idLength, stringTableBytes, 'field id');
    rangeWithin(nameOffset, nameLength, stringTableBytes, 'field name');
    const id = decodeUtf8(bytes, stringTableStart + idOffset, idLength, 'field id');
    const name = decodeUtf8(bytes, stringTableStart + nameOffset, nameLength, 'field name');
    if (!id.trim() || !name.trim() || ordinal !== index || ids.has(id)) fail('has an invalid field schema');
    ids.add(id);
    const field = { id, name, ordinal, type: fieldType } satisfies ColumnarBlockField;
    const sectionOffset = view.getUint32(descriptorOffset + 24, true);
    const sectionLength = view.getUint32(descriptorOffset + 28, true);
    rangeWithin(sectionOffset, sectionLength, payloadBytes, `field ${id} section`);
    fields.push(field);
    sections.push({ field, offset: sectionOffset, length: sectionLength });
  }
  validateExpectedFields(fields, options.expectedFields);
  if (options.expectedRowCount !== undefined && options.expectedRowCount !== rowCount) {
    fail('row count does not match the expected row count');
  }

  const rows = Array.from({ length: rowCount }, () => new Array<TableScalar>(fieldCount).fill(null));
  let nextSectionOffset = 0;
  for (const entry of sections) {
    if (entry.offset !== nextSectionOffset) fail('column sections are not contiguous');
    const section = bytes.subarray(payloadStart + entry.offset, payloadStart + entry.offset + entry.length);
    const values = decodeFieldValues(entry.field, section, rowCount);
    for (let row = 0; row < rowCount; row += 1) rows[row]![entry.field.ordinal] = values[row]!;
    nextSectionOffset = checkedAdd(nextSectionOffset, entry.length, 'column payload');
  }
  if (nextSectionOffset !== payloadBytes) fail('column payload has an unreferenced region');
  return {
    schema: 'ColumnarBlock',
    formatVersion: COLUMNAR_BLOCK_FORMAT_VERSION,
    encoding: COLUMNAR_BLOCK_ENCODING,
    rowCount,
    fields,
    rows,
    checksum: actualChecksum,
  };
}

/** Validate a sparse overlay without materializing or replacing block cells. */
export function validateSparseCellOverlay(
  value: unknown,
  bounds?: SparseCellOverlayBounds,
): asserts value is SparseCellOverlay {
  if (!isRecord(value) || value.schema !== SPARSE_CELL_OVERLAY_SCHEMA
    || !Number.isSafeInteger(value.revision) || Number(value.revision) < 0
    || !Array.isArray(value.cells)) {
    fail('has an invalid sparse cell overlay');
  }
  if (bounds !== undefined && (!Number.isSafeInteger(bounds.rowCount) || bounds.rowCount < 0
    || !Number.isSafeInteger(bounds.columnCount) || bounds.columnCount < 0)) {
    fail('has invalid sparse overlay bounds');
  }
  const coordinates = new Set<string>();
  for (const cell of value.cells) {
    if (!isRecord(cell)) fail('contains an invalid sparse cell coordinate');
    const row = cell.row;
    const column = cell.column;
    if (!Number.isSafeInteger(row) || Number(row) < 0
      || !Number.isSafeInteger(column) || Number(column) < 0) {
      fail('contains an invalid sparse cell coordinate');
    }
    if (bounds !== undefined && (Number(row) >= bounds.rowCount || Number(column) >= bounds.columnCount)) {
      fail('contains a sparse cell outside the block bounds');
    }
    const key = `${String(row)}:${String(column)}`;
    if (coordinates.has(key)) fail(`contains duplicate sparse cell ${key}`);
    coordinates.add(key);
    assertScalar(cell.value, `sparse cell ${key}`);
  }
}
