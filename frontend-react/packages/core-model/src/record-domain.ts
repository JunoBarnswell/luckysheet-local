import { formatFormula, mapAstTableReferences, parseFormula } from '@react-sheets/formula-engine';
import type { WorkbookModel, WorkbookTableModel, DataRelationship, CellData, RecordFieldAddress } from './index';

export class RecordDomainError extends Error {
  readonly code = 'RECORD_DOMAIN_INVALID';
}

export function canonicalRecordFieldFormula(table: WorkbookTableModel, formula: string): string {
  return formatFormula(mapAstTableReferences(parseFormula(formula.startsWith('=') ? formula : `=${formula}`), node => {
    if (node.tableName && node.tableName !== table.name) return node;
    if (node.specifier || node.columnEndName || !node.columnName) throw new RecordDomainError('Record formulas require scalar field references');
    const field = table.fields.find(field => field.id === node.columnName || field.name === node.columnName);
    if (!field) throw new RecordDomainError(`Field is missing: ${node.columnName}`);
    return { ...node, tableName: '', columnName: field.id, thisRow: true };
  }));
}

/** Derived identity-to-source-position index. Authored identity lives in the source ID column. */
export function recordRows(workbook: WorkbookModel, table: WorkbookTableModel): ReadonlyMap<string, number> {
  if (!table.recordIdFieldId || !table.sourceRange) throw new RecordDomainError(`Record table ${table.id} requires an identity field and source range`);
  const identity = table.fields.find(field => field.id === table.recordIdFieldId);
  if (!identity || identity.calculation || identity.type !== 'text') throw new RecordDomainError('Record identity must be a stored text field');
  const range = table.sourceRange, sheet = workbook.getSheet(range.sheetId), rows = new Map<string, number>();
  for (let row = range.startRow + 1; row <= range.endRow; row++) {
    const id = sheet.cells.getWithoutHydration(row, range.startColumn + identity.ordinal)?.value;
    if (typeof id !== 'string' || !id.trim() || id.length > 200 || rows.has(id)) throw new RecordDomainError(`Record identity is missing or duplicated at ${sheet.id}:${row}`);
    rows.set(id, row);
  }
  return rows;
}

export function assertRecordTable(workbook: WorkbookModel, table: WorkbookTableModel): void {
  if (!table.recordIdFieldId && table.fields.every(field => !field.calculation)) return;
  if (!table.sourceRange || table.sourceRange.endRow - table.sourceRange.startRow > 100000) throw new RecordDomainError('Record table exceeds the bounded worksheet source contract');
  const range = table.sourceRange, source = workbook.getSheet(range.sheetId);
  if (source.kind !== 'worksheet' || table.sourceSheetId !== undefined && table.sourceSheetId !== range.sheetId || table.sourceId || table.blocks.length || !Number.isSafeInteger(table.rowCount) || table.rowCount !== range.endRow - range.startRow || table.fields.length !== range.endColumn - range.startColumn + 1) throw new RecordDomainError('Record source dimensions and field ownership are inconsistent');
  for (const other of workbook.dataModel.tables.values()) {
    const otherRange = other.sourceRange;
    if (other.id !== table.id && other.recordIdFieldId && otherRange?.sheetId === range.sheetId
      && otherRange.startRow <= range.endRow && otherRange.endRow >= range.startRow
      && otherRange.startColumn <= range.endColumn && otherRange.endColumn >= range.startColumn) throw new RecordDomainError('Record source ranges must have one table owner');
  }
  const ids = new Set<string>(), ordinals = new Set<number>();
  for (const field of table.fields) {
    if (!field.id || ids.has(field.id) || !Number.isSafeInteger(field.ordinal) || field.ordinal < 0 || ordinals.has(field.ordinal)
      || table.sourceRange.startColumn + field.ordinal > table.sourceRange.endColumn) throw new RecordDomainError('Record field identities/ordinals are invalid');
    ids.add(field.id); ordinals.add(field.ordinal);
    const calculation = field.calculation;
    if (!calculation) {
      const source = workbook.getSheet(table.sourceRange.sheetId);
      for (let row = table.sourceRange.startRow + 1; row <= table.sourceRange.endRow; row++) if (source.cells.getWithoutHydration(row, table.sourceRange.startColumn + field.ordinal)?.formula) throw new RecordDomainError('Stored Record fields cannot own cell formulas');
      continue;
    }
    if (calculation.kind === 'formula') {
      if (canonicalRecordFieldFormula(table, calculation.formula) !== calculation.formula) throw new RecordDomainError('Record field formula must bind canonical field IDs');
    } else if ((calculation.kind !== 'lookup' && calculation.kind !== 'rollup') || !calculation.relationshipId || !calculation.targetFieldId
      || !['forward', 'reverse'].includes(calculation.direction) || calculation.kind === 'rollup' && !['SUM', 'AVERAGE', 'COUNT', 'COUNTA', 'MIN', 'MAX', 'PRODUCT'].includes(calculation.aggregate)) throw new RecordDomainError('Record field calculation is invalid');
    const sheet = workbook.getSheet(table.sourceRange.sheetId);
    for (let row = table.sourceRange.startRow + 1; row <= table.sourceRange.endRow; row++) {
      const cell = sheet.cells.getWithoutHydration(row, table.sourceRange.startColumn + field.ordinal);
      if (cell?.formula !== undefined || cell?.value != null && cell.value !== '') throw new RecordDomainError('Computed fields cannot overlap stored inputs');
    }
  }
  if (table.fields.some((_, ordinal) => !ordinals.has(ordinal))) throw new RecordDomainError("Record field ordinals must be contiguous");
  recordRows(workbook, table);
}

export function assertRecordRelationship(workbook: WorkbookModel, relationship: DataRelationship): void {
  const from = workbook.getTable(relationship.fromTableId), to = workbook.getTable(relationship.toTableId);
  assertRecordTable(workbook, from); assertRecordTable(workbook, to);
  const field = from.fields.find(field => field.id === relationship.fromFieldId);
  if (!relationship.id || !from.recordIdFieldId || !to.recordIdFieldId || !field || field.calculation
    || relationship.toFieldId !== to.recordIdFieldId || !['many-to-one', 'one-to-one'].includes(relationship.cardinality)) throw new RecordDomainError('Relations must reference immutable target record IDs');
  const target = recordRows(workbook, to), used = new Set<string>();
  const sheet = workbook.getSheet(from.sourceRange!.sheetId);
  for (const row of recordRows(workbook, from).values()) {
    const id = sheet.cells.getWithoutHydration(row, from.sourceRange!.startColumn + field.ordinal)?.value;
    if (id == null || id === '') continue;
    if (typeof id !== 'string' || !target.has(id) || relationship.cardinality === 'one-to-one' && used.has(id)) throw new RecordDomainError('Relation has a missing or duplicate target record');
    used.add(id);
  }
}

export function resolveRecordField(workbook: WorkbookModel, address: RecordFieldAddress): { sheetId: string; row: number; column: number; cell?: CellData; writable: boolean } {
  const table = workbook.getTable(address.tableId), range = table.sourceRange;
  const field = table.fields.find(field => field.id === address.fieldId), row = recordRows(workbook, table).get(address.recordId);
  if (!range || !field || row === undefined) throw new RecordDomainError('Record/Field identity is missing');
  const column = range.startColumn + field.ordinal;
  return { sheetId: range.sheetId, row, column, cell: workbook.getSheet(range.sheetId).cells.getWithoutHydration(row, column), writable: field.id !== table.recordIdFieldId && !field.calculation };
}

export function guardRecordWorksheetWrites(workbook: WorkbookModel, mutation: string, ranges: readonly import('./index').RangeRef[]): void {
  if (!['cell.set', 'cell.restore', 'range.set', 'range.paste', 'range.clear', 'range.clear.restore', 'fill.applied', 'fill.restored', 'query.load.range', 'range.move'].includes(mutation)) return;
  for (const table of workbook.dataModel.tables.values()) {
    if (!table.recordIdFieldId || !table.sourceRange) continue;
    for (const field of table.fields) {

      const column = table.sourceRange.startColumn + field.ordinal;
      if (ranges.some(range => range.sheetId === table.sourceRange!.sheetId && range.startRow <= table.sourceRange!.endRow && range.endRow > table.sourceRange!.startRow && range.startColumn <= column && range.endColumn >= column)) throw new RecordDomainError('Record fields require writes addressed by Record/Field IDs');
    }
  }
}

export function assertRecordCalculations(workbook: WorkbookModel): void {
  for (const table of workbook.dataModel.tables.values()) for (const field of table.fields) {
    const calculation = field.calculation;
    if (!calculation || calculation.kind === 'formula') continue;
    const relation = workbook.dataModel.relationships.get(calculation.relationshipId);
    const forward = calculation.direction === 'forward';
    if (!relation || table.id !== (forward ? relation.fromTableId : relation.toTableId)) throw new RecordDomainError('Calculated field relationship owner is invalid');
    const target = workbook.getTable(forward ? relation.toTableId : relation.fromTableId);
    if (!target.recordIdFieldId || !target.fields.some(candidate => candidate.id === calculation.targetFieldId)) throw new RecordDomainError('Calculated field target is missing');
  }
}

export function assertRecordFieldWrite(workbook: WorkbookModel, address: RecordFieldAddress, value: import('./index').TableScalar): void {
  const target = resolveRecordField(workbook, address);
  if (!target.writable) throw new RecordDomainError('Record field is read-only');
  for (const relation of workbook.dataModel.relationships.values()) {
    if (relation.fromTableId !== address.tableId || relation.fromFieldId !== address.fieldId) continue;
    if (value === null || value === '') continue;
    if (typeof value !== 'string' || !recordRows(workbook, workbook.getTable(relation.toTableId)).has(value)) throw new RecordDomainError('Relation target Record is missing');
    if (relation.cardinality === 'one-to-one') {
      const table = workbook.getTable(address.tableId), field = table.fields.find(field => field.id === address.fieldId)!;
      for (const [id, row] of recordRows(workbook, table)) if (id !== address.recordId && workbook.getSheet(table.sourceRange!.sheetId).cells.getWithoutHydration(row, table.sourceRange!.startColumn + field.ordinal)?.value === value) throw new RecordDomainError('Relation target Record is duplicated');
    }
  }
}
