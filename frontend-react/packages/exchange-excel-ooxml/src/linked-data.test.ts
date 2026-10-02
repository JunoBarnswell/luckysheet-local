import assert from 'node:assert/strict';
import test from 'node:test';
import { WorkbookModel, canonicalRecordFieldFormula } from '@react-sheets/core-model';
import { exportSnapshotToOoxmlBuffer } from './archive';
import { importOoxmlDocument } from './import';

test('native package metadata preserves external bindings and Record field definitions without cached values', async () => {
  const workbook = new WorkbookModel('linked-native', 'Linked native');
  const source = workbook.getSheet(workbook.primarySheetId);
  source.cells.set(0, 0, { value: 'ID' }); source.cells.set(0, 1, { value: 'Price' }); source.cells.set(0, 2, { value: 'Total' });
  source.cells.set(1, 0, { value: 'record-1' }); source.cells.set(1, 1, { value: 7 });
  const table = { id: 'records', name: 'Records', recordIdFieldId: 'record-id', sourceSheetId: source.id, sourceRange: { sheetId: source.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 2 }, fields: [{ id: 'record-id', name: 'ID', ordinal: 0, type: 'text' as const }, { id: 'record-price', name: 'Price', ordinal: 1, type: 'number' as const }, { id: 'record-total', name: 'Total', ordinal: 2, type: 'number' as const, calculation: { kind: 'formula' as const, formula: '' } }], rowCount: 1, blockSize: 1024, blocks: [], revision: 0 };
  table.fields[2]!.calculation!.formula = canonicalRecordFieldFormula(table, '=[Price]*2');
  workbook.addTable(table);
  const link = { id: 'link-1', token: 'Source.xlsx', sourceUnitId: 'source-workbook', sheets: [{ token: 'Sales', sheetId: 'source-sales-id' }] };
  workbook.dataModel.externalLinks.set(link.id, link);
  const imported = await importOoxmlDocument({ fileName: 'linked.xlsx', buffer: exportSnapshotToOoxmlBuffer(workbook.snapshot()), options: { compatibilityTarget: 'B' } });
  assert.deepEqual(imported.snapshot.dataModel.externalLinks, [link]);
  assert.deepEqual(imported.snapshot.dataModel.tables, [table]);
  assert.equal(imported.snapshot.sheets[0]?.cells['1']?.['2']?.formula, undefined);
});
