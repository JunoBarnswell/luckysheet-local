import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { createPivotMemberKey, defaultChartSubtype, planConnectorRoute, WorkbookModel } from '@react-sheets/core-model';
import { exportXlsx } from './export';
import { importXlsx } from './import';
import { scanFormulaPreserveIssues, scanSnapshotFeatures } from './feature-scan';
import { exportSnapshotToXlsxBuffer } from './archive';
import { loadOpcPackageGraph, parseLoadedXlsx, zipXlsxPartsBuffer } from './archive';
import { mapNativePivotDefinition, readNativePivotGraph } from './native-pivot';
import type { NativePivotCacheDefinition, NativePivotTableDefinition } from './types';
import { strFromU8, strToU8 } from 'fflate';

describe('exchange-excel-ooxml', () => {
  it('preserves native Pivot error cache items as typed error members', () => {
    const main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    const graph = readNativePivotGraph({
      files: {
        'xl/workbook.xml': strToU8(`<workbook xmlns="${main}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1"/></sheets><pivotCaches count="1"><pivotCache cacheId="1" r:id="rIdCache"/></pivotCaches></workbook>`),
        'xl/pivotCache/pivotCacheDefinition1.xml': strToU8(`<pivotCacheDefinition xmlns="${main}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><cacheSource type="worksheet"><worksheetSource ref="A1:A2" sheet="Sheet1"/></cacheSource><cacheFields count="1"><cacheField name="Error"><sharedItems containsError="1"><e v="#N/A"/></sharedItems></cacheField></cacheFields></pivotCacheDefinition>`),
        'xl/pivotCache/pivotCacheRecords1.xml': strToU8(`<pivotCacheRecords xmlns="${main}" count="1"><r><e v="#N/A"/></r></pivotCacheRecords>`),
        'xl/worksheets/sheet1.xml': strToU8(`<worksheet xmlns="${main}"/>`),
      },
      relationships: {
        'xl/workbook.xml': [{ id: 'rIdCache', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition', target: 'pivotCache/pivotCacheDefinition1.xml' }],
        'xl/pivotCache/pivotCacheDefinition1.xml': [{ id: 'rIdRecords', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords', target: 'pivotCacheRecords1.xml' }],
        'xl/worksheets/sheet1.xml': [],
      },
      sheetPartById: { 'sheet-1': 'xl/worksheets/sheet1.xml' },
      dateSystem: '1900',
    });
    assert.equal(graph.caches[0]?.fields[0]?.dataType, 'error');
    assert.deepEqual(graph.caches[0]?.fields[0]?.sharedItems, [{ kind: 'error', code: '#N/A' }]);
    assert.equal(graph.caches[0]?.recordCount, 1);
  });

  it('scans workbook features for compatibility reporting', () => {
    const workbook = new WorkbookModel('wb-xlsx', 'XLSX');
    const sheet = workbook.getSheet(workbook.primarySheetId);
    sheet.cells.set(0, 0, { value: 10, formula: '=SUM(A2:A3)' });
    sheet.merges.push({
      range: { sheetId: sheet.id, startRow: 1, endRow: 1, startColumn: 0, endColumn: 1 },
      anchor: { row: 1, column: 0 },
    });
    const features = scanSnapshotFeatures(workbook.snapshot());
    assert.ok(features.includes('cells'));
    assert.ok(features.includes('formulas'));
    assert.ok(features.includes('merges'));
  });

  it('reports a preserved Excel data-table formula even when OOXML has no formula body', () => {
    const workbook = new WorkbookModel('wb-data-table-diagnostic', 'Data table diagnostic');
    const sheet = workbook.getSheet(workbook.primarySheetId);
    sheet.cells.set(4, 2, {
      value: 42,
      formulaMetadata: {
        kind: 'dataTable', range: 'C5:D7', preservedOnly: true,
        reason: 'Excel data-table formulas are preserved from the source package',
      },
    });
    assert.deepEqual(scanFormulaPreserveIssues(workbook.snapshot()), [
      {
        level: 'C', severity: 'warning', feature: 'data-table-formula', location: 'Sheet1!C5',
        message: 'Formula is preserved-only: Excel data-table formulas are preserved from the source package',
        preserved: true, status: 'preserved-only', reason: 'Excel data-table formulas are preserved from the source package',
      },
    ]);
  });

  it('scans canonical drawing payloads instead of removed per-kind collections', () => {
    const workbook = new WorkbookModel('wb-drawing-features', 'Drawing Features');
    const sheet = workbook.getSheet(workbook.primarySheetId);
    sheet.drawings.push(
      {
        id: 'chart-drawing',
        sheetId: sheet.id,
        kind: 'chart',
        anchor: { kind: 'absolute' },
        transform: { x: 0, y: 0, width: 200, height: 120 },
        zIndex: 0,
        payloadId: 'chart-payload',
      },
      {
        id: 'image-drawing',
        sheetId: sheet.id,
        kind: 'image',
        anchor: { kind: 'one-cell', row: 0, column: 0 },
        transform: { x: 0, y: 0, width: 80, height: 40 },
        zIndex: 1,
        payloadId: 'image-payload',
      },
    );
    sheet.drawingPayloads.set('chart-payload', {
      kind: 'chart',
      chartId: 'chart-payload',
      chartType: 'line',
      subtype: 'line',
      source: { kind: 'worksheet-ranges', ranges: [{ sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 }] },
      elements: { hiddenData: 'show' },
    });
    sheet.drawingPayloads.set('image-payload', {
      kind: 'image',
      asset: { schema: 'AssetRef', assetId: 'asset-test', contentHash: 'a'.repeat(64), mimeType: 'image/png', byteLength: 2 },
    });

    const features = scanSnapshotFeatures(workbook.snapshot());
    assert.ok(features.includes('charts'));
    assert.ok(features.includes('images'));
  });

  it('round-trips snapshot through xlsx archive import/export', async () => {
    const workbook = new WorkbookModel('wb-roundtrip', 'Roundtrip');
    workbook.getSheet(workbook.primarySheetId).cells.set(0, 0, { value: 'hello' });
    workbook.getSheet(workbook.primarySheetId).cells.set(1, 0, { value: 42, formula: '=A1&"!"' });
    const snapshot = workbook.snapshot();
    const buffer = exportSnapshotToXlsxBuffer(snapshot);
    const imported = await importXlsx({
      fileName: 'roundtrip.xlsx',
      buffer,
      options: { compatibilityTarget: 'B' },
    });
    assert.equal(imported.snapshot.sheets.length, snapshot.sheets.length);
    assert.equal(imported.report.schema, 'CompatibilityReport');
    assert.equal(imported.nativePackage.schema, 'NativePackageState');
    assert.equal(imported.nativePackage.checksum.length, 64);
    const exported = await exportXlsx({
      snapshot: imported.snapshot,
      fileName: 'roundtrip.xlsx',
      options: { compatibilityTarget: 'B' },
    });
    assert.ok(exported.buffer.byteLength > 0);
    assert.equal(exported.fileName, 'roundtrip.xlsx');
  });

  it('preserves canonical connector routes, group ownership, and worksheet snap settings in metadata', async () => {
    const workbook = new WorkbookModel('wb-shape-contracts', 'Shape contracts');
    const sheet = workbook.getSheet(workbook.primarySheetId);
    sheet.drawings.push(
      { id: 'shape-a', sheetId: sheet.id, kind: 'shape', anchor: { kind: 'absolute' }, transform: { x: 20, y: 20, width: 80, height: 40, rotation: 0 }, zIndex: 0, payloadId: 'shape-a-payload' },
      { id: 'shape-b', sheetId: sheet.id, kind: 'shape', anchor: { kind: 'absolute' }, transform: { x: 180, y: 80, width: 80, height: 40, rotation: 0 }, zIndex: 1, payloadId: 'shape-b-payload' },
      { id: 'connector-a-b', sheetId: sheet.id, kind: 'connector', anchor: { kind: 'absolute' }, transform: { x: 0, y: 0, width: 0, height: 0, rotation: 0 }, zIndex: 2, payloadId: 'connector-a-b-payload' },
    );
    sheet.drawingPayloads.set('shape-a-payload', { kind: 'shape', type: 'rectangle', fill: '#fff', stroke: '#000' });
    sheet.drawingPayloads.set('shape-b-payload', { kind: 'shape', type: 'ellipse', fill: '#fff', stroke: '#000' });
    const connector = sheet.drawings[2]!;
    const connectorPayload = { kind: 'connector' as const, connectorType: 'elbow' as const, start: { drawingId: 'shape-a', connectionPoint: 'right' as const }, end: { drawingId: 'shape-b', connectionPoint: 'left' as const }, stroke: '#2563eb', startArrowhead: 'none' as const, endArrowhead: 'triangle' as const, route: { points: [{ x: 0, y: 0 }, { x: 1, y: 1 }] } };
    const planned = planConnectorRoute(sheet, connector, connectorPayload);
    connector.transform = planned.transform;
    sheet.drawingPayloads.set(connector.payloadId, planned.payload);
    sheet.drawingGroups.push({ id: 'shape-group', sheetId: sheet.id, memberDrawingIds: ['shape-a', 'shape-b'] });
    sheet.snapSettings = { enabled: true, snapToGrid: false, snapToShape: true, gridSize: 12 };
    const imported = await importXlsx({ fileName: 'shape-contracts.xlsx', buffer: exportSnapshotToXlsxBuffer(workbook.snapshot()), options: { compatibilityTarget: 'B' } });
    const importedSheet = imported.snapshot.sheets[0]!;
    assert.deepEqual(importedSheet.drawingGroups, sheet.drawingGroups);
    assert.deepEqual(importedSheet.snapSettings, sheet.snapSettings);
    assert.deepEqual(importedSheet.drawings.find((drawing) => drawing.id === connector.id)?.transform, connector.transform);
    assert.deepEqual(importedSheet.drawingPayloads[connector.payloadId], planned.payload);
  });

  it('round-trips canonical and imported font families through OOXML styles', async () => {
    const workbook = new WorkbookModel('wb-font-family-roundtrip', 'Font families');
    const sheet = workbook.getSheet(workbook.primarySheetId);
    sheet.cells.set(0, 0, { value: 'listed', style: { fontFamily: '  sEgOe Ui  ' } });
    sheet.cells.set(1, 0, { value: 'imported', style: { fontFamily: '  Imported Local Font  ' } });

    const imported = await importXlsx({
      fileName: 'font-family.xlsx',
      buffer: exportSnapshotToXlsxBuffer(workbook.snapshot()),
      options: { compatibilityTarget: 'B' },
    });
    const importedSheet = imported.snapshot.sheets[0]!;
    assert.equal(importedSheet.cells['0']?.['0']?.style?.fontFamily, 'Segoe UI');
    assert.equal(importedSheet.cells['1']?.['0']?.style?.fontFamily, 'Imported Local Font');

    const exported = await exportXlsx({
      snapshot: imported.snapshot,
      fileName: 'font-family.xlsx',
      options: { compatibilityTarget: 'B' },
    });
    const packageGraph = loadOpcPackageGraph(exported.buffer);
    const stylesXml = strFromU8(packageGraph.files['xl/styles.xml']!);
    assert.match(stylesXml, /name val="Imported Local Font"/);
    assert.match(stylesXml, /name val="Segoe UI"/);
  });

  it('round-trips side-aware border topology without changing non-border style', async () => {
    const workbook = new WorkbookModel('wb-border-roundtrip', 'Borders');
    const sheet = workbook.getSheet(workbook.primarySheetId);
    const line = { style: 'thin' as const, color: '#334155' };
    sheet.cells.set(0, 0, { value: 'keep', style: { bold: true, borders: { top: line, left: line } } });
    sheet.cells.set(0, 1, { value: 2, style: { borders: { top: line, right: line } } });
    sheet.cells.set(1, 0, { value: 3, style: { borders: { bottom: line, left: line } } });
    sheet.cells.set(1, 1, { value: 4, style: { borders: { bottom: line, right: line } } });
    const imported = await importXlsx({
      fileName: 'borders.xlsx',
      buffer: exportSnapshotToXlsxBuffer(workbook.snapshot()),
      options: { compatibilityTarget: 'B' },
    });
    const importedSheet = imported.snapshot.sheets[0]!;
    const topLeft = importedSheet.cells['0']?.['0'];
    const bottomRight = importedSheet.cells['1']?.['1'];
    assert.deepEqual(topLeft?.style?.borders, { top: line, left: line });
    assert.equal(topLeft?.style?.bold, true);
    assert.deepEqual(bottomRight?.style?.borders, { bottom: line, right: line });
  });

  it('round-trips the canonical extended alignment contract without lossy flattening', async () => {
    const workbook = new WorkbookModel('wb-alignment-roundtrip', 'Alignment');
    const sheet = workbook.getSheet(workbook.primarySheetId);
    sheet.cells.set(0, 0, {
      value: 'Across',
      style: {
        horizontalAlignment: 'centerContinuous',
        verticalAlignment: 'distributed',
        shrinkToFit: true,
        indent: 3,
        readingOrder: 'rtl',
        textOrientation: 'stacked',
      },
    });

    const imported = await importXlsx({
      fileName: 'alignment.xlsx',
      buffer: exportSnapshotToXlsxBuffer(workbook.snapshot()),
      options: { compatibilityTarget: 'B' },
    });
    const style = imported.snapshot.sheets[0]?.cells['0']?.['0']?.style;
    assert.equal(style?.horizontalAlignment, 'centerContinuous');
    assert.equal(style?.verticalAlignment, 'distributed');
    assert.equal(style?.shrinkToFit, true);
    assert.equal(style?.indent, 3);
    assert.equal(style?.readingOrder, 'rtl');
    assert.equal(style?.textOrientation, 'stacked');
  });

  it('preserves unsupported native alignment attributes and rejects malformed alignment values', async () => {
    const workbook = new WorkbookModel('wb-alignment-unsupported', 'Alignment Unsupported');
    const sheet = workbook.getSheet(workbook.primarySheetId);
    sheet.cells.set(0, 0, { value: 'native', style: { horizontalAlignment: 'left' } });
    const output = loadOpcPackageGraph(exportSnapshotToXlsxBuffer(workbook.snapshot()));
    const stylesXml = strFromU8(output.packageGraph.parts['xl/styles.xml']!);
    const alignment = stylesXml.match(/<alignment\s+[^>]*\/>/)?.[0];
    assert.ok(alignment, 'export must contain a cell alignment node');
    output.packageGraph.parts['xl/styles.xml'] = strToU8(stylesXml.replace(
      alignment,
      '<alignment horizontal="vendorAlignment" vertical="vendorVertical" foo="preserve-me" shrinkToFit="1"/>',
    ));

    const imported = await importXlsx({
      fileName: 'alignment-unsupported.xlsx',
      buffer: zipXlsxPartsBuffer(output.packageGraph.parts),
      options: { compatibilityTarget: 'B' },
    });
    const style = imported.snapshot.sheets[0]?.cells['0']?.['0']?.style;
    assert.equal(style?.unsupportedAlignment?.horizontal, 'vendorAlignment');
    assert.equal(style?.unsupportedAlignment?.vertical, 'vendorVertical');
    assert.equal(style?.unsupportedAlignment?.attributes?.foo, 'preserve-me');
    assert.equal(style?.shrinkToFit, true);
    const preserved = await exportXlsx({
      snapshot: imported.snapshot,
      nativePackage: imported.nativePackage,
      fileName: 'alignment-unsupported-roundtrip.xlsx',
      options: { compatibilityTarget: 'B' },
    });
    const preservedStyles = strFromU8(loadOpcPackageGraph(preserved.buffer).files['xl/styles.xml']!);
    assert.match(preservedStyles, /horizontal="vendorAlignment"/);
    assert.match(preservedStyles, /vertical="vendorVertical"/);
    assert.match(preservedStyles, /foo="preserve-me"/);

    const malformed = strFromU8(output.packageGraph.parts['xl/styles.xml']!).replace('shrinkToFit="1"', 'shrinkToFit="maybe"');
    output.packageGraph.parts['xl/styles.xml'] = strToU8(malformed);
    await assert.rejects(
      importXlsx({ fileName: 'alignment-malformed.xlsx', buffer: zipXlsxPartsBuffer(output.packageGraph.parts), options: { compatibilityTarget: 'B' } }),
      /Invalid OOXML alignment shrinkToFit/,
    );
  });

  it('round-trips native sheet protection allow flags and cell protection styles', async () => {
    const workbook = new WorkbookModel('wb-protection-roundtrip', 'Protection');
    const sheet = workbook.getSheet(workbook.primarySheetId);
    sheet.protectionRules.push({
      id: 'sheet-protection', scope: 'sheet', sheetId: sheet.id, locked: true,
      allow: { selectLocked: true, selectUnlocked: true, formatCells: true, sort: true, autoFilter: false },
    });
    sheet.cells.set(0, 0, { value: '=SUM(A2:A3)', formula: '=SUM(A2:A3)', style: { locked: false, formulaHidden: true } });
    const imported = await importXlsx({
      fileName: 'protection.xlsx',
      buffer: exportSnapshotToXlsxBuffer(workbook.snapshot()),
      options: { compatibilityTarget: 'B' },
    });
    const importedSheet = imported.snapshot.sheets[0]!;
    assert.deepEqual(importedSheet.protectionRules?.[0]?.allow, {
      selectLocked: true, selectUnlocked: true, formatCells: true, insertRows: false,
      insertColumns: false, deleteRows: false, deleteColumns: false, sort: true, autoFilter: false, editObjects: false,
    });
    assert.equal(importedSheet.cells['0']?.['0']?.style?.locked, false);
    assert.equal(importedSheet.cells['0']?.['0']?.style?.formulaHidden, true);
  });

  it('rejects custom metadata drawing ranges that exceed bounded render work', async () => {
    const workbook = new WorkbookModel('wb-malicious-metadata', 'Malicious metadata');
    const sheet = workbook.getSheet(workbook.primarySheetId);
    sheet.drawings.push({
      id: 'camera', sheetId: sheet.id, kind: 'camera', payloadId: 'camera', anchor: { kind: 'absolute' },
      transform: { x: 0, y: 0, width: 100, height: 100 }, zIndex: 0,
    });
    sheet.drawingPayloads.set('camera', {
      kind: 'camera', sourceRange: { sheetId: sheet.id, startRow: 0, endRow: 999_999, startColumn: 0, endColumn: 999_999 }, refreshPolicy: 'live',
    });
    const buffer = exportSnapshotToXlsxBuffer(workbook.snapshot());

    await assert.rejects(
      importXlsx({ fileName: 'malicious.xlsx', buffer, options: { compatibilityTarget: 'B' } }),
      /Camera source range/,
    );
  });

  it('materializes icon filter metadata from the numeric cells in its conditional-format range', async () => {
    const workbook = new WorkbookModel('wb-icon-filter', 'Icon Filter');
    const sheet = workbook.getSheet(workbook.primarySheetId);
    sheet.cells.set(0, 0, { value: 'Amount' });
    sheet.cells.set(1, 0, { value: 10 });
    sheet.cells.set(2, 0, { value: 20 });
    const output = loadOpcPackageGraph(exportSnapshotToXlsxBuffer(workbook.snapshot()));
    output.packageGraph.parts['xl/worksheets/sheet1.xml'] = strToU8(strFromU8(output.packageGraph.parts['xl/worksheets/sheet1.xml']!).replace(
      '</worksheet>',
      '<conditionalFormatting sqref="A2:A3"><cfRule type="iconSet" priority="1"><iconSet iconSet="3TrafficLights1"><cfvo type="percent" val="0"/><cfvo type="percent" val="50"/><cfvo type="percent" val="100"/></iconSet></cfRule></conditionalFormatting><autoFilter ref="A1:A3"><filterColumn colId="0"><iconFilter iconSet="3TrafficLights1" iconId="1"/></filterColumn></autoFilter></worksheet>',
    ));

    const imported = await importXlsx({ fileName: 'icon-filter.xlsx', buffer: zipXlsxPartsBuffer(output.packageGraph.parts), options: { compatibilityTarget: 'B' } });

    assert.deepEqual(imported.snapshot.sheets[0]?.cells['1']?.['0']?.filterMetadata?.icon, { iconSet: '3TrafficLights1', iconId: 0 });
    assert.deepEqual(imported.snapshot.sheets[0]?.cells['2']?.['0']?.filterMetadata?.icon, { iconSet: '3TrafficLights1', iconId: 2 });
  });

  it('bounds icon filter metadata work by materialized cells for a full-sheet conditional-format range', async () => {
    const workbook = new WorkbookModel('wb-icon-filter-full-sheet', 'Icon Filter Full Sheet');
    const sheet = workbook.getSheet(workbook.primarySheetId);
    sheet.cells.set(0, 0, { value: 'Amount' });
    sheet.cells.set(1, 0, { value: 10 });
    const output = loadOpcPackageGraph(exportSnapshotToXlsxBuffer(workbook.snapshot()));
    output.packageGraph.parts['xl/worksheets/sheet1.xml'] = strToU8(strFromU8(output.packageGraph.parts['xl/worksheets/sheet1.xml']!).replace(
      '</worksheet>',
      '<conditionalFormatting sqref="A1:XFD1048576"><cfRule type="iconSet" priority="1"><iconSet iconSet="3TrafficLights1"><cfvo type="percent" val="0"/><cfvo type="percent" val="50"/><cfvo type="percent" val="100"/></iconSet></cfRule></conditionalFormatting><autoFilter ref="A1:A2"><filterColumn colId="0"><iconFilter iconSet="3TrafficLights1" iconId="0"/></filterColumn></autoFilter></worksheet>',
    ));

    const imported = await importXlsx({ fileName: 'full-sheet-icon-filter.xlsx', buffer: zipXlsxPartsBuffer(output.packageGraph.parts), options: { compatibilityTarget: 'B' } });

    assert.deepEqual(imported.snapshot.sheets[0]?.cells['1']?.['0']?.filterMetadata?.icon, { iconSet: '3TrafficLights1', iconId: 0 });
  });

  it('round-trips worksheet tables and emits table relationship/content-type parts', async () => {
    const workbook = new WorkbookModel('wb-table', 'Table');
    const sheet = workbook.getSheet(workbook.primarySheetId);
    sheet.cells.set(0, 0, { value: 'Category' });
    sheet.cells.set(0, 1, { value: 'Amount' });
    sheet.cells.set(1, 0, { value: 'A' });
    sheet.cells.set(1, 1, { value: 10 });
    sheet.sheetTables.push({
      id: 'table-1',
      sheetId: sheet.id,
      name: 'SalesTable',
      range: { sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 },
      hasHeaderRow: true,
      hasTotalRow: false,
      showBandedRows: true,
      showBandedColumns: false,
      showFirstColumn: false,
      showLastColumn: false,
      showFilterButton: true,
      autoExpand: 'both',
      columns: [{ id: 'category', name: 'Category' }, { id: 'amount', name: 'Amount' }],
      styleName: 'TableStyleMedium2',
    });
    const imported = await importXlsx({ fileName: 'table.xlsx', buffer: exportSnapshotToXlsxBuffer(workbook.snapshot()), options: { compatibilityTarget: 'B' } });
    assert.equal(imported.snapshot.sheets[0]?.sheetTables?.[0]?.name, 'SalesTable');
    assert.equal(imported.report.issues.find((issue) => issue.feature === 'tables')?.status, 'editable');
    const exported = await exportXlsx({ snapshot: imported.snapshot, nativePackage: imported.nativePackage, fileName: 'table.xlsx', options: { compatibilityTarget: 'B' } });
    const output = loadOpcPackageGraph(exported.buffer);
    assert.ok(output.files['xl/tables/table1.xml']);
    assert.match(strFromU8(output.files['[Content_Types].xml']!), /spreadsheetml\.table\+xml/);
    assert.match(strFromU8(output.files['xl/worksheets/_rels/sheet1.xml.rels']!), /\/table/);
  });

  it('writes hyperlinks from the canonical worksheet hyperlink collection', async () => {
    const workbook = new WorkbookModel('wb-links', 'Links');
    const sheet = workbook.getSheet(workbook.primarySheetId);
    const targetSheet = workbook.addSheet('sheet-target', 'Target', 20, 20);
    workbook.setDefinedName({ name: 'SalesTotal', formula: '=Sheet1!A1', scope: 'workbook' });
    sheet.cells.set(0, 0, { value: 'OpenAI' });
    sheet.hyperlinks.set('0:0', { id: 'link-1', target: { kind: 'url', url: 'https://openai.com/' }, tooltip: 'Open' });
    sheet.hyperlinks.set('1:0', { id: 'link-2', target: { kind: 'email', address: 'team@example.com', subject: 'Review' } });
    sheet.hyperlinks.set('2:0', { id: 'link-3', target: { kind: 'sheet', sheetId: targetSheet.id, address: 'B2' } });
    sheet.hyperlinks.set('3:0', { id: 'link-4', target: { kind: 'name', name: 'SalesTotal' } });
    const buffer = exportSnapshotToXlsxBuffer(workbook.snapshot());
    const emitted = loadOpcPackageGraph(buffer);
    assert.match(strFromU8(emitted.files['xl/worksheets/sheet1.xml']!), /<hyperlink ref="A1"/);
    assert.match(strFromU8(emitted.files['xl/worksheets/_rels/sheet1.xml.rels']!), /https:\/\/openai\.com\//);
    const imported = await importXlsx({ fileName: 'links.xlsx', buffer, options: { compatibilityTarget: 'B' } });
    assert.equal(imported.snapshot.sheets[0]?.hyperlinks?.[0]?.hyperlink.target.kind, 'url');
    assert.deepEqual(imported.snapshot.sheets[0]?.hyperlinks?.[1]?.hyperlink.target, { kind: 'email', address: 'team@example.com', subject: 'Review' });
    assert.deepEqual(imported.snapshot.sheets[0]?.hyperlinks?.[2]?.hyperlink.target, { kind: 'sheet', sheetId: 'sheet-target', address: 'B2' });
    assert.deepEqual(imported.snapshot.sheets[0]?.hyperlinks?.[3]?.hyperlink.target, { kind: 'name', name: 'SalesTotal' });
  });

  it('writes a canonical table Pivot as a native cache and table graph', async () => {
    const workbook = new WorkbookModel('wb-new-pivot', 'New Pivot');
    const sheet = workbook.getSheet(workbook.primarySheetId);
    sheet.cells.set(0, 0, { value: 'Category' });
    sheet.cells.set(0, 1, { value: 'Amount' });
    sheet.cells.set(1, 0, { value: 'A' });
    sheet.cells.set(1, 1, { value: 10 });
    sheet.cells.set(2, 0, { value: 'B' });
    sheet.cells.set(2, 1, { value: 20 });
    sheet.sheetTables.push({
      id: 'table-1', sheetId: sheet.id, name: 'SalesTable',
      range: { sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 },
      hasHeaderRow: true, hasTotalRow: false, showBandedRows: true, showBandedColumns: false, showFirstColumn: false, showLastColumn: false, showFilterButton: true, autoExpand: 'both',
      columns: [{ id: 'category', name: 'Category' }, { id: 'amount', name: 'Amount' }], styleName: 'TableStyleMedium2',
    });
    sheet.pivots.push({
      schema: 'PivotDefinition', id: 'sales-pivot', source: { kind: 'table', tableId: 'table-1' }, target: { sheetId: sheet.id, anchor: { row: 5, column: 0 } },
      fieldCatalog: { schema: 'PivotFieldCatalog', fields: [
        { fieldId: 'category', name: 'Category', dataType: 'text', ordinal: 0 },
        { fieldId: 'amount', name: 'Amount', dataType: 'number', ordinal: 1 },
      ] },
      layout: { rows: [{ fieldId: 'category' }], columns: [], filters: [], allowMultipleFiltersPerField: true, collation: { locale: 'en-US', sensitivity: 'variant', numeric: false, caseFirst: 'false' }, values: [{ valueId: `value:${'amount'}`, fieldId: 'amount', summarizeBy: 'sum', numberFormat: '0.000' }], subtotalLocation: 'bottom', showRowGrandTotals: true, showColumnGrandTotals: true, reportLayout: 'compact' },
      refreshPolicy: { mode: 'on-change', preserveFormatting: true, refreshOnLoad: true },
      presentation: {
        styleName: 'PivotStyleMedium4',
        styleOptions: { showRowHeaders: false, showColumnHeaders: true, showRowStripes: true, showColumnStripes: true, showLastColumn: true },
        displayOptions: { fillEmptyCells: true, emptyCellText: '—', showErrorValues: false, errorCellText: 'ERR', showFieldHeaders: false, autoFitColumnsOnUpdate: false },
      },
    });
    const output = loadOpcPackageGraph(exportSnapshotToXlsxBuffer(workbook.snapshot()));
    assert.equal(output.packageGraph.nativePivotGraph?.caches.length, 1);
    assert.equal(output.packageGraph.nativePivotGraph?.tables.length, 1);
    const cacheXml = strFromU8(output.files['xl/pivotCache/pivotCacheDefinition1.xml']!);
    assert.match(cacheXml, /worksheetSource name="SalesTable"/);
    assert.match(cacheXml, /refreshOnLoad="1"/);
    assert.match(cacheXml, /refreshOnSave="1"/);
    assert.match(strFromU8(output.files['xl/pivotCache/pivotCacheDefinition1.xml']!), /<cacheSource type="worksheet">/);
    assert.match(strFromU8(output.files['xl/worksheets/sheet1.xml']!), /pivotTableParts/);
    assert.match(strFromU8(output.files['xl/pivotTables/pivotTable1.xml']!), /name="PivotStyleMedium4" showRowHeaders="0" showColHeaders="1" showRowStripes="1" showColStripes="1" showLastColumn="1"/);
    assert.match(strFromU8(output.files['xl/pivotTables/pivotTable1.xml']!), /showHeaders="0"/);
    assert.match(strFromU8(output.files['xl/pivotTables/pivotTable1.xml']!), /showMissing="1" missingCaption="—" showError="0" errorCaption="ERR" preserveFormatting="1"/);
    assert.match(strFromU8(output.files['xl/pivotTables/pivotTable1.xml']!), /<dataField[^>]*numFmtId="164"/);
    assert.match(strFromU8(output.files['xl/styles.xml']!), /<numFmt numFmtId="164" formatCode="0\.000"\/>/);
    const imported = parseLoadedXlsx(output).snapshot;
    assert.equal(imported.sheets[0]?.pivots[0]?.source.kind, 'table');
    assert.equal(imported.sheets[0]?.pivots[0]?.layout.values[0]?.numberFormat, '0.000');
    assert.deepEqual(imported.sheets[0]?.pivots[0]?.presentation, {
      styleName: 'PivotStyleMedium4',
      styleOptions: { showRowHeaders: false, showColumnHeaders: true, showRowStripes: true, showColumnStripes: true, showLastColumn: true },
      displayOptions: { fillEmptyCells: true, emptyCellText: '—', showErrorValues: false, errorCellText: 'ERR', showFieldHeaders: false, autoFitColumnsOnUpdate: true },
    });
  });

  it('rejects malformed Pivot value number formats before native package mutation', () => {
    const workbook = new WorkbookModel('wb-invalid-pivot-number-format', 'Invalid Pivot Number Format');
    const sheet = workbook.getSheet(workbook.primarySheetId);
    sheet.cells.set(0, 0, { value: 'Amount' });
    sheet.cells.set(1, 0, { value: 10 });
    sheet.pivots.push({
      schema: 'PivotDefinition', id: 'invalid-pivot-number-format',
      source: { kind: 'worksheet-range', range: { sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 } },
      target: { sheetId: sheet.id, anchor: { row: 3, column: 0 } },
      fieldCatalog: { schema: 'PivotFieldCatalog', fields: [{ fieldId: 'amount', name: 'Amount', dataType: 'number', ordinal: 0 }] },
      layout: { rows: [], columns: [], filters: [], allowMultipleFiltersPerField: true, collation: { locale: 'en-US', sensitivity: 'variant', numeric: false, caseFirst: 'false' }, values: [{ valueId: `value:${'amount'}`, fieldId: 'amount', summarizeBy: 'sum', numberFormat: '[Red' }], subtotalLocation: 'bottom', showRowGrandTotals: true, showColumnGrandTotals: true, reportLayout: 'compact' },
      refreshPolicy: { mode: 'manual', preserveFormatting: true, refreshOnLoad: false },
    });
    assert.throws(() => exportSnapshotToXlsxBuffer(workbook.snapshot()), /unterminated/);
  });

  it('round-trips native Difference From base field and typed base item coordinates', () => {
    const workbook = new WorkbookModel('wb-pivot-difference', 'Pivot Difference');
    const sheet = workbook.getSheet(workbook.primarySheetId);
    [['Category', 'Amount'], ['A', 10], ['B', 20]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
    sheet.pivots.push({
      schema: 'PivotDefinition', id: 'pivot-difference', source: { kind: 'worksheet-range', range: { sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 } }, target: { sheetId: sheet.id, anchor: { row: 5, column: 0 } },
      fieldCatalog: { schema: 'PivotFieldCatalog', fields: [{ fieldId: 'category', name: 'Category', dataType: 'text', ordinal: 0, values: ['A', 'B'] }, { fieldId: 'amount', name: 'Amount', dataType: 'number', ordinal: 1 }] },
      layout: { rows: [{ fieldId: 'category' }], columns: [], filters: [], allowMultipleFiltersPerField: true, collation: { locale: 'en-US', sensitivity: 'variant', numeric: false, caseFirst: 'false' }, values: [{ valueId: 'value:amount', fieldId: 'amount', summarizeBy: 'sum', showAs: { kind: 'difference', baseFieldId: 'category', baseItem: { type: 'text', value: 'A' } } }], subtotalLocation: 'bottom', showRowGrandTotals: true, showColumnGrandTotals: true, reportLayout: 'compact' },
      refreshPolicy: { mode: 'on-change', preserveFormatting: true, refreshOnLoad: true },
    });
    const output = loadOpcPackageGraph(exportSnapshotToXlsxBuffer(workbook.snapshot()));
    const tableXml = strFromU8(output.files['xl/pivotTables/pivotTable1.xml']!);
    assert.match(tableXml, /showDataAs="difference"/);
    assert.match(tableXml, /baseField="0"/);
    assert.match(tableXml, /baseItem="0"/);
    const imported = parseLoadedXlsx(output).snapshot;
    assert.deepEqual(imported.sheets[0]?.pivots[0]?.layout.values[0]?.showAs, { kind: 'difference', baseFieldId: 'native:cache:1:field:0', baseItem: { type: 'text', value: 'A' } });
  });

  it('round-trips date, numeric and manual Pivot cache grouping through native fieldGroup metadata', () => {
    const workbook = new WorkbookModel('wb-pivot-groups', 'Pivot Groups');
    const sheet = workbook.getSheet(workbook.primarySheetId);
    const rows = [
      ['Date', 'Amount', 'Category'],
      [45292, 10, 'A'],
      [45323, 20, 'B'],
      [45657, 30, 'C'],
    ];
    rows.forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
    sheet.pivots.push({
      schema: 'PivotDefinition', id: 'grouped-pivot', source: { kind: 'worksheet-range', range: { sheetId: sheet.id, startRow: 0, endRow: 3, startColumn: 0, endColumn: 2 } }, target: { sheetId: sheet.id, anchor: { row: 6, column: 0 } },
      fieldCatalog: { schema: 'PivotFieldCatalog', fields: [
        { fieldId: 'date', name: 'Date', dataType: 'date', ordinal: 0 },
        { fieldId: 'amount', name: 'Amount', dataType: 'number', ordinal: 1 },
        { fieldId: 'category', name: 'Category', dataType: 'text', ordinal: 2 },
      ] },
      layout: {
        rows: [{ fieldId: 'date', group: { kind: 'date', unit: 'month', start: 45292, end: 45657 } }, { fieldId: 'category', group: { kind: 'manual', groups: [{ groupId: 'ab', name: 'AB', items: [createPivotMemberKey('A'), createPivotMemberKey('B')] }, { groupId: 'c', name: 'C', items: [createPivotMemberKey('C')] }] } }],
        columns: [{ fieldId: 'amount', group: { kind: 'number', interval: 10, start: 0, end: 100 } }],
        filters: [], allowMultipleFiltersPerField: true, collation: { locale: 'en-US', sensitivity: 'variant', numeric: false, caseFirst: 'false' }, values: [{ valueId: `value:${'amount'}`, fieldId: 'amount', summarizeBy: 'sum' }], subtotalLocation: 'bottom', showRowGrandTotals: true, showColumnGrandTotals: true, reportLayout: 'compact',
      },
      refreshPolicy: { mode: 'on-change', preserveFormatting: true, refreshOnLoad: true },
    });
    const output = loadOpcPackageGraph(exportSnapshotToXlsxBuffer(workbook.snapshot()));
    const cacheXml = strFromU8(output.files['xl/pivotCache/pivotCacheDefinition1.xml']!);
    assert.match(cacheXml, /<fieldGroup base="0"><rangePr groupBy="months" startDate="2024-01-01T00:00:00.000Z" endDate="2024-12-31T00:00:00.000Z"\/>/);
    assert.match(cacheXml, /<fieldGroup base="1"><rangePr groupBy="range" groupInterval="10" startNum="0" endNum="100"\/>/);
    assert.match(cacheXml, /<fieldGroup base="2"><discretePr count="3"><x v="0"\/><x v="0"\/><x v="1"\/><\/discretePr><groupItems count="2"><s v="AB"\/><s v="C"\/><\/groupItems><\/fieldGroup>/);
    const imported = parseLoadedXlsx(output).snapshot;
    const pivot = imported.sheets[0]?.pivots[0];
    assert.equal(pivot?.layout.rows[0]?.group?.kind, 'date');
    assert.deepEqual(pivot?.layout.rows[0]?.group, { kind: 'date', unit: 'month', start: '2024-01-01T00:00:00.000Z', end: '2024-12-31T00:00:00.000Z' });
    assert.deepEqual(pivot?.layout.columns[0]?.group, { kind: 'number', interval: 10, start: 0, end: 100 });
    assert.deepEqual(pivot?.layout.rows[1]?.group, { kind: 'manual', groups: [
      { groupId: 'native:cache:1:field:2:group:0', name: 'AB', items: [createPivotMemberKey('A'), createPivotMemberKey('B')] },
      { groupId: 'native:cache:1:field:2:group:1', name: 'C', items: [createPivotMemberKey('C')] },
    ] });
  });

  it('writes and validates canonical Slicer and Timeline native parts', async () => {
    const workbook = new WorkbookModel('wb-native-controls', 'Native Controls');
    const sheet = workbook.getSheet(workbook.primarySheetId);
    sheet.cells.set(0, 0, { value: 'Category' });
    sheet.cells.set(0, 1, { value: 'Date' });
    sheet.cells.set(0, 2, { value: 'Amount' });
    sheet.cells.set(1, 0, { value: 'A' });
    sheet.cells.set(1, 1, { value: 45292 });
    sheet.cells.set(1, 2, { value: 10 });
    sheet.pivots.push({
      schema: 'PivotDefinition', id: 'control-pivot', source: { kind: 'worksheet-range', range: { sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 2 } }, target: { sheetId: sheet.id, anchor: { row: 5, column: 0 } },
      fieldCatalog: { schema: 'PivotFieldCatalog', fields: [
        { fieldId: 'category', name: 'Category', dataType: 'text', ordinal: 0 },
        { fieldId: 'date', name: 'Date', dataType: 'date', ordinal: 1, values: ['1890-01-01T00:00:00Z', '2201-12-31T00:00:00Z'] },
        { fieldId: 'amount', name: 'Amount', dataType: 'number', ordinal: 2 },
      ] },
      layout: { rows: [{ fieldId: 'category' }], columns: [], filters: [], allowMultipleFiltersPerField: true, collation: { locale: 'en-US', sensitivity: 'variant', numeric: false, caseFirst: 'false' }, values: [{ valueId: `value:${'amount'}`, fieldId: 'amount', summarizeBy: 'sum' }], subtotalLocation: 'bottom', showRowGrandTotals: true, showColumnGrandTotals: true, reportLayout: 'compact' },
      refreshPolicy: { mode: 'on-change', preserveFormatting: true, refreshOnLoad: true },
    });
    sheet.drawings.push(
      { id: 'category-slicer', sheetId: sheet.id, kind: 'slicer', anchor: { kind: 'one-cell', row: 0, column: 4 }, transform: { x: 0, y: 0, width: 220, height: 180 }, zIndex: 1, payloadId: 'category-slicer' },
      { id: 'date-timeline', sheetId: sheet.id, kind: 'timeline', anchor: { kind: 'one-cell', row: 10, column: 0 }, transform: { x: 0, y: 0, width: 420, height: 120 }, zIndex: 1, payloadId: 'date-timeline' },
    );
    sheet.drawingPayloads.set('category-slicer', { kind: 'slicer', pivotId: 'control-pivot', fieldId: 'category', filter: { mode: 'all', memberKeys: [] }, style: { theme: 'light', fill: '#fff', border: '#ddd', textColor: '#111', accentColor: '#2563eb' }, settings: { showHeader: true, caption: 'Category', multiSelect: true, sort: 'ascending', showNoDataItems: true, noDataItemsLast: true, showNoDataStyle: true, columnCount: 1, itemHeight: 20 } });
    sheet.drawingPayloads.set('date-timeline', { kind: 'timeline', pivotId: 'control-pivot', fieldId: 'date', period: { start: '2024-01-01T00:00:00Z', end: '2024-12-31T00:00:00Z' }, level: 'years', selectionLevel: 'days', showHeader: false, showSelectionLabel: true, showTimeLevel: false, showHorizontalScrollbar: true, scrollPosition: '1895-01-01T00:00:00Z', bounds: { start: '1890-01-01T00:00:00Z', end: '2201-12-31T00:00:00Z' }, filterType: 'dateBetween', caption: 'Fiscal Window', styleName: 'TimelineStyleDark3', style: { theme: 'dark', fill: '#111827', border: '#374151', textColor: '#f9fafb', accentColor: '#60a5fa' } });
    const output = loadOpcPackageGraph(exportSnapshotToXlsxBuffer(workbook.snapshot()));
    assert.ok(output.files['xl/slicerCaches/slicerCache1.xml']);
    assert.ok(output.files['xl/slicers/slicer1.xml']);
    assert.ok(output.files['xl/timelineCaches/timelineCache1.xml']);
    assert.ok(output.files['xl/timelines/timeline1.xml']);
    const timelineXml = strFromU8(output.files['xl/timelines/timeline1.xml']!);
    assert.match(timelineXml, /caption="Fiscal Window"/);
    assert.match(timelineXml, /showHeader="0"/);
    assert.match(timelineXml, /showSelectionLabel="1"/);
    assert.match(timelineXml, /showTimeLevel="0"/);
    assert.match(timelineXml, /showHorizontalScrollbar="1"/);
    assert.match(timelineXml, /level="0"/);
    assert.match(timelineXml, /selectionLevel="3"/);
    assert.match(timelineXml, /scrollPosition="1895-01-01T00:00:00Z"/);
    assert.match(timelineXml, /style="TimelineStyleDark3"/);
    const timelineCacheXml = strFromU8(output.files['xl/timelineCaches/timelineCache1.xml']!);
    assert.match(timelineCacheXml, /selection startDate="2024-01-01T00:00:00Z" endDate="2024-12-31T00:00:00Z"/);
    assert.match(timelineCacheXml, /bounds startDate="1890-01-01T00:00:00Z" endDate="2201-12-31T00:00:00Z"/);
    const contentTypes = strFromU8(output.files['[Content_Types].xml']!);
    assert.match(contentTypes, /application\/vnd\.openxmlformats-officedocument\.drawing\+xml/);
    assert.match(contentTypes, /application\/vnd\.ms-excel\.slicerCache"/);
    assert.doesNotMatch(contentTypes, /application\/vnd\.ms-excel\.slicerCache\+xml/);
    assert.match(strFromU8(output.files['xl/workbook.xml']!), /slicerCaches/);
    assert.match(strFromU8(output.files['xl/workbook.xml']!), /timelineCacheRefs/);
    assert.match(strFromU8(output.files['xl/_rels/workbook.xml.rels']!), /relationships\/timelineCache/);
    assert.match(strFromU8(output.files['xl/worksheets/sheet1.xml']!), /slicerList/);
    assert.match(strFromU8(output.files['xl/worksheets/sheet1.xml']!), /timelineRefs/);
    const imported = parseLoadedXlsx(output).snapshot;
    assert.equal(Object.values(imported.sheets[0]?.drawingPayloads ?? {}).filter((payload) => payload.kind === 'slicer').length, 1);
    assert.equal(Object.values(imported.sheets[0]?.drawingPayloads ?? {}).filter((payload) => payload.kind === 'timeline').length, 1);
    const importedTimeline = Object.values(imported.sheets[0]?.drawingPayloads ?? {}).find((payload) => payload.kind === 'timeline');
    assert.equal(importedTimeline?.kind, 'timeline');
    if (!importedTimeline || importedTimeline.kind !== 'timeline') throw new Error('Timeline payload is missing');
    assert.equal(importedTimeline.level, 'years');
    assert.equal(importedTimeline.selectionLevel, 'days');
    assert.equal(importedTimeline.showHeader, false);
    assert.equal(importedTimeline.showSelectionLabel, true);
    assert.equal(importedTimeline.showTimeLevel, false);
    assert.equal(importedTimeline.showHorizontalScrollbar, true);
    assert.equal(importedTimeline.scrollPosition, '1895-01-01T00:00:00Z');
    assert.deepEqual(importedTimeline.bounds, { start: '1890-01-01T00:00:00Z', end: '2201-12-31T00:00:00Z' });
    assert.equal(importedTimeline.caption, 'Fiscal Window');
    assert.equal(importedTimeline.styleName, 'TimelineStyleDark3');
    const preservedFiles: Record<string, Uint8Array> = Object.fromEntries(Object.entries(output.files).map(([name, bytes]) => [name, bytes.slice()]));
    preservedFiles['xl/timelines/timeline1.xml'] = strToU8(strFromU8(preservedFiles['xl/timelines/timeline1.xml']!).replace('<timeline ', '<timeline futureTimelineAttr="keep-me" ').replace('/></timelines>', '><extLst><ext uri="future"><futureTimelineNode/></ext></extLst></timeline></timelines>'));
    const preservedInput = loadOpcPackageGraph(zipXlsxPartsBuffer(preservedFiles));
    const preservedImported = parseLoadedXlsx(preservedInput).snapshot;
    const preservedExport = loadOpcPackageGraph(exportSnapshotToXlsxBuffer(preservedImported, preservedInput.packageGraph));
    assert.match(strFromU8(preservedExport.files['xl/timelines/timeline1.xml']!), /futureTimelineAttr="keep-me"/);
    assert.match(strFromU8(preservedExport.files['xl/timelines/timeline1.xml']!), /futureTimelineNode/);
    const importedSlicerAnchor = imported.sheets[0]?.drawings.find((drawing) => drawing.kind === 'slicer')?.anchor;
    const importedTimelineAnchor = imported.sheets[0]?.drawings.find((drawing) => drawing.kind === 'timeline')?.anchor;
    assert.equal(importedSlicerAnchor?.kind, 'one-cell');
    assert.equal(importedSlicerAnchor?.kind === 'one-cell' ? importedSlicerAnchor.column : undefined, 4);
    assert.equal(importedTimelineAnchor?.kind === 'one-cell' ? importedTimelineAnchor.row : undefined, 10);

    const controlExport = await exportXlsx({ snapshot: workbook.snapshot(), fileName: 'native-controls.xlsx', options: { compatibilityTarget: 'B' } });
    assert.equal(controlExport.report.issues.some((issue) => issue.feature === 'images' && issue.status === 'unsupported'), false);

    const withoutControls = structuredClone(imported);
    const controlSheet = withoutControls.sheets[0]!;
    const controlDrawingIds = new Set(controlSheet.drawings.filter((drawing) => drawing.kind === 'slicer' || drawing.kind === 'timeline').map((drawing) => drawing.id));
    controlSheet.drawings = controlSheet.drawings.filter((drawing) => !controlDrawingIds.has(drawing.id));
    for (const id of controlDrawingIds) delete controlSheet.drawingPayloads[id];
    const removedControls = loadOpcPackageGraph(exportSnapshotToXlsxBuffer(withoutControls, output.packageGraph));
    const removedDrawing = removedControls.files['xl/drawings/drawing1.xml'];
    assert.ok(!removedDrawing || !/category[_-]slicer|date[_-]timeline/.test(strFromU8(removedDrawing)));
  });

  it('round-trips report filters and manual row/column item visibility through native Pivot XML', async () => {
    const workbook = new WorkbookModel('wb-pivot-visibility', 'Pivot Visibility');
    const sheet = workbook.getSheet(workbook.primarySheetId);
    [
      ['Category', 'Region', 'Country', 'Amount'],
      ['A', 'East', 'US', 10],
      ['A', 'West', 'US', 20],
      ['B', 'East', 'CA', 30],
      ['B', 'West', 'CA', 40],
    ].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
    const member = (value: string) => createPivotMemberKey(value);
    sheet.pivots.push({
      schema: 'PivotDefinition', id: 'pivot-visibility',
      source: { kind: 'worksheet-range', range: { sheetId: sheet.id, startRow: 0, endRow: 4, startColumn: 0, endColumn: 3 } },
      target: { sheetId: sheet.id, anchor: { row: 7, column: 0 } },
      fieldCatalog: { schema: 'PivotFieldCatalog', fields: [
        { fieldId: 'category', name: 'Category', dataType: 'text', ordinal: 0 },
        { fieldId: 'region', name: 'Region', dataType: 'text', ordinal: 1 },
        { fieldId: 'country', name: 'Country', dataType: 'text', ordinal: 2 },
        { fieldId: 'amount', name: 'Amount', dataType: 'number', ordinal: 3 },
      ] },
      layout: {
        rows: [{ fieldId: 'category' }], columns: [{ fieldId: 'region' }],
        filters: [
          { kind: 'manual', family: 'manual', scope: 'report', fieldId: 'amount', mode: 'include', memberKeys: [createPivotMemberKey(10)] },
          { kind: 'condition', family: 'label', scope: 'report', fieldId: 'country', operator: 'equals', value: 'US' },
          { kind: 'manual', family: 'manual', scope: 'field', fieldId: 'category', mode: 'exclude', memberKeys: [member('B')] },
          { kind: 'manual', family: 'manual', scope: 'field', fieldId: 'region', mode: 'exclude', memberKeys: [member('West')] },
        ],
        allowMultipleFiltersPerField: true,
        collation: { locale: 'en-US', sensitivity: 'variant', numeric: false, caseFirst: 'false' },
        values: [{ valueId: `value:${'amount'}`, fieldId: 'amount', summarizeBy: 'sum' }], subtotalLocation: 'bottom', showRowGrandTotals: true, showColumnGrandTotals: true, reportLayout: 'compact',
      },
      refreshPolicy: { mode: 'manual', preserveFormatting: true, refreshOnLoad: false },
    });
    const output = loadOpcPackageGraph(exportSnapshotToXlsxBuffer(workbook.snapshot()));
    const nativeXml = strFromU8(output.files['xl/pivotTables/pivotTable1.xml']!);
    assert.match(nativeXml, /<pageFields count="2"><pageField fld="3"\/><pageField fld="2"\/><\/pageFields>/);
    assert.match(nativeXml, /<pivotField axis="axisRow"[^>]*><items count="1"><item x="1" h="1"\/><\/items>/);
    assert.match(nativeXml, /<pivotField axis="axisCol"[^>]*><items count="1"><item x="1" h="1"\/><\/items>/);
    assert.match(nativeXml, /<pivotField axis="axisPage"[^>]*><items count="3"><item x="1" h="1"\/><item x="2" h="1"\/><item x="3" h="1"\/><\/items>/);
    assert.match(nativeXml, /<pivotFilters count="1"><filter fld="2" type="captionEqual" stringValue1="US"/);

    const imported = await importXlsx({ fileName: 'pivot-visibility.xlsx', buffer: zipXlsxPartsBuffer(output.files), options: { compatibilityTarget: 'B' } });
    const importedPivot = imported.snapshot.sheets[0]?.pivots[0];
    assert.ok(importedPivot);
    assert.deepEqual(importedPivot.layout.filters, [
      { kind: 'manual', family: 'manual', fieldId: 'native:cache:1:field:0', scope: 'field', mode: 'exclude', memberKeys: [member('B')] },
      { kind: 'manual', family: 'manual', fieldId: 'native:cache:1:field:1', scope: 'field', mode: 'exclude', memberKeys: [member('West')] },
      { kind: 'manual', family: 'manual', fieldId: 'native:cache:1:field:3', scope: 'report', mode: 'exclude', memberKeys: [createPivotMemberKey(20), createPivotMemberKey(30), createPivotMemberKey(40)] },
      { kind: 'condition', family: 'label', fieldId: 'native:cache:1:field:2', scope: 'report', operator: 'equals', value: 'US' },
    ]);
    const roundTrip = await exportXlsx({ snapshot: imported.snapshot, nativePackage: imported.nativePackage, fileName: 'pivot-visibility-roundtrip.xlsx', options: { compatibilityTarget: 'B' } });
    const roundTripXml = strFromU8(loadOpcPackageGraph(roundTrip.buffer).files['xl/pivotTables/pivotTable1.xml']!);
    assert.match(roundTripXml, /<pageFields count="2"><pageField fld="3"\/><pageField fld="2"\/><\/pageFields>/);
    assert.equal((roundTripXml.match(/ h="1"/g) ?? []).length, 5);
  });

  it('rejects native manual visibility indexes outside the cache field range', () => {
    const workbook = new WorkbookModel('wb-pivot-invalid-visibility', 'Invalid Pivot Visibility');
    const sheet = workbook.getSheet(workbook.primarySheetId);
    const cache = {
      cacheId: 1, part: 'xl/pivotCache/pivotCacheDefinition1.xml',
      source: { kind: 'worksheet-range' as const, sheetName: sheet.name, sheetPart: 'xl/worksheets/sheet1.xml', ref: 'A1:A2' },
      fields: [{ index: 0, name: 'Category', dataType: 'string' as const, sharedItems: ['A'] }],
    };
    const table = {
      name: 'PivotTable1', part: 'xl/pivotTables/pivotTable1.xml', sheetPart: 'xl/worksheets/sheet1.xml', relationshipId: 'rIdPivot', cacheId: 1,
      locationRef: 'C1:C3', fields: [{ index: 0, axis: 'row' as const, hiddenItemIndexes: [1] }], rowFields: [0], columnFields: [], pageFields: [], dataFields: [],
    };
    assert.throws(() => mapNativePivotDefinition(table, cache, workbook.snapshot(), { [sheet.id]: 'xl/worksheets/sheet1.xml' }), /hidden item 1 is outside sharedItems bounds/);

    const invalidExport = new WorkbookModel('wb-pivot-invalid-filter', 'Invalid Pivot Filter');
    const invalidSheet = invalidExport.getSheet(invalidExport.primarySheetId);
    [['Category'], ['A']].forEach((row, rowIndex) => row.forEach((value, columnIndex) => invalidSheet.cells.set(rowIndex, columnIndex, { value })));
    invalidSheet.pivots.push({
      schema: 'PivotDefinition', id: 'invalid-filter', source: { kind: 'worksheet-range', range: { sheetId: invalidSheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 } }, target: { sheetId: invalidSheet.id, anchor: { row: 4, column: 0 } },
      fieldCatalog: { schema: 'PivotFieldCatalog', fields: [{ fieldId: 'category', name: 'Category', dataType: 'text', ordinal: 0 }] },
      layout: { rows: [], columns: [], filters: [{ kind: 'manual', family: 'manual', scope: 'report', fieldId: 'missing', mode: 'all', memberKeys: [] }], allowMultipleFiltersPerField: true, collation: { locale: 'en-US', sensitivity: 'variant', numeric: false, caseFirst: 'false' }, values: [], subtotalLocation: 'bottom', showRowGrandTotals: true, showColumnGrandTotals: true, reportLayout: 'compact' },
      refreshPolicy: { mode: 'manual', preserveFormatting: true, refreshOnLoad: false },
    });
    assert.throws(() => exportSnapshotToXlsxBuffer(invalidExport.snapshot()), /Pivot filter references missing field missing/);
  });

  it('round-trips the canonical compact, outline, and tabular report layouts', async () => {
    for (const reportLayout of ['compact', 'outline', 'tabular'] as const) {
      const workbook = new WorkbookModel(`wb-report-layout-${reportLayout}`, `Report Layout ${reportLayout}`);
      const sheet = workbook.getSheet(workbook.primarySheetId);
      [['Region', 'Category', 'Amount'], ['East', 'Widget', 10], ['West', 'Gadget', 20]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
      sheet.pivots.push({
        schema: 'PivotDefinition', id: `pivot-${reportLayout}`,
        source: { kind: 'worksheet-range', range: { sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 2 } },
        target: { sheetId: sheet.id, anchor: { row: 5, column: 0 } },
        fieldCatalog: { schema: 'PivotFieldCatalog', fields: [
          { fieldId: 'region', name: 'Region', dataType: 'text', ordinal: 0 },
          { fieldId: 'category', name: 'Category', dataType: 'text', ordinal: 1 },
          { fieldId: 'amount', name: 'Amount', dataType: 'number', ordinal: 2 },
        ] },
        layout: { rows: [{ fieldId: 'region' }, { fieldId: 'category' }], columns: [], filters: [], allowMultipleFiltersPerField: true, collation: { locale: 'en-US', sensitivity: 'variant', numeric: false, caseFirst: 'false' }, values: [{ valueId: `value:${'amount'}`, fieldId: 'amount', summarizeBy: 'sum' }], subtotalLocation: 'bottom', showRowGrandTotals: true, showColumnGrandTotals: true, reportLayout },
        refreshPolicy: { mode: 'manual', preserveFormatting: true, refreshOnLoad: false },
      });
      const output = loadOpcPackageGraph(exportSnapshotToXlsxBuffer(workbook.snapshot()));
      const nativeXml = strFromU8(output.files['xl/pivotTables/pivotTable1.xml']!);
      assert.match(nativeXml, new RegExp(`compactData="${reportLayout === 'compact' ? '1' : '0'}"`));
      assert.match(nativeXml, new RegExp(`repeatAllLabels="${reportLayout === 'tabular' ? '1' : '0'}"`));
      assert.match(nativeXml, new RegExp(`axisRow[^>]*compact="${reportLayout === 'compact' ? '1' : '0'}"[^>]*outline="${reportLayout === 'outline' ? '1' : '0'}"`));
      const imported = await importXlsx({ fileName: `${reportLayout}.xlsx`, buffer: zipXlsxPartsBuffer(output.files), options: { compatibilityTarget: 'B' } });
      assert.equal(imported.snapshot.sheets[0]?.pivots[0]?.layout.reportLayout, reportLayout);
    }
  });

  it('rejects conflicting native report-layout flags instead of guessing', () => {
    const workbook = new WorkbookModel('wb-invalid-report-layout', 'Invalid Report Layout');
    const sheet = workbook.getSheet(workbook.primarySheetId);
    const cache = {
      cacheId: 1, part: 'xl/pivotCache/pivotCacheDefinition1.xml',
      source: { kind: 'worksheet-range' as const, sheetName: sheet.name, sheetPart: 'xl/worksheets/sheet1.xml', ref: 'A1:A2' },
      fields: [{ index: 0, name: 'Region', dataType: 'string' as const, sharedItems: ['East'] }],
    };
    const table = {
      name: 'PivotTable1', part: 'xl/pivotTables/pivotTable1.xml', sheetPart: 'xl/worksheets/sheet1.xml', relationshipId: 'rIdPivot', cacheId: 1,
      locationRef: 'C1:C3', fields: [{ index: 0, axis: 'row' as const, compact: true, outline: false }], rowFields: [0], columnFields: [], pageFields: [], dataFields: [], compactData: true, repeatLabels: true,
    };
    assert.throws(() => mapNativePivotDefinition(table, cache, workbook.snapshot(), { [sheet.id]: 'xl/worksheets/sheet1.xml' }), /conflicting compactData and repeatAllLabels/);
  });

  it('fails closed for invalid native Timeline levels and bounds', () => {
    const workbook = new WorkbookModel('wb-invalid-timeline-state', 'Invalid Timeline State');
    const sheet = workbook.getSheet(workbook.primarySheetId);
    sheet.cells.set(0, 0, { value: 'Date' });
    sheet.cells.set(1, 0, { value: '2024-01-01T00:00:00Z' });
    sheet.pivots.push({
      schema: 'PivotDefinition', id: 'invalid-timeline-pivot', source: { kind: 'worksheet-range', range: { sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 } }, target: { sheetId: sheet.id, anchor: { row: 4, column: 0 } },
      fieldCatalog: { schema: 'PivotFieldCatalog', fields: [{ fieldId: 'date', name: 'Date', dataType: 'date', ordinal: 0, values: ['2024-01-01T00:00:00Z'] }] },
      layout: { rows: [{ fieldId: 'date' }], columns: [], filters: [], allowMultipleFiltersPerField: true, collation: { locale: 'en-US', sensitivity: 'variant', numeric: false, caseFirst: 'false' }, values: [], subtotalLocation: 'bottom', showRowGrandTotals: true, showColumnGrandTotals: true, reportLayout: 'compact' },
      refreshPolicy: { mode: 'on-change', preserveFormatting: true, refreshOnLoad: true },
    });
    sheet.drawings.push({ id: 'invalid-timeline', sheetId: sheet.id, kind: 'timeline', anchor: { kind: 'one-cell', row: 8, column: 0 }, transform: { x: 0, y: 0, width: 420, height: 120 }, zIndex: 1, payloadId: 'invalid-timeline' });
    sheet.drawingPayloads.set('invalid-timeline', { kind: 'timeline', pivotId: 'invalid-timeline-pivot', fieldId: 'date', period: {}, level: 'months', selectionLevel: 'months', showHeader: true, showSelectionLabel: true, showTimeLevel: true, showHorizontalScrollbar: true, bounds: { start: '2024-01-01T00:00:00Z', end: '2024-01-01T00:00:00Z' }, filterType: 'unknown', styleName: 'TimelineStyleLight2', style: { theme: 'light', fill: '#fff', border: '#ddd', textColor: '#111', accentColor: '#2563eb' } });
    const output = loadOpcPackageGraph(exportSnapshotToXlsxBuffer(workbook.snapshot()));
    const timelinePart = 'xl/timelines/timeline1.xml';
    const cachePart = 'xl/timelineCaches/timelineCache1.xml';
    const invalidLevel = loadOpcPackageGraph(exportSnapshotToXlsxBuffer(workbook.snapshot()));
    invalidLevel.files[timelinePart] = strToU8(strFromU8(invalidLevel.files[timelinePart]!).replace('level="2"', 'level="9"'));
    assert.throws(() => parseLoadedXlsx(loadOpcPackageGraph(zipXlsxPartsBuffer(invalidLevel.files))), /level must be one of/);
    const invalidBounds = loadOpcPackageGraph(exportSnapshotToXlsxBuffer(workbook.snapshot()));
    invalidBounds.files[cachePart] = strToU8(strFromU8(invalidBounds.files[cachePart]!).replace(/ endDate="[^"]+"/, ''));
    assert.throws(() => parseLoadedXlsx(loadOpcPackageGraph(zipXlsxPartsBuffer(invalidBounds.files))), /bounds requires startDate and endDate/);
    assert.ok(output.files[timelinePart]);
  });

  it('keeps native Slicer selections typed when member labels collide', () => {
    const workbook = new WorkbookModel('wb-typed-slicer', 'Typed Slicer');
    const sheet = workbook.getSheet(workbook.primarySheetId);
    [["Member", "Amount"], [1, 10], ["1", 20], [true, 30], ["true", 40], [null, 50]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
    sheet.pivots.push({
      schema: 'PivotDefinition', id: 'typed-slicer-pivot', source: { kind: 'worksheet-range', range: { sheetId: sheet.id, startRow: 0, endRow: 5, startColumn: 0, endColumn: 1 } }, target: { sheetId: sheet.id, anchor: { row: 8, column: 0 } },
      fieldCatalog: { schema: 'PivotFieldCatalog', fields: [{ fieldId: 'member', name: 'Member', dataType: 'mixed', ordinal: 0, values: [1, '1', true, 'true', null] }, { fieldId: 'amount', name: 'Amount', dataType: 'number', ordinal: 1 }] },
      layout: { rows: [{ fieldId: 'member' }], columns: [], filters: [], allowMultipleFiltersPerField: true, collation: { locale: 'en-US', sensitivity: 'variant', numeric: false, caseFirst: 'false' }, values: [{ valueId: `value:${'amount'}`, fieldId: 'amount', summarizeBy: 'sum' }], subtotalLocation: 'bottom', showRowGrandTotals: true, showColumnGrandTotals: true, reportLayout: 'compact' },
      refreshPolicy: { mode: 'on-change', preserveFormatting: true, refreshOnLoad: true },
    });
    sheet.drawings.push({ id: 'typed-slicer', sheetId: sheet.id, kind: 'slicer', anchor: { kind: 'one-cell', row: 0, column: 4 }, transform: { x: 0, y: 0, width: 220, height: 180 }, zIndex: 1, payloadId: 'typed-slicer-payload' });
    sheet.drawingPayloads.set('typed-slicer-payload', { kind: 'slicer', pivotId: 'typed-slicer-pivot', fieldId: 'member', filter: { mode: 'include', memberKeys: [createPivotMemberKey('1')] }, style: { theme: 'light', fill: '#fff', border: '#ddd', textColor: '#111', accentColor: '#2563eb' }, settings: { showHeader: true, caption: 'Member', multiSelect: true, sort: 'ascending', showNoDataItems: true, noDataItemsLast: true, showNoDataStyle: true, columnCount: 1, itemHeight: 20 } });
    const output = loadOpcPackageGraph(exportSnapshotToXlsxBuffer(workbook.snapshot()));
    const cacheXml = strFromU8(output.files['xl/slicerCaches/slicerCache1.xml']!);
    assert.match(cacheXml, /<i x="0"\/>/);
    assert.match(cacheXml, /<i x="1" s="1"\/>/);
    const imported = parseLoadedXlsx(output).snapshot;
    const payload = Object.values(imported.sheets[0]?.drawingPayloads ?? {}).find((entry) => entry.kind === 'slicer');
    assert.equal(payload?.kind, 'slicer');
    if (!payload || payload.kind !== 'slicer') throw new Error('Typed Slicer payload is missing');
    assert.deepEqual(payload.filter.memberKeys, [createPivotMemberKey('1')]);
  });

  it('writes canonical PivotCharts as native chart parts linked to the PivotTable', async () => {
    const workbook = new WorkbookModel('wb-native-pivot-chart', 'Native PivotChart');
    const sheet = workbook.getSheet(workbook.primarySheetId);
    sheet.cells.set(0, 0, { value: 'Category' });
    sheet.cells.set(0, 1, { value: 'Amount' });
    sheet.cells.set(1, 0, { value: 'A' });
    sheet.cells.set(1, 1, { value: 10 });
    sheet.cells.set(2, 0, { value: 'B' });
    sheet.cells.set(2, 1, { value: 20 });
    sheet.pivots.push({
      schema: 'PivotDefinition', id: 'pivot-chart-1',
      source: { kind: 'worksheet-range', range: { sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 } },
      target: { sheetId: sheet.id, anchor: { row: 5, column: 0 } },
      fieldCatalog: { schema: 'PivotFieldCatalog', fields: [
        { fieldId: 'category', name: 'Category', dataType: 'text', ordinal: 0 },
        { fieldId: 'amount', name: 'Amount', dataType: 'number', ordinal: 1 },
      ] },
      layout: { rows: [{ fieldId: 'category' }], columns: [], filters: [], allowMultipleFiltersPerField: true, collation: { locale: 'en-US', sensitivity: 'variant', numeric: false, caseFirst: 'false' }, values: [{ valueId: `value:${'amount'}`, fieldId: 'amount', summarizeBy: 'sum' }], subtotalLocation: 'bottom', showRowGrandTotals: true, showColumnGrandTotals: true, reportLayout: 'compact' },
      refreshPolicy: { mode: 'on-change', preserveFormatting: true, refreshOnLoad: true },
    });
    sheet.drawings.push({ id: 'pivot-chart-drawing', sheetId: sheet.id, kind: 'chart', anchor: { kind: 'one-cell', row: 0, column: 4 }, transform: { x: 0, y: 0, width: 320, height: 220 }, zIndex: 1, payloadId: 'pivot-chart-payload' });
    sheet.drawingPayloads.set('pivot-chart-payload', { kind: 'chart', chartId: 'pivot-chart-1', chartType: 'column', subtype: 'clustered', source: { kind: 'pivot', pivotId: 'pivot-chart-1' }, elements: { hiddenData: 'show', title: 'Amounts', legend: { visible: true, position: 'bottom' } } });

    const output = loadOpcPackageGraph(exportSnapshotToXlsxBuffer(workbook.snapshot()));
    const chartParts = Object.keys(output.files).filter((name) => name.startsWith('xl/charts/react-chart-'));
    assert.equal(chartParts.length, 1);
    const chartXml = strFromU8(output.files[chartParts[0]!]!);
    assert.match(chartXml, /<c:pivotSource>/);
    assert.match(chartXml, /<c:barChart>/);
    assert.match(chartXml, /PivotTable|pivot_chart_1/);
    assert.match(strFromU8(output.files['xl/drawings/_rels/drawing1.xml.rels']!), /relationships\/chart/);
    assert.match(strFromU8(output.files['xl/worksheets/sheet1.xml']!), /<drawing r:id=/);
    assert.match(strFromU8(output.files['[Content_Types].xml']!), /drawingml\.chart\+xml/);
    assert.ok(scanSnapshotFeatures(workbook.snapshot()).includes('pivot-chart'));
    const report = await exportXlsx({ snapshot: workbook.snapshot(), fileName: 'pivot-chart.xlsx', options: { compatibilityTarget: 'B' } });
    assert.equal(report.report.issues.some((issue) => issue.feature === 'pivot-chart' && issue.status === 'editable'), true);
    assert.equal(report.report.issues.some((issue) => issue.feature === 'charts' && issue.status === 'unsupported'), false);

    for (const [type, xmlName] of [['bar', 'barDir val="bar"'], ['line', '<c:lineChart>'], ['area', '<c:areaChart>'] ] as const) {
      const typed = structuredClone(workbook.snapshot());
      const payload = typed.sheets[0]!.drawingPayloads['pivot-chart-payload'];
      if (!payload || payload.kind !== 'chart') throw new Error('PivotChart test fixture is missing its chart payload');
      payload.chartType = type;
      payload.subtype = defaultChartSubtype(type);
      const typedOutput = loadOpcPackageGraph(exportSnapshotToXlsxBuffer(typed));
      const typedChartPart = Object.keys(typedOutput.files).find((name) => name.startsWith('xl/charts/react-chart-'));
      assert.ok(typedChartPart);
      assert.match(strFromU8(typedOutput.files[typedChartPart!]!), new RegExp(xmlName.replace(/[<>]/g, '')));
    }

    const withoutChart = structuredClone(workbook.snapshot());
    withoutChart.sheets[0]!.drawings = [];
    withoutChart.sheets[0]!.drawingPayloads = {};
    const removed = loadOpcPackageGraph(exportSnapshotToXlsxBuffer(withoutChart, output.packageGraph));
    assert.equal(Object.keys(removed.files).some((name) => name.startsWith('xl/charts/react-chart-')), false);
    assert.doesNotMatch(strFromU8(removed.files['xl/worksheets/sheet1.xml']!), /<drawing r:id=/);
  });

  it('rejects unsupported native PivotChart types instead of emitting a fake chart', () => {
    const workbook = new WorkbookModel('wb-native-pivot-chart-invalid', 'Invalid PivotChart');
    const sheet = workbook.getSheet(workbook.primarySheetId);
    sheet.pivots.push({
      schema: 'PivotDefinition', id: 'pivot-invalid-chart',
      source: { kind: 'worksheet-range', range: { sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 } },
      target: { sheetId: sheet.id, anchor: { row: 3, column: 0 } },
      fieldCatalog: { schema: 'PivotFieldCatalog', fields: [{ fieldId: 'value', name: 'Value', dataType: 'number', ordinal: 0 }] },
      layout: { rows: [], columns: [], filters: [], allowMultipleFiltersPerField: true, collation: { locale: 'en-US', sensitivity: 'variant', numeric: false, caseFirst: 'false' }, values: [{ valueId: `value:${'value'}`, fieldId: 'value', summarizeBy: 'sum' }], subtotalLocation: 'bottom', showRowGrandTotals: true, showColumnGrandTotals: true, reportLayout: 'compact' },
      refreshPolicy: { mode: 'manual', preserveFormatting: true, refreshOnLoad: false },
    });
    sheet.drawings.push({ id: 'invalid-pivot-chart', sheetId: sheet.id, kind: 'chart', anchor: { kind: 'one-cell', row: 0, column: 4 }, transform: { x: 0, y: 0, width: 200, height: 120 }, zIndex: 1, payloadId: 'invalid-pivot-chart' });
    sheet.drawingPayloads.set('invalid-pivot-chart', { kind: 'chart', chartId: 'invalid-pivot-chart', chartType: 'pie', subtype: 'pie', source: { kind: 'pivot', pivotId: 'pivot-invalid-chart' }, elements: { hiddenData: 'show' } });
    assert.throws(() => exportSnapshotToXlsxBuffer(workbook.snapshot()), /unsupported native chart type/);
  });

  it('writes ordinary XY/Bubble charts through the same native chart chain and preserves unknown chart-space nodes', () => {
    const workbook = new WorkbookModel('wb-native-xy-chart', 'Native XY Chart');
    const sheet = workbook.getSheet(workbook.primarySheetId);
    [['Label', 'X', 'Y', 'Size'], ['A', 1, 10, 5], ['B', 2, 20, 8], ['C', 4, 40, 12]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
    sheet.drawings.push({ id: 'xy-chart-drawing', sheetId: sheet.id, kind: 'chart', anchor: { kind: 'one-cell', row: 5, column: 4 }, transform: { x: 0, y: 0, width: 360, height: 240 }, zIndex: 1, payloadId: 'xy-chart-payload' });
    sheet.drawingPayloads.set('xy-chart-payload', {
      kind: 'chart', chartId: 'xy-chart-payload', chartType: 'bubble', subtype: 'bubble', source: { kind: 'worksheet-ranges', ranges: [{ sheetId: sheet.id, startRow: 0, endRow: 3, startColumn: 0, endColumn: 3 }] },
      series: [{ id: 'bubble-series', name: 'Y', range: { sheetId: sheet.id, startRow: 1, endRow: 3, startColumn: 2, endColumn: 2 }, xRange: { sheetId: sheet.id, startRow: 1, endRow: 3, startColumn: 1, endColumn: 1 }, yRange: { sheetId: sheet.id, startRow: 1, endRow: 3, startColumn: 2, endColumn: 2 }, sizeRange: { sheetId: sheet.id, startRow: 1, endRow: 3, startColumn: 3, endColumn: 3 }, chartType: 'bubble' }],
      categoryRange: { sheetId: sheet.id, startRow: 1, endRow: 3, startColumn: 0, endColumn: 0 }, elements: { hiddenData: 'show', title: 'Bubble' },
    });
    const output = loadOpcPackageGraph(exportSnapshotToXlsxBuffer(workbook.snapshot()));
    const chartPart = Object.keys(output.files).find((name) => name.startsWith('xl/charts/react-chart-'));
    assert.ok(chartPart);
    const chartXml = strFromU8(output.files[chartPart!]!);
    assert.match(chartXml, /<c:bubbleChart>/);
    assert.match(chartXml, /<c:xVal>.*\$B\$2:\$B\$4/);
    assert.match(chartXml, /<c:yVal>.*\$C\$2:\$C\$4/);
    assert.match(chartXml, /<c:bubbleSize>.*\$D\$2:\$D\$4/);

    const preserved = loadOpcPackageGraph(exportSnapshotToXlsxBuffer(workbook.snapshot()));
    const preservedChartPart = Object.keys(preserved.files).find((name) => name.startsWith('xl/charts/react-chart-'));
    assert.ok(preservedChartPart);
    preserved.packageGraph.parts[preservedChartPart!] = strToU8(strFromU8(preserved.packageGraph.parts[preservedChartPart!]!).replace('</c:chartSpace>', '<c:unknownChartNode val="keep"/><c:extLst><c:ext uri="{test}"><c:unknown val="keep"/></c:ext></c:extLst></c:chartSpace>'));
    const rewritten = loadOpcPackageGraph(exportSnapshotToXlsxBuffer(workbook.snapshot(), preserved.packageGraph));
    const rewrittenXml = strFromU8(rewritten.files[preservedChartPart!]!);
    assert.match(rewrittenXml, /unknownChartNode/);
    assert.match(rewrittenXml, /uri="\{test\}"/);
  });

  it('serializes and imports native Sparkline design and group semantics', async () => {
    const workbook = new WorkbookModel('wb-native-sparkline', 'Native Sparkline');
    const sheet = workbook.getSheet(workbook.primarySheetId);
    [[2, null, -4, 8], [1, 3, 5, 7]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex + 1, columnIndex + 1, { value })));
    sheet.sparklines.push({ id: 'spark-native', sheetId: sheet.id, anchor: { row: 1, column: 5 }, sourceRange: { sheetId: sheet.id, startRow: 1, endRow: 1, startColumn: 1, endColumn: 4 }, type: 'line', color: '#2563eb', negativeColor: '#ef4444', showAxis: true, showMarkers: true, lineWeight: 2, dateAxis: true, rightToLeft: true, hiddenCells: 'hide', emptyCells: 'connect', verticalAxis: { mode: 'custom', minimum: -10, maximum: 10 }, axisColor: '#475569', firstColor: '#f59e0b', lastColor: '#8b5cf6', highColor: '#16a34a', lowColor: '#dc2626', markerColor: '#0ea5e9', groupId: 'spark-group' });
    sheet.sparklineGroups.push({ id: 'spark-group', sheetId: sheet.id, type: 'line', sparklineIds: ['spark-native'], showAxis: true, showMarkers: true, lineWeight: 2, dateAxis: true, rightToLeft: true, hiddenCells: 'hide', emptyCells: 'connect', verticalAxis: { mode: 'custom', minimum: -10, maximum: 10 }, axisColor: '#475569', firstColor: '#f59e0b', lastColor: '#8b5cf6', highColor: '#16a34a', lowColor: '#dc2626', negativeColor: '#ef4444', markerColor: '#0ea5e9' });
    const output = loadOpcPackageGraph(exportSnapshotToXlsxBuffer(workbook.snapshot()));
    const sheetXml = strFromU8(output.files['xl/worksheets/sheet1.xml']!);
    assert.match(sheetXml, /sparklineGroup[^>]+lineWeight="2"/);
    assert.match(sheetXml, /displayEmptyCellsAs="span"/);
    assert.match(sheetXml, /manualMin="-10"/);
    const imported = await importXlsx({ fileName: 'native-sparkline.xlsx', buffer: zipXlsxPartsBuffer(output.files), options: { compatibilityTarget: 'B' } });
    const importedSparkline = imported.snapshot.sheets[0]?.sparklines.find((entry) => entry.id === 'spark-native');
    assert.equal(importedSparkline?.emptyCells, 'connect');
    assert.equal(importedSparkline?.rightToLeft, true);
    assert.equal(imported.snapshot.sheets[0]?.sparklineGroups?.[0]?.verticalAxis?.maximum, 10);
  });

  it('reads and rewrites native Pivot cache/table relationship graphs', async () => {
    const workbook = new WorkbookModel('wb-native-pivot', 'Native Pivot');
    const generated = loadOpcPackageGraph(exportSnapshotToXlsxBuffer(workbook.snapshot()));
    const parts = generated.packageGraph.parts;
    parts['xl/worksheets/sheet1.xml'] = strToU8(strFromU8(parts['xl/worksheets/sheet1.xml']!).replace(
      '</sheetData>',
      '<row r="1"><c r="A1" t="inlineStr"><is><t>Category</t></is></c><c r="B1" t="inlineStr"><is><t>Amount</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>A</t></is></c><c r="B2"><v>10</v></c></row></sheetData>',
    ));
    parts['xl/workbook.xml'] = strToU8(strFromU8(parts['xl/workbook.xml']!).replace('</workbook>', '<pivotCaches count="1"><pivotCache cacheId="1" r:id="rIdPivotCache"/></pivotCaches></workbook>'));
    parts['xl/_rels/workbook.xml.rels'] = strToU8(strFromU8(parts['xl/_rels/workbook.xml.rels']!).replace('</Relationships>', '<Relationship Id="rIdPivotCache" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/></Relationships>'));
    parts['xl/pivotCache/pivotCacheDefinition1.xml'] = strToU8('<?xml version="1.0"?><pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><cacheSource type="worksheet"><worksheetSource ref="A1:B2" sheet="Sheet1"/></cacheSource><cacheFields count="2"><cacheField name="Category"><sharedItems containsString="1"><s v="A"/></sharedItems></cacheField><cacheField name="Amount"><sharedItems containsNumber="1"><n v="10"/></sharedItems></cacheField></cacheFields></pivotCacheDefinition>');
    parts['xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels'] = strToU8('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdRecords" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords1.xml"/></Relationships>');
    parts['xl/pivotCache/pivotCacheRecords1.xml'] = strToU8('<?xml version="1.0"?><pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1"><r><s v="0"/><n v="10"/></r></pivotCacheRecords>');
    parts['xl/worksheets/sheet1.xml'] = strToU8(strFromU8(parts['xl/worksheets/sheet1.xml']!).replace('</worksheet>', '<pivotTableParts count="1"><pivotTablePart r:id="rIdPivotTable"/></pivotTableParts></worksheet>'));
    parts['xl/worksheets/_rels/sheet1.xml.rels'] = strToU8(strFromU8(parts['xl/worksheets/_rels/sheet1.xml.rels']!).replace('</Relationships>', '<Relationship Id="rIdPivotTable" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/></Relationships>'));
    parts['xl/pivotTables/pivotTable1.xml'] = strToU8('<?xml version="1.0"?><pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PivotTable1" cacheId="1" rowGrandTotals="0" colGrandTotals="1" subtotalTop="1"><location ref="D1:E3"/><pivotFields count="2"><pivotField axis="axisRow" defaultSubtotal="0" sortType="descending"><autoSortScope><pivotArea dataOnly="0" fieldPosition="0"><references count="1"><reference field="4294967294" count="1" selected="0"><x v="0"/></reference></references></pivotArea></autoSortScope></pivotField><pivotField/></pivotFields><rowFields count="1"><field x="0"/></rowFields><dataFields count="1"><dataField fld="1" name="Sum of Amount" subtotal="sum" showDataAs="difference" baseField="0" baseItem="0" numFmtId="2"/></dataFields><pivotFilters count="4"><filter fld="0" type="captionEqual" stringValue1="A"/><filter fld="0" type="valueGreaterThan" iMeasureFld="0" val="10"/><filter fld="0" type="valueTop10" iMeasureFld="0" val="3" top="1"/><filter fld="1" type="futureFilter" id="7" stringValue1="preserve"/></pivotFilters></pivotTableDefinition>');
    const imported = await importXlsx({ fileName: 'native-pivot.xlsx', buffer: zipXlsxPartsBuffer(parts), options: { compatibilityTarget: 'B' } });
    assert.equal(imported.nativePackage.packageGraph.nativePivotGraph?.caches[0]?.source.kind, 'worksheet-range');
    assert.equal(imported.nativePackage.packageGraph.nativePivotGraph?.caches[0]?.fields[1]?.name, 'Amount');
    assert.equal(imported.nativePackage.packageGraph.nativePivotGraph?.tables[0]?.dataFields[0]?.field, 1);
    assert.equal(imported.nativePackage.packageGraph.nativePivotGraph?.tables[0]?.dataFields[0]?.numberFormat, '0.00');
    assert.equal(imported.nativePackage.packageGraph.nativePivotGraph?.tables[0]?.dataFields[0]?.baseField, 0);
    assert.equal(imported.nativePackage.packageGraph.nativePivotGraph?.tables[0]?.dataFields[0]?.baseItem, 0);
    assert.equal(imported.snapshot.sheets[0]?.pivots.length, 1);
    assert.equal(imported.snapshot.sheets[0]?.pivots[0]?.source.kind, 'worksheet-range');
    assert.equal(imported.snapshot.sheets[0]?.pivots[0]?.layout.rows[0]?.fieldId, 'native:cache:1:field:0');
    assert.equal(imported.snapshot.sheets[0]?.pivots[0]?.layout.values[0]?.numberFormat, '0.00');
    assert.deepEqual(imported.snapshot.sheets[0]?.pivots[0]?.layout.values[0]?.showAs, { kind: 'difference', baseFieldId: 'native:cache:1:field:0', baseItem: createPivotMemberKey('A') });
    assert.equal(imported.snapshot.sheets[0]?.pivots[0]?.refreshPolicy.mode, 'manual');
    assert.equal(imported.snapshot.sheets[0]?.pivots[0]?.refreshPolicy.refreshOnLoad, false);
    assert.deepEqual(imported.snapshot.sheets[0]?.pivots[0]?.nativeMetadata?.cacheFlags, {});
    assert.deepEqual(imported.snapshot.sheets[0]?.pivots[0]?.layout.rows[0]?.subtotal, { mode: 'none' });
    assert.deepEqual(imported.snapshot.sheets[0]?.pivots[0]?.layout.rows[0]?.sort, { direction: 'descending', by: 'label' });
    assert.deepEqual(imported.snapshot.sheets[0]?.pivots[0]?.layout.filters, [
      { kind: 'condition', family: 'label', fieldId: 'native:cache:1:field:0', operator: 'equals', value: 'A', scope: 'field' },
      { kind: 'condition', family: 'value', fieldId: 'native:cache:1:field:0', valueId: 'native:cache:1:value:0', operator: 'greater-than', value: 10, scope: 'field' },
      { kind: 'top-items', family: 'top-items', fieldId: 'native:cache:1:field:0', valueId: 'native:cache:1:value:0', mode: 'items', threshold: 3, direction: 'top', scope: 'field' },
    ]);
    assert.deepEqual(imported.snapshot.sheets[0]?.pivots[0]?.nativeMetadata?.preservedPivotFilters, [{ fieldIndex: 1, type: 'futureFilter', attributes: { fld: '1', type: 'futureFilter', id: '7', stringValue1: 'preserve' } }]);
    assert.equal(imported.snapshot.sheets[0]?.pivots[0]?.layout.subtotalLocation, 'top');
    assert.equal(imported.snapshot.sheets[0]?.pivots[0]?.layout.showRowGrandTotals, false);
    assert.equal(imported.snapshot.sheets[0]?.pivots[0]?.layout.showColumnGrandTotals, true);
    assert.equal(imported.report.issues.find((issue) => issue.feature === 'pivot')?.status, 'editable');
    const exported = await exportXlsx({ snapshot: imported.snapshot, nativePackage: imported.nativePackage, fileName: 'native-pivot.xlsx', options: { compatibilityTarget: 'B' } });
    const output = loadOpcPackageGraph(exported.buffer);
    assert.ok(output.packageGraph.nativePivotGraph?.tables.length === 1);
    assert.match(strFromU8(output.files['xl/workbook.xml']!), /pivotCaches/);
    assert.match(strFromU8(output.files['xl/worksheets/sheet1.xml']!), /pivotTableParts/);
    assert.match(strFromU8(output.files['[Content_Types].xml']!), /pivotTable/);
    const noFlagCacheXml = strFromU8(output.files['xl/pivotCache/pivotCacheDefinition1.xml']!);
    assert.doesNotMatch(noFlagCacheXml, /refreshOnLoad=/);
    assert.doesNotMatch(noFlagCacheXml, /refreshOnSave=/);
    assert.match(strFromU8(output.files['xl/pivotTables/pivotTable1.xml']!), /defaultSubtotal="0"/);
    const nativeAdvancedXml = strFromU8(output.files['xl/pivotTables/pivotTable1.xml']!);
    assert.match(nativeAdvancedXml, /rowGrandTotals="0"/);
    assert.match(nativeAdvancedXml, /colGrandTotals="1"/);
    assert.match(nativeAdvancedXml, /sortType="descending"/);
    assert.match(nativeAdvancedXml, /<autoSortScope><pivotArea[^>]*dataOnly="0"/);
    assert.match(nativeAdvancedXml, /<dataField[^>]*showDataAs="difference"[^>]*baseField="0"[^>]*baseItem="0"/);
    assert.match(nativeAdvancedXml, /<pivotFilters count="4">/);
    assert.match(nativeAdvancedXml, /type="futureFilter"[^>]*id="7"[^>]*stringValue1="preserve"/);

    const editedSnapshot = structuredClone(imported.snapshot);
    const editedPivot = editedSnapshot.sheets[0]?.pivots[0];
    if (!editedPivot) throw new Error('Imported Pivot fixture is missing');
    editedPivot.refreshPolicy = { mode: 'on-change', preserveFormatting: true, refreshOnLoad: true };
    const edited = await exportXlsx({ snapshot: editedSnapshot, nativePackage: imported.nativePackage, fileName: 'native-pivot-edited.xlsx', options: { compatibilityTarget: 'B' } });
    const editedOutput = loadOpcPackageGraph(edited.buffer);
    const editedCacheXml = strFromU8(editedOutput.files['xl/pivotCache/pivotCacheDefinition1.xml']!);
    assert.match(editedCacheXml, /refreshOnLoad="1"/);
    assert.match(editedCacheXml, /refreshOnSave="1"/);
    assert.match(editedCacheXml, /enableRefresh="1"/);

    const flaggedParts = structuredClone(parts);
    flaggedParts['xl/pivotCache/pivotCacheDefinition1.xml'] = strToU8(strFromU8(flaggedParts['xl/pivotCache/pivotCacheDefinition1.xml']!).replace(
      '<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"',
      '<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" refreshOnLoad="1" refreshOnSave="0" saveData="0" enableRefresh="0"',
    ));
    const flaggedImported = await importXlsx({ fileName: 'native-pivot-flags.xlsx', buffer: zipXlsxPartsBuffer(flaggedParts), options: { compatibilityTarget: 'B' } });
    const flaggedPivot = flaggedImported.snapshot.sheets[0]?.pivots[0];
    assert.equal(flaggedPivot?.refreshPolicy.mode, 'on-open');
    assert.deepEqual(flaggedPivot?.nativeMetadata?.cacheFlags, { refreshOnLoad: true, refreshOnSave: false, saveData: false, enableRefresh: false });
    const flaggedEdited = structuredClone(flaggedImported.snapshot);
    const flaggedEditedPivot = flaggedEdited.sheets[0]?.pivots[0];
    if (!flaggedEditedPivot) throw new Error('Flagged imported Pivot fixture is missing');
    flaggedEditedPivot.refreshPolicy = { mode: 'on-change', preserveFormatting: true, refreshOnLoad: true };
    const flaggedExport = await exportXlsx({ snapshot: flaggedEdited, nativePackage: flaggedImported.nativePackage, fileName: 'native-pivot-flags-edited.xlsx', options: { compatibilityTarget: 'B' } });
    const flaggedOutput = loadOpcPackageGraph(flaggedExport.buffer);
    const flaggedCacheXml = strFromU8(flaggedOutput.files['xl/pivotCache/pivotCacheDefinition1.xml']!);
    assert.match(flaggedCacheXml, /refreshOnSave="1"/);
    assert.match(flaggedCacheXml, /saveData="0"/);
    assert.match(flaggedCacheXml, /enableRefresh="1"/);

    const malformedParts = structuredClone(parts);
    malformedParts['xl/pivotTables/pivotTable1.xml'] = strToU8(strFromU8(malformedParts['xl/pivotTables/pivotTable1.xml']!).replace('<filter fld="0" type="captionEqual"', '<filter type="captionEqual"'));
    await assert.rejects(importXlsx({ fileName: 'native-pivot-malformed-filter.xlsx', buffer: zipXlsxPartsBuffer(malformedParts), options: { compatibilityTarget: 'B' } }), /pivotFilters\.filter\[0\]\.fld/);
    const missingBaseFieldParts = structuredClone(parts);
    missingBaseFieldParts['xl/pivotTables/pivotTable1.xml'] = strToU8(strFromU8(missingBaseFieldParts['xl/pivotTables/pivotTable1.xml']!).replace(' baseField="0"', ''));
    await assert.rejects(importXlsx({ fileName: 'native-pivot-missing-base-field.xlsx', buffer: zipXlsxPartsBuffer(missingBaseFieldParts), options: { compatibilityTarget: 'B' } }), /missing baseField/);
  });

  it('round-trips native Difference/%Difference/Running Total operands and Previous/Next sentinels', () => {
    const workbook = new WorkbookModel('wb-native-show-as-operands', 'Native Show As Operands');
    const sheet = workbook.getSheet(workbook.primarySheetId);
    [['Month', 'Amount'], ['Jan', 10], ['Feb', 20]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
    const cache: NativePivotCacheDefinition = {
      cacheId: 7,
      part: 'xl/pivotCache/pivotCacheDefinition7.xml',
      source: { kind: 'worksheet-range', sheetName: sheet.name, sheetPart: 'xl/worksheets/sheet1.xml', ref: 'A1:B3' },
      fields: [
        { index: 0, name: 'Month', dataType: 'string', sharedItems: ['Jan', 'Feb'] },
        { index: 1, name: 'Amount', dataType: 'number', sharedItems: [10, 20] },
      ],
    };
    const baseTable: Omit<NativePivotTableDefinition, 'dataFields'> = {
      name: 'NativeShowAs', part: 'xl/pivotTables/pivotTable7.xml', sheetPart: 'xl/worksheets/sheet1.xml', relationshipId: 'rIdPivotTable', cacheId: 7,
      locationRef: 'D1:E4', fields: [{ index: 0, axis: 'row' }, { index: 1 }], rowFields: [0], columnFields: [], pageFields: [],
    };
    const map = (dataField: NativePivotTableDefinition['dataFields'][number]) => mapNativePivotDefinition({ ...baseTable, dataFields: [dataField] }, cache, workbook.snapshot(), { [sheet.id]: 'xl/worksheets/sheet1.xml' })!;

    const difference = map({ field: 1, subtotal: 'sum', showDataAs: 'difference', baseField: 0, baseItem: 0 });
    assert.deepEqual(difference.layout.values[0], {
      valueId: 'native:cache:7:value:0', fieldId: 'native:cache:7:field:1', summarizeBy: 'sum',
      showAs: { kind: 'difference', baseFieldId: 'native:cache:7:field:0', baseItem: createPivotMemberKey('Jan') },
    });
    const percentDifference = map({ field: 1, subtotal: 'sum', showDataAs: 'percentDiff', baseField: 0, baseItem: 'next' });
    assert.deepEqual(percentDifference.layout.values[0]?.showAs, { kind: 'percentage-difference', baseFieldId: 'native:cache:7:field:0', baseItem: 'next' });
    const runningTotal = map({ field: 1, subtotal: 'sum', showDataAs: 'runTotal', baseField: 0 });
    assert.deepEqual(runningTotal.layout.values[0]?.showAs, { kind: 'running-total', baseFieldId: 'native:cache:7:field:0' });

    const snapshot = workbook.snapshot();
    snapshot.sheets[0]!.pivots.push(percentDifference);
    const output = loadOpcPackageGraph(exportSnapshotToXlsxBuffer(snapshot));
    const xml = strFromU8(output.files['xl/pivotTables/pivotTable1.xml']!);
    assert.match(xml, /<dataField[^>]*showDataAs="percentDiff"[^>]*baseField="0"[^>]*baseItem="1048829"/);

    const invalid = { ...baseTable, dataFields: [{ field: 1, subtotal: 'sum', showDataAs: 'difference', baseField: 0, baseItem: 99 }] };
    assert.throws(() => mapNativePivotDefinition(invalid, cache, workbook.snapshot(), { [sheet.id]: 'xl/worksheets/sheet1.xml' }), /baseItem 99 is outside/);
    assert.throws(() => map({ field: 1, subtotal: 'sum', showDataAs: 'difference', baseField: 0 }), /missing baseField\/baseItem/);

    const reordered = new WorkbookModel('wb-native-show-as-reordered', 'Native Show As Reordered');
    const reorderedSheet = reordered.getSheet(reordered.primarySheetId);
    [['Amount', 'Month'], [10, 'Jan'], [20, 'Feb']].forEach((row, rowIndex) => row.forEach((value, columnIndex) => reorderedSheet.cells.set(rowIndex, columnIndex, { value })));
    reorderedSheet.pivots.push({
      schema: 'PivotDefinition', id: 'pivot-reordered',
      source: { kind: 'worksheet-range', range: { sheetId: reorderedSheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 } },
      target: { sheetId: reorderedSheet.id, anchor: { row: 5, column: 0 } },
      fieldCatalog: { schema: 'PivotFieldCatalog', fields: [
        { fieldId: 'amount', name: 'Amount', dataType: 'number', ordinal: 0 },
        { fieldId: 'month', name: 'Month', dataType: 'text', ordinal: 1 },
      ] },
      layout: { rows: [{ fieldId: 'month' }], columns: [], filters: [], allowMultipleFiltersPerField: true, collation: { locale: 'en-US', sensitivity: 'variant', numeric: false, caseFirst: 'false' }, values: [{ valueId: 'value:amount', fieldId: 'amount', summarizeBy: 'sum', showAs: { kind: 'difference', baseFieldId: 'month', baseItem: createPivotMemberKey('Jan') } }], subtotalLocation: 'bottom', showRowGrandTotals: true, showColumnGrandTotals: true, reportLayout: 'compact' },
      refreshPolicy: { mode: 'manual', preserveFormatting: true, refreshOnLoad: false },
    });
    const reorderedOutput = loadOpcPackageGraph(exportSnapshotToXlsxBuffer(reordered.snapshot()));
    assert.match(strFromU8(reorderedOutput.files['xl/pivotTables/pivotTable1.xml']!), /<dataField[^>]*fld="0"[^>]*showDataAs="difference"[^>]*baseField="1"[^>]*baseItem="0"/);
  });

  it('rejects unsupported native Show Values As instead of projecting it as normal', () => {
    const workbook = new WorkbookModel('wb-native-show-as-invalid', 'Native Show As Invalid');
    const sheet = workbook.getSheet(workbook.primarySheetId);
    [['Month', 'Amount'], ['Jan', 10]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
    const cache: NativePivotCacheDefinition = {
      cacheId: 8, part: 'xl/pivotCache/pivotCacheDefinition8.xml', source: { kind: 'worksheet-range', sheetName: sheet.name, ref: 'A1:B2' },
      fields: [{ index: 0, name: 'Month', dataType: 'string', sharedItems: ['Jan'] }, { index: 1, name: 'Amount', dataType: 'number', sharedItems: [10] }],
    };
    const table: NativePivotTableDefinition = {
      name: 'InvalidShowAs', part: 'xl/pivotTables/pivotTable8.xml', sheetPart: 'xl/worksheets/sheet1.xml', relationshipId: 'rIdPivotTable', cacheId: 8,
      locationRef: 'D1:E3', fields: [{ index: 0, axis: 'row' }, { index: 1 }], rowFields: [0], columnFields: [], pageFields: [],
      dataFields: [{ field: 1, subtotal: 'sum', showDataAs: 'unknownCalculation', baseField: 0, baseItem: 0 }],
    };
    assert.throws(() => mapNativePivotDefinition(table, cache, workbook.snapshot(), { [sheet.id]: 'xl/worksheets/sheet1.xml' }), /dataField\.showDataAs is unsupported/);
  });

  it('rejects native dataField operands that are not semantically attached to a calculation', () => {
    const workbook = new WorkbookModel('wb-native-show-as-orphan', 'Native Show As Orphan');
    const sheet = workbook.getSheet(workbook.primarySheetId);
    [['Month', 'Amount'], ['Jan', 10]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
    const cache: NativePivotCacheDefinition = {
      cacheId: 9, part: 'xl/pivotCache/pivotCacheDefinition9.xml', source: { kind: 'worksheet-range', sheetName: sheet.name, ref: 'A1:B2' },
      fields: [{ index: 0, name: 'Month', dataType: 'string', sharedItems: ['Jan'] }, { index: 1, name: 'Amount', dataType: 'number', sharedItems: [10] }],
    };
    const table: NativePivotTableDefinition = {
      name: 'OrphanShowAs', part: 'xl/pivotTables/pivotTable9.xml', sheetPart: 'xl/worksheets/sheet1.xml', relationshipId: 'rIdPivotTable', cacheId: 9,
      locationRef: 'D1:E3', fields: [{ index: 0, axis: 'row' }, { index: 1 }], rowFields: [0], columnFields: [], pageFields: [],
      dataFields: [{ field: 1, subtotal: 'sum', showDataAs: 'normal', baseField: 0, baseItem: 0 }],
    };
    assert.throws(() => mapNativePivotDefinition(table, cache, workbook.snapshot(), { [sheet.id]: 'xl/worksheets/sheet1.xml' }), /has operands for normal/);
  });

  it('rejects malformed native baseField/baseItem attributes at the parser boundary', () => {
    const main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    const files = {
      'xl/workbook.xml': strToU8(`<workbook xmlns="${main}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1"/></sheets><pivotCaches count="1"><pivotCache cacheId="1" r:id="rIdCache"/></pivotCaches></workbook>`),
      'xl/pivotCache/pivotCacheDefinition1.xml': strToU8(`<pivotCacheDefinition xmlns="${main}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><cacheSource type="worksheet"><worksheetSource ref="A1:B2" sheet="Sheet1"/></cacheSource><cacheFields count="2"><cacheField name="Month"><sharedItems containsString="1"><s v="Jan"/></sharedItems></cacheField><cacheField name="Amount"><sharedItems containsNumber="1"><n v="10"/></sharedItems></cacheField></cacheFields></pivotCacheDefinition>`),
      'xl/worksheets/sheet1.xml': strToU8(`<worksheet xmlns="${main}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><pivotTableParts count="1"><pivotTablePart r:id="rIdTable"/></pivotTableParts></worksheet>`),
      'xl/pivotTables/pivotTable1.xml': strToU8(`<pivotTableDefinition xmlns="${main}" name="Bad" cacheId="1"><location ref="D1:E3"/><pivotFields count="2"><pivotField/><pivotField/></pivotFields><rowFields count="0"/><dataFields count="1"><dataField fld="1" showDataAs="difference" baseField="NaN" baseItem="0"/></dataFields></pivotTableDefinition>`),
    };
    const relationships = {
      'xl/workbook.xml': [{ id: 'rIdCache', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition', target: 'pivotCache/pivotCacheDefinition1.xml' }],
      'xl/pivotCache/pivotCacheDefinition1.xml': [],
      'xl/worksheets/sheet1.xml': [{ id: 'rIdTable', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable', target: '../pivotTables/pivotTable1.xml' }],
    };
    assert.throws(() => readNativePivotGraph({ files, relationships, sheetPartById: { 'sheet-1': 'xl/worksheets/sheet1.xml' }, dateSystem: '1900' }), /dataField\.baseField/);
    const unsupportedAttributeFiles = { ...files, 'xl/pivotTables/pivotTable1.xml': strToU8(strFromU8(files['xl/pivotTables/pivotTable1.xml']!).replace('baseField="NaN"', 'extra="1"')) };
    assert.throws(() => readNativePivotGraph({ files: unsupportedAttributeFiles, relationships, sheetPartById: { 'sheet-1': 'xl/worksheets/sheet1.xml' }, dateSystem: '1900' }), /dataField attribute is unsupported/);
  });

  it('rejects conflicting refresh modes before rebuilding a shared native cache', () => {
    const workbook = new WorkbookModel('wb-shared-cache-conflict', 'Shared Cache Conflict');
    const sheet = workbook.getSheet(workbook.primarySheetId);
    sheet.cells.set(0, 0, { value: 'Category' });
    sheet.cells.set(0, 1, { value: 'Amount' });
    sheet.cells.set(1, 0, { value: 'A' });
    sheet.cells.set(1, 1, { value: 10 });
    const layout = { rows: [{ fieldId: 'category' }], columns: [], filters: [], allowMultipleFiltersPerField: true, collation: { locale: 'en-US' as const, sensitivity: 'variant' as const, numeric: false, caseFirst: 'false' as const }, values: [{ valueId: `value:${'amount'}`, fieldId: 'amount', summarizeBy: 'sum' as const }], subtotalLocation: 'bottom' as const, showRowGrandTotals: true, showColumnGrandTotals: true, reportLayout: 'compact' as const };
    const fields = { schema: 'PivotFieldCatalog' as const, fields: [
      { fieldId: 'category', name: 'Category', dataType: 'text' as const, ordinal: 0 },
      { fieldId: 'amount', name: 'Amount', dataType: 'number' as const, ordinal: 1 },
    ] };
    const source = { kind: 'worksheet-range' as const, range: { sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 } };
    sheet.pivots.push(
      { schema: 'PivotDefinition', id: 'shared-manual', source, target: { sheetId: sheet.id, anchor: { row: 4, column: 0 } }, fieldCatalog: fields, layout, refreshPolicy: { mode: 'manual', preserveFormatting: true, refreshOnLoad: false } },
      { schema: 'PivotDefinition', id: 'shared-open', source, target: { sheetId: sheet.id, anchor: { row: 10, column: 0 } }, fieldCatalog: fields, layout, refreshPolicy: { mode: 'on-open', preserveFormatting: true, refreshOnLoad: true } },
    );
    assert.throws(() => exportSnapshotToXlsxBuffer(workbook.snapshot()), /conflicting refresh policies/);
  });

  it('writes canonical value sorting and value filters with stable data-field identity', () => {
    const workbook = new WorkbookModel('wb-pivot-advanced-write', 'Pivot Advanced Write');
    const sheet = workbook.getSheet(workbook.primarySheetId);
    [['Category', 'Amount'], ['A', 10], ['B', 20]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
    sheet.pivots.push({
      schema: 'PivotDefinition', id: 'advanced-write', source: { kind: 'worksheet-range', range: { sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 } }, target: { sheetId: sheet.id, anchor: { row: 5, column: 0 } },
      fieldCatalog: { schema: 'PivotFieldCatalog', fields: [{ fieldId: 'category', name: 'Category', dataType: 'text', ordinal: 0 }, { fieldId: 'amount', name: 'Amount', dataType: 'number', ordinal: 1 }] },
      layout: {
        rows: [{ fieldId: 'category', sort: { direction: 'ascending', by: 'value', valueId: `value:${'amount'}` } }], columns: [],
        filters: [
          { kind: 'condition', family: 'value', fieldId: 'category', valueId: `value:${'amount'}`, operator: 'greater-than', value: 5 },
          { kind: 'top-items', family: 'top-items', fieldId: 'category', valueId: `value:${'amount'}`, direction: 'top', mode: 'sum', threshold: 25 },
        ],
        allowMultipleFiltersPerField: true,
        collation: { locale: 'en-US', sensitivity: 'variant', numeric: false, caseFirst: 'false' },
        values: [{ valueId: `value:${'amount'}`, fieldId: 'amount', summarizeBy: 'sum' }], subtotalLocation: 'bottom', showRowGrandTotals: true, showColumnGrandTotals: true, reportLayout: 'compact',
      },
      refreshPolicy: { mode: 'on-change', preserveFormatting: true, refreshOnLoad: true },
    });
    const output = loadOpcPackageGraph(exportSnapshotToXlsxBuffer(workbook.snapshot()));
    const xml = strFromU8(output.files['xl/pivotTables/pivotTable1.xml']!);
    assert.match(xml, /sortType="ascending"/);
    assert.match(xml, /<autoSortScope><pivotArea[^>]*dataOnly="1"[^>]*fieldPosition="0"/);
    assert.match(xml, /<reference field="1"/);
    assert.match(xml, /type="valueGreaterThan"[^>]*iMeasureFld="0"[^>]*val="5"/);
    assert.match(xml, /type="valueTop10"[^>]*iMeasureFld="0"[^>]*val="25"[^>]*mode="sum"/);
    const pivot = sheet.pivots[0]!;
    pivot.layout.filters = pivot.layout.filters.map((filter) => filter.kind === 'top-items' ? { ...filter, mode: 'percent', threshold: 50 } : filter);
    const percentXml = strFromU8(loadOpcPackageGraph(exportSnapshotToXlsxBuffer(workbook.snapshot())).files['xl/pivotTables/pivotTable1.xml']!);
    assert.match(percentXml, /type="valueTop10"[^>]*val="50"[^>]*mode="percent"/);
    assert.match(percentXml, /type="valueTop10"[^>]*percent="1"/);
  });

  it('round-trips repeated source fields as distinct native Values placements', async () => {
    const workbook = new WorkbookModel('wb-repeated-values', 'Repeated Values');
    const sheet = workbook.getSheet(workbook.primarySheetId);
    [['Category', 'Amount'], ['A', 10], ['B', 20]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
    sheet.pivots.push({
      schema: 'PivotDefinition', id: 'repeated-values',
      source: { kind: 'worksheet-range', range: { sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 } },
      target: { sheetId: sheet.id, anchor: { row: 5, column: 0 } },
      fieldCatalog: { schema: 'PivotFieldCatalog', fields: [{ fieldId: 'category', name: 'Category', dataType: 'text', ordinal: 0 }, { fieldId: 'amount', name: 'Amount', dataType: 'number', ordinal: 1 }] },
      layout: {
        rows: [{ fieldId: 'category' }], columns: [], filters: [], allowMultipleFiltersPerField: true,
        collation: { locale: 'en-US', sensitivity: 'variant', numeric: false, caseFirst: 'false' },
        values: [
          { valueId: 'amount:sum', fieldId: 'amount', summarizeBy: 'sum' },
          { valueId: 'amount:count', fieldId: 'amount', summarizeBy: 'count' },
        ],
        subtotalLocation: 'bottom', showRowGrandTotals: true, showColumnGrandTotals: true, reportLayout: 'compact',
      },
      refreshPolicy: { mode: 'on-change', preserveFormatting: true, refreshOnLoad: true },
    });
    const buffer = exportSnapshotToXlsxBuffer(workbook.snapshot());
    const graph = loadOpcPackageGraph(buffer);
    const tableXml = strFromU8(graph.files['xl/pivotTables/pivotTable1.xml']!);
    assert.match(tableXml, /<dataFields count="2">/);
    assert.equal((tableXml.match(/<dataField /g) ?? []).length, 2);
    const imported = await importXlsx({ fileName: 'repeated-values.xlsx', buffer, options: { compatibilityTarget: 'B' } });
    const values = imported.snapshot.sheets[0]?.pivots[0]?.layout.values ?? [];
    assert.equal(values.length, 2);
    assert.equal(new Set(values.map((value) => value.valueId)).size, 2);
    assert.deepEqual(values.map((value) => value.fieldId), ['native:cache:1:field:1', 'native:cache:1:field:1']);
  });

  it('round-trips styles, merges, scoped names, freeze and the 1904 date system', async () => {
    const workbook = new WorkbookModel('wb-rich', 'Rich');
    const sheet = workbook.getSheet(workbook.primarySheetId);
    sheet.cells.set(0, 0, {
      value: 45292,
      style: { bold: true, background: '#ff0000', horizontalAlignment: 'center', indent: 2 },
      numberFormat: 'm/d/yy',
    });
    sheet.cells.set(1, 1, { value: 'merged' });
    sheet.merges.push({
      range: { sheetId: sheet.id, startRow: 1, endRow: 1, startColumn: 1, endColumn: 2 },
      anchor: { row: 1, column: 1 },
    });
    sheet.pane = { kind: 'frozen', xSplit: 1, ySplit: 1, startRow: 1, startColumn: 1, state: 'frozen' };
    workbook.setCellStyleTemplate({ id: 'input-style', name: 'Input Style', style: { background: '#e3f1ff', indent: 2, numberFormat: '#,##0.00' } });
    workbook.setDefinedName({ name: 'LocalValue', formula: `=${sheet.name}!$A$1`, scope: 'sheet', sheetId: sheet.id });
    const original = workbook.snapshot();
    const buffer = exportSnapshotToXlsxBuffer(original, undefined, { dateSystem: '1904' });
    const imported = await importXlsx({ fileName: 'rich.xlsx', buffer, options: { compatibilityTarget: 'A' } });
    const restored = imported.snapshot.sheets[0]!;
    assert.equal(imported.report.dateSystem, '1904');
    assert.equal(restored.merges.length, 1);
    assert.deepEqual(restored.pane, original.sheets[0]!.pane);
    assert.equal(restored.cells['0']?.['0']?.numberFormat, 'm/d/yy');
    assert.equal(restored.cells['0']?.['0']?.style?.bold, true);
    assert.equal(restored.cells['0']?.['0']?.style?.indent, 2);
    assert.equal(imported.snapshot.cellStyleTemplates?.find((template) => template.name === 'Input Style')?.style.indent, 2);
    assert.equal(imported.snapshot.definedNameModels?.[0]?.scope, 'sheet');
  });

  it('keeps OOXML cell types authoritative for numeric-looking text, numbers, and date-formatted serials', async () => {
    const workbook = new WorkbookModel('wb-cell-types', 'Cell types');
    const sheet = workbook.getSheet(workbook.primarySheetId);
    sheet.cells.set(0, 0, { value: '260' });
    sheet.cells.set(0, 1, { value: 260 });
    sheet.cells.set(0, 2, { value: 45292, numberFormat: 'm/d/yy' });
    const imported = await importXlsx({
      fileName: 'cell-types.xlsx',
      buffer: exportSnapshotToXlsxBuffer(workbook.snapshot()),
      options: { compatibilityTarget: 'A' },
    });
    const cells = imported.snapshot.sheets[0]!.cells['0']!;
    assert.equal(cells['0']?.value, '260');
    assert.equal(typeof cells['0']?.value, 'string');
    assert.equal(cells['1']?.value, 260);
    assert.equal(typeof cells['1']?.value, 'number');
    assert.equal(cells['2']?.value, '2024-01-01T00:00:00.000Z');
    assert.equal(typeof cells['2']?.value, 'string');
    assert.equal(cells['2']?.numberFormat, 'm/d/yy');
  });

  it('preserves opaque chart/binary parts and relationships across an editable export', async () => {
    const workbook = new WorkbookModel('wb-preserve', 'Preserve');
    workbook.getSheet(workbook.primarySheetId).cells.set(0, 0, { value: 1 });
    const generated = loadOpcPackageGraph(exportSnapshotToXlsxBuffer(workbook.snapshot()));
    generated.packageGraph.parts['xl/charts/chart1.xml'] = strToU8('<chartSpace xmlns="http://schemas.openxmlformats.org/drawingml/2006/chart"><title>Keep</title></chartSpace>');
    generated.packageGraph.parts['xl/drawings/drawing1.xml'] = strToU8('<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>');
    generated.packageGraph.parts['customXml/item1.bin'] = Uint8Array.from([0, 1, 2, 255]);
    generated.packageGraph.parts['xl/worksheets/_rels/sheet1.xml.rels'] = strToU8('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>');
    generated.packageGraph.relationships['xl/worksheets/sheet1.xml'] = [{
      id: 'rIdChart',
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
      target: '../drawings/drawing1.xml',
    }];
    generated.packageGraph.parts['xl/worksheets/sheet1.xml'] = strToU8(strFromU8(generated.packageGraph.parts['xl/worksheets/sheet1.xml']!).replace('</worksheet>', '<drawing r:id="rIdChart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></worksheet>'));
    // Rebuild through the public package writer so this test exercises the
    // same ZIP limits and relationship reader used by production imports.
    const imported = await importXlsx({ fileName: 'opaque.xlsx', buffer: zipXlsxPartsBuffer(generated.packageGraph.parts), options: { compatibilityTarget: 'B', preserveMacros: true } });
    const exported = await exportXlsx({ snapshot: imported.snapshot, nativePackage: imported.nativePackage, fileName: 'opaque.xlsx', options: { compatibilityTarget: 'B' } });
    const restored = loadOpcPackageGraph(exported.buffer);
    assert.deepEqual([...restored.files['customXml/item1.bin']!], [0, 1, 2, 255]);
    assert.equal(strFromU8(restored.files['xl/charts/chart1.xml']!).includes('Keep'), true);
    assert.equal(strFromU8(restored.files['xl/worksheets/sheet1.xml']!).includes('rIdChart'), true);
    assert.equal(imported.report.issues.some((issue) => issue.feature === 'charts' && issue.preserved), true);
    assert.equal(exported.report.issues.some((issue) => issue.feature === 'charts' && issue.preserved), true);
  });

  it('rejects oversized and unsafe ZIP entries before inflation', () => {
    const workbook = new WorkbookModel('wb-limit', 'Limit');
    const buffer = exportSnapshotToXlsxBuffer(workbook.snapshot());
    assert.throws(() => loadOpcPackageGraph(buffer, { maxArchiveBytes: 10 }), /archive exceeds/);
    assert.throws(() => loadOpcPackageGraph(buffer, { maxEntries: 1 }), /too many entries/);
    assert.throws(() => loadOpcPackageGraph(buffer, { maxCompressionRatio: 1 }), /compression ratio/);
  });

  it('imports native OOXML geometry, shared formulas, split panes and worksheet capabilities', async () => {
    const workbook = new WorkbookModel('wb-native-geometry', 'Native Geometry');
    const generated = loadOpcPackageGraph(exportSnapshotToXlsxBuffer(workbook.snapshot()));
    generated.packageGraph.parts['xl/styles.xml'] = strToU8('<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><sz val="18"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>');
    generated.packageGraph.parts['xl/worksheets/sheet1.xml'] = strToU8('<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:B2"/><sheetViews><sheetView workbookViewId="0"><pane xSplit="2000" ySplit="3000" topLeftCell="B2" activePane="bottomRight" state="split"/></sheetView></sheetViews><sheetFormatPr defaultColWidth="8.7109375" defaultRowHeight="15"/><cols><col min="1" max="1" width="9.625" customWidth="1"/></cols><sheetData><row r="1" ht="15" customHeight="1"><c r="A1"><v>2</v></c><c r="B1" s="1"><f t="shared" si="0" ref="B1:B2">A1*2</f><v>4</v></c></row><row r="2"><c r="A2"><v>3</v></c><c r="B2"><f t="shared" si="0"/><v>6</v></c></row></sheetData><autoFilter ref="A1:B2"/><conditionalFormatting sqref="A1:A2"><cfRule type="cellIs" operator="greaterThan" priority="1"><formula>1</formula></cfRule></conditionalFormatting><dataValidations count="1"><dataValidation type="whole" operator="greaterThan" sqref="A1:A2"><formula1>0</formula1></dataValidation></dataValidations></worksheet>');
    const imported = await importXlsx({ fileName: 'CAN配置.xlsx', buffer: zipXlsxPartsBuffer(generated.packageGraph.parts), options: { compatibilityTarget: 'B', compatibilityMode: 'balanced' } });
    const sheet = imported.snapshot.sheets[0]!;
    assert.equal(imported.snapshot.name, 'CAN配置');
    assert.ok((sheet.columnWidthsPx?.[0] ?? 0) >= 67);
    assert.equal(sheet.defaultColumnWidthPx, 61);
    assert.equal(sheet.rowHeightsPx?.[0], 20);
    assert.equal(sheet.cells['0']?.['1']?.style?.fontSizePx, 24);
    assert.equal(sheet.cells['1']?.['1']?.formula, '=A2*2');
    assert.equal(sheet.pane.kind, 'split');
    assert.equal(sheet.conditionalFormats?.length, 1);
    assert.equal(sheet.dataValidations?.length, 1);
    assert.equal(sheet.autoFilter?.range.endColumn, 1);
  });

  it('normalizes Excel and WPS formula namespaces only for runtime and restores their source spelling on export', async () => {
    const workbook = new WorkbookModel('wb-formula-namespace', 'Formula namespace');
    const generated = loadOpcPackageGraph(exportSnapshotToXlsxBuffer(workbook.snapshot()));
    generated.packageGraph.parts['xl/worksheets/sheet1.xml'] = strToU8('<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetFormatPr defaultRowHeight="15"/><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>_xlfn.SUM(A1:A1)</f><v>1</v></c><c r="C1"><f>_xlfn._xlws.FILTER(A1:A1,A1:A1&gt;0)</f><v>1</v></c><c r="D1"><f>_xlfn.SINGLE(A1)</f><v>1</v></c></row></sheetData></worksheet>');
    const imported = await importXlsx({ fileName: 'formula-namespace.xlsx', buffer: zipXlsxPartsBuffer(generated.packageGraph.parts), options: { compatibilityTarget: 'B' } });
    const sheet = imported.snapshot.sheets[0]!;
    assert.equal(sheet.cells['0']?.['1']?.formula, '=SUM(A1:A1)');
    assert.equal(sheet.cells['0']?.['1']?.formulaMetadata?.sourceFormula, '=_xlfn.SUM(A1:A1)');
    assert.equal(sheet.cells['0']?.['2']?.formula, '=FILTER(A1:A1,A1:A1>0)');
    assert.equal(sheet.cells['0']?.['2']?.formulaMetadata?.sourceFormula, '=_xlfn._xlws.FILTER(A1:A1,A1:A1>0)');
    assert.equal(sheet.cells['0']?.['3']?.formula, '=@(A1)');
    assert.equal(sheet.cells['0']?.['3']?.formulaMetadata?.sourceFormula, '=_xlfn.SINGLE(A1)');

    const exported = loadOpcPackageGraph(exportSnapshotToXlsxBuffer(imported.snapshot));
    const worksheet = strFromU8(exported.files['xl/worksheets/sheet1.xml']!);
    assert.match(worksheet, /_xlfn\.SUM\(A1:A1\)/);
    assert.match(worksheet, /_xlfn\._xlws\.FILTER\(A1:A1,A1:A1&gt;0\)/);
    assert.match(worksheet, /_xlfn\.SINGLE\(A1\)/);
  });

  it('accepts supported dynamic AutoFilters and rejects unknown OOXML types', () => {
    const workbook = new WorkbookModel('wb-dynamic-filter', 'Dynamic Filter');
    const generated = loadOpcPackageGraph(exportSnapshotToXlsxBuffer(workbook.snapshot()));
    const worksheet = strFromU8(generated.packageGraph.parts['xl/worksheets/sheet1.xml']!);
    const withType = (type: string): ArrayBuffer => zipXlsxPartsBuffer({
      ...generated.packageGraph.parts,
      'xl/worksheets/sheet1.xml': strToU8(worksheet.replace('</worksheet>', `<autoFilter ref="A1:A2"><filterColumn colId="0"><dynamicFilter type="${type}"/></filterColumn></autoFilter></worksheet>`)),
    } as Record<string, Uint8Array>);

    const supported = parseLoadedXlsx(loadOpcPackageGraph(withType('today')), { canonicalReferenceDate: { year: 2026, month: 8, day: 26, hour: 12, minute: 0, second: 0, millisecond: 0 } }).snapshot;
    assert.deepEqual(supported.sheets[0]?.autoFilter?.columns[0]?.criterion, { kind: 'dynamic', type: 'today' });
    assert.throws(
      () => parseLoadedXlsx(loadOpcPackageGraph(withType('attackerUnknown'))),
      /UNSUPPORTED_FEATURE: dynamic AutoFilter type "attackerUnknown" is not supported/,
    );
  });

  it('round-trips typed date-group criteria and preserves unsupported date-group nodes', () => {
    const workbook = new WorkbookModel('wb-date-group-filter', 'Date Group Filter');
    const generated = loadOpcPackageGraph(exportSnapshotToXlsxBuffer(workbook.snapshot()));
    const worksheet = strFromU8(generated.packageGraph.parts['xl/worksheets/sheet1.xml']!);
    const autoFilter = '<autoFilter ref="A1:A2"><filterColumn colId="0"><filters blank="1"><filter val="literal"/><dateGroupItem dateTimeGrouping="year" year="2026"/><dateGroupItem dateTimeGrouping="month" year="2026" month="8"/><dateGroupItem dateTimeGrouping="day" year="2026" month="8" day="15"/><dateGroupItem dateTimeGrouping="hour" year="2026" month="8" day="15" hour="13"/><dateGroupItem dateTimeGrouping="minute" year="2026" month="8" day="15" hour="13" minute="14"/><dateGroupItem dateTimeGrouping="second" year="2026" month="8" day="15" hour="13" minute="14" second="15"/><dateGroupItem dateTimeGrouping="quarter" year="2026" quarter="3"/><dateGroupItem dateTimeGrouping="year" year="2026" month="8"/></filters></filterColumn></autoFilter>';
    const parts = { ...generated.packageGraph.parts, 'xl/worksheets/sheet1.xml': strToU8(worksheet.replace('</worksheet>', `${autoFilter}</worksheet>`)) };
    const imported = parseLoadedXlsx(loadOpcPackageGraph(zipXlsxPartsBuffer(parts))).snapshot;
    const column = imported.sheets[0]?.autoFilter?.columns[0];
    assert.deepEqual(column?.criterion, {
      kind: 'values',
      values: ['literal'],
      includeBlank: true,
      dateGroups: [
        { year: 2026 },
        { year: 2026, month: 8 },
        { year: 2026, month: 8, day: 15 },
        { year: 2026, month: 8, day: 15, hour: 13 },
        { year: 2026, month: 8, day: 15, hour: 13, minute: 14 },
        { year: 2026, month: 8, day: 15, hour: 13, minute: 14, second: 15 },
      ],
    });
    const exported = loadOpcPackageGraph(exportSnapshotToXlsxBuffer(imported));
    const exportedWorksheet = strFromU8(exported.files['xl/worksheets/sheet1.xml']!);
    assert.match(exportedWorksheet, /dateTimeGrouping="quarter" year="2026" quarter="3"/);
    assert.match(exportedWorksheet, /dateTimeGrouping="year" year="2026" month="8"/);
  });

  it('resolves and preserves Strict relationship kinds and a non-standard workbook part path', async () => {
    const parts: Record<string, Uint8Array> = {
      '[Content_Types].xml': strToU8('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'),
      '_rels/.rels': strToU8('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument" Target="custom/book.xml"/></Relationships>'),
      'custom/book.xml': strToU8('<?xml version="1.0"?><workbook xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main" xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"><sheets><sheet name="Strict" sheetId="1" r:id="rIdSheet"/></sheets></workbook>'),
      'custom/_rels/book.xml.rels': strToU8('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdSheet" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet" Target="worksheets/main.xml"/></Relationships>'),
      'custom/worksheets/main.xml': strToU8('<?xml version="1.0"?><worksheet xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"><sheetFormatPr defaultRowHeight="15" defaultColWidth="8.7109375"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>ok</t></is></c></row></sheetData></worksheet>'),
    };
    const imported = await importXlsx({ fileName: 'strict.xlsx', buffer: zipXlsxPartsBuffer(parts), options: { compatibilityTarget: 'B' } });
    assert.equal(imported.nativePackage.format.family, 'ooxml');
    assert.equal(imported.nativePackage.format.profile, 'strict');
    assert.equal(imported.nativePackage.packageGraph.workbookPart, 'custom/book.xml');
    assert.equal(imported.snapshot.sheets[0]?.cells['0']?.['0']?.value, 'ok');
    const exported = await exportXlsx({ snapshot: imported.snapshot, nativePackage: imported.nativePackage, fileName: 'strict.xlsx', options: { compatibilityTarget: 'B' } });
    const output = loadOpcPackageGraph(exported.buffer, {}, 'strict.xlsx');
    assert.equal(output.packageGraph.workbookPart, 'custom/book.xml');
    assert.equal(output.packageGraph.profile, 'strict');
    assert.match(strFromU8(output.files['custom/book.xml']!), /purl\.oclc\.org\/ooxml\/spreadsheetml\/main/);
  });

  it('round-trips editable rich-text run formatting including superscript and subscript', async () => {
    const workbook = new WorkbookModel('rich-text-editing', 'Rich Text Editing');
    const sheet = workbook.getSheet(workbook.primarySheetId);
    sheet.cells.set(0, 0, { value: 'H2O', richText: [
      { text: 'H', style: { bold: true, textColor: '#2563EB' } },
      { text: '2', style: { verticalAlignment: 'subscript' } },
      { text: 'O', style: { italic: true } },
    ] });
    const imported = await importXlsx({ fileName: 'rich-text.xlsx', buffer: exportSnapshotToXlsxBuffer(workbook.snapshot()), options: { compatibilityTarget: 'B' } });
    const richText = imported.snapshot.sheets[0]?.cells['0']?.['0']?.richText;
    assert.equal(richText?.map((run) => run.text).join(''), 'H2O');
    assert.equal(richText?.[1]?.style?.verticalAlignment, 'subscript');
    assert.equal(richText?.[0]?.style?.bold, true);
  });

  it('round-trips the canonical workbook editing options through owned OOXML metadata', async () => {
    const workbook = new WorkbookModel('editing-options', 'Editing Options');
    workbook.setEditingOptions({ allowEditDirectly: false, moveAfterEnter: true, enterDirection: 'right', formulaAutoComplete: false, valueAutoComplete: true, fixedDecimalPlaces: 3 });
    const imported = await importXlsx({ fileName: 'editing-options.xlsx', buffer: exportSnapshotToXlsxBuffer(workbook.snapshot()), options: { compatibilityTarget: 'B' } });
    assert.deepEqual(imported.snapshot.editingOptions, workbook.editingOptions);
  });

  it('round-trips East Asian phonetic runs without changing canonical cell text', async () => {
    const workbook = new WorkbookModel('phonetic-roundtrip', 'Phonetic Roundtrip');
    const sheet = workbook.getSheet(workbook.primarySheetId);
    sheet.cells.set(0, 0, { value: '東京', phonetic: { visible: true, type: 'hiragana', alignment: 'center', fontFamily: 'Microsoft YaHei', fontSizePx: 8, runs: [{ text: 'とうきょう', start: 0, end: 2 }] } });
    const imported = await importXlsx({ fileName: 'phonetic.xlsx', buffer: exportSnapshotToXlsxBuffer(workbook.snapshot()), options: { compatibilityTarget: 'B' } });
    const cell = imported.snapshot.sheets[0]?.cells['0']?.['0'];
    assert.equal(cell?.value, '東京');
    assert.deepEqual(cell?.phonetic?.runs, [{ text: 'とうきょう', start: 0, end: 2 }]);
    assert.equal(cell?.phonetic?.type, 'hiragana');
  });
});
