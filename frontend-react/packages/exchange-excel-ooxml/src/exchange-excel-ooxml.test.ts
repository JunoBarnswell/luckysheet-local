import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { createPivotMemberKey, WorkbookModel } from '@react-sheets/core-model';
import { exportXlsx } from './export';
import { importXlsx } from './import';
import { scanSnapshotFeatures } from './feature-scan';
import { exportSnapshotToXlsxBuffer } from './archive';
import { loadOpcPackageGraph, parseLoadedXlsx, zipXlsxPartsBuffer } from './archive';
import { readNativePivotGraph } from './native-pivot';
import { strFromU8, strToU8 } from 'fflate';
import { readFile } from 'node:fs/promises';

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
      sourceRanges: [{ sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 }],
      elements: { hiddenData: 'show' },
    });
    sheet.drawingPayloads.set('image-payload', { kind: 'image', src: 'data:image/png;base64,AA==' });

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
    sheet.cells.set(0, 0, { value: 'OpenAI' });
    sheet.hyperlinks.set('0:0', { id: 'link-1', target: { kind: 'url', url: 'https://openai.com/' }, tooltip: 'Open' });
    const buffer = exportSnapshotToXlsxBuffer(workbook.snapshot());
    const emitted = loadOpcPackageGraph(buffer);
    assert.match(strFromU8(emitted.files['xl/worksheets/sheet1.xml']!), /<hyperlink ref="A1"/);
    assert.match(strFromU8(emitted.files['xl/worksheets/_rels/sheet1.xml.rels']!), /https:\/\/openai\.com\//);
    const imported = await importXlsx({ fileName: 'links.xlsx', buffer, options: { compatibilityTarget: 'B' } });
    assert.equal(imported.snapshot.sheets[0]?.hyperlinks?.[0]?.hyperlink.target.kind, 'url');
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
      layout: { rows: [{ fieldId: 'category' }], columns: [], filters: [], allowMultipleFiltersPerField: true, collation: { locale: 'en-US', sensitivity: 'variant', numeric: false, caseFirst: 'false' }, values: [{ fieldId: 'amount', summarizeBy: 'sum' }], subtotalLocation: 'bottom', showGrandTotals: true, compact: true, repeatLabels: false },
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
    const imported = parseLoadedXlsx(output).snapshot;
    assert.equal(imported.sheets[0]?.pivots[0]?.source.kind, 'table');
    assert.deepEqual(imported.sheets[0]?.pivots[0]?.presentation, {
      styleName: 'PivotStyleMedium4',
      styleOptions: { showRowHeaders: false, showColumnHeaders: true, showRowStripes: true, showColumnStripes: true, showLastColumn: true },
      displayOptions: { fillEmptyCells: true, emptyCellText: '—', showErrorValues: false, errorCellText: 'ERR', showFieldHeaders: false, autoFitColumnsOnUpdate: true },
    });
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
        filters: [], allowMultipleFiltersPerField: true, collation: { locale: 'en-US', sensitivity: 'variant', numeric: false, caseFirst: 'false' }, values: [{ fieldId: 'amount', summarizeBy: 'sum' }], subtotalLocation: 'bottom', showGrandTotals: true, compact: true, repeatLabels: false,
      },
      refreshPolicy: { mode: 'on-change', preserveFormatting: true, refreshOnLoad: true },
    });
    const output = loadOpcPackageGraph(exportSnapshotToXlsxBuffer(workbook.snapshot()));
    const cacheXml = strFromU8(output.files['xl/pivotCache/pivotCacheDefinition1.xml']!);
    assert.match(cacheXml, /<fieldGroup base="0"><rangePr groupBy="months" startNum="45292" endNum="45657"\/>/);
    assert.match(cacheXml, /<fieldGroup base="1"><rangePr groupBy="range" groupInterval="10" startNum="0" endNum="100"\/>/);
    assert.match(cacheXml, /<fieldGroup base="2"><discretePr count="3"><x v="0"\/><x v="0"\/><x v="1"\/><\/discretePr><groupItems count="2"><s v="AB"\/><s v="C"\/><\/groupItems><\/fieldGroup>/);
    const imported = parseLoadedXlsx(output).snapshot;
    const pivot = imported.sheets[0]?.pivots[0];
    assert.equal(pivot?.layout.rows[0]?.group?.kind, 'date');
    assert.deepEqual(pivot?.layout.rows[0]?.group, { kind: 'date', unit: 'month', start: 45292, end: 45657 });
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
        { fieldId: 'date', name: 'Date', dataType: 'date', ordinal: 1 },
        { fieldId: 'amount', name: 'Amount', dataType: 'number', ordinal: 2 },
      ] },
      layout: { rows: [{ fieldId: 'category' }], columns: [], filters: [], allowMultipleFiltersPerField: true, collation: { locale: 'en-US', sensitivity: 'variant', numeric: false, caseFirst: 'false' }, values: [{ fieldId: 'amount', summarizeBy: 'sum' }], subtotalLocation: 'bottom', showGrandTotals: true, compact: true, repeatLabels: false },
      refreshPolicy: { mode: 'on-change', preserveFormatting: true, refreshOnLoad: true },
    });
    sheet.drawings.push(
      { id: 'category-slicer', sheetId: sheet.id, kind: 'slicer', anchor: { kind: 'one-cell', row: 0, column: 4 }, transform: { x: 0, y: 0, width: 220, height: 180 }, zIndex: 1, payloadId: 'category-slicer' },
      { id: 'date-timeline', sheetId: sheet.id, kind: 'timeline', anchor: { kind: 'one-cell', row: 10, column: 0 }, transform: { x: 0, y: 0, width: 420, height: 120 }, zIndex: 1, payloadId: 'date-timeline' },
    );
    sheet.drawingPayloads.set('category-slicer', { kind: 'slicer', pivotId: 'control-pivot', fieldId: 'category', filter: { mode: 'all', memberKeys: [] }, style: { theme: 'light', fill: '#fff', border: '#ddd', textColor: '#111', accentColor: '#2563eb' } });
    sheet.drawingPayloads.set('date-timeline', { kind: 'timeline', pivotId: 'control-pivot', fieldId: 'date', period: { start: '2024-01-01T00:00:00Z', end: '2024-12-31T00:00:00Z' }, style: { theme: 'light', fill: '#fff', border: '#ddd', textColor: '#111', accentColor: '#2563eb' } });
    const output = loadOpcPackageGraph(exportSnapshotToXlsxBuffer(workbook.snapshot()));
    assert.ok(output.files['xl/slicerCaches/slicerCache1.xml']);
    assert.ok(output.files['xl/slicers/slicer1.xml']);
    assert.ok(output.files['xl/timelineCaches/timelineCache1.xml']);
    assert.ok(output.files['xl/timelines/timeline1.xml']);
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

  it('keeps native Slicer selections typed when member labels collide', () => {
    const workbook = new WorkbookModel('wb-typed-slicer', 'Typed Slicer');
    const sheet = workbook.getSheet(workbook.primarySheetId);
    [["Member", "Amount"], [1, 10], ["1", 20], [true, 30], ["true", 40], [null, 50]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
    sheet.pivots.push({
      schema: 'PivotDefinition', id: 'typed-slicer-pivot', source: { kind: 'worksheet-range', range: { sheetId: sheet.id, startRow: 0, endRow: 5, startColumn: 0, endColumn: 1 } }, target: { sheetId: sheet.id, anchor: { row: 8, column: 0 } },
      fieldCatalog: { schema: 'PivotFieldCatalog', fields: [{ fieldId: 'member', name: 'Member', dataType: 'mixed', ordinal: 0, values: [1, '1', true, 'true', null] }, { fieldId: 'amount', name: 'Amount', dataType: 'number', ordinal: 1 }] },
      layout: { rows: [{ fieldId: 'member' }], columns: [], filters: [], allowMultipleFiltersPerField: true, collation: { locale: 'en-US', sensitivity: 'variant', numeric: false, caseFirst: 'false' }, values: [{ fieldId: 'amount', summarizeBy: 'sum' }], subtotalLocation: 'bottom', showGrandTotals: true, compact: true, repeatLabels: false },
      refreshPolicy: { mode: 'on-change', preserveFormatting: true, refreshOnLoad: true },
    });
    sheet.drawings.push({ id: 'typed-slicer', sheetId: sheet.id, kind: 'slicer', anchor: { kind: 'one-cell', row: 0, column: 4 }, transform: { x: 0, y: 0, width: 220, height: 180 }, zIndex: 1, payloadId: 'typed-slicer-payload' });
    sheet.drawingPayloads.set('typed-slicer-payload', { kind: 'slicer', pivotId: 'typed-slicer-pivot', fieldId: 'member', filter: { mode: 'include', memberKeys: [createPivotMemberKey('1')] }, style: { theme: 'light', fill: '#fff', border: '#ddd', textColor: '#111', accentColor: '#2563eb' } });
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
      layout: { rows: [{ fieldId: 'category' }], columns: [], filters: [], allowMultipleFiltersPerField: true, collation: { locale: 'en-US', sensitivity: 'variant', numeric: false, caseFirst: 'false' }, values: [{ fieldId: 'amount', summarizeBy: 'sum' }], subtotalLocation: 'bottom', showGrandTotals: true, compact: true, repeatLabels: false },
      refreshPolicy: { mode: 'on-change', preserveFormatting: true, refreshOnLoad: true },
    });
    sheet.drawings.push({ id: 'pivot-chart-drawing', sheetId: sheet.id, kind: 'chart', anchor: { kind: 'one-cell', row: 0, column: 4 }, transform: { x: 0, y: 0, width: 320, height: 220 }, zIndex: 1, payloadId: 'pivot-chart-payload' });
    sheet.drawingPayloads.set('pivot-chart-payload', { kind: 'chart', chartId: 'pivot-chart-1', chartType: 'column', pivotId: 'pivot-chart-1', sourceRanges: [], elements: { hiddenData: 'show', title: 'Amounts', legend: { visible: true, position: 'bottom' } } });

    const output = loadOpcPackageGraph(exportSnapshotToXlsxBuffer(workbook.snapshot()));
    const chartParts = Object.keys(output.files).filter((name) => name.startsWith('xl/charts/react-pivot-chart-'));
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
      const typedOutput = loadOpcPackageGraph(exportSnapshotToXlsxBuffer(typed));
      const typedChartPart = Object.keys(typedOutput.files).find((name) => name.startsWith('xl/charts/react-pivot-chart-'));
      assert.ok(typedChartPart);
      assert.match(strFromU8(typedOutput.files[typedChartPart!]!), new RegExp(xmlName.replace(/[<>]/g, '')));
    }

    const withoutChart = structuredClone(workbook.snapshot());
    withoutChart.sheets[0]!.drawings = [];
    withoutChart.sheets[0]!.drawingPayloads = {};
    const removed = loadOpcPackageGraph(exportSnapshotToXlsxBuffer(withoutChart, output.packageGraph));
    assert.equal(Object.keys(removed.files).some((name) => name.startsWith('xl/charts/react-pivot-chart-')), false);
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
      layout: { rows: [], columns: [], filters: [], allowMultipleFiltersPerField: true, collation: { locale: 'en-US', sensitivity: 'variant', numeric: false, caseFirst: 'false' }, values: [{ fieldId: 'value', summarizeBy: 'sum' }], subtotalLocation: 'bottom', showGrandTotals: true, compact: true, repeatLabels: false },
      refreshPolicy: { mode: 'manual', preserveFormatting: true, refreshOnLoad: false },
    });
    sheet.drawings.push({ id: 'invalid-pivot-chart', sheetId: sheet.id, kind: 'chart', anchor: { kind: 'one-cell', row: 0, column: 4 }, transform: { x: 0, y: 0, width: 200, height: 120 }, zIndex: 1, payloadId: 'invalid-pivot-chart' });
    sheet.drawingPayloads.set('invalid-pivot-chart', { kind: 'chart', chartId: 'invalid-pivot-chart', chartType: 'pie', pivotId: 'pivot-invalid-chart', sourceRanges: [], elements: { hiddenData: 'show' } });
    assert.throws(() => exportSnapshotToXlsxBuffer(workbook.snapshot()), /unsupported native chart type/);
  });

  it('reads and rewrites native Pivot cache/table relationship graphs', async () => {
    const workbook = new WorkbookModel('wb-native-pivot', 'Native Pivot');
    const generated = loadOpcPackageGraph(exportSnapshotToXlsxBuffer(workbook.snapshot()));
    const parts = generated.packageGraph.parts;
    parts['xl/workbook.xml'] = strToU8(strFromU8(parts['xl/workbook.xml']!).replace('</workbook>', '<pivotCaches count="1"><pivotCache cacheId="1" r:id="rIdPivotCache"/></pivotCaches></workbook>'));
    parts['xl/_rels/workbook.xml.rels'] = strToU8(strFromU8(parts['xl/_rels/workbook.xml.rels']!).replace('</Relationships>', '<Relationship Id="rIdPivotCache" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/></Relationships>'));
    parts['xl/pivotCache/pivotCacheDefinition1.xml'] = strToU8('<?xml version="1.0"?><pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><cacheSource type="worksheet"><worksheetSource ref="A1:B2" sheet="Sheet1"/></cacheSource><cacheFields count="2"><cacheField name="Category"><sharedItems containsString="1"><s v="A"/></sharedItems></cacheField><cacheField name="Amount"><sharedItems containsNumber="1"><n v="10"/></sharedItems></cacheField></cacheFields></pivotCacheDefinition>');
    parts['xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels'] = strToU8('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdRecords" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords1.xml"/></Relationships>');
    parts['xl/pivotCache/pivotCacheRecords1.xml'] = strToU8('<?xml version="1.0"?><pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1"><r><s v="0"/><n v="10"/></r></pivotCacheRecords>');
    parts['xl/worksheets/sheet1.xml'] = strToU8(strFromU8(parts['xl/worksheets/sheet1.xml']!).replace('</worksheet>', '<pivotTableParts count="1"><pivotTablePart r:id="rIdPivotTable"/></pivotTableParts></worksheet>'));
    parts['xl/worksheets/_rels/sheet1.xml.rels'] = strToU8(strFromU8(parts['xl/worksheets/_rels/sheet1.xml.rels']!).replace('</Relationships>', '<Relationship Id="rIdPivotTable" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/></Relationships>'));
    parts['xl/pivotTables/pivotTable1.xml'] = strToU8('<?xml version="1.0"?><pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PivotTable1" cacheId="1" subtotalTop="1"><location ref="D1:E3"/><pivotFields count="2"><pivotField axis="axisRow" defaultSubtotal="0" sortType="descending"><autoSortScope><pivotArea dataOnly="0" fieldPosition="0"><references count="1"><reference field="4294967294" count="1" selected="0"><x v="0"/></reference></references></pivotArea></autoSortScope></pivotField><pivotField/></pivotFields><rowFields count="1"><field x="0"/></rowFields><dataFields count="1"><dataField fld="1" name="Sum of Amount" subtotal="sum"/></dataFields><pivotFilters count="4"><filter fld="0" type="captionEqual" stringValue1="A"/><filter fld="0" type="valueGreaterThan" iMeasureFld="0" val="10"/><filter fld="0" type="valueTop10" iMeasureFld="0" val="3" top="1"/><filter fld="1" type="futureFilter" id="7" stringValue1="preserve"/></pivotFilters></pivotTableDefinition>');
    const imported = await importXlsx({ fileName: 'native-pivot.xlsx', buffer: zipXlsxPartsBuffer(parts), options: { compatibilityTarget: 'B' } });
    assert.equal(imported.nativePackage.packageGraph.nativePivotGraph?.caches[0]?.source.kind, 'worksheet-range');
    assert.equal(imported.nativePackage.packageGraph.nativePivotGraph?.caches[0]?.fields[1]?.name, 'Amount');
    assert.equal(imported.nativePackage.packageGraph.nativePivotGraph?.tables[0]?.dataFields[0]?.field, 1);
    assert.equal(imported.snapshot.sheets[0]?.pivots.length, 1);
    assert.equal(imported.snapshot.sheets[0]?.pivots[0]?.source.kind, 'worksheet-range');
    assert.equal(imported.snapshot.sheets[0]?.pivots[0]?.layout.rows[0]?.fieldId, 'native:cache:1:field:0');
    assert.equal(imported.snapshot.sheets[0]?.pivots[0]?.refreshPolicy.mode, 'manual');
    assert.equal(imported.snapshot.sheets[0]?.pivots[0]?.refreshPolicy.refreshOnLoad, false);
    assert.deepEqual(imported.snapshot.sheets[0]?.pivots[0]?.nativeMetadata?.cacheFlags, {});
    assert.deepEqual(imported.snapshot.sheets[0]?.pivots[0]?.layout.rows[0]?.subtotal, { mode: 'none' });
    assert.deepEqual(imported.snapshot.sheets[0]?.pivots[0]?.layout.rows[0]?.sort, { direction: 'descending', by: 'label' });
    assert.deepEqual(imported.snapshot.sheets[0]?.pivots[0]?.layout.filters, [
      { kind: 'condition', family: 'label', fieldId: 'native:cache:1:field:0', operator: 'equals', value: 'A', scope: 'field' },
      { kind: 'condition', family: 'value', fieldId: 'native:cache:1:field:0', valueFieldId: 'native:cache:1:field:1', operator: 'greater-than', value: 10, scope: 'field' },
      { kind: 'top-items', family: 'top-items', fieldId: 'native:cache:1:field:0', valueFieldId: 'native:cache:1:field:1', count: 3, direction: 'top', scope: 'field' },
    ]);
    assert.deepEqual(imported.snapshot.sheets[0]?.pivots[0]?.nativeMetadata?.preservedPivotFilters, [{ fieldIndex: 1, type: 'futureFilter', attributes: { fld: '1', type: 'futureFilter', id: '7', stringValue1: 'preserve' } }]);
    assert.equal(imported.snapshot.sheets[0]?.pivots[0]?.layout.subtotalLocation, 'top');
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
    assert.match(nativeAdvancedXml, /sortType="descending"/);
    assert.match(nativeAdvancedXml, /<autoSortScope><pivotArea[^>]*dataOnly="0"/);
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
  });

  it('rejects conflicting refresh modes before rebuilding a shared native cache', () => {
    const workbook = new WorkbookModel('wb-shared-cache-conflict', 'Shared Cache Conflict');
    const sheet = workbook.getSheet(workbook.primarySheetId);
    sheet.cells.set(0, 0, { value: 'Category' });
    sheet.cells.set(0, 1, { value: 'Amount' });
    sheet.cells.set(1, 0, { value: 'A' });
    sheet.cells.set(1, 1, { value: 10 });
    const layout = { rows: [{ fieldId: 'category' }], columns: [], filters: [], allowMultipleFiltersPerField: true, collation: { locale: 'en-US' as const, sensitivity: 'variant' as const, numeric: false, caseFirst: 'false' as const }, values: [{ fieldId: 'amount', summarizeBy: 'sum' as const }], subtotalLocation: 'bottom' as const, showGrandTotals: true, compact: true, repeatLabels: false };
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
        rows: [{ fieldId: 'category', sort: { direction: 'ascending', by: 'value', valueFieldId: 'amount' } }], columns: [],
        filters: [{ kind: 'condition', family: 'value', fieldId: 'category', valueFieldId: 'amount', operator: 'greater-than', value: 5 }],
        allowMultipleFiltersPerField: true,
        collation: { locale: 'en-US', sensitivity: 'variant', numeric: false, caseFirst: 'false' },
        values: [{ fieldId: 'amount', summarizeBy: 'sum' }], subtotalLocation: 'bottom', showGrandTotals: true, compact: true, repeatLabels: false,
      },
      refreshPolicy: { mode: 'on-change', preserveFormatting: true, refreshOnLoad: true },
    });
    const output = loadOpcPackageGraph(exportSnapshotToXlsxBuffer(workbook.snapshot()));
    const xml = strFromU8(output.files['xl/pivotTables/pivotTable1.xml']!);
    assert.match(xml, /sortType="ascending"/);
    assert.match(xml, /<autoSortScope><pivotArea[^>]*dataOnly="1"[^>]*fieldPosition="0"/);
    assert.match(xml, /<reference field="1"/);
    assert.match(xml, /type="valueGreaterThan"[^>]*iMeasureFld="0"[^>]*val="5"/);
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

  it('accepts supported dynamic AutoFilters and rejects unknown OOXML types', () => {
    const workbook = new WorkbookModel('wb-dynamic-filter', 'Dynamic Filter');
    const generated = loadOpcPackageGraph(exportSnapshotToXlsxBuffer(workbook.snapshot()));
    const worksheet = strFromU8(generated.packageGraph.parts['xl/worksheets/sheet1.xml']!);
    const withType = (type: string): ArrayBuffer => zipXlsxPartsBuffer({
      ...generated.packageGraph.parts,
      'xl/worksheets/sheet1.xml': strToU8(worksheet.replace('</worksheet>', `<autoFilter ref="A1:A2"><filterColumn colId="0"><dynamicFilter type="${type}"/></filterColumn></autoFilter></worksheet>`)),
    } as Record<string, Uint8Array>);

    const supported = parseLoadedXlsx(loadOpcPackageGraph(withType('today'))).snapshot;
    assert.deepEqual(supported.sheets[0]?.autoFilter?.columns[0]?.criterion, { kind: 'dynamic', type: 'today' });
    assert.throws(
      () => parseLoadedXlsx(loadOpcPackageGraph(withType('attackerUnknown'))),
      /UNSUPPORTED_FEATURE: dynamic AutoFilter type "attackerUnknown" is not supported/,
    );
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

  it('imports the repository House cleaning sample with Excel widths converted to CSS pixels', async () => {
    const bytes = await readFile(new URL('../../../../luckyexcel-node/House cleaning checklist.xlsx', import.meta.url));
    const imported = await importXlsx({ fileName: 'House cleaning checklist.xlsx', buffer: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength), options: { compatibilityTarget: 'B', compatibilityMode: 'balanced' } });
    const widths = imported.snapshot.sheets.flatMap((sheet) => Object.values(sheet.columnWidthsPx ?? {}));
    assert.ok(widths.some((width) => width >= 67 && width <= 68), '9.625 Excel characters should render near 68 CSS pixels');
    assert.equal(widths.some((width) => Math.abs(width - 9.625) < 0.001), false);
  });
});
