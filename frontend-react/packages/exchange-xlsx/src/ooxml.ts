import type {
  CellData,
  CellStyle,
  CellHyperlink,
  DefinedNameModel,
  DrawingObject,
  FreezeModel,
  MergeSpan,
  WorkbookSnapshot,
  SheetSnapshot,
  RangeRef,
} from '@react-sheets/core-model';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import {
  child,
  children,
  descendants,
  encodeXml,
  localName,
  parseXml,
  serializeXml,
  textContent,
  type XmlNode,
} from './xml';
import {
  DEFAULT_XLSX_ZIP_LIMITS,
  type DateSystem,
  type XlsxPackage,
  type XlsxRelationship,
  type XlsxZipLimits,
} from './types';
import { mapNativePivotDefinition, readNativePivotGraph, serializeNativePivotCaches, synchronizeNativePivotPackage } from './native-pivot';
import type { NativePivotControlDefinition, NativePivotGraph } from './types';

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const NS_DOC_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL_OFFICE_DOCUMENT = `${NS_DOC_REL}/officeDocument`;
const REL_WORKSHEET = `${NS_DOC_REL}/worksheet`;
const REL_STYLES = `${NS_DOC_REL}/styles`;
const REL_SHARED_STRINGS = `${NS_DOC_REL}/sharedStrings`;
const REL_HYPERLINK = `${NS_DOC_REL}/hyperlink`;
const REL_DRAWING = `${NS_DOC_REL}/drawing`;

export interface LoadedXlsxPackage {
  package: XlsxPackage;
  files: Record<string, Uint8Array>;
}

export interface ParsedXlsxPackage {
  package: XlsxPackage;
  snapshot: WorkbookSnapshot;
  features: string[];
}

interface StyleRecord {
  numberFormat?: string;
  style?: CellStyle;
}

interface SheetDescriptor {
  id: string;
  name: string;
  part: string;
  hidden: boolean;
}

export function loadXlsxPackage(input: ArrayBuffer | Uint8Array, limits: Partial<XlsxZipLimits> = {}): LoadedXlsxPackage {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
  const effective = { ...DEFAULT_XLSX_ZIP_LIMITS, ...limits };
  if (bytes.byteLength > effective.maxArchiveBytes) {
    throw new Error(`XLSX archive exceeds ${effective.maxArchiveBytes} byte limit`);
  }

  const names = new Set<string>();
  let entryCount = 0;
  let totalUncompressed = 0;
  const files = unzipSync(bytes, {
    filter(file) {
      entryCount += 1;
      const name = normalizePartName(file.name);
      if (names.has(name)) throw new Error(`XLSX archive contains duplicate part: ${name}`);
      names.add(name);
      if (entryCount > effective.maxEntries) throw new Error('XLSX archive has too many entries');
      if (file.originalSize > effective.maxEntryBytes) throw new Error(`XLSX part exceeds size limit: ${name}`);
      totalUncompressed += file.originalSize;
      if (totalUncompressed > effective.maxUncompressedBytes) throw new Error('XLSX archive exceeds total uncompressed size limit');
      if (file.originalSize > effective.maxCompressionRatio * Math.max(file.size, 1)) {
        throw new Error(`XLSX part has an unsafe compression ratio: ${name}`);
      }
      if (name.endsWith('/')) throw new Error(`XLSX archive contains a directory part: ${name}`);
      return true;
    },
  }) as Record<string, Uint8Array>;

  const normalizedFiles: Record<string, Uint8Array> = {};
  for (const [name, data] of Object.entries(files)) normalizedFiles[normalizePartName(name)] = data;
  if (!normalizedFiles['xl/workbook.xml']) throw new Error('Not a valid XLSX package: xl/workbook.xml is missing');

  const relationships = readRelationships(normalizedFiles);
  const dateSystem = parseWorkbookDateSystem(strFromU8(normalizedFiles['xl/workbook.xml']));
  const sheetPartById = readSheetPartMap(normalizedFiles, relationships);
  const coreParts = new Set<string>([
    '[Content_Types].xml', '_rels/.rels', 'xl/workbook.xml', 'xl/_rels/workbook.xml.rels',
    'xl/styles.xml', 'xl/sharedStrings.xml',
    ...Object.values(sheetPartById),
  ]);
  for (const source of Object.keys(relationships)) {
    if (source) coreParts.add(relationshipPartName(source));
  }
  const opaqueParts: Record<string, Uint8Array> = {};
  for (const [name, data] of Object.entries(normalizedFiles)) {
    if (!coreParts.has(name)) opaqueParts[name] = data.slice();
  }

  const nativePivotGraph = hasNativePivotMarkers(normalizedFiles)
    ? readNativePivotGraph({ files: normalizedFiles, relationships, sheetPartById })
    : undefined;
  return {
    files: normalizedFiles,
    package: {
      schema: 'XlsxPackage',
      parts: cloneParts(normalizedFiles),
      opaqueParts,
      relationships,
      sheetPartById,
      contentTypesXml: normalizedFiles['[Content_Types].xml']?.slice(),
      dateSystem,
      ...(nativePivotGraph ? { nativePivotGraph } : {}),
    },
  };
}

export function parseLoadedXlsx(loaded: LoadedXlsxPackage): ParsedXlsxPackage {
  const files = loaded.files;
  const workbookXml = parseXml(strFromU8(files['xl/workbook.xml']!));
  const workbook = firstElement(workbookXml, 'workbook');
  const relationships = loaded.package.relationships;
  const descriptors: SheetDescriptor[] = [];
  const workbookRels = relationships['xl/workbook.xml'] ?? [];
  const sheetNodes = children(child(workbook, 'sheets'), 'sheet');
  for (let index = 0; index < sheetNodes.length; index += 1) {
    const node = sheetNodes[index]!;
    const relationId = node.attrs['r:id'] ?? node.attrs.id ?? '';
    const relation = workbookRels.find((candidate) => candidate.id === relationId && candidate.type === REL_WORKSHEET);
    const part = relation ? resolveTarget('xl/workbook.xml', relation.target) : `xl/worksheets/sheet${index + 1}.xml`;
    if (!files[part]) throw new Error(`Workbook sheet relation points to missing part: ${part}`);
    const sheetId = `sheet-${node.attrs.sheetId ?? index + 1}`;
    descriptors.push({
      id: sheetId,
      name: node.attrs.name ?? `Sheet${index + 1}`,
      part,
      hidden: node.attrs.state === 'hidden' || node.attrs.state === 'veryHidden',
    });
    loaded.package.sheetPartById[sheetId] = part;
  }
  if (descriptors.length === 0) throw new Error('XLSX workbook has no worksheets');

  const sharedStrings = parseSharedStrings(files['xl/sharedStrings.xml']);
  const styles = parseStyles(files['xl/styles.xml']);
  const sheets = descriptors.map((descriptor) => parseSheet(descriptor, files, loaded.package, sharedStrings, styles));
  const definedNameModels = parseDefinedNames(child(workbook, 'definedNames'), descriptors);
  const definedNames: Record<string, string> = {};
  for (const name of definedNameModels) if (name.scope === 'workbook') definedNames[name.name] = name.formula;
  const snapshot: WorkbookSnapshot = {
    schema: 'WorkbookSnapshot',
    version: 2,
    unitId: `imported-${randomId()}`,
    name: workbook.attrs.name ?? 'Imported Workbook',
    definedNames,
    definedNameModels,
    dataSources: [],
    sheets,
  };
  attachNativePivots(snapshot, loaded.package.nativePivotGraph, loaded.package.sheetPartById);
  return { package: loaded.package, snapshot, features: detectPackageFeatures(loaded.package, snapshot) };
}

export function exportSnapshotToXlsxPackage(
  snapshot: WorkbookSnapshot,
  options: { dateSystem: DateSystem; includeCachedValues?: boolean; preserveMacros?: boolean },
  preserved?: XlsxPackage,
): ArrayBuffer {
  const files = new Map<string, Uint8Array>();
  if (preserved) {
    for (const [name, data] of Object.entries(preserved.parts)) {
      if (options.preserveMacros === false && isMacroPart(name)) continue;
      files.set(normalizePartName(name), data.slice());
    }
  }
  const sourceFiles = preserved?.parts ?? {};
  const sheetParts = snapshot.sheets.map((sheet, index) => preserved?.sheetPartById[sheet.id] ?? `xl/worksheets/sheet${index + 1}.xml`);
  const sheetPartById = Object.fromEntries(snapshot.sheets.map((sheet, index) => [sheet.id, sheetParts[index]!])) as Record<string, string>;
  const nativeUpdate = synchronizeNativePivotPackage({
    files: Object.fromEntries([...files.entries()]),
    relationships: preserved?.relationships ?? {},
    graph: preserved?.nativePivotGraph,
    snapshot,
    sheetPartById,
  });
  for (const [name, data] of Object.entries(nativeUpdate.files)) files.set(name, data);
  const styleOutput = buildStyles(snapshot);
  const styleIndexes = collectStyleIndexes(snapshot);
  const sharedOutput = buildSharedStrings(snapshot);
  const sheetRelationships: Record<string, XlsxRelationship[]> = {};
  for (let index = 0; index < snapshot.sheets.length; index += 1) {
    const sheet = snapshot.sheets[index]!;
    const part = sheetParts[index]!;
    const original = preserved ? strFromU8(sourceFiles[part] ?? new Uint8Array()) : '';
    const originalRoot = original ? firstElement(parseXml(original), 'worksheet') : undefined;
    const requiredHyperlinks = collectHyperlinkRelationships(sheet, preserved?.relationships[part] ?? []);
    const tableParts = prepareTableParts(sheet, part, preserved, files);
    for (const [tablePart, tableXml] of tableParts.parts) files.set(tablePart, strToU8(tableXml));
    const relationships = mergeRelationships(nativeUpdate.relationships[part] ?? preserved?.relationships[part] ?? [], [...requiredHyperlinks, ...tableParts.required]);
    sheetRelationships[part] = relationships;
    files.set(part, strToU8(buildWorksheetXml(sheet, part, relationships, originalRoot, files, styleIndexes, options.includeCachedValues ?? true, nativeUpdate.displayCellsBySheetPart[part], nativeUpdate.graph.controls ?? [])));
  }

  const workbookRelations = mergeRelationships(
    nativeUpdate.relationships['xl/workbook.xml'] ?? preserved?.relationships['xl/workbook.xml'] ?? [],
    [
      { id: '', type: REL_STYLES, target: 'styles.xml' },
      { id: '', type: REL_SHARED_STRINGS, target: 'sharedStrings.xml' },
      ...sheetParts.map((part) => ({ id: '', type: REL_WORKSHEET, target: relativeTarget('xl/workbook.xml', part) })),
    ],
  );
  files.set('xl/workbook.xml', strToU8(buildWorkbookXml(snapshot, workbookRelations, descriptorsForSnapshot(snapshot), options.dateSystem, nativeUpdate.graph, preserved)));
  files.set('xl/_rels/workbook.xml.rels', strToU8(buildRelationshipsXml(workbookRelations)));
  files.set('_rels/.rels', strToU8(buildRootRelationshipsXml(preserved?.relationships[''] ?? [])));
  files.set('xl/styles.xml', strToU8(styleOutput));
  files.set('xl/sharedStrings.xml', strToU8(sharedOutput));
  for (const [source, relationships] of Object.entries(sheetRelationships)) {
    files.set(relationshipPartName(source), strToU8(buildRelationshipsXml(relationships)));
  }
  for (const [source, relationships] of Object.entries(nativeUpdate.relationships)) {
    if (!source || source === 'xl/workbook.xml' || sheetRelationships[source]) continue;
    files.set(relationshipPartName(source), strToU8(buildRelationshipsXml(relationships)));
  }
  files.set('[Content_Types].xml', strToU8(buildContentTypesXml(files, preserved)));

  const zipped: Record<string, Uint8Array> = {};
  for (const [name, data] of files) zipped[name] = data;
  const zippedBytes = zipSync(zipped, { level: 6 });
  return zippedBytes.buffer.slice(zippedBytes.byteOffset, zippedBytes.byteOffset + zippedBytes.byteLength) as ArrayBuffer;
}

export function detectPackageFeatures(pkg: XlsxPackage, snapshot?: WorkbookSnapshot): string[] {
  const features = new Set<string>(snapshot ? ['cells', 'styles'] : []);
  for (const name of Object.keys(pkg.parts)) {
    const lower = name.toLowerCase();
    if (lower.includes('/charts/')) features.add('charts');
    if (lower.includes('/pivot') || lower.includes('pivottableparts')) features.add('pivot');
    if (lower.includes('vba') || lower.endsWith('.bin')) features.add('vba');
    if (lower.includes('externalconnections') || lower.includes('connections.xml')) features.add('external-connection');
    if (lower.includes('/slicers/') || lower.includes('slicer')) features.add('slicer');
    if (lower.includes('/timelines/') || lower.includes('timeline')) features.add('timeline');
    if (lower.includes('/theme/')) features.add('theme');
    if (lower.includes('/comments')) features.add('comments');
    if (lower.includes('/drawings/')) features.add('images');
    if (lower.includes('/tables/')) features.add('tables');
  }
  if (snapshot) {
    for (const sheet of snapshot.sheets) {
      if (sheet.merges.length) features.add('merges');
      if (sheet.freeze.xSplit || sheet.freeze.ySplit) features.add('freeze');
      if (Object.values(sheet.cells).some((row) => Object.values(row).some((cell) => Boolean(cell.formula)))) features.add('formulas');
      if (sheet.conditionalFormats?.length) features.add('conditional-format');
      if (sheet.dataValidations?.length) features.add('validation');
      if (sheet.sheetTables?.length) features.add('tables');
      if (sheet.filter) features.add('filters');
      if (sheet.notes?.length || sheet.commentThreads?.length) features.add('comments');
      for (const payload of Object.values(sheet.drawingPayloads)) {
        if (payload.kind === 'chart') features.add('charts');
        else if (payload.kind === 'slicer') features.add('slicer');
        else if (payload.kind === 'timeline') features.add('timeline');
        else if (payload.kind === 'image' || payload.kind === 'shape' || payload.kind === 'textbox') features.add('images');
      }
    }
    if (snapshot.definedNameModels?.length || Object.keys(snapshot.definedNames ?? {}).length) features.add('defined-names');
  }
  return [...features];
}

function parseSheet(
  descriptor: SheetDescriptor,
  files: Record<string, Uint8Array>,
  pkg: XlsxPackage,
  sharedStrings: string[],
  styles: StyleRecord[],
): SheetSnapshot {
  const xml = strFromU8(files[descriptor.part]!);
  const root = firstElement(parseXml(xml), 'worksheet');
  const cells: Record<string, Record<string, CellData>> = {};
  const hiddenRows: number[] = [];
  const dimensions = parseRange(child(root, 'dimension')?.attrs.ref);
  const sheetData = child(root, 'sheetData');
  let maxRow = dimensions?.endRow ?? 999;
  let maxColumn = dimensions?.endColumn ?? 25;
  for (const rowNode of children(sheetData, 'row')) {
    const rowNumber = parsePositiveInt(rowNode.attrs.r, 1) - 1;
    maxRow = Math.max(maxRow, rowNumber);
    if (rowNode.attrs.hidden === '1' || rowNode.attrs.hidden === 'true') hiddenRows.push(rowNumber);
    for (const cellNode of children(rowNode, 'c')) {
      const address = parseA1(cellNode.attrs.r ?? 'A1');
      if (!address) continue;
      maxColumn = Math.max(maxColumn, address.column);
      const styleId = cellNode.attrs.s;
      const style = styleId === undefined ? undefined : styles[Number(styleId)];
      const formula = child(cellNode, 'f');
      const value = readCellValue(cellNode, sharedStrings);
      const cell: CellData = {
        value,
        ...(formula ? { formula: `=${textContent(formula)}` } : {}),
        ...(styleId === undefined ? {} : { styleId }),
        ...(style?.style ? { style: structuredClone(style.style) } : {}),
        ...(style?.numberFormat ? { numberFormat: style.numberFormat } : {}),
      };
      const hyperlink = hyperlinkForCell(root, pkg.relationships[descriptor.part] ?? [], address.row, address.column);
      if (hyperlink) cell.hyperlinkDetail = hyperlink;
      cells[String(address.row)] ??= {};
      cells[String(address.row)]![String(address.column)] = cell;
    }
  }
  const freeze = parseFreeze(root);
  const merges = children(child(root, 'mergeCells'), 'mergeCell')
    .map((node) => parseRange(node.attrs.ref))
    .filter((range): range is RangeRef => Boolean(range))
    .map((range) => ({ range: { ...range, sheetId: descriptor.id }, anchor: { row: range.startRow, column: range.startColumn } } satisfies MergeSpan));
  const hiddenColumns = parseHiddenColumns(root);
  const rowHeights = parseRowHeights(root);
  const columnWidths = parseColumnWidths(root);
  const tabColor = child(child(root, 'sheetPr'), 'tabColor')?.attrs.rgb;
  const notes = parseNotes(root, descriptor, files, pkg);
  const sheetTables = parseSheetTables(root, descriptor, files, pkg);
  const rowCount = Math.max(1000, maxRow + 1);
  const columnCount = Math.max(26, maxColumn + 1);
  return {
    id: descriptor.id,
    name: descriptor.name,
    rowCount,
    columnCount,
    cells,
    merges,
    freeze,
    pivots: [],
    sparklines: [],
    drawings: [],
    drawingPayloads: {},
    conditionalFormats: [],
    dataValidations: [],
    rowHeights,
    columnWidths,
    hiddenRows,
    hiddenColumns,
    tabColor,
    notes,
    ...(sheetTables.length ? { sheetTables } : {}),
    hidden: descriptor.hidden,
  };
}

function attachNativePivots(snapshot: WorkbookSnapshot, graph: NativePivotGraph | undefined, sheetPartById: Record<string, string>): void {
  if (!graph) return;
  const caches = new Map(graph.caches.map((cache) => [cache.cacheId, cache]));
  for (const table of graph.tables) {
    const cache = caches.get(table.cacheId);
    if (!cache) continue;
    const definition = mapNativePivotDefinition(table, cache, snapshot, sheetPartById);
    if (!definition) continue;
    const sheet = snapshot.sheets.find((candidate) => candidate.id === definition.target.sheetId);
    if (!sheet) continue;
    sheet.pivots.push(definition);
    const location = table.locationRef ? parseRange(table.locationRef) : undefined;
    if (!location) continue;
    for (const row of Object.keys(sheet.cells)) {
      const rowIndex = Number(row);
      if (rowIndex < location.startRow || rowIndex > location.endRow) continue;
      for (const column of Object.keys(sheet.cells[row] ?? {})) {
        const columnIndex = Number(column);
        if (columnIndex >= location.startColumn && columnIndex <= location.endColumn) {
          const rowCells = sheet.cells[row];
          if (rowCells) delete rowCells[column];
        }
      }
      if (!Object.keys(sheet.cells[row] ?? {}).length) delete sheet.cells[row];
    }
  }
  attachNativePivotControls(snapshot, graph.controls ?? [], sheetPartById);
}

function attachNativePivotControls(snapshot: WorkbookSnapshot, controls: NativePivotControlDefinition[], sheetPartById: Record<string, string>): void {
  const style = { theme: 'light' as const, fill: '#ffffff', border: '#d1d5db', textColor: '#111827', accentColor: '#2563eb' };
  for (const control of controls) {
    if (!control.valid || !control.pivotId || !control.fieldId) continue;
    const sheet = snapshot.sheets.find((candidate) => sheetPartById[candidate.id] === control.sheetPart);
    const pivot = snapshot.sheets.flatMap((candidate) => candidate.pivots).find((candidate) => candidate.id === control.pivotId);
    if (!sheet || !pivot || !pivot.fieldCatalog.fields.some((field) => field.fieldId === control.fieldId)) continue;
    const existing = sheet.drawings.find((drawing) => drawing.id === control.id);
    if (existing) continue;
    const payloadId = control.id;
    if (control.kind === 'slicer') {
      const field = pivot.fieldCatalog.fields.find((candidate) => candidate.fieldId === control.fieldId);
      const memberKeys = (control.selectedItemIndexes ?? []).flatMap((index) => field?.values?.[index] === undefined ? [] : [nativeMemberKey(field.values[index] ?? null)]);
      sheet.drawingPayloads[payloadId] = {
        kind: 'slicer',
        pivotId: control.pivotId,
        fieldId: control.fieldId,
        filter: { mode: memberKeys.length ? 'include' : 'all', memberKeys },
        style,
        ...(control.connectedPivotIds?.length ? { connectedPivotIds: control.connectedPivotIds } : {}),
      };
    } else {
      sheet.drawingPayloads[payloadId] = {
        kind: 'timeline',
        pivotId: control.pivotId,
        fieldId: control.fieldId,
        period: control.selection ?? {},
        style,
        ...(control.connectedPivotIds?.length ? { connectedPivotIds: control.connectedPivotIds } : {}),
      };
    }
    const drawing: DrawingObject = {
      id: control.id,
      sheetId: sheet.id,
      kind: control.kind,
      anchor: { kind: 'one-cell', row: 0, column: 0 },
      transform: { x: 0, y: 0, width: control.kind === 'slicer' ? 220 : 420, height: control.kind === 'slicer' ? 180 : 120 },
      zIndex: 0,
      payloadId,
    };
    sheet.drawings.push(drawing);
  }
}

function nativeMemberKey(value: string | number | boolean | null): { type: 'text' | 'number' | 'boolean' | 'blank'; value: string | number | boolean | null } {
  if (value === null || value === '') return { type: 'blank', value: null };
  if (typeof value === 'number') return { type: 'number', value };
  if (typeof value === 'boolean') return { type: 'boolean', value };
  return { type: 'text', value };
}

function parseSheetTables(
  root: XmlNode,
  descriptor: SheetDescriptor,
  files: Record<string, Uint8Array>,
  pkg: XlsxPackage,
): NonNullable<SheetSnapshot['sheetTables']> {
  return children(child(root, 'tableParts'), 'tablePart').flatMap((partNode) => {
    const relationId = partNode.attrs['r:id'] ?? partNode.attrs.id;
    if (!relationId) throw new Error(`Worksheet ${descriptor.part} tablePart is missing r:id`);
    const relation = (pkg.relationships[descriptor.part] ?? []).find((candidate) => candidate.id === relationId && candidate.type.endsWith('/table'));
    if (!relation) throw new Error(`Worksheet ${descriptor.part} table relation ${relationId} is missing`);
    const part = resolveTarget(descriptor.part, relation.target);
    const bytes = files[part];
    if (!bytes) throw new Error(`Worksheet table relation points to missing part: ${part}`);
    const table = firstElement(parseXml(strFromU8(bytes)), 'table');
    const range = parseRange(table.attrs.ref);
    if (!range) throw new Error(`Table ${part} has an invalid ref`);
    const columns = children(child(table, 'tableColumns'), 'tableColumn').map((column, index) => ({
      id: `column-${table.attrs.id ?? descriptor.id}-${index}`,
      name: column.attrs.name ?? `Column${index + 1}`,
      ...(column.attrs.totalsRowFunction ? { totalsFunction: column.attrs.totalsRowFunction as NonNullable<SheetSnapshot['sheetTables']>[number]['columns'][number]['totalsFunction'] } : {}),
    }));
    const tableNumber = table.attrs.id ?? (part.replace(/[^0-9]/g, '') || descriptor.id);
    return [{
      id: `table-${tableNumber}`,
      sheetId: descriptor.id,
      name: table.attrs.displayName ?? table.attrs.name ?? `Table${table.attrs.id ?? '1'}`,
      range: { ...range, sheetId: descriptor.id },
      hasHeaderRow: table.attrs.headerRowCount !== '0',
      hasTotalRow: table.attrs.totalsRowCount === '1',
      showBandedRows: child(table, 'tableStyleInfo')?.attrs.showRowStripes !== '0',
      showBandedColumns: child(table, 'tableStyleInfo')?.attrs.showColumnStripes === '1',
      showFilterButton: table.attrs.headerRowCount !== '0',
      columns,
      ...(child(table, 'tableStyleInfo')?.attrs.name ? { styleName: child(table, 'tableStyleInfo')!.attrs.name } : {}),
    }];
  });
}

function parseSharedStrings(bytes: Uint8Array | undefined): string[] {
  if (!bytes) return [];
  const root = firstElement(parseXml(strFromU8(bytes)), 'sst');
  return children(root, 'si').map((item) => descendants(item, 't').map(textContent).join(''));
}

function parseStyles(bytes: Uint8Array | undefined): StyleRecord[] {
  if (!bytes) return [{}];
  const root = firstElement(parseXml(strFromU8(bytes)), 'styleSheet');
  const customFormats = new Map<number, string>();
  for (const node of children(child(root, 'numFmts'), 'numFmt')) {
    const id = Number(node.attrs.numFmtId);
    const format = node.attrs.formatCode;
    if (Number.isFinite(id) && format) customFormats.set(id, format);
  }
  const xfs = children(child(root, 'cellXfs'), 'xf');
  return xfs.map((xf, index) => {
    const numberFormat = customFormats.get(Number(xf.attrs.numFmtId)) ?? builtInNumberFormat(Number(xf.attrs.numFmtId));
    const font = child(root, 'fonts')?.children[Number(xf.attrs.fontId) || 0];
    const fill = child(root, 'fills')?.children[Number(xf.attrs.fillId) || 0];
    const alignment = child(xf, 'alignment');
    const style: CellStyle = {
      ...(child(font, 'name')?.attrs.val ? { fontFamily: child(font, 'name')!.attrs.val } : {}),
      ...(child(font, 'sz')?.attrs.val ? { fontSize: Number(child(font, 'sz')!.attrs.val) } : {}),
      ...(child(font, 'b') ? { bold: true } : {}),
      ...(child(font, 'i') ? { italic: true } : {}),
      ...(child(font, 'u') ? { underline: true } : {}),
      ...(font && descendants(font, 'color')[0]?.attrs.rgb ? { textColor: descendants(font, 'color')[0]!.attrs.rgb } : {}),
      ...(descendants(fill, 'fgColor')[0]?.attrs.rgb ? { background: descendants(fill, 'fgColor')[0]!.attrs.rgb } : {}),
      ...(alignment?.attrs.horizontal ? { horizontalAlignment: normalizeHorizontal(alignment.attrs.horizontal) } : {}),
      ...(alignment?.attrs.vertical ? { verticalAlignment: normalizeVertical(alignment.attrs.vertical) } : {}),
      ...(alignment?.attrs.wrapText ? { wrapText: alignment.attrs.wrapText === '1' || alignment.attrs.wrapText === 'true' } : {}),
      ...(alignment?.attrs.textRotation ? { textRotate: Number(alignment.attrs.textRotation) } : {}),
    };
    return { numberFormat, style: Object.keys(style).length ? style : undefined, ...(index === 0 ? {} : {}) };
  });
}

function buildSharedStrings(snapshot: WorkbookSnapshot): string {
  const values: string[] = [];
  const lookup = new Map<string, number>();
  for (const sheet of snapshot.sheets) {
    for (const row of Object.values(sheet.cells)) {
      for (const cell of Object.values(row)) {
        if (cell.formula || typeof cell.value !== 'string') continue;
        if (!lookup.has(cell.value)) {
          lookup.set(cell.value, values.length);
          values.push(cell.value);
        }
      }
    }
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="${NS_MAIN}" count="${values.length}" uniqueCount="${values.length}">${values.map((value) => `<si><t>${encodeXml(value)}</t></si>`).join('')}</sst>`;
}

function buildStyles(snapshot: WorkbookSnapshot): string {
  const records: StyleRecord[] = [{}];
  const indexes = new Map<string, number>();
  for (const sheet of snapshot.sheets) {
    for (const row of Object.values(sheet.cells)) {
      for (const cell of Object.values(row)) {
        if (!cell.style && !cell.numberFormat) continue;
        const key = JSON.stringify({ style: cell.style, numberFormat: cell.numberFormat });
        if (!indexes.has(key)) {
          indexes.set(key, records.length);
          records.push({ style: cell.style, numberFormat: cell.numberFormat });
        }
      }
    }
  }
  const custom = new Map<string, number>();
  let nextFormat = 164;
  for (const record of records) {
    if (record.numberFormat && !builtInNumberFormatId(record.numberFormat)) {
      if (!custom.has(record.numberFormat)) custom.set(record.numberFormat, nextFormat++);
    }
  }
  const numFmts = [...custom.entries()].map(([code, id]) => `<numFmt numFmtId="${id}" formatCode="${encodeXml(code)}"/>`).join('');
  const fontIndexes = new Map<string, number>();
  const fillIndexes = new Map<string, number>();
  const fontRecords = ['<font><sz val="11"/><name val="Calibri"/></font>'];
  const fillRecords = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
  for (const record of records) {
    const style = record.style;
    const fontKey = JSON.stringify({ fontFamily: style?.fontFamily, fontSize: style?.fontSize, bold: style?.bold, italic: style?.italic, underline: style?.underline, textColor: style?.textColor });
    if (!fontIndexes.has(fontKey) && style) {
      const index = fontRecords.length;
      fontIndexes.set(fontKey, index);
      fontRecords.push(`<font><sz val="${style.fontSize ?? 11}"/><name val="${encodeXml(style.fontFamily ?? 'Calibri')}"/>${style.bold ? '<b/>' : ''}${style.italic ? '<i/>' : ''}${style.underline ? '<u/>' : ''}${style.textColor ? `<color rgb="${encodeXml(style.textColor)}"/>` : ''}</font>`);
    }
    const fillKey = style?.background ?? '';
    if (fillKey && !fillIndexes.has(fillKey)) {
      const index = fillRecords.length;
      fillIndexes.set(fillKey, index);
      fillRecords.push(`<fill><patternFill patternType="solid"><fgColor rgb="${encodeXml(fillKey)}"/><bgColor indexed="64"/></patternFill></fill>`);
    }
  }
  const borders = ['<border><left/><right/><top/><bottom/><diagonal/></border>'];
  const xfs = records.map((record) => {
    const style = record.style;
    const numFmtId = record.numberFormat ? (builtInNumberFormatId(record.numberFormat) ?? custom.get(record.numberFormat) ?? 0) : 0;
    const fontKey = JSON.stringify({ fontFamily: style?.fontFamily, fontSize: style?.fontSize, bold: style?.bold, italic: style?.italic, underline: style?.underline, textColor: style?.textColor });
    const fontId = style ? (fontIndexes.get(fontKey) ?? 0) : 0;
    const fillId = style?.background ? (fillIndexes.get(style.background) ?? 0) : 0;
    const attrs = [`numFmtId="${numFmtId}"`, `fontId="${fontId}"`, `fillId="${fillId}"`, 'borderId="0"', 'xfId="0"'];
    const alignment = style && (style.horizontalAlignment || style.verticalAlignment || style.wrapText || style.textRotate !== undefined)
      ? `<alignment${style.horizontalAlignment ? ` horizontal="${style.horizontalAlignment}"` : ''}${style.verticalAlignment ? ` vertical="${style.verticalAlignment}"` : ''}${style.wrapText ? ' wrapText="1"' : ''}${style.textRotate !== undefined ? ` textRotation="${style.textRotate}"` : ''}/>`
      : '';
    return `<xf ${attrs.join(' ')}${alignment ? ` applyAlignment="1">${alignment}</xf>` : '/>'}`;
  }).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="${NS_MAIN}"><numFmts count="${custom.size}">${numFmts}</numFmts><fonts count="${fontRecords.length}">${fontRecords.join('')}</fonts><fills count="${fillRecords.length}">${fillRecords.join('')}</fills><borders count="${borders.length}">${borders.join('')}</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${records.length}">${xfs}</cellXfs></styleSheet>`;
}

function prepareTableParts(
  sheet: SheetSnapshot,
  sourcePart: string,
  preserved: XlsxPackage | undefined,
  files: Map<string, Uint8Array>,
): { parts: Map<string, string>; required: Array<Pick<XlsxRelationship, 'type' | 'target'>> } {
  const tables = sheet.sheetTables ?? [];
  if (!tables.length) return { parts: new Map(), required: [] };
  const existing = (preserved?.relationships[sourcePart] ?? []).filter((relation) => relation.type.endsWith('/table'));
  const usedParts = new Set(files.keys());
  let nextNumber = 1;
  const parts = new Map<string, string>();
  const required: Array<Pick<XlsxRelationship, 'type' | 'target'>> = [];
  for (let index = 0; index < tables.length; index += 1) {
    const table = tables[index]!;
    const existingRelation = existing.find((relation) => {
      const target = resolveTarget(sourcePart, relation.target);
      const bytes = files.get(target);
      if (!bytes) return false;
      const root = firstElement(parseXml(strFromU8(bytes)), 'table');
      return root.attrs.displayName === table.name || root.attrs.name === table.name;
    });
    let part = existingRelation ? resolveTarget(sourcePart, existingRelation.target) : '';
    if (!part) {
      do {
        part = `xl/tables/table${nextNumber++}.xml`;
      } while (usedParts.has(part));
    }
    usedParts.add(part);
    parts.set(part, buildTableXml(table, index + 1));
    required.push({ type: `${NS_DOC_REL}/table`, target: relativeTarget(sourcePart, part) });
  }
  return { parts, required };
}

function buildTableXml(table: NonNullable<SheetSnapshot['sheetTables']>[number], fallbackId: number): string {
  const numericId = Number(table.id.replace(/\D/g, '')) || fallbackId;
  const ref = rangeToA1(table.range);
  const columns = table.columns.map((column, index) => `<tableColumn id="${index + 1}" name="${encodeXml(column.name)}"${column.totalsFunction && column.totalsFunction !== 'none' ? ` totalsRowFunction="${encodeXml(column.totalsFunction)}"` : ''}/>`).join('');
  const style = table.styleName
    ? `<tableStyleInfo name="${encodeXml(table.styleName)}" showFirstColumn="0" showLastColumn="0" showRowStripes="${table.showBandedRows ? '1' : '0'}" showColumnStripes="${table.showBandedColumns ? '1' : '0'}"/>`
    : '';
  const autoFilter = table.showFilterButton && table.hasHeaderRow ? `<autoFilter ref="${encodeXml(ref)}"/>` : '';
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><table xmlns="${NS_MAIN}" id="${numericId}" name="${encodeXml(table.name)}" displayName="${encodeXml(table.name)}" ref="${encodeXml(ref)}" headerRowCount="${table.hasHeaderRow ? '1' : '0'}" totalsRowCount="${table.hasTotalRow ? '1' : '0'}">${autoFilter}<tableColumns count="${table.columns.length}">${columns}</tableColumns>${style}</table>`;
}

function collectStyleIndexes(snapshot: WorkbookSnapshot): Map<string, number> {
  const indexes = new Map<string, number>();
  let next = 1;
  for (const sheet of snapshot.sheets) {
    for (const row of Object.values(sheet.cells)) {
      for (const cell of Object.values(row)) {
        if (!cell.style && !cell.numberFormat) continue;
        const key = JSON.stringify({ style: cell.style, numberFormat: cell.numberFormat });
        if (!indexes.has(key)) indexes.set(key, next++);
      }
    }
  }
  return indexes;
}

function buildWorksheetXml(
  sheet: SheetSnapshot,
  sourcePart: string,
  relationships: XlsxRelationship[],
  originalRoot: XmlNode | undefined,
  files: Map<string, Uint8Array>,
  styleIndexes: Map<string, number>,
  includeCachedValues: boolean,
  nativeDisplayCells?: Record<string, Record<string, CellData>>,
  nativeControls: NativePivotControlDefinition[] = [],
): string {
  let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_DOC_REL}">`;
  if (sheet.tabColor) xml += `<sheetPr><tabColor rgb="${encodeXml(sheet.tabColor)}"/></sheetPr>`;
  const dimension = inferDimension(sheet);
  if (dimension) xml += `<dimension ref="${dimension}"/>`;
  if (sheet.freeze.xSplit || sheet.freeze.ySplit) {
    const topLeft = `${columnToLetter(sheet.freeze.startColumn)}${sheet.freeze.startRow + 1}`;
    xml += `<sheetViews><sheetView workbookViewId="0"><pane xSplit="${sheet.freeze.xSplit}" ySplit="${sheet.freeze.ySplit}" topLeftCell="${topLeft}" activePane="bottomRight" state="frozen"/></sheetView></sheetViews>`;
  }
  if (Object.keys(sheet.columnWidths ?? {}).length || (sheet.hiddenColumns?.length ?? 0)) {
    xml += `<cols>${Object.entries(sheet.columnWidths ?? {}).map(([column, width]) => `<col min="${Number(column) + 1}" max="${Number(column) + 1}" width="${width}" customWidth="1"/>`).join('')}${(sheet.hiddenColumns ?? []).map((column) => `<col min="${column + 1}" max="${column + 1}" hidden="1"/>`).join('')}</cols>`;
  }
  xml += '<sheetData>';
  const rowKeys = [...new Set([...Object.keys(sheet.cells), ...Object.keys(nativeDisplayCells ?? {})])].map(Number).sort((a, b) => a - b);
  for (const row of rowKeys) {
    const cells = { ...(nativeDisplayCells?.[String(row)] ?? {}), ...(sheet.cells[String(row)] ?? {}) };
    const columns = Object.keys(cells).map(Number).sort((a, b) => a - b);
    const hidden = sheet.hiddenRows?.includes(row) ? ' hidden="1"' : '';
    const height = sheet.rowHeights?.[row];
    xml += `<row r="${row + 1}"${hidden}${height === undefined ? '' : ` ht="${height}" customHeight="1"`}>`;
    for (const column of columns) {
      const cell = cells[String(column)];
      if (!cell) continue;
      xml += buildCellXml(cell, row, column, styleIndexes, includeCachedValues);
    }
    xml += '</row>';
  }
  xml += '</sheetData>';
  if (sheet.merges.length) {
    xml += `<mergeCells count="${sheet.merges.length}">${sheet.merges.map((merge) => `<mergeCell ref="${rangeToA1(merge.range)}"/>`).join('')}</mergeCells>`;
  }
  const hyperlinks = Object.entries(sheet.cells).flatMap(([row, columns]) => Object.entries(columns).flatMap(([column, cell]) => {
    if (!cell.hyperlinkDetail) return [];
    const link = cell.hyperlinkDetail;
    const address = `${columnToLetter(Number(column))}${Number(row) + 1}`;
    const relation = relationships.find((candidate) => candidate.type === REL_HYPERLINK && candidate.target === hyperlinkTarget(link));
    const target = link.target.kind === 'url' || link.target.kind === 'email';
    return [`<hyperlink ref="${address}"${target && relation ? ` r:id="${relation.id}"` : ''}${link.target.kind === 'sheet' ? ` location="${encodeXml(link.target.address ?? '')}"` : ''}${link.tooltip ? ` tooltip="${encodeXml(link.tooltip)}"` : ''}/>`];
  }));
  if (hyperlinks.length) xml += `<hyperlinks>${hyperlinks.join('')}</hyperlinks>`;
  const tableRelations = relationships.filter((relation) => relation.type.endsWith('/table'));
  if (sheet.sheetTables?.length && tableRelations.length) {
    const tableParts = tableRelations.map((relation) => {
      const target = resolveTarget(sourcePart, relation.target);
      if (!files.has(target)) throw new Error(`Worksheet table relationship points to missing part: ${target}`);
      return `<tablePart r:id="${encodeXml(relation.id)}"/>`;
    });
    xml += `<tableParts count="${tableParts.length}">${tableParts.join('')}</tableParts>`;
  }
  const pivotRelations = relationships.filter((relation) => relation.type.endsWith('/pivotTable'));
  if (pivotRelations.length) {
    const pivotParts = pivotRelations.map((relation) => {
      const target = resolveTarget(sourcePart, relation.target);
      if (!files.has(target)) throw new Error(`Worksheet PivotTable relationship points to missing part: ${target}`);
      return `<pivotTablePart r:id="${encodeXml(relation.id)}"/>`;
    });
    xml += `<pivotTableParts count="${pivotParts.length}">${pivotParts.join('')}</pivotTableParts>`;
  }
  const drawingRelations = relationships.filter((relation) => relation.type === REL_DRAWING || relation.type.endsWith('/drawing'));
  if (drawingRelations.length && !child(originalRoot, 'drawing')) xml += `<drawing r:id="${encodeXml(drawingRelations[0]!.id)}"/>`;
  // A drawing or other unsupported worksheet node is still emitted through its
  // original relationship. Native Pivot parts are rebuilt above from the
  // canonical graph and therefore never copied from a stale worksheet node.
  if (originalRoot) {
    const preservedNodes = new Map<string, XmlNode>();
    for (const node of originalRoot.children) {
      const name = localName(node.name);
      if (name === 'drawing' || name === 'legacyDrawing' || name === 'oleObjects' || name === 'extLst' || name === 'picture' || (name === 'tableParts' && !sheet.sheetTables?.length)) preservedNodes.set(name, node);
    }
    for (const name of ['drawing', 'legacyDrawing', 'oleObjects', 'picture', 'tableParts']) {
      const node = preservedNodes.get(name);
      if (node) xml += serializeXml(node);
    }
    const extension = preservedNodes.get('extLst');
    if (extension) xml += serializeWorksheetControlExtensions(extension, nativeControls.filter((control) => control.sheetPart === sourcePart));
    else if (nativeControls.some((control) => control.sheetPart === sourcePart && control.valid)) xml += serializeWorksheetControlExtensions(undefined, nativeControls.filter((control) => control.sheetPart === sourcePart));
  } else if (nativeControls.some((control) => control.sheetPart === sourcePart && control.valid)) {
    xml += serializeWorksheetControlExtensions(undefined, nativeControls.filter((control) => control.sheetPart === sourcePart));
  }
  xml += '</worksheet>';
  return xml;
}

function buildCellXml(cell: CellData, row: number, column: number, styleIndexes: Map<string, number>, includeCachedValues: boolean): string {
  const ref = `${columnToLetter(column)}${row + 1}`;
  const styleKey = JSON.stringify({ style: cell.style, numberFormat: cell.numberFormat });
  const style = cell.style || cell.numberFormat ? (styleIndexes.get(styleKey) ?? 1) : undefined;
  const styleAttr = style === undefined ? '' : ` s="${style}"`;
  if (cell.formula) {
    const formula = cell.formula.startsWith('=') ? cell.formula.slice(1) : cell.formula;
    const cached = includeCachedValues && isScalar(cell.formulaValue) ? `<v>${encodeXml(String(cell.formulaValue))}</v>` : '';
    return `<c r="${ref}"${styleAttr}${cached && typeof cell.formulaValue === 'string' ? ' t="str"' : ''}><f>${encodeXml(formula)}</f>${cached}</c>`;
  }
  if (cell.value === null || cell.value === undefined) return `<c r="${ref}"${styleAttr}/>`;
  if (typeof cell.value === 'boolean') return `<c r="${ref}"${styleAttr} t="b"><v>${cell.value ? '1' : '0'}</v></c>`;
  if (typeof cell.value === 'number') return `<c r="${ref}"${styleAttr}><v>${Number.isFinite(cell.value) ? cell.value : 0}</v></c>`;
  return `<c r="${ref}"${styleAttr} t="inlineStr"><is><t>${encodeXml(cell.value)}</t></is></c>`;
}

function serializeWorksheetControlExtensions(original: XmlNode | undefined, controls: NativePivotControlDefinition[]): string {
  const root = original ? structuredClone(original) : firstElement(parseXml('<extLst/>'), 'extLst');
  if (!controls.some((control) => !control.valid)) {
    root.children = root.children.flatMap((extension) => {
      const hasNativeControl = descendants(extension, 'slicerList').length > 0 || descendants(extension, 'timelineRefs').length > 0;
      return hasNativeControl ? [] : [extension];
    });
  }
  const slicers = controls.filter((control) => control.kind === 'slicer' && control.valid && control.relationshipId);
  const timelines = controls.filter((control) => control.kind === 'timeline' && control.valid && control.relationshipId);
  if (slicers.length) {
    const node = firstElement(parseXml(`<ext uri="{A8765BA9-456A-4DAB-B4F3-ACF838C121DE}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:slicerList>${slicers.map((control) => `<x14:slicer r:id="${encodeXml(control.relationshipId)}"/>`).join('')}</x14:slicerList></ext>`), 'ext');
    root.children.push(node);
  }
  if (timelines.length) {
    const node = firstElement(parseXml(`<ext uri="{7E03D99C-DC04-49D9-9315-930204A7B6E9}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><x15:timelineRefs>${timelines.map((control) => `<x15:timelineRef r:id="${encodeXml(control.relationshipId)}"/>`).join('')}</x15:timelineRefs></ext>`), 'ext');
    root.children.push(node);
  }
  return serializeXml(root);
}

function serializeWorkbookControlExtensions(original: XmlNode | undefined, controls: NativePivotControlDefinition[], relationships: XlsxRelationship[]): string {
  const root = original ? structuredClone(original) : firstElement(parseXml('<extLst/>'), 'extLst');
  if (!controls.some((control) => !control.valid)) {
    root.children = root.children.flatMap((extension) => {
      const hasNativeControl = descendants(extension, 'slicerCaches').length > 0 || descendants(extension, 'timelineCacheRefs').length > 0;
      return hasNativeControl ? [] : [extension];
    });
  }
  const slicerCaches = [...new Map(controls.filter((control) => control.kind === 'slicer' && control.valid && control.cacheRelationshipId).map((control) => [control.cachePart, control])).values()];
  const timelines = [...new Map(controls.filter((control) => control.kind === 'timeline' && control.valid && control.cacheRelationshipId).map((control) => [control.cachePart, control])).values()];
  if (slicerCaches.length) {
    const node = firstElement(parseXml(`<ext uri="{BBE1A952-AA13-448E-AADC-164F8A28A991}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:slicerCaches>${slicerCaches.map((control) => `<x14:slicerCache r:id="${encodeXml(control.cacheRelationshipId)}"/>`).join('')}</x14:slicerCaches></ext>`), 'ext');
    root.children.push(node);
  }
  if (timelines.length) {
    const node = firstElement(parseXml(`<ext uri="{D0CA8CA8-9F24-4464-BF8E-62219DCF47F9}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><x15:timelineCacheRefs>${timelines.map((control) => `<x15:timelineCacheRef r:id="${encodeXml(control.cacheRelationshipId)}"/>`).join('')}</x15:timelineCacheRefs></ext>`), 'ext');
    root.children.push(node);
  }
  void relationships;
  return serializeXml(root);
}

function buildWorkbookXml(snapshot: WorkbookSnapshot, relationships: XlsxRelationship[], descriptors: SheetDescriptor[], dateSystem: DateSystem, nativePivotGraph?: NativePivotGraph, preserved?: XlsxPackage): string {
  const relationFor = (target: string, type: string) => relationships.find((relation) => relation.type === type && resolveTarget('xl/workbook.xml', relation.target) === target)?.id ?? '';
  let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="${NS_MAIN}" xmlns:r="${NS_DOC_REL}"><workbookPr date1904="${dateSystem === '1904' ? '1' : '0'}"/><sheets>`;
  for (const descriptor of descriptors) {
    const target = relativeTarget('xl/workbook.xml', descriptor.part);
    const id = relationFor(descriptor.part, REL_WORKSHEET);
    const sheet = snapshot.sheets.find((candidate) => candidate.id === descriptor.id);
    xml += `<sheet name="${encodeXml(sheet?.name ?? descriptor.name)}" sheetId="${encodeXml(descriptor.id.replace(/^sheet-/, ''))}" r:id="${id}"${sheet?.hidden ? ' state="hidden"' : ''}/>`;
  }
  xml += '</sheets>';
  const names: DefinedNameModel[] = snapshot.definedNameModels ?? Object.entries(snapshot.definedNames ?? {}).map(([name, formula]) => ({ name, formula, scope: 'workbook' as const }));
  if (names.length) {
    xml += '<definedNames>';
    for (const name of names) {
      const localSheetId = name.scope === 'sheet' && name.sheetId ? descriptors.findIndex((descriptor) => descriptor.id === name.sheetId) : -1;
      xml += `<definedName name="${encodeXml(name.name)}"${localSheetId >= 0 ? ` localSheetId="${localSheetId}"` : ''}${name.hidden ? ' hidden="1"' : ''}>${encodeXml(name.formula.startsWith('=') ? name.formula.slice(1) : name.formula)}</definedName>`;
    }
    xml += '</definedNames>';
  }
  // OOXML places pivotCaches after definedNames.  Keeping the canonical
  // child order avoids Excel repair prompts on otherwise valid preserved
  // native Pivot packages.
  if (nativePivotGraph?.caches.length) {
    xml += serializeNativePivotCaches(nativePivotGraph, relationships);
  }
  // Preserve workbook-level extension/calculation metadata that this package
  // does not edit.  Sheet references and defined names above remain canonical.
  const originalRoot = preserved?.parts['xl/workbook.xml'] ? firstElement(parseXml(strFromU8(preserved.parts['xl/workbook.xml'])), 'workbook') : undefined;
  if (originalRoot) {
    let hasExtensionList = false;
    for (const node of originalRoot.children) {
      const name = localName(node.name);
      if (name === 'bookViews' || name === 'calcPr' || name === 'fileVersion' || name === 'fileSharing' || name === 'workbookProtection') xml += serializeXml(node);
      else if (name === 'extLst') { hasExtensionList = true; xml += serializeWorkbookControlExtensions(node, nativePivotGraph?.controls ?? [], relationships); }
    }
    if (!hasExtensionList && nativePivotGraph?.controls?.some((control) => control.valid)) xml += serializeWorkbookControlExtensions(undefined, nativePivotGraph.controls, relationships);
  } else if (nativePivotGraph?.controls?.some((control) => control.valid)) {
    xml += serializeWorkbookControlExtensions(undefined, nativePivotGraph.controls, relationships);
  }
  return `${xml}</workbook>`;
}

function buildContentTypesXml(files: Map<string, Uint8Array>, preserved?: XlsxPackage): string {
  const defaults = new Map<string, string>([['rels', 'application/vnd.openxmlformats-package.relationships+xml'], ['xml', 'application/xml']]);
  const overrides = new Map<string, string>([
    ['/xl/workbook.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'],
    ['/xl/styles.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'],
    ['/xl/sharedStrings.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'],
  ]);
  const original = preserved?.contentTypesXml ? firstElement(parseXml(strFromU8(preserved.contentTypesXml)), 'Types') : undefined;
  for (const node of children(original, 'Default')) if (node.attrs.Extension && node.attrs.ContentType) defaults.set(node.attrs.Extension, node.attrs.ContentType);
  for (const node of children(original, 'Override')) if (node.attrs.PartName && node.attrs.ContentType) overrides.set(node.attrs.PartName, node.attrs.ContentType);
  for (const name of files.keys()) {
    if (!name.startsWith('xl/worksheets/') || !name.endsWith('.xml')) continue;
    overrides.set(`/${name}`, 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml');
  }
  for (const name of files.keys()) {
    if (name.startsWith('xl/tables/') && name.endsWith('.xml')) {
      overrides.set(`/${name}`, 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml');
    } else if (name.startsWith('xl/pivotTables/') && name.endsWith('.xml')) {
      overrides.set(`/${name}`, 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml');
    } else if (name.startsWith('xl/pivotCache/') && name.toLowerCase().includes('definition') && name.endsWith('.xml')) {
      overrides.set(`/${name}`, 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml');
    } else if (name.startsWith('xl/pivotCache/') && name.toLowerCase().includes('records') && name.endsWith('.xml')) {
      overrides.set(`/${name}`, 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml');
    } else if (name.startsWith('xl/slicerCaches/') && name.endsWith('.xml')) {
      overrides.set(`/${name}`, 'application/vnd.ms-excel.slicerCache+xml');
    } else if (name.startsWith('xl/slicers/') && name.endsWith('.xml')) {
      overrides.set(`/${name}`, 'application/vnd.ms-excel.slicer+xml');
    } else if (name.startsWith('xl/timelineCaches/') && name.endsWith('.xml')) {
      overrides.set(`/${name}`, 'application/vnd.ms-excel.TimelineCache+xml');
    } else if (name.startsWith('xl/timelines/') && name.endsWith('.xml')) {
      overrides.set(`/${name}`, 'application/vnd.ms-excel.timeline+xml');
    }
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${[...defaults.entries()].map(([extension, type]) => `<Default Extension="${encodeXml(extension)}" ContentType="${encodeXml(type)}"/>`).join('')}${[...overrides.entries()].filter(([part]) => files.has(part.slice(1))).map(([part, type]) => `<Override PartName="${encodeXml(part)}" ContentType="${encodeXml(type)}"/>`).join('')}</Types>`;
}

function buildRelationshipsXml(relationships: XlsxRelationship[]): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${NS_REL}">${relationships.map((relation) => `<Relationship Id="${encodeXml(relation.id)}" Type="${encodeXml(relation.type)}" Target="${encodeXml(relation.target)}"${relation.targetMode ? ` TargetMode="${encodeXml(relation.targetMode)}"` : ''}/>`).join('')}</Relationships>`;
}

function buildRootRelationshipsXml(existing: XlsxRelationship[]): string {
  const relationships = mergeRelationships(existing, [{ id: '', type: REL_OFFICE_DOCUMENT, target: 'xl/workbook.xml' }]);
  return buildRelationshipsXml(relationships);
}

function mergeRelationships(existing: XlsxRelationship[], required: Array<Pick<XlsxRelationship, 'type' | 'target' | 'targetMode'> & { id?: string }>): XlsxRelationship[] {
  const result = existing.map((relation) => ({ ...relation }));
  const used = new Set(result.map((relation) => relation.id));
  let next = 1;
  for (const request of required) {
    const found = result.find((relation) => relation.type === request.type && relation.target === request.target && relation.targetMode === request.targetMode);
    if (found) continue;
    let id = request.id ?? '';
    if (!id || used.has(id)) {
      while (used.has(`rId${next}`)) next += 1;
      id = `rId${next++}`;
    }
    used.add(id);
    result.push({ id, type: request.type, target: request.target, ...(request.targetMode ? { targetMode: request.targetMode } : {}) });
  }
  return result;
}

function collectHyperlinkRelationships(sheet: SheetSnapshot, existing: XlsxRelationship[]): Array<Pick<XlsxRelationship, 'type' | 'target' | 'targetMode'>> {
  const links: Array<Pick<XlsxRelationship, 'type' | 'target' | 'targetMode'>> = [];
  for (const row of Object.values(sheet.cells)) {
    for (const cell of Object.values(row)) {
      const target = cell.hyperlinkDetail?.target;
      if (!target || (target.kind !== 'url' && target.kind !== 'email')) continue;
      const href = hyperlinkTarget(cell.hyperlinkDetail!);
      if (!existing.some((relation) => relation.type === REL_HYPERLINK && relation.target === href)) {
        links.push({ type: REL_HYPERLINK, target: href, targetMode: 'External' });
      }
    }
  }
  return links;
}

function hyperlinkTarget(link: CellHyperlink): string {
  switch (link.target.kind) {
    case 'url': return link.target.url;
    case 'email': return `mailto:${link.target.address}${link.target.subject ? `?subject=${encodeURIComponent(link.target.subject)}` : ''}`;
    default: return '';
  }
}

function hyperlinkForCell(root: XmlNode, relationships: XlsxRelationship[], row: number, column: number): CellHyperlink | undefined {
  const hyperlinks = child(root, 'hyperlinks');
  const reference = `${columnToLetter(column)}${row + 1}`;
  const node = children(hyperlinks, 'hyperlink').find((candidate) => candidate.attrs.ref === reference);
  if (!node) return undefined;
  const relationId = node.attrs['r:id'];
  if (relationId) {
    const relation = relationships.find((candidate) => candidate.id === relationId);
    if (relation) {
      if (relation.target.startsWith('mailto:')) {
        return { id: `hyperlink-${row}-${column}`, target: { kind: 'email', address: relation.target.slice(7) }, ...(node.attrs.tooltip ? { tooltip: node.attrs.tooltip } : {}) };
      }
      return { id: `hyperlink-${row}-${column}`, target: { kind: 'url', url: relation.target }, ...(node.attrs.tooltip ? { tooltip: node.attrs.tooltip } : {}) };
    }
  }
  if (node.attrs.location) return { id: `hyperlink-${row}-${column}`, target: { kind: 'name', name: node.attrs.location }, ...(node.attrs.tooltip ? { tooltip: node.attrs.tooltip } : {}) };
  return undefined;
}

function parseNotes(root: XmlNode, descriptor: SheetDescriptor, files: Record<string, Uint8Array>, pkg: XlsxPackage): SheetSnapshot['notes'] {
  const relation = (pkg.relationships[descriptor.part] ?? []).find((candidate) => candidate.type.endsWith('/comments'));
  if (!relation) return [];
  const part = resolveTarget(descriptor.part, relation.target);
  const bytes = files[part];
  if (!bytes) return [];
  const commentsRoot = firstElement(parseXml(strFromU8(bytes)), 'comments');
  const authors = children(child(commentsRoot, 'authors'), 'author').map(textContent);
  return children(child(commentsRoot, 'commentList'), 'comment').flatMap((comment) => {
    const ref = parseA1(comment.attrs.ref ?? '');
    if (!ref) return [];
    return [{ row: ref.row, column: ref.column, note: { id: `note-${descriptor.id}-${ref.row}-${ref.column}`, author: authors[Number(comment.attrs.author) || 0] ?? 'Unknown', text: descendants(comment, 't').map(textContent).join(''), createdAt: new Date(0).toISOString(), visible: false } }];
  });
}

function parseDefinedNames(node: XmlNode | undefined, descriptors: SheetDescriptor[]): DefinedNameModel[] {
  return children(node, 'definedName').flatMap((name) => {
    const formula = textContent(name).trim();
    if (!name.attrs.name || !formula) return [];
    const localSheetId = name.attrs.localSheetId === undefined ? -1 : Number(name.attrs.localSheetId);
    const descriptor = localSheetId >= 0 ? descriptors[localSheetId] : undefined;
    return [{ name: name.attrs.name, formula: `=${formula}`, scope: descriptor ? 'sheet' : 'workbook', ...(descriptor ? { sheetId: descriptor.id } : {}), ...(name.attrs.hidden ? { hidden: name.attrs.hidden === '1' || name.attrs.hidden === 'true' } : {}) }];
  });
}

function parseFreeze(root: XmlNode): FreezeModel {
  const pane = child(child(child(root, 'sheetViews'), 'sheetView'), 'pane');
  const xSplit = Number(pane?.attrs.xSplit ?? 0) || 0;
  const ySplit = Number(pane?.attrs.ySplit ?? 0) || 0;
  const topLeft = pane?.attrs.topLeftCell ? parseA1(pane.attrs.topLeftCell) : undefined;
  return { xSplit, ySplit, startRow: topLeft?.row ?? ySplit, startColumn: topLeft?.column ?? xSplit };
}

function parseHiddenColumns(root: XmlNode): number[] {
  const result: number[] = [];
  for (const col of children(child(root, 'cols'), 'col')) {
    if (col.attrs.hidden !== '1' && col.attrs.hidden !== 'true') continue;
    const start = parsePositiveInt(col.attrs.min, 1) - 1;
    const end = parsePositiveInt(col.attrs.max, start + 1) - 1;
    for (let index = start; index <= end; index += 1) result.push(index);
  }
  return result;
}

function parseRowHeights(root: XmlNode): Record<number, number> {
  const result: Record<number, number> = {};
  for (const row of children(child(root, 'sheetData'), 'row')) {
    if (row.attrs.ht === undefined) continue;
    result[parsePositiveInt(row.attrs.r, 1) - 1] = Number(row.attrs.ht);
  }
  return result;
}

function parseColumnWidths(root: XmlNode): Record<number, number> {
  const result: Record<number, number> = {};
  for (const col of children(child(root, 'cols'), 'col')) {
    if (col.attrs.width === undefined) continue;
    const start = parsePositiveInt(col.attrs.min, 1) - 1;
    const end = parsePositiveInt(col.attrs.max, start + 1) - 1;
    for (let index = start; index <= end; index += 1) result[index] = Number(col.attrs.width);
  }
  return result;
}

function readCellValue(cell: XmlNode, sharedStrings: string[]): string | number | boolean | null {
  const type = cell.attrs.t;
  if (type === 'inlineStr') return descendants(cell, 't').map(textContent).join('');
  const raw = textContent(child(cell, 'v'));
  if (raw === '') return null;
  if (type === 's') return sharedStrings[Number(raw)] ?? '';
  if (type === 'b') return raw === '1' || raw.toLowerCase() === 'true';
  if (type === 'e') return raw;
  const number = Number(raw);
  return Number.isNaN(number) ? raw : number;
}

function parseWorkbookDateSystem(xml: string): DateSystem {
  const workbook = firstElement(parseXml(xml), 'workbook');
  const value = child(workbook, 'workbookPr')?.attrs.date1904;
  return value === '1' || value === 'true' ? '1904' : '1900';
}

function hasNativePivotMarkers(files: Record<string, Uint8Array>): boolean {
  for (const [name, bytes] of Object.entries(files)) {
    if (!name.endsWith('.xml')) continue;
    const lower = name.toLowerCase();
    if (lower.includes('/pivottable') || lower.includes('/pivotcache/') || lower.includes('/slicers/') || lower.includes('/slicercaches/') || lower.includes('/timelines/') || lower.includes('/timelinecaches/')) return true;
    const xml = strFromU8(bytes);
    if (xml.includes('<pivotCaches') || xml.includes(':pivotCaches') || xml.includes('<pivotTableParts') || xml.includes(':pivotTableParts') || xml.includes('slicerCaches') || xml.includes('timelineCacheRefs') || xml.includes('slicerList') || xml.includes('timelineRefs')) return true;
  }
  return false;
}

function readRelationships(files: Record<string, Uint8Array>): Record<string, XlsxRelationship[]> {
  const result: Record<string, XlsxRelationship[]> = {};
  for (const name of Object.keys(files)) {
    if (!name.endsWith('.rels')) continue;
    const source = sourcePartFromRelationships(name);
    const root = firstElement(parseXml(strFromU8(files[name]!)), 'Relationships');
    result[source] = children(root, 'Relationship').flatMap((node) => {
      if (!node.attrs.Id || !node.attrs.Type || !node.attrs.Target) return [];
      return [{ id: node.attrs.Id, type: node.attrs.Type, target: node.attrs.Target, ...(node.attrs.TargetMode ? { targetMode: node.attrs.TargetMode } : {}) }];
    });
  }
  return result;
}

function readSheetPartMap(files: Record<string, Uint8Array>, relationships: Record<string, XlsxRelationship[]>): Record<string, string> {
  const map: Record<string, string> = {};
  const workbook = firstElement(parseXml(strFromU8(files['xl/workbook.xml']!)), 'workbook');
  const rels = relationships['xl/workbook.xml'] ?? [];
  children(child(workbook, 'sheets'), 'sheet').forEach((node, index) => {
    const id = `sheet-${node.attrs.sheetId ?? index + 1}`;
    const relation = rels.find((candidate) => candidate.id === (node.attrs['r:id'] ?? node.attrs.id));
    map[id] = relation ? resolveTarget('xl/workbook.xml', relation.target) : `xl/worksheets/sheet${index + 1}.xml`;
  });
  return map;
}

function firstElement(root: XmlNode, name: string): XmlNode {
  const found = root.name === name || localName(root.name) === name ? root : descendants(root, name)[0];
  if (!found) throw new Error(`OOXML part is missing <${name}>`);
  return found;
}

function sourcePartFromRelationships(name: string): string {
  if (name === '_rels/.rels') return '';
  const marker = '/_rels/';
  const index = name.lastIndexOf(marker);
  if (index < 0) return '';
  const directory = name.slice(0, index);
  const file = name.slice(index + marker.length, -'.rels'.length);
  return directory ? `${directory}/${file}` : file;
}

function relationshipPartName(source: string): string {
  const slash = source.lastIndexOf('/');
  if (slash < 0) return `_rels/${source}.rels`;
  return `${source.slice(0, slash)}/_rels/${source.slice(slash + 1)}.rels`;
}

function normalizePartName(name: string): string {
  const value = name.replaceAll('\\', '/');
  if (!value || value.startsWith('/') || value.includes('\0') || /^[A-Za-z]:/.test(value)) throw new Error(`Unsafe XLSX part name: ${name}`);
  const pieces: string[] = [];
  for (const piece of value.split('/')) {
    if (!piece || piece === '.') continue;
    if (piece === '..') {
      if (!pieces.length) throw new Error(`Unsafe XLSX part name: ${name}`);
      pieces.pop();
    } else pieces.push(piece);
  }
  const normalized = pieces.join('/');
  if (!normalized) throw new Error(`Unsafe XLSX part name: ${name}`);
  return normalized;
}

function resolveTarget(source: string, target: string): string {
  if (target.startsWith('/')) return normalizePartName(target.slice(1));
  const base = source.includes('/') ? source.slice(0, source.lastIndexOf('/') + 1) : '';
  return normalizePartName(`${base}${target}`);
}

function relativeTarget(source: string, target: string): string {
  const sourceDir = source.includes('/') ? source.slice(0, source.lastIndexOf('/') + 1) : '';
  const sourceParts = sourceDir.split('/').filter(Boolean);
  const targetParts = target.split('/').filter(Boolean);
  while (sourceParts.length && targetParts.length && sourceParts[0] === targetParts[0]) {
    sourceParts.shift();
    targetParts.shift();
  }
  return `${'../'.repeat(sourceParts.length)}${targetParts.join('/')}`;
}

function cloneParts(parts: Record<string, Uint8Array>): Record<string, Uint8Array> {
  return Object.fromEntries(Object.entries(parts).map(([name, data]) => [name, data.slice()]));
}

function isMacroPart(name: string): boolean {
  return name.toLowerCase().includes('vba') || name.toLowerCase().endsWith('.bin');
}

function parseRange(value: string | undefined): RangeRef | undefined {
  if (!value) return undefined;
  const parts = value.split(':');
  const start = parseA1(parts[0] ?? '');
  const end = parseA1(parts[1] ?? parts[0] ?? '');
  if (!start || !end) return undefined;
  return { sheetId: '', startRow: start.row, endRow: end.row, startColumn: start.column, endColumn: end.column };
}

function parseA1(value: string): { row: number; column: number } | undefined {
  const trimmed = value.replace(/\$/g, '').split('!').pop() ?? '';
  const match = /^([A-Za-z]+)(\d+)$/.exec(trimmed);
  if (!match) return undefined;
  return { column: columnFromLetter(match[1]!), row: Number(match[2]) - 1 };
}

function rangeToA1(range: RangeRef): string {
  const start = `${columnToLetter(range.startColumn)}${range.startRow + 1}`;
  const end = `${columnToLetter(range.endColumn)}${range.endRow + 1}`;
  return start === end ? start : `${start}:${end}`;
}

function inferDimension(sheet: SheetSnapshot): string {
  let maxRow = 0;
  let maxColumn = 0;
  let hasCell = false;
  for (const [row, columns] of Object.entries(sheet.cells)) {
    for (const column of Object.keys(columns)) {
      hasCell = true;
      maxRow = Math.max(maxRow, Number(row));
      maxColumn = Math.max(maxColumn, Number(column));
    }
  }
  if (!hasCell) return 'A1';
  return `A1:${columnToLetter(maxColumn)}${maxRow + 1}`;
}

function columnToLetter(index: number): string {
  let value = index + 1;
  let result = '';
  while (value > 0) {
    const remainder = (value - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    value = Math.floor((value - 1) / 26);
  }
  return result;
}

function columnFromLetter(value: string): number {
  let result = 0;
  for (const character of value.toUpperCase()) result = result * 26 + character.charCodeAt(0) - 64;
  return result - 1;
}

function parsePositiveInt(value: string | undefined, fallback: number): number {
  const parsed = Number(value);
  return Number.isInteger(parsed) && parsed > 0 ? parsed : fallback;
}

function isScalar(value: unknown): value is string | number | boolean {
  return typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean';
}

function randomId(): string {
  return Math.random().toString(36).slice(2, 10);
}

function builtInNumberFormat(id: number): string | undefined {
  const formats: Record<number, string> = { 0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 9: '0%', 10: '0.00%', 14: 'm/d/yy', 20: 'h:mm', 21: 'h:mm:ss', 22: 'm/d/yy h:mm' };
  return formats[id];
}

function builtInNumberFormatId(format: string): number | undefined {
  for (const id of [0, 1, 2, 3, 4, 9, 10, 14, 20, 21, 22]) if (builtInNumberFormat(id) === format) return id;
  return undefined;
}

function normalizeHorizontal(value: string): CellStyle['horizontalAlignment'] | undefined {
  return value === 'centerContinuous' ? 'center' : value === 'general' ? undefined : value as CellStyle['horizontalAlignment'];
}

function normalizeVertical(value: string): CellStyle['verticalAlignment'] | undefined {
  return value === 'center' ? 'middle' : value as CellStyle['verticalAlignment'];
}

function descriptorsForSnapshot(snapshot: WorkbookSnapshot): SheetDescriptor[] {
  return snapshot.sheets.map((sheet, index) => ({ id: sheet.id, name: sheet.name, part: `xl/worksheets/sheet${index + 1}.xml`, hidden: Boolean(sheet.hidden) }));
}
