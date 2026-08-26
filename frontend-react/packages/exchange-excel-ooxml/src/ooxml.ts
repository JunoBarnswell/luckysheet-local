import type {
  CellData,
  CellStyleTemplate,
  CellStyle,
  CellHyperlink,
  ConditionalFormatRule,
  DataValidationRule,
  DefinedNameModel,
  DrawingObject,
  AutoFilterModel,
  FilterScalar,
  MergeSpan,
  OutlineModel,
  ProtectionRule,
  RichTextRun,
  WorkbookSnapshot,
  SheetSnapshot,
  RangeRef,
  WorksheetPane,
} from '@react-sheets/core-model';
import { assertCanonicalWorkbookSnapshot, createPivotMemberKey, isDynamicFilterType, normalizeFontFamily, pivotSourceIdentity, resolveFilterCellValue } from '@react-sheets/core-model';
import {
  canonicalExcelDateDayOfWeek,
  canonicalExcelDateFromParts,
  canonicalExcelDateFromValue,
  canonicalExcelDateFromUtcDate,
  canonicalExcelDateToUtcDate,
  shiftCanonicalExcelDate,
  type CanonicalExcelDate,
  type CanonicalExcelDateParts,
  formatFormula,
  offsetAst,
  parseFormula,
} from '@react-sheets/formula-engine';
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
  type ExcelDocumentFormat,
  type OpcPackageGraph,
  type XlsxRelationship,
  type XlsxZipLimits,
} from './types';
import { mapNativePivotDefinition, readNativePivotGraph, serializeNativePivotCaches, synchronizeNativePivotPackage } from './native-pivot';
import { synchronizeNativePivotCharts } from './native-chart';
import type { NativePivotControlDefinition, NativePivotGraph } from './types';
import { builtInNumberFormat, builtInNumberFormatId, collectCustomNumberFormatIds } from './native-number-format';
import { canonicalDateToSerial, isExcelDateFormat, parseDateSystem, serialToCanonicalDate } from './date-system';

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}
import {
  DEFAULT_EXCEL_FONT_FAMILY,
  DEFAULT_EXCEL_FONT_SIZE_PT,
  DEFAULT_OOXML_FONT_MEASURER,
  excelColumnWidthToPixels,
  pixelsToExcelColumnWidth,
  pixelsToPoints,
  pointsToPixels,
  type OoxmlFontMeasurer,
  type OoxmlNormalFont,
} from './ooxml-metrics';

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const NS_DOC_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL_OFFICE_DOCUMENT = `${NS_DOC_REL}/officeDocument`;
const REL_WORKSHEET = `${NS_DOC_REL}/worksheet`;
const REL_STYLES = `${NS_DOC_REL}/styles`;
const REL_SHARED_STRINGS = `${NS_DOC_REL}/sharedStrings`;
const REL_HYPERLINK = `${NS_DOC_REL}/hyperlink`;
const REL_DRAWING = `${NS_DOC_REL}/drawing`;
const REL_CUSTOM_XML = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml';
const REACT_SHEETS_METADATA_PART = 'customXml/react-sheets-workbook.xml';
const OOXML_MAX_ROW_INDEX = 1_048_575;
const OOXML_MAX_COLUMN_INDEX = 16_383;

export interface LoadedOpcPackageGraph {
  packageGraph: OpcPackageGraph;
  files: Record<string, Uint8Array>;
}

export interface ParsedOpcPackageGraph {
  packageGraph: OpcPackageGraph;
  snapshot: WorkbookSnapshot;
  features: string[];
}

export interface ParseLoadedXlsxOptions {
  fontMeasurer?: OoxmlFontMeasurer;
  workbookName?: string;
  /** Required when an imported AutoFilter contains a dynamic date criterion. */
  canonicalReferenceDate?: CanonicalExcelDateParts;
}

interface StyleRecord {
  numberFormat?: string;
  style?: CellStyle;
}

interface SharedStringRecord {
  value: string;
  richText?: RichTextRun[];
}

interface StyleContext {
  records: StyleRecord[];
  namedCellStyles: CellStyleTemplate[];
  differentialStyles: Array<CellStyle | undefined>;
  normalFont: OoxmlNormalFont;
  maximumDigitWidthPx: number;
  themeColors: string[];
}

interface SheetDescriptor {
  id: string;
  name: string;
  part: string;
  hidden: boolean;
}

export function loadOpcPackageGraph(input: ArrayBuffer | Uint8Array, limits: Partial<XlsxZipLimits> = {}, fileName = 'workbook.xlsx'): LoadedOpcPackageGraph {
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
  const relationships = readRelationships(normalizedFiles);
  const rootOfficeDocument = (relationships[''] ?? []).find((relation) => isRelationshipKind(relation.type, 'officeDocument'));
  const workbookPart = rootOfficeDocument ? resolveTarget('', rootOfficeDocument.target) : 'xl/workbook.xml';
  if (!normalizedFiles[workbookPart]) throw new Error(`Not a valid XLSX package: workbook part is missing (${workbookPart})`);
  const dateSystem = parseDateSystem(strFromU8(normalizedFiles[workbookPart]));
  const format = detectOoxmlFormat(normalizedFiles, workbookPart, fileName);
  const sheetPartById = readSheetPartMap(normalizedFiles, relationships, workbookPart);
  const coreParts = new Set<string>([
    '[Content_Types].xml', '_rels/.rels', workbookPart, relationshipPartName(workbookPart),
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

  const nativePivotGraph = workbookPart === 'xl/workbook.xml' && hasNativePivotMarkers(normalizedFiles)
    ? readNativePivotGraph({ files: normalizedFiles, relationships, sheetPartById, dateSystem })
    : undefined;
  return {
    files: normalizedFiles,
    packageGraph: {
      schema: 'OpcPackageGraph',
      workbookPart,
      parts: cloneParts(normalizedFiles),
      opaqueParts,
      relationships,
      sheetPartById,
      contentTypesXml: normalizedFiles['[Content_Types].xml']?.slice(),
    dateSystem,
      format,
      profile: format.family === 'ooxml' ? format.profile : 'transitional',
      ...(nativePivotGraph ? { nativePivotGraph } : {}),
    },
  };
}

export function parseLoadedXlsx(loaded: LoadedOpcPackageGraph, options: ParseLoadedXlsxOptions = {}): ParsedOpcPackageGraph {
  const files = loaded.files;
  const workbookPart = loaded.packageGraph.workbookPart;
  const workbookXml = parseXml(strFromU8(files[workbookPart]!));
  const workbook = firstElement(workbookXml, 'workbook');
  const relationships = loaded.packageGraph.relationships;
  const descriptors: SheetDescriptor[] = [];
  const workbookRels = relationships[workbookPart] ?? [];
  const sheetNodes = children(child(workbook, 'sheets'), 'sheet');
  for (let index = 0; index < sheetNodes.length; index += 1) {
    const node = sheetNodes[index]!;
    const relationId = node.attrs['r:id'] ?? node.attrs.id ?? '';
    const relation = workbookRels.find((candidate) => candidate.id === relationId && isRelationshipKind(candidate.type, 'worksheet'));
    const part = relation ? resolveTarget(workbookPart, relation.target) : resolveTarget(workbookPart, `worksheets/sheet${index + 1}.xml`);
    if (!files[part]) throw new Error(`Workbook sheet relation points to missing part: ${part}`);
    const sheetId = `sheet-${node.attrs.sheetId ?? index + 1}`;
    descriptors.push({
      id: sheetId,
      name: node.attrs.name ?? `Sheet${index + 1}`,
      part,
      hidden: node.attrs.state === 'hidden' || node.attrs.state === 'veryHidden',
    });
    loaded.packageGraph.sheetPartById[sheetId] = part;
  }
  if (descriptors.length === 0) throw new Error('XLSX workbook has no worksheets');

  const sharedStringsPart = resolveWorkbookRelatedPart(workbookPart, workbookRels, 'sharedStrings', resolveTarget(workbookPart, 'sharedStrings.xml'));
  const stylesPart = resolveWorkbookRelatedPart(workbookPart, workbookRels, 'styles', resolveTarget(workbookPart, 'styles.xml'));
  const themePart = resolveWorkbookRelatedPart(workbookPart, workbookRels, 'theme', resolveTarget(workbookPart, 'theme/theme1.xml'));
  const styles = parseStyles(files[stylesPart], files[themePart], options.fontMeasurer ?? DEFAULT_OOXML_FONT_MEASURER);
  const sharedStrings = parseSharedStrings(files[sharedStringsPart], styles.themeColors);
  const sheets = descriptors.map((descriptor) => parseSheet(descriptor, files, loaded.packageGraph, sharedStrings, styles, options.canonicalReferenceDate, descriptors));
  const definedNameModels = parseDefinedNames(child(workbook, 'definedNames'), descriptors);
  const definedNames: Record<string, string> = {};
  for (const name of definedNameModels) if (name.scope === 'workbook') definedNames[name.name] = name.formula;
  const unitId = `imported-${randomId()}`;
  const snapshot: WorkbookSnapshot = {
    schema: 'WorkbookSnapshot',
    version: 5,
    unitId,
    name: options.workbookName ?? 'Imported Workbook',
    dimensionMetrics: { normalFontFamily: styles.normalFont.family, normalFontSizePx: pointsToPixels(styles.normalFont.sizePt), maximumDigitWidthPx: styles.maximumDigitWidthPx },
    definedNames,
    definedNameModels,
    dataModel: { sources: [], tables: [], relationships: [], views: [] },
    cellStyleTemplates: styles.namedCellStyles,
    printDocuments: sheets.flatMap((sheet, index) => parsePrintDocument(
      firstElement(parseXml(strFromU8(files[descriptors[index]!.part]!)), 'worksheet'),
      unitId,
      sheet.id,
    )),
    sheets,
  };
  applyReactSheetsMetadata(snapshot, files[REACT_SHEETS_METADATA_PART], loaded.packageGraph);
  assertCanonicalWorkbookSnapshot(snapshot);
  applyPrintDefinedNames(snapshot);
  attachNativePivots(snapshot, loaded.packageGraph.nativePivotGraph, loaded.packageGraph.sheetPartById);
  return { packageGraph: loaded.packageGraph, snapshot, features: detectPackageFeatures(loaded.packageGraph, snapshot) };
}

function detectOoxmlFormat(files: Record<string, Uint8Array>, workbookPart: string, fileName: string): ExcelDocumentFormat {
  const lowerName = fileName.toLocaleLowerCase();
  const workbookXml = strFromU8(files[workbookPart] ?? new Uint8Array());
  const contentTypes = strFromU8(files['[Content_Types].xml'] ?? new Uint8Array());
  const strict = workbookXml.includes('purl.oclc.org/ooxml/spreadsheetml') || contentTypes.includes('purl.oclc.org/ooxml/officeDocument');
  const hasMacros = Object.keys(files).some((name) => isVbaPart(name, undefined, files));
  const variant = lowerName.endsWith('.xlam') ? 'xlam' : lowerName.endsWith('.xltm') ? 'xltm' : lowerName.endsWith('.xltx') ? 'xltx' : lowerName.endsWith('.xlsm') || hasMacros ? 'xlsm' : 'xlsx';
  return { family: 'ooxml', profile: strict ? 'strict' : 'transitional', variant };
}

export function exportSnapshotToOpcPackageGraph(
  snapshot: WorkbookSnapshot,
  options: { dateSystem: DateSystem; includeCachedValues?: boolean; preserveMacros?: boolean },
  preserved?: OpcPackageGraph,
): ArrayBuffer {
  const files = new Map<string, Uint8Array>();
  if (preserved) {
    for (const [name, data] of Object.entries(preserved.parts)) {
      if (options.preserveMacros === false && isVbaPart(name, preserved)) continue;
      files.set(normalizePartName(name), data.slice());
    }
  }
  const sourceFiles = preserved?.parts ?? {};
  const workbookPart = preserved?.workbookPart ?? 'xl/workbook.xml';
  const workbookRelationships = options.preserveMacros === false
    ? filterMacroRelationships(workbookPart, preserved?.relationships[workbookPart] ?? [], preserved)
    : preserved?.relationships[workbookPart] ?? [];
  const stylesPart = relationshipTarget(preserved, workbookPart, REL_STYLES) ?? 'xl/styles.xml';
  const sharedStringsPart = relationshipTarget(preserved, workbookPart, REL_SHARED_STRINGS) ?? 'xl/sharedStrings.xml';
  const sheetParts = snapshot.sheets.map((sheet, index) => preserved?.sheetPartById[sheet.id] ?? `xl/worksheets/sheet${index + 1}.xml`);
  const sheetPartById = Object.fromEntries(snapshot.sheets.map((sheet, index) => [sheet.id, sheetParts[index]!])) as Record<string, string>;
  const nativeUpdate = synchronizeNativePivotPackage({
    files: Object.fromEntries([...files.entries()]),
    relationships: preserved?.relationships ?? {},
    graph: preserved?.nativePivotGraph,
    snapshot,
    sheetPartById,
    dateSystem: options.dateSystem,
  });
  const chartUpdate = synchronizeNativePivotCharts({
    files: nativeUpdate.files,
    relationships: nativeUpdate.relationships,
    snapshot,
    graph: nativeUpdate.graph,
    sheetPartById,
    displayCellsBySheetPart: nativeUpdate.displayCellsBySheetPart,
  });
  nativeUpdate.files = chartUpdate.files;
  nativeUpdate.relationships = chartUpdate.relationships;
  files.clear();
  for (const [name, data] of Object.entries(nativeUpdate.files)) files.set(name, data);
  const originalStylesXml = preserved && files.get(stylesPart) ? strFromU8(files.get(stylesPart)!) : undefined;
  const styleOutput = buildStyles(snapshot, originalStylesXml);
  const styleIndexes = collectStyleIndexes(snapshot);
  const differentialStyleIndexes = collectDifferentialStyleIndexes(snapshot);
  const sharedOutput = buildSharedStrings(snapshot);
  const sheetRelationships: Record<string, XlsxRelationship[]> = {};
  for (let index = 0; index < snapshot.sheets.length; index += 1) {
    const sheet = snapshot.sheets[index]!;
    const part = sheetParts[index]!;
    const original = preserved ? strFromU8(sourceFiles[part] ?? new Uint8Array()) : '';
    const originalRoot = original ? firstElement(parseXml(original), 'worksheet') : undefined;
    const requiredHyperlinks = collectHyperlinkRelationships(sheet, preserved?.relationships[part] ?? []);
    const tableParts = prepareTableParts(sheet, part, preserved, files, differentialStyleIndexes);
    for (const [tablePart, tableXml] of tableParts.parts) files.set(tablePart, strToU8(tableXml));
    const originalRelationships = nativeUpdate.relationships[part] ?? preserved?.relationships[part] ?? [];
    const relationships = mergeRelationships(
      options.preserveMacros === false ? filterMacroRelationships(part, originalRelationships, preserved) : originalRelationships,
      [...requiredHyperlinks, ...tableParts.required],
    );
    sheetRelationships[part] = relationships;
    files.set(part, strToU8(buildWorksheetXml(sheet, part, relationships, originalRoot, files, styleIndexes, differentialStyleIndexes, snapshot.dimensionMetrics.maximumDigitWidthPx, options.includeCachedValues ?? true, options.dateSystem, nativeUpdate.displayCellsBySheetPart[part], nativeUpdate.graph.controls ?? [], snapshot.printDocuments?.find((document) => document.sheetId === sheet.id), new Map(snapshot.sheets.map((entry) => [entry.id, entry.name])))));
  }

  const workbookRelationsSource = nativeUpdate.relationships[workbookPart] ?? workbookRelationships;
  const workbookRelations = mergeRelationships(
    options.preserveMacros === false ? filterMacroRelationships(workbookPart, workbookRelationsSource, preserved) : workbookRelationsSource,
    [
      { id: '', type: REL_STYLES, target: relativeTarget(workbookPart, stylesPart) },
      { id: '', type: REL_SHARED_STRINGS, target: relativeTarget(workbookPart, sharedStringsPart) },
      ...sheetParts.map((part) => ({ id: '', type: REL_WORKSHEET, target: relativeTarget(workbookPart, part) })),
    ],
  );
  files.set(workbookPart, strToU8(buildWorkbookXml(snapshot, workbookPart, workbookRelations, descriptorsForSnapshot(snapshot), options.dateSystem, nativeUpdate.graph, preserved)));
  files.set(relationshipPartName(workbookPart), strToU8(buildRelationshipsXml(workbookRelations)));
  files.set(REACT_SHEETS_METADATA_PART, strToU8(buildReactSheetsMetadata(snapshot)));
  const rootRelationships = mergeRelationships(
    (preserved?.relationships[''] ?? []).filter((relationship) => !isRelationshipKind(relationship.type, 'officeDocument') && relationship.type !== REL_CUSTOM_XML),
    [{ id: '', type: REL_OFFICE_DOCUMENT, target: relativeTarget('', workbookPart) }, { id: '', type: REL_CUSTOM_XML, target: REACT_SHEETS_METADATA_PART }],
  );
  files.set('_rels/.rels', strToU8(buildRelationshipsXml(rootRelationships)));
  files.set(stylesPart, strToU8(styleOutput));
  files.set(sharedStringsPart, strToU8(sharedOutput));
  for (const [source, relationships] of Object.entries(sheetRelationships)) {
    files.set(relationshipPartName(source), strToU8(buildRelationshipsXml(relationships)));
  }
  for (const [source, relationships] of Object.entries(nativeUpdate.relationships)) {
    if (!source || source === workbookPart || sheetRelationships[source]) continue;
    const outputRelationships = options.preserveMacros === false ? filterMacroRelationships(source, relationships, preserved) : relationships;
    files.set(relationshipPartName(source), strToU8(buildRelationshipsXml(outputRelationships)));
  }
  files.set('[Content_Types].xml', strToU8(buildContentTypesXml(files, preserved, workbookPart, stylesPart, sharedStringsPart)));

  if (preserved?.profile === 'strict') {
    const strictParts = new Set<string>([
      workbookPart,
      stylesPart,
      sharedStringsPart,
      ...sheetParts,
      relationshipPartName(workbookPart),
      '_rels/.rels',
      '[Content_Types].xml',
      ...Object.keys(sheetRelationships).map(relationshipPartName),
      ...Object.keys(nativeUpdate.relationships).map(relationshipPartName),
    ]);
    for (const name of strictParts) {
      const data = files.get(name);
      if (!data) continue;
      files.set(name, strToU8(strFromU8(data)
        .replaceAll(NS_MAIN, 'http://purl.oclc.org/ooxml/spreadsheetml/main')
        .replaceAll(NS_DOC_REL, 'http://purl.oclc.org/ooxml/officeDocument/relationships')));
    }
  }

  const zipped: Record<string, Uint8Array> = {};
  for (const [name, data] of files) zipped[name] = data;
  const zippedBytes = zipSync(zipped, { level: 6 });
  return zippedBytes.buffer.slice(zippedBytes.byteOffset, zippedBytes.byteOffset + zippedBytes.byteLength) as ArrayBuffer;
}

export function detectPackageFeatures(pkg: OpcPackageGraph, snapshot?: WorkbookSnapshot): string[] {
  const features = new Set<string>(snapshot ? ['cells', 'styles'] : []);
  for (const name of Object.keys(pkg.parts)) {
    const lower = name.toLowerCase();
    if (lower.includes('/charts/')) features.add('charts');
    if (lower.includes('/pivot') || lower.includes('pivottableparts')) features.add('pivot');
    if (isVbaPart(name, pkg)) features.add('vba');
    if (lower.includes('externalconnections') || lower.includes('connections.xml')) features.add('external-connection');
    if (lower.includes('/slicers/') || lower.includes('slicer')) features.add('slicer');
    if (lower.includes('/timelines/') || lower.includes('timeline')) features.add('timeline');
    if (lower.includes('/theme/')) features.add('theme');
    if (lower.includes('/comments')) features.add('comments');
    // A drawing part is also the container for charts and native
    // Slicer/Timeline controls; it is not evidence of an image by itself.
    if (lower.includes('/media/') || lower.includes('/images/')) features.add('images');
    if (lower.includes('/tables/')) features.add('tables');
  }
  if (snapshot) {
    for (const sheet of snapshot.sheets) {
      if (sheet.merges.length) features.add('merges');
      if (sheet.pane.kind === 'frozen') features.add('freeze');
      if (sheet.pane.kind === 'split') features.add('split');
      if (Object.values(sheet.cells).some((row) => Object.values(row).some((cell) => Boolean(cell.formula)))) features.add('formulas');
      if (sheet.conditionalFormats?.length) features.add('conditional-format');
      if (sheet.dataValidations?.length) features.add('validation');
      if (sheet.sheetTables?.length) features.add('tables');
      if (sheet.autoFilter) features.add('filters');
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
  pkg: OpcPackageGraph,
  sharedStrings: SharedStringRecord[],
  styles: StyleContext,
  canonicalReferenceDate?: CanonicalExcelDateParts,
  sheetDescriptors: readonly SheetDescriptor[] = [],
): SheetSnapshot {
  const xml = strFromU8(files[descriptor.part]!);
  const root = firstElement(parseXml(xml), 'worksheet');
  const cells: Record<string, Record<string, CellData>> = {};
  const hyperlinks: NonNullable<SheetSnapshot['hyperlinks']> = [];
  const hiddenRows: number[] = [];
  const dimensions = parseRange(child(root, 'dimension')?.attrs.ref);
  if (dimensions) {
    if (dimensions.startRow < 0 || dimensions.startColumn < 0 || dimensions.endRow < dimensions.startRow || dimensions.endColumn < dimensions.startColumn
      || dimensions.endRow > OOXML_MAX_ROW_INDEX || dimensions.endColumn > OOXML_MAX_COLUMN_INDEX) {
      throw new Error(`UNSUPPORTED_FEATURE: Worksheet ${descriptor.name} dimension exceeds the OOXML worksheet boundary`);
    }
  }
  const sheetData = child(root, 'sheetData');
  const sheetFormat = child(root, 'sheetFormatPr');
  const defaultRowHeightPt = finitePositive(sheetFormat?.attrs.defaultRowHeight, 15);
  const defaultColumnWidthChars = finitePositive(sheetFormat?.attrs.defaultColWidth ?? sheetFormat?.attrs.baseColWidth, 8.7109375);
  const defaultRowHeightPx = pointsToPixels(defaultRowHeightPt);
  const defaultColumnWidthPx = excelColumnWidthToPixels(defaultColumnWidthChars, styles.maximumDigitWidthPx);
  const rowHeightsPx = parseRowHeights(root);
  const columnWidthsPx = parseColumnWidths(root, styles.maximumDigitWidthPx);
  const pane = parsePane(root);
  let maxRow = dimensions?.endRow ?? -1;
  let maxColumn = dimensions?.endColumn ?? -1;
  const cellEntries: Array<{ node: XmlNode; row: number; column: number }> = [];
  for (const rowNode of children(sheetData, 'row')) {
    const rowNumber = parsePositiveInt(rowNode.attrs.r, 1) - 1;
    assertOoxmlAddress(rowNumber, 0, `${descriptor.name}!row ${rowNode.attrs.r ?? ''}`);
    maxRow = Math.max(maxRow, rowNumber);
    if (rowNode.attrs.hidden === '1' || rowNode.attrs.hidden === 'true') hiddenRows.push(rowNumber);
    for (const cellNode of children(rowNode, 'c')) {
      const address = parseA1(cellNode.attrs.r ?? 'A1');
      if (!address) throw new Error(`Worksheet ${descriptor.name} contains an invalid cell reference: ${cellNode.attrs.r ?? ''}`);
      assertOoxmlAddress(address.row, address.column, `${descriptor.name}!${cellNode.attrs.r ?? ''}`);
      maxColumn = Math.max(maxColumn, address.column);
      maxRow = Math.max(maxRow, address.row);
      cellEntries.push({ node: cellNode, row: address.row, column: address.column });
    }
  }

  const sharedFormulaMasters = collectSharedFormulaMasters(cellEntries, descriptor);
  for (const entry of cellEntries) {
      const cellNode = entry.node;
      const address = { row: entry.row, column: entry.column };
      const styleId = cellNode.attrs.s;
      const style = styleId === undefined ? undefined : styles.records[Number(styleId)];
      const valueRecord = readCellValue(cellNode, sharedStrings, styles.themeColors);
      const canonicalValue = canonicalizeImportedDate(valueRecord.value, style?.numberFormat, pkg.dateSystem, `${descriptor.name}!${cellNode.attrs.r ?? 'unknown'}`);
      const formulaRecord = readFormula(cellNode, address, sharedFormulaMasters, descriptor);
      const cell: CellData = {
        value: canonicalValue,
        ...(valueRecord.richText ? { richText: valueRecord.richText } : {}),
        ...(formulaRecord.formula ? { formula: formulaRecord.formula } : {}),
        ...(formulaRecord.metadata ? { formulaMetadata: formulaRecord.metadata } : {}),
        ...(formulaRecord.formula && isScalar(canonicalValue) ? { formulaValue: canonicalValue } : {}),
        ...(styleId === undefined ? {} : { styleId }),
        ...(style?.style ? { style: structuredClone(style.style) } : {}),
        ...(style?.numberFormat ? { numberFormat: style.numberFormat } : {}),
      };
      cells[String(address.row)] ??= {};
      cells[String(address.row)]![String(address.column)] = cell;
  }
  for (const hyperlinkNode of children(child(root, 'hyperlinks'), 'hyperlink')) {
    const address = parseA1(hyperlinkNode.attrs.ref ?? '');
    if (!address) throw new Error(`Worksheet ${descriptor.name} contains an invalid hyperlink reference: ${hyperlinkNode.attrs.ref ?? ''}`);
    const hyperlink = hyperlinkForCell(root, pkg.relationships[descriptor.part] ?? [], address.row, address.column, sheetDescriptors);
    if (hyperlink) hyperlinks.push({ row: address.row, column: address.column, hyperlink });
  }
  const merges = children(child(root, 'mergeCells'), 'mergeCell')
    .map((node) => requireSheetRange(node.attrs.ref, descriptor, 'merge'))
    .map((range) => ({ range: { ...range, sheetId: descriptor.id }, anchor: { row: range.startRow, column: range.startColumn } } satisfies MergeSpan));
  validateNonOverlappingMerges(merges, descriptor);
  const hiddenColumns = parseHiddenColumns(root);
  const tabColor = resolveColor(child(child(root, 'sheetPr'), 'tabColor'), styles.themeColors);
  const notes = parseNotes(root, descriptor, files, pkg);
  const sheetTables = parseSheetTables(root, descriptor, files, pkg, styles);
  const conditionalFormats = parseConditionalFormats(root, descriptor, styles);
  const dataValidations = parseDataValidations(root, descriptor);
  const autoFilter = parseAutoFilter(root, descriptor, styles);
  const importedFilters = [
    ...(autoFilter ? [autoFilter] : []),
    ...sheetTables.flatMap((table) => table.autoFilter ? [table.autoFilter] : []),
  ];
  if (importedFilters.some((filter) => Object.values(filter.columns).some((column) => column.criterion?.kind === 'dynamic')) && !canonicalReferenceDate) {
    throw new Error('Dynamic date AutoFilter requires an explicit canonical workbook reference date');
  }
  materializeFilterMetadata(cells, descriptor.id, importedFilters, conditionalFormats, pkg.dateSystem);
  const filterOwnedRows = new Set<number>();
  for (const filter of importedFilters) {
    const table = sheetTables.find((candidate) => candidate.autoFilter === filter);
    const endRow = table?.hasTotalRow ? filter.range.endRow - 1 : filter.range.endRow;
    for (const row of hiddenRows) {
      if (row <= filter.range.startRow || row > endRow || !filterRowMatches(cells, filter, row, pkg.dateSystem, canonicalReferenceDate)) continue;
      filterOwnedRows.add(row);
    }
  }
  const manualHiddenRows = hiddenRows.filter((row) => !filterOwnedRows.has(row));
  const outline = parseOutline(root);
  const protectionRules = parseProtection(root, descriptor);
  const sheetView = child(child(root, 'sheetViews'), 'sheetView');
  const rowCount = Math.max(1, maxRow + 1);
  const columnCount = Math.max(1, maxColumn + 1);
  return {
    kind: 'worksheet',
    id: descriptor.id,
    name: descriptor.name,
    rowCount,
    columnCount,
    cells,
    merges,
    pane,
    pivots: [],
    sparklines: [],
    drawings: [],
    drawingPayloads: {},
    conditionalFormats,
    dataValidations,
    defaultRowHeightPx,
    defaultColumnWidthPx,
    rowHeightsPx,
    columnWidthsPx,
    hiddenRows: manualHiddenRows,
    hiddenColumns,
    tabColor,
    ...(hyperlinks.length ? { hyperlinks } : {}),
    notes,
    ...(autoFilter ? { autoFilter } : {}),
    ...(outline ? { outline } : {}),
    protectionRules,
    showGridlines: sheetView?.attrs.showGridLines !== '0' && sheetView?.attrs.showGridLines !== 'false',
    showHeaders: sheetView?.attrs.showRowColHeaders !== '0' && sheetView?.attrs.showRowColHeaders !== 'false',
    zoom: finitePositive(sheetView?.attrs.zoomScale, 100),
    ...(sheetTables.length ? { sheetTables } : {}),
    hidden: descriptor.hidden,
  };
}

function filterRowMatches(
  cells: Record<string, Record<string, CellData>>,
  filter: AutoFilterModel,
  row: number,
  dateSystem: DateSystem,
  canonicalReferenceDate?: CanonicalExcelDateParts,
): boolean {
  for (const column of Object.values(filter.columns)) {
    const criterion = column.criterion;
    if (!criterion) continue;
    const cell = cells[String(row)]?.[String(column.column)];
    const resolved = resolveFilterCellValue(cell, undefined, dateSystem);
    const value = resolved.value;
    const text = resolved.text;
    if (criterion.kind === 'values') {
      const valueMatches = criterion.values.some((candidate) => String(candidate ?? '').toLocaleLowerCase() === text.toLocaleLowerCase())
        || (text === '' && criterion.includeBlank);
      const dateMatches = (criterion.dateGroups ?? []).some((group) => dateMatchesGroup(toFilterDate(value, dateSystem), group));
      if (!valueMatches && !dateMatches) return false;
    } else if (criterion.kind === 'custom') {
      const results = criterion.conditions.filter((condition): condition is NonNullable<typeof condition> => Boolean(condition)).map((condition) => compareFilterText(text, String(condition.value ?? ''), condition.operator, value));
      if (criterion.join === 'and' ? !results.every(Boolean) : !results.some(Boolean)) return false;
    } else if (criterion.kind === 'dynamic') {
      if (!matchesDynamicFilter(toFilterDate(value, dateSystem), criterion.type, dateSystem, canonicalReferenceDate)) return false;
    } else if (criterion.kind === 'top10') {
      const numeric = Number(value);
      if (!Number.isFinite(numeric)) return false;
      const values = Array.from({ length: filter.range.endRow - filter.range.startRow }, (_, index) => {
        const candidate = resolveFilterCellValue(cells[String(filter.range.startRow + 1 + index)]?.[String(column.column)], undefined, dateSystem).value;
        return typeof candidate === 'number' ? candidate : Number.NaN;
      }).filter(Number.isFinite).sort((left, right) => criterion.top ? right - left : left - right);
      const requested = criterion.percent ? Math.max(1, Math.ceil(values.length * criterion.rank / 100)) : criterion.rank;
      const cutoff = values[Math.min(requested, values.length) - 1];
      if (cutoff === undefined || (criterion.top ? numeric < cutoff : numeric > cutoff)) return false;
    } else if (criterion.kind === 'color') {
      if (!(cell?.filterMetadata?.color?.dxfId === criterion.dxfId
        || (criterion.style && (criterion.target === 'cell' ? cell?.style?.background === criterion.style.background : cell?.style?.textColor === criterion.style.textColor)))) return false;
    } else if (criterion.kind === 'icon') {
      if (cell?.filterMetadata?.icon?.iconSet !== criterion.iconSet || cell.filterMetadata.icon.iconId !== criterion.iconId) return false;
    }
  }
  return true;
}

function compareFilterText(text: string, operand: string, operator: string, value: FilterScalar): boolean {
  const left = text.toLocaleLowerCase();
  const right = operand.toLocaleLowerCase();
  const numericLeft = typeof value === 'number' && Number.isFinite(value) ? value : Number.NaN;
  const numericRight = Number(operand);
  const numeric = Number.isFinite(numericLeft) && Number.isFinite(numericRight);
  switch (operator.toLocaleLowerCase()) {
    case 'equals': case '=': return left === right;
    case 'notequals': case 'not equal': case '<>': return left !== right;
    case 'lessthan': case '<': return numeric ? numericLeft < numericRight : left < right;
    case 'lessthanorequal': case '<=': return numeric ? numericLeft <= numericRight : left <= right;
    case 'greaterthan': case '>': return numeric ? numericLeft > numericRight : left > right;
    case 'greaterthanorequal': case '>=': return numeric ? numericLeft >= numericRight : left >= right;
    case 'contains': return left.includes(right);
    case 'notcontains': return !left.includes(right);
    case 'beginswith': return left.startsWith(right);
    case 'endswith': return left.endsWith(right);
    default: return false;
  }
}

function dateMatchesGroup(date: CanonicalExcelDate | null, group: { year: number; month?: number; day?: number; hour?: number; minute?: number; second?: number }): boolean {
  return Boolean(date) && date!.year === group.year
    && (group.month === undefined || date!.month === group.month)
    && (group.day === undefined || date!.day === group.day)
    && (group.hour === undefined || date!.hour === group.hour)
    && (group.minute === undefined || date!.minute === group.minute)
    && (group.second === undefined || date!.second === group.second);
}

function toFilterDate(value: unknown, dateSystem: DateSystem): CanonicalExcelDate | null {
  if (value === null || value === undefined || value === '') return null;
  if (typeof value === 'number' && !Number.isFinite(value)) throw new Error('Filter date serial must be finite');
  return canonicalExcelDateFromValue(value, dateSystem);
}

function matchesDynamicFilter(date: CanonicalExcelDate | null, type: import('@react-sheets/core-model').DynamicFilterType, dateSystem: DateSystem, canonicalReferenceDate?: CanonicalExcelDateParts): boolean {
  if (!date) return false;
  if (!canonicalReferenceDate) throw new Error('Dynamic date AutoFilter requires an explicit canonical workbook reference date');
  const today = canonicalExcelDateFromParts({ ...canonicalReferenceDate, hour: 0, minute: 0, second: 0, millisecond: 0 }, dateSystem);
  const monday = shiftCanonicalExcelDate(today, -((canonicalExcelDateDayOfWeek(today) + 6) % 7), dateSystem);
  const startOfMonth = monthBoundary(today, 0, dateSystem);
  const startOfQuarter = monthBoundary(today, -((today.month - 1) % 3), dateSystem);
  const startOfYear = monthBoundary(today, -(today.month - 1), dateSystem);
  const ranges: Record<import('@react-sheets/core-model').DynamicFilterType, [CanonicalExcelDate, CanonicalExcelDate]> = {
    today: [today, shiftCanonicalExcelDate(today, 1, dateSystem)], yesterday: [shiftCanonicalExcelDate(today, -1, dateSystem), today], tomorrow: [shiftCanonicalExcelDate(today, 1, dateSystem), shiftCanonicalExcelDate(today, 2, dateSystem)],
    thisWeek: [monday, shiftCanonicalExcelDate(monday, 7, dateSystem)], lastWeek: [shiftCanonicalExcelDate(monday, -7, dateSystem), monday], nextWeek: [shiftCanonicalExcelDate(monday, 7, dateSystem), shiftCanonicalExcelDate(monday, 14, dateSystem)],
    thisMonth: [startOfMonth, monthBoundary(startOfMonth, 1, dateSystem)], lastMonth: [monthBoundary(startOfMonth, -1, dateSystem), startOfMonth], nextMonth: [monthBoundary(startOfMonth, 1, dateSystem), monthBoundary(startOfMonth, 2, dateSystem)],
    thisQuarter: [startOfQuarter, monthBoundary(startOfQuarter, 3, dateSystem)], lastQuarter: [monthBoundary(startOfQuarter, -3, dateSystem), startOfQuarter], nextQuarter: [monthBoundary(startOfQuarter, 3, dateSystem), monthBoundary(startOfQuarter, 6, dateSystem)],
    thisYear: [startOfYear, monthBoundary(startOfYear, 12, dateSystem)], lastYear: [monthBoundary(startOfYear, -12, dateSystem), startOfYear], nextYear: [monthBoundary(startOfYear, 12, dateSystem), monthBoundary(startOfYear, 24, dateSystem)], yearToDate: [startOfYear, shiftCanonicalExcelDate(today, 1, dateSystem)],
  };
  const [from, to] = ranges[type];
  return date.serial >= from.serial && date.serial < to.serial;
}

function monthBoundary(value: CanonicalExcelDate, offset: number, dateSystem: DateSystem): CanonicalExcelDate {
  const date = canonicalExcelDateToUtcDate(value);
  date.setUTCDate(1);
  date.setUTCMonth(date.getUTCMonth() + offset);
  return canonicalExcelDateFromUtcDate(date, dateSystem);
}

function materializeFilterMetadata(
  cells: Record<string, Record<string, CellData>>,
  sheetId: string,
  filters: AutoFilterModel[],
  conditionalFormats: ConditionalFormatRule[],
  dateSystem: DateSystem,
): void {
  const inRange = (range: RangeRef, row: number, column: number): boolean => range.sheetId === sheetId
    && row >= range.startRow && row <= range.endRow && column >= range.startColumn && column <= range.endColumn;
  const numericValuesByRange = new Map<RangeRef, number[]>();
  for (const filter of filters) {
    for (const column of Object.values(filter.columns)) {
      const criterion = column.criterion;
      if (!criterion || (criterion.kind !== 'color' && criterion.kind !== 'icon')) continue;
      for (let row = filter.range.startRow + 1; row <= filter.range.endRow; row += 1) {
        const cell = cells[String(row)]?.[String(column.column)];
        if (!cell) continue;
        if (criterion.kind === 'color' && criterion.style) {
          const matches = criterion.target === 'cell'
            ? criterion.style.background !== undefined && cell.style?.background === criterion.style.background
            : criterion.style.textColor !== undefined && cell.style?.textColor === criterion.style.textColor;
          if (matches) cell.filterMetadata = { ...cell.filterMetadata, color: { target: criterion.target, dxfId: criterion.dxfId, value: criterion.target === 'cell' ? criterion.style.background : criterion.style.textColor } };
        }
        if (criterion.kind === 'icon') {
          const rule = conditionalFormats.find((candidate) => candidate.type === 'iconSet' && candidate.ranges.some((range) => inRange(range, row, column.column)));
          const numericValue = resolveFilterCellValue(cell, undefined, dateSystem).value;
          const numeric = typeof numericValue === 'number' ? numericValue : Number.NaN;
          if (rule && Number.isFinite(numeric)) {
            const ruleRange = rule.ranges.find((range) => inRange(range, row, column.column));
            let values: number[] = [];
            if (ruleRange) {
              values = numericValuesByRange.get(ruleRange) ?? numericValues(cells, ruleRange, dateSystem);
              numericValuesByRange.set(ruleRange, values);
            }
            const thresholds = resolveIconThresholds(rule.iconThresholds, values);
            const iconId = thresholds.reduce((identity, threshold, index) => numeric >= threshold ? index : identity, 0);
            cell.filterMetadata = { ...cell.filterMetadata, icon: { iconSet: criterion.iconSet, iconId } };
          }
        }
      }
    }
  }
}

function numericValues(cells: Record<string, Record<string, CellData>>, range: RangeRef, dateSystem: DateSystem): number[] {
  const values: number[] = [];
  for (const [rowKey, rowCells] of Object.entries(cells)) {
    const row = Number(rowKey);
    if (row < range.startRow || row > range.endRow) continue;
    for (const [columnKey, cell] of Object.entries(rowCells)) {
      const column = Number(columnKey);
      if (column < range.startColumn || column > range.endColumn) continue;
      const value = resolveFilterCellValue(cell, undefined, dateSystem).value;
      const numeric = typeof value === 'number' ? value : Number.NaN;
      if (Number.isFinite(numeric)) values.push(numeric);
    }
  }
  return values.sort((left, right) => left - right);
}

function resolveIconThresholds(
  definitions: Array<{ type: 'percent' | 'percentile' | 'num' | 'formula'; value?: number }> | undefined,
  values: number[],
): number[] {
  if (!values.length) return [];
  const min = values[0]!;
  const max = values.at(-1)!;
  const span = max - min || 1;
  return (definitions?.length ? definitions : [{ type: 'percent', value: 0 }, { type: 'percent', value: 33 }, { type: 'percent', value: 67 }]).map((definition) => {
    const value = definition.value ?? 0;
    if (definition.type === 'num' || definition.type === 'formula') return value;
    if (definition.type === 'percentile') {
      const index = Math.min(values.length - 1, Math.max(0, Math.ceil((value / 100) * values.length) - 1));
      return values[index]!;
    }
    return min + span * value / 100;
  });
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
    const field = pivot.fieldCatalog.fields.find((candidate) => candidate.fieldId === control.fieldId);
    if (!field) continue;
    const sourceKey = pivotSourceIdentity(pivot.source);
    const connections = (control.connectionPivotIds ?? []).filter((pivotId) => pivotId !== control.pivotId).flatMap((pivotId) => {
      const target = snapshot.sheets.flatMap((candidate) => candidate.pivots).find((candidate) => candidate.id === pivotId);
      const targetField = target?.fieldCatalog.fields.find((candidate) => candidate.ordinal === field.ordinal && candidate.name === field.name && candidate.dataType === field.dataType);
      return target && targetField && (control.kind !== 'timeline' || targetField.dataType === 'date') ? [{ pivotId, sourceKey, fieldId: targetField.fieldId }] : [];
    });
    if (control.kind === 'slicer') {
      const memberKeys = (control.selectedItemIndexes ?? []).map((index) => {
        const value = field?.values?.[index];
        if (!field || value === undefined) throw new Error(`Native Slicer ${control.id} selected item index ${index} is outside field ${control.fieldId} member domain`);
        return createPivotMemberKey(value);
      });
      sheet.drawingPayloads[payloadId] = {
        kind: 'slicer',
        pivotId: control.pivotId,
        fieldId: control.fieldId,
        filter: { mode: memberKeys.length ? 'include' : 'all', memberKeys },
        style,
        settings: { showHeader: true, caption: field?.name ?? 'Slicer', multiSelect: true, sort: 'ascending', showNoDataItems: true, noDataItemsLast: true, showNoDataStyle: true, columnCount: 1, itemHeight: 20 },
        ...(connections.length ? { connections } : {}),
      };
    } else {
      if (!control.level || !control.selectionLevel || control.showHeader === undefined || control.showSelectionLabel === undefined || control.showTimeLevel === undefined || control.showHorizontalScrollbar === undefined || !control.filterType) {
        throw new Error(`Native Timeline ${control.id} is missing canonical level/display state`);
      }
      sheet.drawingPayloads[payloadId] = {
        kind: 'timeline',
        pivotId: control.pivotId,
        fieldId: control.fieldId,
        period: control.selection ?? {},
        level: control.level,
        selectionLevel: control.selectionLevel,
        showHeader: control.showHeader,
        showSelectionLabel: control.showSelectionLabel,
        showTimeLevel: control.showTimeLevel,
        showHorizontalScrollbar: control.showHorizontalScrollbar,
        ...(control.scrollPosition === undefined ? {} : { scrollPosition: control.scrollPosition }),
        bounds: control.bounds ?? {},
        filterType: control.filterType,
        ...(control.caption === undefined ? {} : { caption: control.caption }),
        ...(control.styleName === undefined ? {} : { styleName: control.styleName }),
        style,
        ...(connections.length ? { connections } : {}),
      };
    }
    const drawing: DrawingObject = {
      id: control.id,
      sheetId: sheet.id,
      kind: control.kind,
      anchor: { kind: 'one-cell', row: control.drawingAnchor?.row ?? 0, column: control.drawingAnchor?.column ?? 0 },
      transform: { x: 0, y: 0, width: control.kind === 'slicer' ? 220 : 420, height: control.kind === 'slicer' ? 180 : 120 },
      zIndex: 0,
      payloadId,
    };
    sheet.drawings.push(drawing);
  }
}

function parseSheetTables(
  root: XmlNode,
  descriptor: SheetDescriptor,
  files: Record<string, Uint8Array>,
  pkg: OpcPackageGraph,
  styles: StyleContext,
): NonNullable<SheetSnapshot['sheetTables']> {
  return children(child(root, 'tableParts'), 'tablePart').flatMap((partNode) => {
    const relationId = partNode.attrs['r:id'] ?? partNode.attrs.id;
    if (!relationId) throw new Error(`Worksheet ${descriptor.part} tablePart is missing r:id`);
    const relation = (pkg.relationships[descriptor.part] ?? []).find((candidate) => candidate.id === relationId && isRelationshipKind(candidate.type, 'table'));
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
  const tableAutoFilter = parseAutoFilter(table, descriptor, styles);
    return [{
      id: `table-${tableNumber}`,
      sheetId: descriptor.id,
      name: table.attrs.displayName ?? table.attrs.name ?? `Table${table.attrs.id ?? '1'}`,
      range: { ...range, sheetId: descriptor.id },
      hasHeaderRow: table.attrs.headerRowCount !== '0',
      hasTotalRow: table.attrs.totalsRowCount === '1',
      showBandedRows: child(table, 'tableStyleInfo')?.attrs.showRowStripes !== '0',
      showBandedColumns: child(table, 'tableStyleInfo')?.attrs.showColumnStripes === '1',
      showFirstColumn: child(table, 'tableStyleInfo')?.attrs.showFirstColumn === '1',
      showLastColumn: child(table, 'tableStyleInfo')?.attrs.showLastColumn === '1',
      showFilterButton: table.attrs.headerRowCount !== '0',
      autoExpand: 'both',
      ...(tableAutoFilter ? { autoFilter: tableAutoFilter } : {}),
      columns,
      ...(child(table, 'tableStyleInfo')?.attrs.name ? { styleName: child(table, 'tableStyleInfo')!.attrs.name } : {}),
    }];
  });
}

function parseSharedStrings(bytes: Uint8Array | undefined, themeColors: string[]): SharedStringRecord[] {
  if (!bytes) return [];
  const root = firstElement(parseXml(strFromU8(bytes)), 'sst');
  return children(root, 'si').map((item) => parseRichTextContainer(item, themeColors));
}

function parseStyles(
  bytes: Uint8Array | undefined,
  themeBytes: Uint8Array | undefined,
  measurer: OoxmlFontMeasurer,
): StyleContext {
  const fallbackFont = { family: DEFAULT_EXCEL_FONT_FAMILY, sizePt: DEFAULT_EXCEL_FONT_SIZE_PT };
  if (!bytes) {
    return { records: [{}], namedCellStyles: [], differentialStyles: [], normalFont: fallbackFont, maximumDigitWidthPx: measurer.maximumDigitWidthPx(fallbackFont), themeColors: [] };
  }
  const root = firstElement(parseXml(strFromU8(bytes)), 'styleSheet');
  const themeColors = parseThemeColors(themeBytes);
  const customFormats = new Map<number, string>();
  for (const node of children(child(root, 'numFmts'), 'numFmt')) {
    const id = Number(node.attrs.numFmtId);
    const format = node.attrs.formatCode;
    if (Number.isFinite(id) && format) customFormats.set(id, format);
  }
  const fonts = children(child(root, 'fonts'), 'font');
  const fills = children(child(root, 'fills'), 'fill');
  const borders = children(child(root, 'borders'), 'border');
  const baseXfs = children(child(root, 'cellStyleXfs'), 'xf');
  const xfs = children(child(root, 'cellXfs'), 'xf');
  const normalStyle = children(child(root, 'cellStyles'), 'cellStyle').find((style) => (style.attrs.name ?? '').toLocaleLowerCase() === 'normal');
  const normalBaseXf = baseXfs[Number(normalStyle?.attrs.xfId ?? 0)] ?? baseXfs[0];
  const normalFontNode = fonts[Number(normalBaseXf?.attrs.fontId ?? 0)] ?? fonts[0];
  const normalFont: OoxmlNormalFont = {
    family: normalizeFontFamily(child(normalFontNode, 'name')?.attrs.val ?? fallbackFont.family),
    sizePt: finitePositive(child(normalFontNode, 'sz')?.attrs.val, fallbackFont.sizePt),
  };
  const records = xfs.map((xf) => {
    const base = baseXfs[Number(xf.attrs.xfId ?? 0)] ?? normalBaseXf;
    const fontId = resolvedXfId(xf, base, 'fontId', 'applyFont');
    const fillId = resolvedXfId(xf, base, 'fillId', 'applyFill');
    const borderId = resolvedXfId(xf, base, 'borderId', 'applyBorder');
    const numberFormatId = resolvedXfId(xf, base, 'numFmtId', 'applyNumberFormat');
    const fontStyle = parseFontStyle(fonts[fontId], themeColors);
    const background = parseFillColor(fills[fillId], themeColors);
    const cellBorders = parseBorders(borders[borderId], themeColors);
    const baseAlignment = child(base, 'alignment');
    const alignment = xf.attrs.applyAlignment === '0' ? baseAlignment : child(xf, 'alignment') ?? baseAlignment;
    const baseProtection = child(base, 'protection');
    const protection = xf.attrs.applyProtection === '0' ? baseProtection : child(xf, 'protection') ?? baseProtection;
    const style: CellStyle = {
      ...fontStyle,
      ...(background ? { background } : {}),
      ...(cellBorders ? { borders: cellBorders } : {}),
      ...parseAlignmentAttributes(alignment),
      ...(protection?.attrs.locked !== undefined ? { locked: xmlBoolean(protection.attrs.locked) } : {}),
      ...(protection?.attrs.hidden !== undefined ? { formulaHidden: xmlBoolean(protection.attrs.hidden) } : {}),
    };
    const numberFormat = customFormats.get(numberFormatId) ?? builtInNumberFormat(numberFormatId);
    return { numberFormat, style: Object.keys(style).length ? style : undefined };
  });
  const namedCellStyles = children(child(root, 'cellStyles'), 'cellStyle').flatMap((named, index) => {
    const name = named.attrs.name?.trim() ?? '';
    if (!name || name.toLocaleLowerCase() === 'normal') return [];
    const xf = baseXfs[Number(named.attrs.xfId ?? 0)];
    if (!xf) return [];
    const fontId = Number(xf.attrs.fontId ?? 0) || 0;
    const fillId = Number(xf.attrs.fillId ?? 0) || 0;
    const borderId = Number(xf.attrs.borderId ?? 0) || 0;
    const numberFormatId = Number(xf.attrs.numFmtId ?? 0) || 0;
    const alignment = child(xf, 'alignment');
    const style: CellStyle = {
      ...parseFontStyle(fonts[fontId], themeColors),
      ...(parseFillColor(fills[fillId], themeColors) ? { background: parseFillColor(fills[fillId], themeColors)! } : {}),
      ...(parseBorders(borders[borderId], themeColors) ? { borders: parseBorders(borders[borderId], themeColors)! } : {}),
      ...parseAlignmentAttributes(alignment),
    };
    const numberFormat = customFormats.get(numberFormatId) ?? builtInNumberFormat(numberFormatId);
    return [{ id: `ooxml-cell-style-${index + 1}`, name, style: { ...style, ...(numberFormat ? { numberFormat } : {}) } } satisfies CellStyleTemplate];
  });
  const differentialStyles = children(child(root, 'dxfs'), 'dxf').map((dxf) => {
    const style: CellStyle = {
      ...parseFontStyle(child(dxf, 'font'), themeColors),
      ...(parseFillColor(child(dxf, 'fill'), themeColors) ? { background: parseFillColor(child(dxf, 'fill'), themeColors)! } : {}),
      ...(parseBorders(child(dxf, 'border'), themeColors) ? { borders: parseBorders(child(dxf, 'border'), themeColors)! } : {}),
      ...parseAlignmentAttributes(child(dxf, 'alignment')),
    };
    return Object.keys(style).length ? style : undefined;
  });
  return {
    records: records.length ? records : [{}],
    namedCellStyles,
    differentialStyles,
    normalFont,
    maximumDigitWidthPx: measurer.maximumDigitWidthPx(normalFont),
    themeColors,
  };
}

function resolvedXfId(xf: XmlNode, base: XmlNode | undefined, key: string, applyKey: string): number {
  const own = xf.attrs[key];
  if (own !== undefined && xf.attrs[applyKey] !== '0' && xf.attrs[applyKey] !== 'false') return Number(own) || 0;
  return Number(base?.attrs[key] ?? own ?? 0) || 0;
}

function parseRichTextContainer(container: XmlNode, themeColors: string[] = []): SharedStringRecord {
  const runs = children(container, 'r');
  if (!runs.length) return { value: descendants(container, 't').map(textContent).join('') };
  const richText = runs.map((run) => {
    const properties = child(run, 'rPr');
    const known = new Set(['rFont', 'name', 'sz', 'b', 'i', 'u', 'strike', 'color']);
    const preservedProperties = properties?.children
      .map((property) => localName(property.name))
      .filter((name) => !known.has(name)) ?? [];
    const style = parseFontStyle(properties, themeColors);
    return {
      text: descendants(run, 't').map(textContent).join(''),
      ...(Object.keys(style).length ? { style } : {}),
      ...(preservedProperties.length ? { preservedProperties } : {}),
    } satisfies RichTextRun;
  });
  return { value: richText.map((run) => run.text).join(''), richText };
}

function parseFontStyle(font: XmlNode | undefined, themeColors: string[]): CellStyle {
  if (!font) return {};
  const sizePt = Number(child(font, 'sz')?.attrs.val);
  const textColor = resolveColor(child(font, 'color'), themeColors);
  const rawFamily = child(font, 'name')?.attrs.val ?? child(font, 'rFont')?.attrs.val;
  return {
    ...(rawFamily === undefined ? {} : { fontFamily: normalizeFontFamily(rawFamily) }),
    ...(Number.isFinite(sizePt) && sizePt > 0 ? { fontSizePx: pointsToPixels(sizePt) } : {}),
    ...(child(font, 'b') ? { bold: xmlBoolean(child(font, 'b')?.attrs.val ?? '1') } : {}),
    ...(child(font, 'i') ? { italic: xmlBoolean(child(font, 'i')?.attrs.val ?? '1') } : {}),
    ...(child(font, 'u') ? { underline: xmlBoolean(child(font, 'u')?.attrs.val ?? '1') } : {}),
    ...(child(font, 'strike') ? { strikethrough: xmlBoolean(child(font, 'strike')?.attrs.val ?? '1') } : {}),
    ...(textColor ? { textColor } : {}),
  };
}

function parseFillColor(fill: XmlNode | undefined, themeColors: string[]): string | undefined {
  const pattern = child(fill, 'patternFill');
  if (!pattern || pattern.attrs.patternType === 'none' || pattern.attrs.patternType === 'gray125') return undefined;
  return resolveColor(child(pattern, 'fgColor') ?? child(pattern, 'bgColor'), themeColors);
}

function parseBorders(border: XmlNode | undefined, themeColors: string[]): CellStyle['borders'] | undefined {
  if (!border) return undefined;
  const result: NonNullable<CellStyle['borders']> = {};
  for (const side of ['top', 'right', 'bottom', 'left'] as const) {
    const node = child(border, side);
    const style = normalizeBorderStyle(node?.attrs.style);
    if (!style) continue;
    result[side] = { style, color: resolveColor(child(node, 'color'), themeColors) ?? '#000000' };
  }
  return Object.keys(result).length ? result : undefined;
}

function normalizeBorderStyle(value: string | undefined): NonNullable<NonNullable<CellStyle['borders']>['top']>['style'] | undefined {
  if (!value || value === 'none') return undefined;
  if (value === 'double') return 'double';
  if (value === 'thick') return 'thick';
  if (value.startsWith('medium')) return 'medium';
  if (value === 'dashed' || value === 'dotted' || value === 'dashDot' || value === 'dashDotDot' || value === 'slantDashDot') return 'dashed';
  return 'thin';
}

const INDEXED_COLORS = [
  '#000000', '#FFFFFF', '#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF',
  '#000000', '#FFFFFF', '#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF',
  '#800000', '#008000', '#000080', '#808000', '#800080', '#008080', '#C0C0C0', '#808080',
  '#9999FF', '#993366', '#FFFFCC', '#CCFFFF', '#660066', '#FF8080', '#0066CC', '#CCCCFF',
  '#000080', '#FF00FF', '#FFFF00', '#00FFFF', '#800080', '#800000', '#008080', '#0000FF',
  '#00CCFF', '#CCFFFF', '#CCFFCC', '#FFFF99', '#99CCFF', '#FF99CC', '#CC99FF', '#FFCC99',
  '#3366FF', '#33CCCC', '#99CC00', '#FFCC00', '#FF9900', '#FF6600', '#666699', '#969696',
  '#003366', '#339966', '#003300', '#333300', '#993300', '#993366', '#333399', '#333333',
] as const;

function parseThemeColors(bytes: Uint8Array | undefined): string[] {
  if (!bytes) return [];
  const root = parseXml(strFromU8(bytes));
  const scheme = descendants(root, 'clrScheme')[0];
  if (!scheme) return [];
  return scheme.children.map((entry) => {
    const color = entry.children[0];
    const value = color?.attrs.val ?? color?.attrs.lastClr;
    return normalizeRgb(value) ?? '#000000';
  });
}

function resolveColor(node: XmlNode | undefined, themeColors: string[]): string | undefined {
  if (!node) return undefined;
  let color = normalizeRgb(node.attrs.rgb);
  if (!color && node.attrs.theme !== undefined) color = themeColors[Number(node.attrs.theme)];
  if (!color && node.attrs.indexed !== undefined) color = INDEXED_COLORS[Number(node.attrs.indexed)];
  if (!color && xmlBoolean(node.attrs.auto ?? '0')) color = '#000000';
  if (!color) return undefined;
  const tint = Number(node.attrs.tint);
  return Number.isFinite(tint) && tint !== 0 ? applyTint(color, tint) : color;
}

function normalizeRgb(value: string | undefined): string | undefined {
  if (!value) return undefined;
  const normalized = value.replace(/^#/, '').trim();
  if (/^[0-9a-f]{8}$/i.test(normalized)) return `#${normalized.slice(2).toUpperCase()}`;
  if (/^[0-9a-f]{6}$/i.test(normalized)) return `#${normalized.toUpperCase()}`;
  return undefined;
}

function applyTint(color: string, tint: number): string {
  const rgb = [1, 3, 5].map((offset) => Number.parseInt(color.slice(offset, offset + 2), 16) / 255);
  const max = Math.max(...rgb);
  const min = Math.min(...rgb);
  let h = 0;
  let s = 0;
  let l = (max + min) / 2;
  const delta = max - min;
  if (delta !== 0) {
    s = l > 0.5 ? delta / (2 - max - min) : delta / (max + min);
    if (max === rgb[0]) h = ((rgb[1]! - rgb[2]!) / delta + (rgb[1]! < rgb[2]! ? 6 : 0)) / 6;
    else if (max === rgb[1]) h = ((rgb[2]! - rgb[0]!) / delta + 2) / 6;
    else h = ((rgb[0]! - rgb[1]!) / delta + 4) / 6;
  }
  l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
  const hue = (p: number, q: number, value: number): number => {
    let t = value;
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  const channels = s === 0
    ? [l, l, l]
    : (() => { const q = l < 0.5 ? l * (1 + s) : l + s - l * s; const p = 2 * l - q; return [hue(p, q, h + 1 / 3), hue(p, q, h), hue(p, q, h - 1 / 3)]; })();
  return `#${channels.map((channel) => Math.round(Math.max(0, Math.min(1, channel)) * 255).toString(16).padStart(2, '0')).join('').toUpperCase()}`;
}

function buildSharedStrings(snapshot: WorkbookSnapshot): string {
  const values: Array<{ value: string; richText?: RichTextRun[] }> = [];
  const lookup = new Map<string, number>();
  for (const sheet of snapshot.sheets) {
    for (const row of Object.values(sheet.cells)) {
      for (const cell of Object.values(row)) {
        if (cell.formula || typeof cell.value !== 'string') continue;
        const key = JSON.stringify({ value: cell.value, richText: cell.richText });
        if (!lookup.has(key)) {
          lookup.set(key, values.length);
          values.push({ value: cell.value, richText: cell.richText });
        }
      }
    }
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="${NS_MAIN}" count="${values.length}" uniqueCount="${values.length}">${values.map((value) => `<si>${serializeRichText(value.value, value.richText)}</si>`).join('')}</sst>`;
}

function buildStyles(snapshot: WorkbookSnapshot, originalStylesXml?: string): string {
  const records: StyleRecord[] = [{}];
  const templateRecords: StyleRecord[] = (snapshot.cellStyleTemplates ?? []).map((template) => ({ style: structuredClone(template.style), numberFormat: template.style.numberFormat }));
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
  const custom = collectCustomNumberFormatIds(snapshot);
  const numFmts = [...custom.entries()].map(([code, id]) => `<numFmt numFmtId="${id}" formatCode="${encodeXml(code)}"/>`).join('');
  const fontIndexes = new Map<string, number>();
  const fillIndexes = new Map<string, number>();
  const borderIndexes = new Map<string, number>();
  const fontRecords = ['<font><sz val="11"/><name val="Calibri"/></font>'];
  const fillRecords = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
  const borderRecords = ['<border><left/><right/><top/><bottom/><diagonal/></border>'];
  for (const record of [...records, ...templateRecords]) {
    const style = record.style;
    const fontKey = JSON.stringify({ fontFamily: style?.fontFamily, fontSizePx: style?.fontSizePx, bold: style?.bold, italic: style?.italic, underline: style?.underline, strikethrough: style?.strikethrough, textColor: style?.textColor });
    if (!fontIndexes.has(fontKey) && style) {
      const index = fontRecords.length;
      fontIndexes.set(fontKey, index);
      fontRecords.push(`<font><sz val="${roundMetric(pixelsToPoints(style.fontSizePx ?? pointsToPixels(11)))}"/><name val="${encodeXml(style.fontFamily ?? 'Calibri')}"/>${style.bold ? '<b/>' : ''}${style.italic ? '<i/>' : ''}${style.underline ? '<u/>' : ''}${style.strikethrough ? '<strike/>' : ''}${style.textColor ? `<color rgb="${ooxmlRgb(style.textColor)}"/>` : ''}</font>`);
    }
    const fillKey = style?.background ?? '';
    if (fillKey && !fillIndexes.has(fillKey)) {
      const index = fillRecords.length;
      fillIndexes.set(fillKey, index);
      fillRecords.push(`<fill><patternFill patternType="solid"><fgColor rgb="${ooxmlRgb(fillKey)}"/><bgColor indexed="64"/></patternFill></fill>`);
    }
    const borderKey = JSON.stringify(style?.borders ?? {});
    if (style?.borders && !borderIndexes.has(borderKey)) {
      const index = borderRecords.length;
      borderIndexes.set(borderKey, index);
      borderRecords.push(serializeBorders(style.borders));
    }
  }
  const serializeXf = (record: StyleRecord, xfId = 0) => {
    const style = record.style;
    const numFmtId = record.numberFormat ? (builtInNumberFormatId(record.numberFormat) ?? custom.get(record.numberFormat) ?? 0) : 0;
    const fontKey = JSON.stringify({ fontFamily: style?.fontFamily, fontSizePx: style?.fontSizePx, bold: style?.bold, italic: style?.italic, underline: style?.underline, strikethrough: style?.strikethrough, textColor: style?.textColor });
    const fontId = style ? (fontIndexes.get(fontKey) ?? 0) : 0;
    const fillId = style?.background ? (fillIndexes.get(style.background) ?? 0) : 0;
    const borderId = style?.borders ? (borderIndexes.get(JSON.stringify(style.borders)) ?? 0) : 0;
    const attrs = [`numFmtId="${numFmtId}"`, `fontId="${fontId}"`, `fillId="${fillId}"`, `borderId="${borderId}"`, `xfId="${xfId}"`, 'applyFont="1"', 'applyFill="1"', 'applyBorder="1"', 'applyNumberFormat="1"'];
    const alignment = style ? serializeAlignment(style) : '';
    const protection = style && (style.locked !== undefined || style.formulaHidden !== undefined)
      ? `<protection${style.locked !== undefined ? ` locked="${style.locked ? '1' : '0'}"` : ''}${style.formulaHidden !== undefined ? ` hidden="${style.formulaHidden ? '1' : '0'}"` : ''}/>`
      : '';
    return `<xf ${attrs.join(' ')}${alignment || protection ? ` applyAlignment="${alignment ? '1' : '0'}" applyProtection="${protection ? '1' : '0'}">${alignment}${protection}</xf>` : '/>'}`;
  };
  const xfs = records.map((record) => serializeXf(record)).join('');
  const templateXfs = templateRecords.map((record) => serializeXf(record)).join('');
  const differentialStyles = [...collectDifferentialStyleIndexes(snapshot).keys()].map((key) => `<dxf>${serializeDifferentialStyle(JSON.parse(key) as CellStyle)}</dxf>`).join('');
  const originalRoot = originalStylesXml ? firstElement(parseXml(originalStylesXml), 'styleSheet') : undefined;
  const originalTableStyles = originalRoot ? child(originalRoot, 'tableStyles') : undefined;
  const tableStyles = originalTableStyles ? serializeXml(originalTableStyles) : '';
  const cellStyleXfs = `<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>${templateXfs}`;
  const cellStyles = `<cellStyle name="Normal" builtinId="0"/>${(snapshot.cellStyleTemplates ?? []).map((template, index) => `<cellStyle name="${encodeXml(template.name)}" xfId="${index + 1}"/>`).join('')}`;
  return `<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="${NS_MAIN}"><numFmts count="${custom.size}">${numFmts}</numFmts><fonts count="${fontRecords.length}">${fontRecords.join('')}</fonts><fills count="${fillRecords.length}">${fillRecords.join('')}</fills><borders count="${borderRecords.length}">${borderRecords.join('')}</borders><cellStyleXfs count="${templateRecords.length + 1}">${cellStyleXfs}</cellStyleXfs><cellXfs count="${records.length}">${xfs}</cellXfs><cellStyles count="${templateRecords.length + 1}">${cellStyles}</cellStyles><dxfs count="${collectDifferentialStyleIndexes(snapshot).size}">${differentialStyles}</dxfs>${tableStyles}</styleSheet>`;
}

function prepareTableParts(
  sheet: SheetSnapshot,
  sourcePart: string,
  preserved: OpcPackageGraph | undefined,
  files: Map<string, Uint8Array>,
  differentialStyleIndexes: Map<string, number>,
): { parts: Map<string, string>; required: Array<Pick<XlsxRelationship, 'type' | 'target'>> } {
  const tables = sheet.sheetTables ?? [];
  if (!tables.length) return { parts: new Map(), required: [] };
  const existing = (preserved?.relationships[sourcePart] ?? []).filter((relation) => isRelationshipKind(relation.type, 'table'));
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
    parts.set(part, buildTableXml(table, index + 1, differentialStyleIndexes));
    required.push({ type: `${NS_DOC_REL}/table`, target: relativeTarget(sourcePart, part) });
  }
  return { parts, required };
}

function buildTableXml(table: NonNullable<SheetSnapshot['sheetTables']>[number], fallbackId: number, differentialStyleIndexes: Map<string, number>): string {
  const numericId = Number(table.id.replace(/\D/g, '')) || fallbackId;
  const ref = rangeToA1(table.range);
  const columns = table.columns.map((column, index) => `<tableColumn id="${index + 1}" name="${encodeXml(column.name)}"${column.totalsFunction && column.totalsFunction !== 'none' ? ` totalsRowFunction="${encodeXml(column.totalsFunction)}"` : ''}/>`).join('');
  const style = table.styleName
    ? `<tableStyleInfo name="${encodeXml(table.styleName)}" showFirstColumn="${table.showFirstColumn ? '1' : '0'}" showLastColumn="${table.showLastColumn ? '1' : '0'}" showRowStripes="${table.showBandedRows ? '1' : '0'}" showColumnStripes="${table.showBandedColumns ? '1' : '0'}"/>`
    : '';
  const autoFilter = table.autoFilter
    ? serializeAutoFilter(table.autoFilter, differentialStyleIndexes)
    : table.showFilterButton && table.hasHeaderRow ? `<autoFilter ref="${encodeXml(ref)}"/>` : '';
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

function collectDifferentialStyleIndexes(snapshot: WorkbookSnapshot): Map<string, number> {
  const indexes = new Map<string, number>();
  for (const sheet of snapshot.sheets) {
    for (const rule of sheet.conditionalFormats ?? []) {
      if (!rule.style) continue;
      const key = JSON.stringify(rule.style);
      if (!indexes.has(key)) indexes.set(key, indexes.size);
    }
    const filters = [sheet.autoFilter, ...(sheet.sheetTables ?? []).map((table) => table.autoFilter)].filter((filter): filter is AutoFilterModel => Boolean(filter));
    for (const filter of filters) {
      for (const column of Object.values(filter.columns)) {
        const criterion = column.criterion;
        if (criterion?.kind !== 'color' || !criterion.style) continue;
        const key = JSON.stringify(criterion.style);
        if (!indexes.has(key)) indexes.set(key, indexes.size);
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
  differentialStyleIndexes: Map<string, number>,
  maximumDigitWidthPx: number,
  includeCachedValues: boolean,
  dateSystem: DateSystem,
  nativeDisplayCells?: Record<string, Record<string, CellData>>,
  nativeControls: NativePivotControlDefinition[] = [],
  printDocument?: NonNullable<WorkbookSnapshot['printDocuments']>[number],
  sheetNames: ReadonlyMap<string, string> = new Map(),
): string {
  validateOoxmlExchangeBoundary(sheet);
  let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_DOC_REL}">`;
  if (sheet.tabColor || sheet.outline?.groups.length) xml += `<sheetPr>${sheet.tabColor ? `<tabColor rgb="${ooxmlRgb(sheet.tabColor)}"/>` : ''}${sheet.outline?.groups.length ? '<outlinePr summaryBelow="1" summaryRight="1"/>' : ''}</sheetPr>`;
  const dimension = inferDimension(sheet);
  if (dimension) xml += `<dimension ref="${dimension}"/>`;
  const pane = sheet.pane;
  const paneXml = pane.kind === 'none' ? '' : `<pane xSplit="${roundMetric(pane.xSplit)}" ySplit="${roundMetric(pane.ySplit)}" topLeftCell="${columnToLetter(pane.startColumn)}${pane.startRow + 1}" activePane="${pane.activePane ?? 'bottomRight'}"${pane.kind === 'frozen' ? ` state="${pane.state ?? 'frozen'}"` : ''}/>`;
  xml += `<sheetViews><sheetView workbookViewId="0" showGridLines="${sheet.showGridlines === false ? '0' : '1'}" showRowColHeaders="${sheet.showHeaders === false ? '0' : '1'}" zoomScale="${sheet.zoom ?? 100}">${paneXml}</sheetView></sheetViews>`;
  xml += `<sheetFormatPr baseColWidth="8" defaultColWidth="${roundMetric(pixelsToExcelColumnWidth(sheet.defaultColumnWidthPx, maximumDigitWidthPx))}" defaultRowHeight="${roundMetric(pixelsToPoints(sheet.defaultRowHeightPx))}"/>`;
  const outlinedColumns = sheet.outline?.groups.filter((group) => group.axis === 'column').flatMap((group) => Array.from({ length: group.end - group.start + 1 }, (_, offset) => group.start + offset)) ?? [];
  const columnIndexes = [...new Set([...Object.keys(sheet.columnWidthsPx ?? {}).map(Number), ...(sheet.hiddenColumns ?? []), ...outlinedColumns])].sort((a, b) => a - b);
  if (columnIndexes.length) {
    xml += `<cols>${columnIndexes.map((column) => {
      const width = sheet.columnWidthsPx?.[column];
      const columnOutline = outlineLevelAt(sheet.outline, 'column', column);
      return `<col min="${column + 1}" max="${column + 1}"${width === undefined ? '' : ` width="${roundMetric(pixelsToExcelColumnWidth(width, maximumDigitWidthPx))}" customWidth="1"`}${sheet.hiddenColumns?.includes(column) ? ' hidden="1"' : ''}${columnOutline ? ` outlineLevel="${columnOutline.level}"${columnOutline.collapsed ? ' collapsed="1"' : ''}` : ''}/>`;
    }).join('')}</cols>`;
  }
  xml += '<sheetData>';
  const outlinedRows = sheet.outline?.groups.filter((group) => group.axis === 'row').flatMap((group) => Array.from({ length: group.end - group.start + 1 }, (_, offset) => group.start + offset)) ?? [];
  const rowKeys = [...new Set([...Object.keys(sheet.cells), ...Object.keys(nativeDisplayCells ?? {}), ...Object.keys(sheet.rowHeightsPx ?? {}), ...(sheet.hiddenRows ?? []), ...outlinedRows])].map(Number).sort((a, b) => a - b);
  for (const row of rowKeys) {
    const cells = { ...(nativeDisplayCells?.[String(row)] ?? {}), ...(sheet.cells[String(row)] ?? {}) };
    const columns = Object.keys(cells).map(Number).sort((a, b) => a - b);
    const hidden = sheet.hiddenRows?.includes(row) ? ' hidden="1"' : '';
    const height = sheet.rowHeightsPx?.[row];
    const rowOutline = outlineLevelAt(sheet.outline, 'row', row);
    xml += `<row r="${row + 1}"${hidden}${height === undefined ? '' : ` ht="${roundMetric(pixelsToPoints(height))}" customHeight="1"`}${rowOutline ? ` outlineLevel="${rowOutline.level}"${rowOutline.collapsed ? ' collapsed="1"' : ''}` : ''}>`;
    for (const column of columns) {
      const cell = cells[String(column)];
      if (!cell) continue;
      xml += buildCellXml(cell, row, column, styleIndexes, includeCachedValues, dateSystem);
    }
    xml += '</row>';
  }
  xml += '</sheetData>';
  if (sheet.protectionRules?.length) xml += serializeProtection(sheet.protectionRules);
  if (sheet.autoFilter) xml += serializeAutoFilter(sheet.autoFilter, differentialStyleIndexes);
  if (sheet.merges.length) {
    xml += `<mergeCells count="${sheet.merges.length}">${sheet.merges.map((merge) => `<mergeCell ref="${rangeToA1(merge.range)}"/>`).join('')}</mergeCells>`;
  }
  if (sheet.conditionalFormats?.length) xml += serializeConditionalFormats(sheet.conditionalFormats, differentialStyleIndexes);
  if (sheet.dataValidations?.length) xml += serializeDataValidations(sheet.dataValidations);
  const hyperlinks = (sheet.hyperlinks ?? []).map((entry) => {
    const link = entry.hyperlink;
    const address = `${columnToLetter(entry.column)}${entry.row + 1}`;
    const relation = relationships.find((candidate) => isRelationshipKind(candidate.type, 'hyperlink') && candidate.target === hyperlinkTarget(link));
    const target = link.target.kind === 'url' || link.target.kind === 'email';
    let location: string | undefined;
    if (link.target.kind === 'sheet') {
      const targetSheetName = sheetNames.get(link.target.sheetId);
      if (!targetSheetName) throw new Error(`Hyperlink target sheet not found: ${link.target.sheetId}`);
      location = `${encodeXml(excelSheetName(targetSheetName))}!${encodeXml(link.target.address ?? (link.target.row !== undefined && link.target.column !== undefined ? `${columnToLetter(link.target.column)}${link.target.row + 1}` : 'A1'))}`;
    } else if (link.target.kind === 'name') {
      location = encodeXml(link.target.name);
    }
    return `<hyperlink ref="${address}"${target && relation ? ` r:id="${relation.id}"` : ''}${location ? ` location="${location}"` : ''}${link.tooltip ? ` tooltip="${encodeXml(link.tooltip)}"` : ''}/>`;
  });
  if (hyperlinks.length) xml += `<hyperlinks>${hyperlinks.join('')}</hyperlinks>`;
  if (printDocument) xml += serializePrintDocument(printDocument);
  const tableRelations = relationships.filter((relation) => isRelationshipKind(relation.type, 'table'));
  if (sheet.sheetTables?.length && tableRelations.length) {
    const tableParts = tableRelations.map((relation) => {
      const target = resolveTarget(sourcePart, relation.target);
      if (!files.has(target)) throw new Error(`Worksheet table relationship points to missing part: ${target}`);
      return `<tablePart r:id="${encodeXml(relation.id)}"/>`;
    });
    xml += `<tableParts count="${tableParts.length}">${tableParts.join('')}</tableParts>`;
  }
  const pivotRelations = relationships.filter((relation) => isRelationshipKind(relation.type, 'pivotTable'));
  if (pivotRelations.length) {
    const pivotParts = pivotRelations.map((relation) => {
      const target = resolveTarget(sourcePart, relation.target);
      if (!files.has(target)) throw new Error(`Worksheet PivotTable relationship points to missing part: ${target}`);
      return `<pivotTablePart r:id="${encodeXml(relation.id)}"/>`;
    });
    xml += `<pivotTableParts count="${pivotParts.length}">${pivotParts.join('')}</pivotTableParts>`;
  }
  const drawingRelations = relationships.filter((relation) => isRelationshipKind(relation.type, 'drawing'));
  if (drawingRelations.length && !child(originalRoot, 'drawing')) xml += `<drawing r:id="${encodeXml(drawingRelations[0]!.id)}"/>`;
  // A drawing or other unsupported worksheet node is still emitted through its
  // original relationship. Native Pivot parts are rebuilt above from the
  // canonical graph and therefore never copied from a stale worksheet node.
  if (originalRoot) {
    const preservedNodes = new Map<string, XmlNode>();
    for (const node of originalRoot.children) {
      const name = localName(node.name);
      if (name === 'drawing' && node.attrs['r:id'] && !drawingRelations.some((relation) => relation.id === node.attrs['r:id'])) continue;
      if (name === 'drawing' || name === 'legacyDrawing' || name === 'oleObjects' || name === 'controls' || name === 'extLst' || name === 'picture' || (name === 'tableParts' && !sheet.sheetTables?.length)) preservedNodes.set(name, node);
    }
    for (const name of ['drawing', 'legacyDrawing', 'oleObjects', 'controls', 'picture', 'tableParts']) {
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

function buildCellXml(cell: CellData, row: number, column: number, styleIndexes: Map<string, number>, includeCachedValues: boolean, dateSystem: DateSystem): string {
  const ref = `${columnToLetter(column)}${row + 1}`;
  const styleKey = JSON.stringify({ style: cell.style, numberFormat: cell.numberFormat });
  const style = cell.style || cell.numberFormat ? (styleIndexes.get(styleKey) ?? 1) : undefined;
  const styleAttr = style === undefined ? '' : ` s="${style}"`;
  const metadata = cell.formulaMetadata;
  if (cell.formula || metadata?.preservedOnly) {
    const sourceFormula = metadata?.sourceFormula ?? cell.formula ?? '';
    const formula = sourceFormula.startsWith('=') ? sourceFormula.slice(1) : sourceFormula;
    const cachedValue = isScalar(cell.formulaValue) ? cell.formulaValue : isScalar(cell.value) ? cell.value : null;
    const cachedSerial = typeof cachedValue === 'string' && isExcelDateFormat(cell.numberFormat) ? canonicalDateToSerial(cachedValue, dateSystem) : undefined;
    const serializedCachedValue = cachedSerial ?? cachedValue;
    const cached = includeCachedValues && serializedCachedValue !== null ? `<v>${encodeXml(typeof serializedCachedValue === 'boolean' ? (serializedCachedValue ? '1' : '0') : String(serializedCachedValue))}</v>` : '';
    const cachedType = typeof serializedCachedValue === 'boolean' ? 'b' : typeof serializedCachedValue === 'string' ? 'str' : undefined;
    const formulaAttrs = metadata?.kind === 'shared' && metadata.preservedOnly
      ? ` t="shared"${metadata.sharedIndex !== undefined ? ` si="${metadata.sharedIndex}"` : ''}${metadata.sharedMaster && metadata.range ? ` ref="${encodeXml(metadata.range)}"` : ''}`
      : metadata?.kind === 'array'
      ? ` t="array"${metadata.range ? ` ref="${encodeXml(metadata.range)}"` : ''}`
      : metadata?.kind === 'dataTable'
        ? ` t="dataTable"${metadata.range ? ` ref="${encodeXml(metadata.range)}"` : ''}`
        : '';
    const formulaBody = metadata?.kind === 'shared' && metadata.preservedOnly && !metadata.sharedMaster ? '' : encodeXml(formula);
    return `<c r="${ref}"${styleAttr}${cachedType ? ` t="${cachedType}"` : ''}><f${formulaAttrs}>${formulaBody}</f>${cached}</c>`;
  }
  if (cell.value === null || cell.value === undefined) return `<c r="${ref}"${styleAttr}/>`;
  if (typeof cell.value === 'boolean') return `<c r="${ref}"${styleAttr} t="b"><v>${cell.value ? '1' : '0'}</v></c>`;
  if (typeof cell.value === 'number') return `<c r="${ref}"${styleAttr}><v>${Number.isFinite(cell.value) ? cell.value : 0}</v></c>`;
  if (isExcelDateFormat(cell.numberFormat)) {
    const serial = canonicalDateToSerial(cell.value, dateSystem);
    return `<c r="${ref}"${styleAttr}><v>${serial}</v></c>`;
  }
  return `<c r="${ref}"${styleAttr} t="inlineStr"><is>${serializeRichText(cell.value, cell.richText)}</is></c>`;
}

function canonicalizeImportedDate(value: CellData['value'], format: string | undefined, dateSystem: DateSystem, label: string): CellData['value'] {
  if (!isExcelDateFormat(format) || typeof value !== 'number') return value;
  try {
    return serialToCanonicalDate(value, dateSystem);
  } catch (error) {
    throw new Error(`Invalid date serial at ${label}: ${String(error instanceof Error ? error.message : error)}`, { cause: error });
  }
}

function serializeRichText(value: string, runs?: RichTextRun[]): string {
  if (!runs?.length) return `<t xml:space="preserve">${encodeXml(value)}</t>`;
  return runs.map((run) => {
    const style = run.style;
    const properties = style ? `<rPr>${style.fontFamily ? `<rFont val="${encodeXml(style.fontFamily)}"/>` : ''}${style.fontSizePx ? `<sz val="${roundMetric(pixelsToPoints(style.fontSizePx))}"/>` : ''}${style.bold ? '<b/>' : ''}${style.italic ? '<i/>' : ''}${style.underline ? '<u/>' : ''}${style.strikethrough ? '<strike/>' : ''}${style.textColor ? `<color rgb="${ooxmlRgb(style.textColor)}"/>` : ''}</rPr>` : '';
    return `<r>${properties}<t xml:space="preserve">${encodeXml(run.text)}</t></r>`;
  }).join('');
}

function serializeBorders(borders: NonNullable<CellStyle['borders']>): string {
  const side = (name: keyof typeof borders): string => {
    const value = borders[name];
    return value ? `<${name} style="${value.style}"><color rgb="${ooxmlRgb(value.color)}"/></${name}>` : `<${name}/>`;
  };
  return `<border>${side('left')}${side('right')}${side('top')}${side('bottom')}<diagonal/></border>`;
}

function serializeDifferentialStyle(style: CellStyle): string {
  const font = style.fontFamily || style.fontSizePx || style.bold || style.italic || style.underline || style.strikethrough || style.textColor
    ? `<font>${style.fontFamily ? `<name val="${encodeXml(style.fontFamily)}"/>` : ''}${style.fontSizePx ? `<sz val="${roundMetric(pixelsToPoints(style.fontSizePx))}"/>` : ''}${style.bold ? '<b/>' : ''}${style.italic ? '<i/>' : ''}${style.underline ? '<u/>' : ''}${style.strikethrough ? '<strike/>' : ''}${style.textColor ? `<color rgb="${ooxmlRgb(style.textColor)}"/>` : ''}</font>` : '';
  const fill = style.background ? `<fill><patternFill patternType="solid"><fgColor rgb="${ooxmlRgb(style.background)}"/><bgColor indexed="64"/></patternFill></fill>` : '';
  const border = style.borders ? serializeBorders(style.borders) : '';
  return `${font}${fill}${border}`;
}

function serializeConditionalFormats(rules: ConditionalFormatRule[], differentialStyles: Map<string, number>): string {
  return rules.map((rule, index) => {
    const sqref = rule.ranges.map(rangeToA1).join(' ');
    const dxfId = rule.style ? differentialStyles.get(JSON.stringify(rule.style)) : undefined;
    const common = ` priority="${rule.priority ?? index + 1}"${dxfId === undefined ? '' : ` dxfId="${dxfId}"`}${rule.stopIfTrue ? ' stopIfTrue="1"' : ''}`;
    let body = '';
    let attrs = '';
    if (rule.type === 'highlight') {
      const expression = rule.operator === 'formula';
      attrs = ` type="${expression ? 'expression' : 'cellIs'}"${expression ? '' : ` operator="${rule.operator ?? 'equal'}"`}`;
      if (rule.value1 !== undefined) body += `<formula>${encodeXml(String(rule.value1).replace(/^=/, ''))}</formula>`;
      if (rule.value2 !== undefined) body += `<formula>${encodeXml(String(rule.value2).replace(/^=/, ''))}</formula>`;
    } else if (rule.type === 'topBottom') {
      attrs = ` type="top10" rank="${rule.topBottom?.rank ?? 10}"${rule.topBottom?.direction === 'bottom' ? ' bottom="1"' : ''}${rule.topBottom?.percent ? ' percent="1"' : ''}`;
    } else if (rule.type === 'dataBar') {
      attrs = ' type="dataBar"';
      body = `<dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="${ooxmlRgb(rule.barColor ?? '#638EC6')}"/></dataBar>`;
    } else if (rule.type === 'colorScale') {
      attrs = ' type="colorScale"';
      const colors = [rule.minColor ?? '#F8696B', ...(rule.midColor ? [rule.midColor] : []), rule.maxColor ?? '#63BE7B'];
      body = `<colorScale>${colors.map((_, colorIndex) => `<cfvo type="${colorIndex === 0 ? 'min' : colorIndex === colors.length - 1 ? 'max' : 'percentile'}"${colorIndex > 0 && colorIndex < colors.length - 1 ? ' val="50"' : ''}/>`).join('')}${colors.map((color) => `<color rgb="${ooxmlRgb(color)}"/>`).join('')}</colorScale>`;
    } else {
      attrs = ' type="iconSet"';
      const thresholds = rule.iconThresholds?.length ? rule.iconThresholds : [
        { type: 'percent' as const, value: 0 },
        { type: 'percent' as const, value: 33 },
        { type: 'percent' as const, value: 67 },
      ];
      body = `<iconSet iconSet="${encodeXml(rule.iconSet ?? '3TrafficLights1')}">${thresholds.map((threshold) => `<cfvo type="${threshold.type}"${threshold.value === undefined ? '' : ` val="${threshold.value}"`}/>`).join('')}</iconSet>`;
    }
    return `<conditionalFormatting sqref="${encodeXml(sqref)}"><cfRule${attrs}${common}>${body}</cfRule></conditionalFormatting>`;
  }).join('');
}

function serializeDataValidations(rules: DataValidationRule[]): string {
  const body = rules.map((rule) => {
    const formula1 = rule.listSource?.kind === 'values'
      ? `"${rule.listSource.values.join(',').replace(/"/g, '""')}"`
      : rule.listSource?.kind === 'formula' ? rule.listSource.formula.replace(/^=/, '') : rule.formula1;
    return `<dataValidation type="${rule.type === 'checkbox' ? 'list' : rule.type}" sqref="${encodeXml(rule.ranges.map(rangeToA1).join(' '))}"${rule.operator ? ` operator="${rule.operator}"` : ''}${rule.allowBlank ? ' allowBlank="1"' : ''}${rule.alertStyle ? ` errorStyle="${rule.alertStyle}"` : ''}${rule.showErrorMessage ? ' showErrorMessage="1"' : ''}${rule.showInputMessage ? ' showInputMessage="1"' : ''}${rule.showDropdown === false ? ' showDropDown="1"' : ''}${rule.promptTitle ? ` promptTitle="${encodeXml(rule.promptTitle)}"` : ''}${rule.promptMessage ? ` prompt="${encodeXml(rule.promptMessage)}"` : ''}${rule.errorTitle ? ` errorTitle="${encodeXml(rule.errorTitle)}"` : ''}${rule.errorMessage ? ` error="${encodeXml(rule.errorMessage)}"` : ''}>${formula1 ? `<formula1>${encodeXml(formula1)}</formula1>` : ''}${rule.formula2 ? `<formula2>${encodeXml(rule.formula2)}</formula2>` : ''}</dataValidation>`;
  }).join('');
  return `<dataValidations count="${rules.length}">${body}</dataValidations>`;
}

function serializeAutoFilter(filter: AutoFilterModel, differentialStyleIndexes?: Map<string, number>): string {
  const columns = Object.values(filter.columns).map((column) => {
    const colId = column.column - filter.range.startColumn;
    const criterion = column.criterion;
    const preservedRecord = isRecord(column.preservedXml) ? column.preservedXml : undefined;
    const preserved = preservedRecord && Array.isArray(preservedRecord.children)
      ? preservedRecord.children.filter((value): value is string => typeof value === 'string').join('')
      : '';
    const preservedFilter = preservedRecord && Array.isArray(preservedRecord.filterChildren)
      ? preservedRecord.filterChildren.filter((value): value is string => typeof value === 'string').join('')
      : '';
    const buttonAttrs = `${column.showButton === false ? ' showButton="0"' : ''}${column.hiddenButton ? ' hiddenButton="1"' : ''}`;
    if (!criterion && column.showButton !== false && !column.hiddenButton && !preserved) return '';
    if (!criterion) return `<filterColumn colId="${colId}"${buttonAttrs}>${preserved}</filterColumn>`;
    if (criterion.kind === 'values') {
      const dateGroups = (criterion.dateGroups ?? []).map((group) => {
        const grouping = group.second !== undefined ? 'second' : group.minute !== undefined ? 'minute' : group.hour !== undefined ? 'hour' : group.day !== undefined ? 'day' : group.month !== undefined ? 'month' : 'year';
        const attrs = [`dateTimeGrouping="${grouping}"`, `year="${group.year}"`];
        if (group.month !== undefined) attrs.push(`month="${group.month}"`);
        if (group.day !== undefined) attrs.push(`day="${group.day}"`);
        if (group.hour !== undefined) attrs.push(`hour="${group.hour}"`);
        if (group.minute !== undefined) attrs.push(`minute="${group.minute}"`);
        if (group.second !== undefined) attrs.push(`second="${group.second}"`);
        return `<dateGroupItem ${attrs.join(' ')}/>`;
      }).join('');
      return `<filterColumn colId="${colId}"${buttonAttrs}><filters${criterion.includeBlank ? ' blank="1"' : ''}>${criterion.values.map((value) => `<filter val="${encodeXml(String(value ?? ''))}"/>`).join('')}${dateGroups}${preservedFilter}</filters>${preserved}</filterColumn>`;
    }
    if (criterion.kind === 'custom') {
      const conditions = criterion.conditions.filter((condition): condition is NonNullable<typeof condition> => Boolean(condition));
      return `<filterColumn colId="${colId}"${buttonAttrs}><customFilters${criterion.join === 'or' ? ' and="0"' : ''}>${conditions.map((condition) => `<customFilter operator="${encodeXml(condition.operator)}" val="${encodeXml(String(condition.value ?? ''))}"/>`).join('')}</customFilters>${preserved}</filterColumn>`;
    }
    if (criterion.kind === 'dynamic') {
      return `<filterColumn colId="${colId}"${buttonAttrs}><dynamicFilter type="${encodeXml(criterion.type)}"${criterion.value === undefined ? '' : ` val="${criterion.value}"`}${criterion.maxValue === undefined ? '' : ` maxVal="${criterion.maxValue}"`}/>${preserved}</filterColumn>`;
    }
    if (criterion.kind === 'top10') {
      return `<filterColumn colId="${colId}"${buttonAttrs}><top10 top="${criterion.top ? '1' : '0'}"${criterion.percent ? ' percent="1"' : ''} rank="${criterion.rank}"${criterion.filterValue === undefined ? '' : ` filterVal="${criterion.filterValue}"`}/>${preserved}</filterColumn>`;
    }
    if (criterion.kind === 'color') {
      const styleKey = criterion.style ? JSON.stringify(criterion.style) : undefined;
      const dxfId = criterion.dxfId >= 0 ? criterion.dxfId : styleKey && differentialStyleIndexes?.get(styleKey);
      if (dxfId === undefined) throw new Error('Color AutoFilter requires a differential style identity');
      return `<filterColumn colId="${colId}"${buttonAttrs}><colorFilter dxfId="${dxfId}" cellColor="${criterion.target === 'cell' ? '1' : '0'}"/>${preserved}</filterColumn>`;
    }
    return `<filterColumn colId="${colId}"${buttonAttrs}><iconFilter iconSet="${encodeXml(criterion.iconSet)}" iconId="${criterion.iconId}"/>${preserved}</filterColumn>`;
  }).join('');
  const sortState = filter.sortState
    ? `<sortState ref="${rangeToA1(filter.sortState.ref)}">${filter.sortState.conditions.map((condition) => `<sortCondition ref="${rangeToA1(condition.ref)}" descending="${condition.descending ? '1' : '0'}"/>`).join('')}</sortState>`
    : '';
  const preserved = isRecord(filter.preservedXml) && typeof filter.preservedXml.extLst === 'string' ? filter.preservedXml.extLst : '';
  return `<autoFilter ref="${rangeToA1(filter.range)}">${columns}${sortState}${preserved}</autoFilter>`;
}

function serializeProtection(rules: ProtectionRule[]): string {
  const rule = rules.find((candidate) => candidate.scope === 'sheet');
  if (!rule) return '';
  const allow = rule.allow;
  return `<sheetProtection sheet="1"${rule.passwordHash ? ` password="${encodeXml(rule.passwordHash)}"` : ''} selectLockedCells="${allow.selectLocked ? '0' : '1'}" selectUnlockedCells="${allow.selectUnlocked ? '0' : '1'}" formatCells="${allow.formatCells ? '0' : '1'}" insertRows="${allow.insertRows ? '0' : '1'}" insertColumns="${allow.insertColumns ? '0' : '1'}" deleteRows="${allow.deleteRows ? '0' : '1'}" deleteColumns="${allow.deleteColumns ? '0' : '1'}" sort="${allow.sort ? '0' : '1'}" autoFilter="${allow.autoFilter ? '0' : '1'}" objects="${allow.editObjects ? '0' : '1'}"/>`;
}

function serializePrintDocument(document: NonNullable<WorkbookSnapshot['printDocuments']>[number]): string {
  const setup = document.pageSetup;
  const paperSize = setup.paperSize === 'letter' ? 1 : setup.paperSize === 'legal' ? 5 : setup.paperSize === 'a3' ? 8 : setup.paperSize === 'a4' ? 9 : 0;
  const printOptions = `<printOptions horizontalCentered="${setup.centerHorizontally ? '1' : '0'}" verticalCentered="${setup.centerVertically ? '1' : '0'}" headings="${setup.printHeadings ? '1' : '0'}" gridLines="${setup.printGridlines ? '1' : '0'}"/>`;
  const margins = `<pageMargins left="${setup.margins.left}" right="${setup.margins.right}" top="${setup.margins.top}" bottom="${setup.margins.bottom}" header="${setup.margins.header}" footer="${setup.margins.footer}"/>`;
  const pageSetup = `<pageSetup${paperSize ? ` paperSize="${paperSize}"` : ''} orientation="${setup.orientation}" scale="${setup.scale}"${setup.fitToWidth !== undefined ? ` fitToWidth="${setup.fitToWidth}"` : ''}${setup.fitToHeight !== undefined ? ` fitToHeight="${setup.fitToHeight}"` : ''}/>`;
  const headerFooter = setup.headerText || setup.footerText ? `<headerFooter>${setup.headerText ? `<oddHeader>${encodeXml(setup.headerText)}</oddHeader>` : ''}${setup.footerText ? `<oddFooter>${encodeXml(setup.footerText)}</oddFooter>` : ''}</headerFooter>` : '';
  const rows = document.pageBreaks.filter((entry) => entry.row !== undefined);
  const columns = document.pageBreaks.filter((entry) => entry.column !== undefined);
  return `${printOptions}${margins}${pageSetup}${headerFooter}${rows.length ? `<rowBreaks count="${rows.length}" manualBreakCount="${rows.length}">${rows.map((entry) => `<brk id="${entry.row}" min="0" max="16383" man="1"/>`).join('')}</rowBreaks>` : ''}${columns.length ? `<colBreaks count="${columns.length}" manualBreakCount="${columns.length}">${columns.map((entry) => `<brk id="${entry.column}" min="0" max="1048575" man="1"/>`).join('')}</colBreaks>` : ''}`;
}

function outlineLevelAt(outline: OutlineModel | undefined, axis: 'row' | 'column', index: number): { level: number; collapsed: boolean } | undefined {
  const matches = outline?.groups.filter((group) => group.axis === axis && index >= group.start && index <= group.end) ?? [];
  if (!matches.length) return undefined;
  return { level: Math.max(...matches.map((group) => group.level)), collapsed: matches.some((group) => group.collapsed) };
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

function buildWorkbookXml(snapshot: WorkbookSnapshot, workbookPart: string, relationships: XlsxRelationship[], descriptors: SheetDescriptor[], dateSystem: DateSystem, nativePivotGraph?: NativePivotGraph, preserved?: OpcPackageGraph): string {
  const relationFor = (target: string, type: string) => relationships.find((relation) => isRelationshipKind(relation.type, relationshipKind(type)) && resolveTarget(workbookPart, relation.target) === target)?.id ?? '';
  let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="${NS_MAIN}" xmlns:r="${NS_DOC_REL}"><workbookPr date1904="${dateSystem === '1904' ? '1' : '0'}"/><sheets>`;
  for (const descriptor of descriptors) {
    const target = relativeTarget(workbookPart, descriptor.part);
    const id = relationFor(descriptor.part, REL_WORKSHEET);
    const sheet = snapshot.sheets.find((candidate) => candidate.id === descriptor.id);
    xml += `<sheet name="${encodeXml(sheet?.name ?? descriptor.name)}" sheetId="${encodeXml(descriptor.id.replace(/^sheet-/, ''))}" r:id="${id}"${sheet?.hidden ? ' state="hidden"' : ''}/>`;
  }
  xml += '</sheets>';
  const names: DefinedNameModel[] = structuredClone(snapshot.definedNameModels ?? Object.entries(snapshot.definedNames ?? {}).map(([name, formula]) => ({ name, formula, scope: 'workbook' as const })));
  for (const document of snapshot.printDocuments ?? []) {
    const sheet = snapshot.sheets.find((candidate) => candidate.id === document.sheetId);
    if (!sheet) continue;
    const qualified = `'${sheet.name.replace(/'/g, "''")}'!`;
    if (document.printAreas.length && !names.some((name) => name.name === '_xlnm.Print_Area' && name.scope === 'sheet' && name.sheetId === sheet.id)) {
      names.push({ name: '_xlnm.Print_Area', scope: 'sheet', sheetId: sheet.id, formula: `=${document.printAreas.map((area) => `${qualified}${rangeToA1(area.range)}`).join(',')}`, hidden: true });
    }
    const titleParts = [
      document.repeatRows ? `${qualified}$${document.repeatRows.start + 1}:$${document.repeatRows.end + 1}` : '',
      document.repeatColumns ? `${qualified}$${columnToLetter(document.repeatColumns.start)}:$${columnToLetter(document.repeatColumns.end)}` : '',
    ].filter(Boolean);
    if (titleParts.length && !names.some((name) => name.name === '_xlnm.Print_Titles' && name.scope === 'sheet' && name.sheetId === sheet.id)) names.push({ name: '_xlnm.Print_Titles', scope: 'sheet', sheetId: sheet.id, formula: `=${titleParts.join(',')}`, hidden: true });
  }
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
  const originalRoot = preserved?.parts[workbookPart] ? firstElement(parseXml(strFromU8(preserved.parts[workbookPart])), 'workbook') : undefined;
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

function buildContentTypesXml(files: Map<string, Uint8Array>, preserved: OpcPackageGraph | undefined, workbookPart: string, stylesPart: string, sharedStringsPart: string): string {
  const defaults = new Map<string, string>([['rels', 'application/vnd.openxmlformats-package.relationships+xml'], ['xml', 'application/xml']]);
  const variant = preserved?.format.family === 'ooxml' ? preserved.format.variant : 'xlsx';
  const mainType = variant === 'xlsm'
    ? 'application/vnd.ms-excel.sheet.macroEnabled.main+xml'
    : variant === 'xltm'
      ? 'application/vnd.ms-excel.template.macroEnabled.main+xml'
      : variant === 'xltx'
        ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml'
        : variant === 'xlam'
          ? 'application/vnd.ms-excel.addin.macroEnabled.main+xml'
          : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml';
  const overrides = new Map<string, string>([
    [`/${workbookPart}`, mainType],
    [`/${stylesPart}`, 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'],
    [`/${sharedStringsPart}`, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'],
  ]);
  const original = preserved?.contentTypesXml ? firstElement(parseXml(strFromU8(preserved.contentTypesXml)), 'Types') : undefined;
  for (const node of children(original, 'Default')) if (node.attrs.Extension && node.attrs.ContentType) defaults.set(node.attrs.Extension, node.attrs.ContentType);
  for (const node of children(original, 'Override')) if (node.attrs.PartName && node.attrs.ContentType) overrides.set(node.attrs.PartName, node.attrs.ContentType);
  for (const name of files.keys()) {
    if (!name.startsWith('xl/worksheets/') || !name.endsWith('.xml')) continue;
    overrides.set(`/${name}`, 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml');
  }
  for (const name of files.keys()) {
    if (name.startsWith('xl/drawings/') && name.endsWith('.xml')) {
      overrides.set(`/${name}`, 'application/vnd.openxmlformats-officedocument.drawing+xml');
    } else if (name.startsWith('xl/charts/') && name.endsWith('.xml')) {
      overrides.set(`/${name}`, 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml');
    } else if (name.startsWith('xl/tables/') && name.endsWith('.xml')) {
      overrides.set(`/${name}`, 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml');
    } else if (name.startsWith('xl/pivotTables/') && name.endsWith('.xml')) {
      overrides.set(`/${name}`, 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml');
    } else if (name.startsWith('xl/pivotCache/') && name.toLowerCase().includes('definition') && name.endsWith('.xml')) {
      overrides.set(`/${name}`, 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml');
    } else if (name.startsWith('xl/pivotCache/') && name.toLowerCase().includes('records') && name.endsWith('.xml')) {
      overrides.set(`/${name}`, 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml');
    } else if (name.startsWith('xl/slicerCaches/') && name.endsWith('.xml')) {
      overrides.set(`/${name}`, 'application/vnd.ms-excel.slicerCache');
    } else if (name.startsWith('xl/slicers/') && name.endsWith('.xml')) {
      overrides.set(`/${name}`, 'application/vnd.ms-excel.slicer');
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

function buildRootRelationshipsXml(existing: XlsxRelationship[], workbookPart = 'xl/workbook.xml'): string {
  const relationships = mergeRelationships(existing.filter((relation) => !isRelationshipKind(relation.type, 'officeDocument')), [{ id: '', type: REL_OFFICE_DOCUMENT, target: relativeTarget('', workbookPart) }]);
  return buildRelationshipsXml(relationships);
}

interface ReactSheetsPackageMetadata {
  schema: 'ReactSheetsWorkbookMetadata';
  version: 1;
  dataModel: WorkbookSnapshot['dataModel'];
  sheets: Array<Pick<SheetSnapshot, 'id' | 'kind' | 'tableSheet' | 'ganttSheet' | 'reportSheet' | 'drawings' | 'drawingPayloads' | 'drawingGroups' | 'snapSettings'> & { cellMetadata: Array<{ row: number; column: number; presentation?: CellData['presentation']; editor?: CellData['editor'] }> }>;
}

function buildReactSheetsMetadata(snapshot: WorkbookSnapshot): string {
  const metadata: ReactSheetsPackageMetadata = {
    schema: 'ReactSheetsWorkbookMetadata', version: 1, dataModel: structuredClone(snapshot.dataModel),
    sheets: snapshot.sheets.map((sheet) => {
      const cellMetadata: ReactSheetsPackageMetadata['sheets'][number]['cellMetadata'] = [];
      for (const [rowKey, columns] of Object.entries(sheet.cells)) for (const [columnKey, cell] of Object.entries(columns)) if (cell.presentation || cell.editor) cellMetadata.push({ row: Number(rowKey), column: Number(columnKey), presentation: cell.presentation ? structuredClone(cell.presentation) : undefined, editor: cell.editor ? structuredClone(cell.editor) : undefined });
      const drawings = sheet.drawings.filter((drawing) => drawing.kind !== 'slicer' && drawing.kind !== 'timeline');
      const payloadIds = new Set(drawings.map((drawing) => drawing.payloadId));
      const drawingPayloads = Object.fromEntries(Object.entries(sheet.drawingPayloads).filter(([payloadId]) => payloadIds.has(payloadId)));
      const drawingIds = new Set(drawings.map((drawing) => drawing.id));
      const drawingGroups = (sheet.drawingGroups ?? []).filter((group) => group.memberDrawingIds.every((id) => drawingIds.has(id)));
      return { id: sheet.id, kind: sheet.kind, tableSheet: sheet.tableSheet ? structuredClone(sheet.tableSheet) : undefined, ganttSheet: sheet.ganttSheet ? structuredClone(sheet.ganttSheet) : undefined, reportSheet: sheet.reportSheet ? structuredClone(sheet.reportSheet) : undefined, drawings: structuredClone(drawings), drawingPayloads: structuredClone(drawingPayloads), drawingGroups: structuredClone(drawingGroups), snapSettings: sheet.snapSettings ? structuredClone(sheet.snapSettings) : undefined, cellMetadata };
    }),
  };
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><reactSheetsWorkbook xmlns="urn:react-sheets:workbook-metadata:v1"><json>${encodeXml(JSON.stringify(metadata))}</json></reactSheetsWorkbook>`;
}

function remapSheetIdentity(value: unknown, fromId: string, toId: string): void {
  if (!value || typeof value !== 'object') return;
  if (Array.isArray(value)) { for (const entry of value) remapSheetIdentity(entry, fromId, toId); return; }
  const record = value as Record<string, unknown>;
  for (const [key, entry] of Object.entries(record)) {
    if ((key === 'sheetId' || key === 'sourceSheetId' || key === 'templateSheetId') && entry === fromId) record[key] = toId;
    else remapSheetIdentity(entry, fromId, toId);
  }
}

function applyReactSheetsMetadata(snapshot: WorkbookSnapshot, bytes: Uint8Array | undefined, graph: OpcPackageGraph): void {
  if (!bytes) return;
  try {
    const root = firstElement(parseXml(strFromU8(bytes)), 'reactSheetsWorkbook');
    const parsed = JSON.parse(textContent(child(root, 'json'))) as ReactSheetsPackageMetadata;
    if (parsed.schema !== 'ReactSheetsWorkbookMetadata' || parsed.version !== 1 || !parsed.dataModel || !Array.isArray(parsed.sheets)) return;
    snapshot.dataModel = structuredClone(parsed.dataModel);
    for (let index = 0; index < parsed.sheets.length; index += 1) {
      const metadata = parsed.sheets[index]!;
      const sheet = snapshot.sheets.find((candidate) => candidate.id === metadata.id) ?? snapshot.sheets[index];
      if (!sheet) continue;
      const importedId = sheet.id;
      if (importedId !== metadata.id) {
        remapSheetIdentity(sheet, importedId, metadata.id);
        sheet.id = metadata.id;
        const part = graph.sheetPartById[importedId];
        if (part) { delete graph.sheetPartById[importedId]; graph.sheetPartById[metadata.id] = part; }
      }
      sheet.kind = metadata.kind;
      sheet.tableSheet = metadata.tableSheet ? structuredClone(metadata.tableSheet) : undefined;
      sheet.ganttSheet = metadata.ganttSheet ? structuredClone(metadata.ganttSheet) : undefined;
      sheet.reportSheet = metadata.reportSheet ? structuredClone(metadata.reportSheet) : undefined;
      sheet.drawings = structuredClone(metadata.drawings);
      sheet.drawingPayloads = structuredClone(metadata.drawingPayloads);
      sheet.drawingGroups = structuredClone(metadata.drawingGroups ?? []);
      sheet.snapSettings = metadata.snapSettings ? structuredClone(metadata.snapSettings) : sheet.snapSettings;
      for (const entry of metadata.cellMetadata ?? []) {
        sheet.cells[String(entry.row)] ??= {};
        const cell = sheet.cells[String(entry.row)]![String(entry.column)] ?? { value: null };
        if (entry.presentation) cell.presentation = structuredClone(entry.presentation);
        if (entry.editor) cell.editor = structuredClone(entry.editor);
        sheet.cells[String(entry.row)]![String(entry.column)] = cell;
      }
    }
  } catch (error) {
    throw new Error(`React Sheets workbook metadata is invalid: ${error instanceof Error ? error.message : String(error)}`);
  }
}

function mergeRelationships(existing: XlsxRelationship[], required: Array<Pick<XlsxRelationship, 'type' | 'target' | 'targetMode'> & { id?: string }>): XlsxRelationship[] {
  const result = existing.map((relation) => ({ ...relation }));
  const used = new Set(result.map((relation) => relation.id));
  let next = 1;
  for (const request of required) {
    const found = result.find((relation) => relationshipKind(relation.type) === relationshipKind(request.type) && relation.target === request.target && relation.targetMode === request.targetMode);
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
  for (const entry of sheet.hyperlinks ?? []) {
    const target = entry.hyperlink.target;
    if (target.kind !== 'url' && target.kind !== 'email') continue;
    const href = hyperlinkTarget(entry.hyperlink);
    if (!existing.some((relation) => isRelationshipKind(relation.type, 'hyperlink') && relation.target === href)
      && !links.some((relation) => relation.target === href)) {
      links.push({ type: REL_HYPERLINK, target: href, targetMode: 'External' });
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

function hyperlinkForCell(root: XmlNode, relationships: XlsxRelationship[], row: number, column: number, sheetDescriptors: readonly SheetDescriptor[]): CellHyperlink | undefined {
  const hyperlinks = child(root, 'hyperlinks');
  const reference = `${columnToLetter(column)}${row + 1}`;
  const node = children(hyperlinks, 'hyperlink').find((candidate) => candidate.attrs.ref === reference);
  if (!node) return undefined;
  const relationId = node.attrs['r:id'];
  if (relationId) {
    const relation = relationships.find((candidate) => candidate.id === relationId);
    if (relation) {
      if (relation.target.startsWith('mailto:')) {
        const [address = '', query = ''] = relation.target.slice(7).split('?');
        if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(address)) throw new Error('Worksheet contains an invalid email hyperlink');
        const subject = new URLSearchParams(query).get('subject') ?? undefined;
        return { id: `hyperlink-${row}-${column}`, target: { kind: 'email', address, ...(subject ? { subject } : {}) }, ...(node.attrs.tooltip ? { tooltip: node.attrs.tooltip } : {}) };
      }
      return { id: `hyperlink-${row}-${column}`, target: { kind: 'url', url: relation.target }, ...(node.attrs.tooltip ? { tooltip: node.attrs.tooltip } : {}) };
    }
  }
  if (node.attrs.location) {
    const location = node.attrs.location;
    const separator = location.lastIndexOf('!');
    if (separator > 0) {
      const rawSheetName = location.slice(0, separator).replace(/^'(.*)'$/, '$1').replace(/''/g, "'");
      const targetSheet = sheetDescriptors.find((descriptor) => descriptor.name === rawSheetName);
      const targetAddress = location.slice(separator + 1);
      if (!targetSheet || !parseA1(targetAddress)) throw new Error(`Hyperlink worksheet location is invalid: ${location}`);
      return { id: `hyperlink-${row}-${column}`, target: { kind: 'sheet', sheetId: targetSheet.id, address: targetAddress }, ...(node.attrs.tooltip ? { tooltip: node.attrs.tooltip } : {}) };
    }
    return { id: `hyperlink-${row}-${column}`, target: { kind: 'name', name: location }, ...(node.attrs.tooltip ? { tooltip: node.attrs.tooltip } : {}) };
  }
  return undefined;
}

function excelSheetName(name: string): string {
  return /[\s!'"(),]/.test(name) ? `'${name.replace(/'/g, "''")}'` : name;
}

function parseNotes(root: XmlNode, descriptor: SheetDescriptor, files: Record<string, Uint8Array>, pkg: OpcPackageGraph): SheetSnapshot['notes'] {
  const relation = (pkg.relationships[descriptor.part] ?? []).find((candidate) => isRelationshipKind(candidate.type, 'comments'));
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

function applyPrintDefinedNames(snapshot: WorkbookSnapshot): void {
  const documents = new Map((snapshot.printDocuments ?? []).map((document) => [document.sheetId, document]));
  const ensureDocument = (sheetId: string) => {
    let document = documents.get(sheetId);
    if (!document) {
      document = { schema: 'PrintDocument', unitId: snapshot.unitId, sheetId, pageSetup: { paperSize: 'a4', orientation: 'portrait', margins: { top: 0.75, right: 0.7, bottom: 0.75, left: 0.7, header: 0.3, footer: 0.3 }, scale: 100, printGridlines: false, printHeadings: false, centerHorizontally: false, centerVertically: false }, printAreas: [], pageBreaks: [] };
      documents.set(sheetId, document);
    }
    return document;
  };
  for (const name of snapshot.definedNameModels ?? []) {
    if (name.scope !== 'sheet' || !name.sheetId) continue;
    const formula = name.formula.replace(/^=/, '');
    if (name.name === '_xlnm.Print_Area') {
      const document = ensureDocument(name.sheetId);
      document.printAreas = formula.split(',').flatMap((part) => {
        const range = parseRange(part.split('!').pop()?.replace(/\$/g, ''));
        return range ? [{ sheetId: name.sheetId!, range: { ...range, sheetId: name.sheetId! } }] : [];
      });
    } else if (name.name === '_xlnm.Print_Titles') {
      const document = ensureDocument(name.sheetId);
      for (const part of formula.split(',')) {
        const reference = part.split('!').pop()?.replace(/\$/g, '') ?? '';
        const rows = /^(\d+):(\d+)$/.exec(reference);
        const columns = /^([A-Za-z]+):([A-Za-z]+)$/.exec(reference);
        if (rows) document.repeatRows = { start: Number(rows[1]) - 1, end: Number(rows[2]) - 1 };
        if (columns) document.repeatColumns = { start: columnFromLetter(columns[1]!), end: columnFromLetter(columns[2]!) };
      }
    }
  }
  snapshot.printDocuments = [...documents.values()];
}

interface SharedFormulaMaster {
  row: number;
  column: number;
  formula: string;
  range?: string;
}

function collectSharedFormulaMasters(
  entries: Array<{ node: XmlNode; row: number; column: number }>,
  descriptor: SheetDescriptor,
): Map<number, SharedFormulaMaster> {
  const result = new Map<number, SharedFormulaMaster>();
  for (const entry of entries) {
    const formula = child(entry.node, 'f');
    if (!formula || formula.attrs.t !== 'shared' || !textContent(formula).trim()) continue;
    const index = Number(formula.attrs.si);
    if (!Number.isSafeInteger(index) || index < 0) throw new Error(`Worksheet ${descriptor.name}!${columnToLetter(entry.column)}${entry.row + 1} has an invalid shared formula index`);
    if (result.has(index)) throw new Error(`Worksheet ${descriptor.name} has duplicate shared formula master si=${index}`);
    result.set(index, {
      row: entry.row,
      column: entry.column,
      formula: `=${textContent(formula)}`,
      ...(formula.attrs.ref ? { range: formula.attrs.ref } : {}),
    });
  }
  return result;
}

function readFormula(
  cell: XmlNode,
  address: { row: number; column: number },
  sharedMasters: Map<number, SharedFormulaMaster>,
  descriptor: SheetDescriptor,
): { formula?: string; metadata?: NonNullable<CellData['formulaMetadata']> } {
  const node = child(cell, 'f');
  if (!node) return {};
  const kind = node.attrs.t ?? 'normal';
  const raw = textContent(node).trim();
  if (kind === 'dataTable') {
    return { metadata: { kind: 'dataTable', ...(node.attrs.ref ? { range: node.attrs.ref } : {}), preservedOnly: true, reason: 'Excel data-table formulas are preserved from the source package', ...(raw ? { sourceFormula: `=${raw}` } : {}) } };
  }
  if (kind === 'shared') {
    const index = Number(node.attrs.si);
    if (!Number.isSafeInteger(index) || index < 0) throw new Error(`Worksheet ${descriptor.name}!${columnToLetter(address.column)}${address.row + 1} has an invalid shared formula index`);
    const master = sharedMasters.get(index);
    if (!master) throw new Error(`Worksheet ${descriptor.name}!${columnToLetter(address.column)}${address.row + 1} references missing shared formula master si=${index}`);
    try {
      const formula = address.row === master.row && address.column === master.column
        ? formatFormula(parseFormula(master.formula))
        : formatFormula(offsetAst(parseFormula(master.formula), address.row - master.row, address.column - master.column));
      return { formula, metadata: { kind: 'shared', sharedIndex: index, sharedMaster: address.row === master.row && address.column === master.column, ...(master.range ? { range: master.range } : {}) } };
    } catch {
      return {
        metadata: {
          kind: 'shared',
          sharedIndex: index,
          sharedMaster: address.row === master.row && address.column === master.column,
          ...(master.range ? { range: master.range } : {}),
          preservedOnly: true,
          reason: 'Shared formula uses syntax outside the canonical formula AST',
          sourceFormula: master.formula,
        },
      };
    }
  }
  if (!raw) return {};
  const formula = `=${raw}`;
  const formulaKind = kind === 'array' ? 'array' : 'normal';
  const unsupported = formulaPreserveReason(formula);
  if (unsupported) return { formula, metadata: { kind: formulaKind, ...(node.attrs.ref ? { range: node.attrs.ref } : {}), preservedOnly: true, reason: unsupported, sourceFormula: formula } };
  try {
    parseFormula(formula);
    return { formula, metadata: formulaKind === 'normal' ? undefined : { kind: formulaKind, ...(node.attrs.ref ? { range: node.attrs.ref } : {}) } };
  } catch {
    return { formula, metadata: { kind: formulaKind, ...(node.attrs.ref ? { range: node.attrs.ref } : {}), preservedOnly: true, reason: 'Formula syntax is not supported by the canonical AST', sourceFormula: formula } };
  }
}

function formulaPreserveReason(formula: string): string | undefined {
  if (/\[[^\]]+\]/.test(formula)) return 'External workbook reference is fail-closed';
  if (/\bCUBE(?:VALUE|SET|MEMBER|RANKEDMEMBER|MEMBERPROPERTY)\s*\(/i.test(formula)) return 'CUBE formula requires an external OLAP calculation contract';
  if (/\bINDIRECT\s*\(/i.test(formula)) return 'INDIRECT is preserved without local recalculation';
  return undefined;
}

function parseConditionalFormats(root: XmlNode, descriptor: SheetDescriptor, styles: StyleContext): ConditionalFormatRule[] {
  const result: ConditionalFormatRule[] = [];
  let sequence = 0;
  for (const container of children(root, 'conditionalFormatting')) {
    const ranges = parseSqref(container.attrs.sqref, descriptor, 'conditional formatting');
    for (const node of children(container, 'cfRule')) {
      const formulas = children(node, 'formula').map(textContent);
      const type = node.attrs.type;
      const base = {
        id: `cf-${descriptor.id}-${node.attrs.priority ?? ++sequence}`,
        sheetId: descriptor.id,
        ranges,
        ...(Number.isFinite(Number(node.attrs.priority)) ? { priority: Number(node.attrs.priority) } : {}),
        ...(node.attrs.stopIfTrue !== undefined ? { stopIfTrue: xmlBoolean(node.attrs.stopIfTrue) } : {}),
        ...(styles.differentialStyles[Number(node.attrs.dxfId)] ? { style: structuredClone(styles.differentialStyles[Number(node.attrs.dxfId)]!) } : {}),
      };
      if (type === 'cellIs' || type === 'expression' || type === 'containsText' || type === 'notContainsText' || type === 'duplicateValues' || type === 'uniqueValues') {
        result.push({
          ...base,
          type: 'highlight',
          operator: normalizeConditionalOperator(type === 'cellIs' ? node.attrs.operator : type),
          ...(formulas[0] !== undefined ? { value1: scalarFormulaValue(formulas[0]) } : {}),
          ...(formulas[1] !== undefined ? { value2: scalarFormulaValue(formulas[1]) } : {}),
        });
      } else if (type === 'top10') {
        result.push({ ...base, type: 'topBottom', operator: xmlBoolean(node.attrs.bottom ?? '0') ? 'bottom' : 'top', topBottom: { direction: xmlBoolean(node.attrs.bottom ?? '0') ? 'bottom' : 'top', rank: finitePositive(node.attrs.rank, 10), ...(xmlBoolean(node.attrs.percent ?? '0') ? { percent: true } : {}) } });
      } else if (type === 'dataBar') {
        const color = resolveColor(child(child(node, 'dataBar'), 'color'), styles.themeColors);
        result.push({ ...base, type: 'dataBar', ...(color ? { barColor: color } : {}) });
      } else if (type === 'colorScale') {
        const colors = children(child(node, 'colorScale'), 'color').map((color) => resolveColor(color, styles.themeColors)).filter((value): value is string => Boolean(value));
        result.push({ ...base, type: 'colorScale', ...(colors[0] ? { minColor: colors[0] } : {}), ...(colors.length === 3 && colors[1] ? { midColor: colors[1] } : {}), ...(colors.at(-1) ? { maxColor: colors.at(-1)! } : {}) });
      } else if (type === 'iconSet') {
        const iconSet = child(node, 'iconSet');
        const iconThresholds = children(iconSet, 'cfvo').map((threshold) => ({
          type: (threshold.attrs.type ?? 'percent') as 'percent' | 'percentile' | 'num' | 'formula',
          ...(threshold.attrs.val === undefined || !Number.isFinite(Number(threshold.attrs.val)) ? {} : { value: Number(threshold.attrs.val) }),
        }));
        result.push({ ...base, type: 'iconSet', ...(iconSet?.attrs.iconSet ? { iconSet: iconSet.attrs.iconSet } : {}), ...(iconThresholds.length ? { iconThresholds } : {}) });
      }
    }
  }
  return result;
}

function normalizeConditionalOperator(value: string | undefined): ConditionalFormatRule['operator'] {
  const map: Record<string, NonNullable<ConditionalFormatRule['operator']>> = {
    greaterThan: 'greaterThan', lessThan: 'lessThan', between: 'between', equal: 'equal', notEqual: 'notEqual',
    containsText: 'containsText', notContainsText: 'notContainsText', duplicateValues: 'duplicate', uniqueValues: 'unique', expression: 'formula',
  };
  return map[value ?? ''] ?? 'formula';
}

function scalarFormulaValue(value: string): string | number {
  const number = Number(value);
  return Number.isFinite(number) && value.trim() !== '' ? number : value;
}

function parseDataValidations(root: XmlNode, descriptor: SheetDescriptor): DataValidationRule[] {
  const validations = children(child(root, 'dataValidations'), 'dataValidation');
  return validations.flatMap((node, index) => {
    const type = normalizeValidationType(node.attrs.type);
    if (!type) return [];
    const formula1 = textContent(child(node, 'formula1'));
    const formula2 = textContent(child(node, 'formula2'));
    const ranges = parseSqref(node.attrs.sqref, descriptor, 'data validation');
    const listSource = type === 'list' && formula1
      ? formula1.startsWith('"') && formula1.endsWith('"')
        ? { kind: 'values' as const, values: formula1.slice(1, -1).split(',') }
        : { kind: 'formula' as const, formula: formula1.startsWith('=') ? formula1 : `=${formula1}` }
      : undefined;
    return [{
      id: `dv-${descriptor.id}-${index + 1}`,
      sheetId: descriptor.id,
      ranges,
      type,
      ...(normalizeValidationOperator(node.attrs.operator) ? { operator: normalizeValidationOperator(node.attrs.operator)! } : {}),
      ...(formula1 ? { formula1 } : {}),
      ...(formula2 ? { formula2 } : {}),
      ...(node.attrs.allowBlank !== undefined ? { allowBlank: xmlBoolean(node.attrs.allowBlank) } : {}),
      ...(node.attrs.errorStyle && ['stop', 'warning', 'information'].includes(node.attrs.errorStyle) ? { alertStyle: node.attrs.errorStyle as DataValidationRule['alertStyle'] } : {}),
      ...(node.attrs.showErrorMessage !== undefined ? { showErrorMessage: xmlBoolean(node.attrs.showErrorMessage) } : {}),
      ...(node.attrs.showInputMessage !== undefined ? { showInputMessage: xmlBoolean(node.attrs.showInputMessage) } : {}),
      ...(node.attrs.showDropDown !== undefined ? { showDropdown: !xmlBoolean(node.attrs.showDropDown) } : {}),
      ...(node.attrs.promptTitle ? { promptTitle: node.attrs.promptTitle, inputTitle: node.attrs.promptTitle } : {}),
      ...(node.attrs.prompt ? { promptMessage: node.attrs.prompt, inputMessage: node.attrs.prompt } : {}),
      ...(node.attrs.errorTitle ? { errorTitle: node.attrs.errorTitle } : {}),
      ...(node.attrs.error ? { errorMessage: node.attrs.error } : {}),
      ...(listSource ? { listSource } : {}),
    }];
  });
}

function normalizeValidationType(value: string | undefined): DataValidationRule['type'] | undefined {
  const map: Record<string, DataValidationRule['type']> = { list: 'list', whole: 'whole', decimal: 'decimal', date: 'date', time: 'time', textLength: 'textLength', custom: 'custom' };
  return map[value ?? ''];
}

function normalizeValidationOperator(value: string | undefined): DataValidationRule['operator'] | undefined {
  const supported = new Set<DataValidationRule['operator']>(['between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan']);
  return supported.has(value as DataValidationRule['operator']) ? value as DataValidationRule['operator'] : undefined;
}

function parseDateGroupItem(item: XmlNode): { year: number; month?: number; day?: number; hour?: number; minute?: number; second?: number } | undefined {
  const grouping = item.attrs.dateTimeGrouping;
  if (!grouping || !['year', 'month', 'day', 'hour', 'minute', 'second'].includes(grouping)) return undefined;
  const allowed = new Set(['dateTimeGrouping', 'year', 'month', 'day', 'hour', 'minute', 'second']);
  if (Object.keys(item.attrs).some((key) => !allowed.has(key))) return undefined;
  const parse = (key: string, minimum: number, maximum?: number): number | undefined => {
    const raw = item.attrs[key];
    if (raw === undefined) return undefined;
    const number = Number(raw);
    if (!Number.isSafeInteger(number) || number < minimum || (maximum !== undefined && number > maximum)) return undefined;
    return number;
  };
  const year = parse('year', 1);
  if (year === undefined) return undefined;
  const month = parse('month', 1, 12);
  const day = parse('day', 1, 31);
  const hour = parse('hour', 0, 23);
  const minute = parse('minute', 0, 59);
  const second = parse('second', 0, 59);
  if ((item.attrs.month !== undefined && month === undefined)
    || (item.attrs.day !== undefined && day === undefined)
    || (item.attrs.hour !== undefined && hour === undefined)
    || (item.attrs.minute !== undefined && minute === undefined)
    || (item.attrs.second !== undefined && second === undefined)) return undefined;
  const expected: Record<string, string | undefined> = { year: item.attrs.year, month: item.attrs.month, day: item.attrs.day, hour: item.attrs.hour, minute: item.attrs.minute, second: item.attrs.second };
  const order = ['year', 'month', 'day', 'hour', 'minute', 'second'];
  const deepest = order.indexOf(grouping);
  if (deepest < 0 || order.slice(1, deepest + 1).some((key) => expected[key] === undefined) || order.slice(deepest + 1).some((key) => expected[key] !== undefined)) return undefined;
  return { year, ...(month === undefined ? {} : { month }), ...(day === undefined ? {} : { day }), ...(hour === undefined ? {} : { hour }), ...(minute === undefined ? {} : { minute }), ...(second === undefined ? {} : { second }) };
}

function parseAutoFilter(root: XmlNode, descriptor: SheetDescriptor, styles?: StyleContext): AutoFilterModel | undefined {
  const node = child(root, 'autoFilter');
  if (!node?.attrs.ref) return undefined;
  const range = requireSheetRange(node.attrs.ref, descriptor, 'auto filter');
  const columns: AutoFilterModel['columns'] = {};
  for (const column of children(node, 'filterColumn')) {
    const relative = Number(column.attrs.colId);
    if (!Number.isSafeInteger(relative) || relative < 0) continue;
    const absolute = range.startColumn + relative;
    const filtersNode = child(column, 'filters');
    const filters: string[] = [];
    const dateGroups: Array<{ year: number; month?: number; day?: number; hour?: number; minute?: number; second?: number }> = [];
    const preservedFilterChildren: string[] = [];
    for (const filterChild of filtersNode?.children ?? []) {
      const name = localName(filterChild.name);
      if (name === 'filter') {
        filters.push(filterChild.attrs.val ?? '');
        continue;
      }
      if (name === 'dateGroupItem') {
        const dateGroup = parseDateGroupItem(filterChild);
        if (dateGroup) dateGroups.push(dateGroup);
        else preservedFilterChildren.push(serializeXml(filterChild));
        continue;
      }
      preservedFilterChildren.push(serializeXml(filterChild));
    }
    const custom = children(child(column, 'customFilters'), 'customFilter');
    const customFiltersNode = child(column, 'customFilters');
    const dynamic = child(column, 'dynamicFilter');
    if (dynamic?.attrs.type && !isDynamicFilterType(dynamic.attrs.type)) {
      throw new Error(`UNSUPPORTED_FEATURE: dynamic AutoFilter type "${dynamic.attrs.type}" is not supported`);
    }
    const top10 = child(column, 'top10');
    const color = child(column, 'colorFilter');
    const icon = child(column, 'iconFilter');
    const known = new Set(['filters', 'customFilters', 'dynamicFilter', 'top10', 'colorFilter', 'iconFilter']);
    const preservedChildren = column.children.filter((candidate) => !known.has(localName(candidate.name))).map(serializeXml);
    columns[absolute] = {
      column: absolute,
      showButton: column.attrs.showButton !== '0',
      hiddenButton: column.attrs.hiddenButton === '1',
      ...((preservedChildren.length || preservedFilterChildren.length) ? { preservedXml: {
        ...(preservedChildren.length ? { children: preservedChildren } : {}),
        ...(preservedFilterChildren.length ? { filterChildren: preservedFilterChildren } : {}),
      } } : {}),
      criterion: filters.length || dateGroups.length || filtersNode?.attrs.blank !== undefined || preservedFilterChildren.length
        ? { kind: 'values', values: filters, includeBlank: filtersNode?.attrs.blank === '1', ...(dateGroups.length ? { dateGroups } : {}) }
        : custom.length
          ? {
            kind: 'custom',
            join: customFiltersNode?.attrs.and === '0' ? 'or' : 'and',
            conditions: [
              { operator: (custom[0]?.attrs.operator ?? 'equals') as import('@react-sheets/core-model').FilterComparisonOperator, value: custom[0]?.attrs.val ?? null },
              custom[1] ? { operator: (custom[1].attrs.operator ?? 'equals') as import('@react-sheets/core-model').FilterComparisonOperator, value: custom[1].attrs.val ?? null } : undefined,
            ],
          }
          : dynamic?.attrs.type
            ? { kind: 'dynamic', type: dynamic.attrs.type as import('@react-sheets/core-model').DynamicFilterType, ...(dynamic.attrs.val === undefined ? {} : { value: Number(dynamic.attrs.val) }), ...(dynamic.attrs.maxVal === undefined ? {} : { maxValue: Number(dynamic.attrs.maxVal) }) }
            : top10
              ? { kind: 'top10', top: top10.attrs.top !== '0', percent: top10.attrs.percent === '1', rank: Number(top10.attrs.rank ?? 10), ...(top10.attrs.filterVal === undefined ? {} : { filterValue: Number(top10.attrs.filterVal) }) }
              : color
                ? { kind: 'color', target: color.attrs.cellColor === '0' ? 'font' : 'cell', dxfId: Number(color.attrs.dxfId ?? -1), ...(styles?.differentialStyles[Number(color.attrs.dxfId)] ? { style: structuredClone(styles.differentialStyles[Number(color.attrs.dxfId)]!) } : {}) }
                : icon
                  ? { kind: 'icon', iconSet: icon.attrs.iconSet ?? '', iconId: Number(icon.attrs.iconId ?? -1) }
                  : undefined,
    };
  }
  const sort = child(node, 'sortState');
  const sortState = sort ? {
    ref: { ...requireSheetRange(sort.attrs.ref ?? node.attrs.ref, descriptor, 'sort state'), sheetId: descriptor.id },
    conditions: children(sort, 'sortCondition').map((condition) => ({
      ref: { ...requireSheetRange(condition.attrs.ref ?? node.attrs.ref, descriptor, 'sort condition'), sheetId: descriptor.id },
      descending: xmlBoolean(condition.attrs.descending ?? '0'),
    })),
  } : undefined;
  const preserved = node.children.filter((candidate) => localName(candidate.name) === 'extLst').map(serializeXml);
  return {
    sheetId: descriptor.id,
    range: { ...range, sheetId: descriptor.id },
    columns,
    sortState,
    ...(preserved.length ? { preservedXml: { extLst: preserved.join('') } } : {}),
  };
}

function parseOutline(root: XmlNode): OutlineModel | undefined {
  const groups: OutlineModel['groups'] = [];
  const appendGroups = (nodes: XmlNode[], axis: 'row' | 'column', startAttr: string, endAttr: string): void => {
    const entries = nodes.flatMap((node) => {
      const level = Number(node.attrs.outlineLevel ?? 0);
      if (!Number.isSafeInteger(level) || level <= 0) return [];
      return [{ start: parsePositiveInt(node.attrs[startAttr], 1) - 1, end: parsePositiveInt(node.attrs[endAttr] ?? node.attrs[startAttr], 1) - 1, level, collapsed: xmlBoolean(node.attrs.collapsed ?? '0') }];
    }).sort((a, b) => a.level - b.level || a.start - b.start);
    for (const entry of entries) {
      const previous = groups.at(-1);
      if (previous && previous.axis === axis && previous.level === entry.level && previous.end + 1 === entry.start && previous.collapsed === entry.collapsed) previous.end = entry.end;
      else groups.push({ id: `outline-${axis}-${entry.level}-${entry.start}`, axis, ...entry });
    }
  };
  appendGroups(children(child(root, 'sheetData'), 'row'), 'row', 'r', 'r');
  appendGroups(children(child(root, 'cols'), 'col'), 'column', 'min', 'max');
  return groups.length ? { groups } : undefined;
}

function parseProtection(root: XmlNode, descriptor: SheetDescriptor): ProtectionRule[] {
  const node = child(root, 'sheetProtection');
  if (!node) return [];
  return [{
    id: `protection-${descriptor.id}`,
    scope: 'sheet',
    sheetId: descriptor.id,
    passwordHash: node.attrs.password ?? node.attrs.hashValue,
    locked: true,
    allow: {
      selectLocked: !xmlBoolean(node.attrs.selectLockedCells ?? '0'),
      selectUnlocked: !xmlBoolean(node.attrs.selectUnlockedCells ?? '0'),
      formatCells: !xmlBoolean(node.attrs.formatCells ?? '1'),
      insertRows: !xmlBoolean(node.attrs.insertRows ?? '1'),
      insertColumns: !xmlBoolean(node.attrs.insertColumns ?? '1'),
      deleteRows: !xmlBoolean(node.attrs.deleteRows ?? '1'),
      deleteColumns: !xmlBoolean(node.attrs.deleteColumns ?? '1'),
      sort: !xmlBoolean(node.attrs.sort ?? '1'),
      autoFilter: !xmlBoolean(node.attrs.autoFilter ?? '1'),
      editObjects: !xmlBoolean(node.attrs.objects ?? '1'),
    },
  }];
}

function parsePrintDocument(root: XmlNode, unitId: string, sheetId: string): NonNullable<WorkbookSnapshot['printDocuments']> {
  const pageSetup = child(root, 'pageSetup');
  const margins = child(root, 'pageMargins');
  const printOptions = child(root, 'printOptions');
  const headerFooter = child(root, 'headerFooter');
  const rowBreaks = children(child(root, 'rowBreaks'), 'brk');
  const columnBreaks = children(child(root, 'colBreaks'), 'brk');
  if (!pageSetup && !margins && !printOptions && !headerFooter && !rowBreaks.length && !columnBreaks.length) return [];
  const paper = Number(pageSetup?.attrs.paperSize);
  return [{
    schema: 'PrintDocument', unitId, sheetId,
    pageSetup: {
      paperSize: paper === 8 ? 'a3' : paper === 9 ? 'a4' : paper === 5 ? 'legal' : paper === 1 ? 'letter' : 'custom',
      orientation: pageSetup?.attrs.orientation === 'landscape' ? 'landscape' : 'portrait',
      margins: { top: finiteNumber(margins?.attrs.top, 0.75), right: finiteNumber(margins?.attrs.right, 0.7), bottom: finiteNumber(margins?.attrs.bottom, 0.75), left: finiteNumber(margins?.attrs.left, 0.7), header: finiteNumber(margins?.attrs.header, 0.3), footer: finiteNumber(margins?.attrs.footer, 0.3) },
      scale: finitePositive(pageSetup?.attrs.scale, 100),
      ...(pageSetup?.attrs.fitToWidth !== undefined ? { fitToWidth: Number(pageSetup.attrs.fitToWidth) } : {}),
      ...(pageSetup?.attrs.fitToHeight !== undefined ? { fitToHeight: Number(pageSetup.attrs.fitToHeight) } : {}),
      printGridlines: xmlBoolean(printOptions?.attrs.gridLines ?? '0'), printHeadings: xmlBoolean(printOptions?.attrs.headings ?? '0'),
      centerHorizontally: xmlBoolean(printOptions?.attrs.horizontalCentered ?? '0'), centerVertically: xmlBoolean(printOptions?.attrs.verticalCentered ?? '0'),
      ...(child(headerFooter, 'oddHeader') ? { headerText: textContent(child(headerFooter, 'oddHeader')) } : {}),
      ...(child(headerFooter, 'oddFooter') ? { footerText: textContent(child(headerFooter, 'oddFooter')) } : {}),
    },
    printAreas: [],
    pageBreaks: [
      ...rowBreaks.flatMap((node) => Number.isSafeInteger(Number(node.attrs.id)) ? [{ sheetId, row: Number(node.attrs.id) }] : []),
      ...columnBreaks.flatMap((node) => Number.isSafeInteger(Number(node.attrs.id)) ? [{ sheetId, column: Number(node.attrs.id) }] : []),
    ],
  }];
}

function parsePane(root: XmlNode): WorksheetPane {
  const pane = child(child(child(root, 'sheetViews'), 'sheetView'), 'pane');
  if (!pane) return { kind: 'none' };
  const xSplit = Number(pane?.attrs.xSplit ?? 0) || 0;
  const ySplit = Number(pane?.attrs.ySplit ?? 0) || 0;
  const topLeft = pane?.attrs.topLeftCell ? parseA1(pane.attrs.topLeftCell) : undefined;
  const activePane = normalizeActivePane(pane.attrs.activePane);
  const state = pane.attrs.state;
  if (state === 'frozen' || state === 'frozenSplit') {
    return {
      kind: 'frozen',
      xSplit: Math.max(0, Math.trunc(xSplit)),
      ySplit: Math.max(0, Math.trunc(ySplit)),
      startRow: topLeft?.row ?? Math.max(0, Math.trunc(ySplit)),
      startColumn: topLeft?.column ?? Math.max(0, Math.trunc(xSplit)),
      ...(activePane ? { activePane } : {}),
      state,
    };
  }
  return {
    kind: 'split',
    xSplit: Math.max(0, xSplit),
    ySplit: Math.max(0, ySplit),
    startRow: topLeft?.row ?? 0,
    startColumn: topLeft?.column ?? 0,
    ...(activePane ? { activePane } : {}),
    state: 'split',
  };
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
    const points = Number(row.attrs.ht);
    if (!Number.isFinite(points) || points < 0) throw new Error(`Worksheet row ${row.attrs.r ?? '?'} has an invalid height`);
    result[parsePositiveInt(row.attrs.r, 1) - 1] = pointsToPixels(points);
  }
  return result;
}

function parseColumnWidths(root: XmlNode, maximumDigitWidthPx: number): Record<number, number> {
  const result: Record<number, number> = {};
  for (const col of children(child(root, 'cols'), 'col')) {
    if (col.attrs.width === undefined) continue;
    const start = parsePositiveInt(col.attrs.min, 1) - 1;
    const end = parsePositiveInt(col.attrs.max, start + 1) - 1;
    const width = Number(col.attrs.width);
    if (!Number.isFinite(width) || width < 0 || width > 255) throw new Error(`Worksheet column ${start + 1}:${end + 1} has an invalid width`);
    for (let index = start; index <= end; index += 1) result[index] = excelColumnWidthToPixels(width, maximumDigitWidthPx);
  }
  return result;
}

function readCellValue(cell: XmlNode, sharedStrings: SharedStringRecord[], themeColors: string[]): { value: string | number | boolean | null; richText?: RichTextRun[] } {
  const type = cell.attrs.t;
  if (type === 'inlineStr') return parseRichTextContainer(child(cell, 'is') ?? cell, themeColors);
  const raw = textContent(child(cell, 'v'));
  if (raw === '') return { value: null };
  if (type === 's') return sharedStrings[Number(raw)] ?? { value: '' };
  if (type === 'b') return { value: raw === '1' || raw.toLowerCase() === 'true' };
  if (type === 'e' || type === 'str') return { value: raw };
  const number = Number(raw);
  return { value: Number.isNaN(number) ? raw : number };
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

function readSheetPartMap(files: Record<string, Uint8Array>, relationships: Record<string, XlsxRelationship[]>, workbookPart: string): Record<string, string> {
  const map: Record<string, string> = {};
  const workbook = firstElement(parseXml(strFromU8(files[workbookPart]!)), 'workbook');
  const rels = relationships[workbookPart] ?? [];
  children(child(workbook, 'sheets'), 'sheet').forEach((node, index) => {
    const id = `sheet-${node.attrs.sheetId ?? index + 1}`;
    const relation = rels.find((candidate) => candidate.id === (node.attrs['r:id'] ?? node.attrs.id));
    map[id] = relation ? resolveTarget(workbookPart, relation.target) : resolveTarget(workbookPart, `worksheets/sheet${index + 1}.xml`);
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

function relationshipTarget(pkg: OpcPackageGraph | undefined, source: string, type: string): string | undefined {
  const relation = pkg?.relationships[source]?.find((candidate) => isRelationshipKind(candidate.type, relationshipKind(type)));
  return relation ? resolveTarget(source, relation.target) : undefined;
}

function cloneParts(parts: Record<string, Uint8Array>): Record<string, Uint8Array> {
  return Object.fromEntries(Object.entries(parts).map(([name, data]) => [name, data.slice()]));
}

function isVbaPart(name: string, pkg?: OpcPackageGraph, files?: Record<string, Uint8Array>): boolean {
  const lower = name.toLowerCase();
  if (lower.endsWith('vbaproject.bin') || lower.endsWith('vbaprojectsignature.bin')) return true;
  if (pkg) {
    for (const [source, relationships] of Object.entries(pkg.relationships)) {
      for (const relation of relationships) {
        const type = relation.type.toLowerCase();
        if (!type.includes('vbaproject')) continue;
        if (resolveTarget(source, relation.target) === name) return true;
      }
    }
    const contentTypes = pkg.contentTypesXml ? strFromU8(pkg.contentTypesXml) : '';
    if (contentTypes.includes(`/` + name) && /vbaProject/i.test(contentTypes)) return true;
  }
  if (files) {
    const contentTypes = files['[Content_Types].xml'] ? strFromU8(files['[Content_Types].xml']!) : '';
    if (contentTypes.includes(`/` + name) && /vbaProject/i.test(contentTypes)) return true;
  }
  return false;
}

function filterMacroRelationships(source: string, relationships: XlsxRelationship[], pkg: OpcPackageGraph | undefined): XlsxRelationship[] {
  return relationships.filter((relationship) => !isVbaPart(resolveTarget(source, relationship.target), pkg));
}

function relationshipKind(type: string): string {
  const normalized = type.replace(/\/+$/, '');
  return normalized.slice(normalized.lastIndexOf('/') + 1);
}

function isRelationshipKind(type: string, kind: string): boolean {
  return relationshipKind(type).toLocaleLowerCase() === kind.toLocaleLowerCase();
}

function resolveWorkbookRelatedPart(workbookPart: string, relationships: XlsxRelationship[], kind: string, fallback: string): string {
  const relation = relationships.find((candidate) => isRelationshipKind(candidate.type, kind));
  return relation ? resolveTarget(workbookPart, relation.target) : fallback;
}

function requireSheetRange(value: string | undefined, descriptor: SheetDescriptor, feature: string): RangeRef {
  const range = parseRange(value);
  if (!range || range.startRow < 0 || range.startColumn < 0 || range.endRow < range.startRow || range.endColumn < range.startColumn
    || range.endRow > OOXML_MAX_ROW_INDEX || range.endColumn > OOXML_MAX_COLUMN_INDEX) {
    throw new Error(`Worksheet ${descriptor.name} has an invalid ${feature} range: ${value ?? ''}`);
  }
  return { ...range, sheetId: descriptor.id };
}

function parseSqref(value: string | undefined, descriptor: SheetDescriptor, feature: string): RangeRef[] {
  if (!value?.trim()) throw new Error(`Worksheet ${descriptor.name} ${feature} is missing sqref`);
  return value.trim().split(/\s+/).map((entry) => requireSheetRange(entry, descriptor, feature));
}

function validateNonOverlappingMerges(merges: MergeSpan[], descriptor: SheetDescriptor): void {
  const sorted = [...merges].sort((left, right) => left.range.startRow - right.range.startRow || left.range.startColumn - right.range.startColumn);
  for (let index = 0; index < sorted.length; index += 1) {
    const current = sorted[index]!;
    for (let candidateIndex = index + 1; candidateIndex < sorted.length; candidateIndex += 1) {
      const candidate = sorted[candidateIndex]!;
      if (candidate.range.startRow > current.range.endRow) break;
      if (current.range.startColumn <= candidate.range.endColumn && candidate.range.startColumn <= current.range.endColumn) {
        throw new Error(`Worksheet ${descriptor.name} has overlapping merges ${rangeToA1(current.range)} and ${rangeToA1(candidate.range)}`);
      }
    }
  }
}

function normalizeActivePane(value: string | undefined): Exclude<WorksheetPane, { kind: 'none' }>['activePane'] | undefined {
  return value === 'topLeft' || value === 'topRight' || value === 'bottomLeft' || value === 'bottomRight' ? value : undefined;
}

function xmlBoolean(value: string): boolean {
  return value === '1' || value.toLocaleLowerCase() === 'true';
}

function finiteNumber(value: string | undefined, fallback: number): number {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function finitePositive(value: string | undefined, fallback: number): number {
  const parsed = Number(value);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : fallback;
}

function roundMetric(value: number): string {
  return Number(value.toFixed(8)).toString();
}

function ooxmlRgb(value: string): string {
  const color = normalizeRgb(value) ?? '#000000';
  return `FF${color.slice(1)}`;
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

function assertOoxmlAddress(row: number, column: number, subject: string): void {
  if (!Number.isSafeInteger(row) || !Number.isSafeInteger(column) || row < 0 || column < 0
    || row > OOXML_MAX_ROW_INDEX || column > OOXML_MAX_COLUMN_INDEX) {
    throw new Error(`UNSUPPORTED_FEATURE: ${subject} exceeds the OOXML worksheet boundary`);
  }
}

function validateOoxmlExchangeBoundary(sheet: SheetSnapshot): void {
  for (const [row, columns] of Object.entries(sheet.cells)) {
    const rowIndex = Number(row);
    for (const column of Object.keys(columns)) assertOoxmlAddress(rowIndex, Number(column), `${sheet.name}!cell`);
  }
  for (const row of Object.keys(sheet.rowHeightsPx ?? {})) assertOoxmlAddress(Number(row), 0, `${sheet.name}!row dimension`);
  for (const row of sheet.hiddenRows ?? []) assertOoxmlAddress(row, 0, `${sheet.name}!hidden row`);
  for (const column of Object.keys(sheet.columnWidthsPx ?? {})) assertOoxmlAddress(0, Number(column), `${sheet.name}!column dimension`);
  for (const column of sheet.hiddenColumns ?? []) assertOoxmlAddress(0, column, `${sheet.name}!hidden column`);
  for (const merge of sheet.merges) validateOoxmlRange(merge.range, `${sheet.name}!merge`);
  for (const range of [sheet.autoFilter?.range, ...(sheet.sheetTables ?? []).map((table) => table.range)].filter((range): range is RangeRef => Boolean(range))) {
    validateOoxmlRange(range, `${sheet.name}!filter`);
  }
  for (const rule of sheet.conditionalFormats ?? []) for (const range of rule.ranges) validateOoxmlRange(range, `${sheet.name}!conditional format`);
  for (const rule of sheet.dataValidations ?? []) for (const range of rule.ranges) validateOoxmlRange(range, `${sheet.name}!data validation`);
  for (const rule of sheet.dataValidations ?? []) {
    if (rule.listSource?.kind === 'range') validateOoxmlRange(rule.listSource.range, `${sheet.name}!validation source`);
  }
  for (const hyperlink of sheet.hyperlinks ?? []) assertOoxmlAddress(hyperlink.row, hyperlink.column, `${sheet.name}!hyperlink`);
}

function validateOoxmlRange(range: RangeRef, subject: string): void {
  assertOoxmlAddress(range.startRow, range.startColumn, subject);
  assertOoxmlAddress(range.endRow, range.endColumn, subject);
  if (range.endRow < range.startRow || range.endColumn < range.startColumn) throw new Error(`UNSUPPORTED_FEATURE: ${subject} is inverted`);
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

function parseAlignmentAttributes(node: XmlNode | undefined): Partial<CellStyle> {
  if (!node) return {};
  const attrs = node.attrs;
  const style: Partial<CellStyle> = {};
  const unsupported: NonNullable<CellStyle['unsupportedAlignment']> = {};
  const preserved: Record<string, string> = {};
  const known = new Set(['horizontal', 'vertical', 'wrapText', 'shrinkToFit', 'indent', 'readingOrder', 'textRotation']);

  const horizontal = attrs.horizontal;
  if (horizontal !== undefined) {
    if (['general', 'left', 'center', 'right', 'centerContinuous', 'justify', 'distributed', 'fill'].includes(horizontal)) style.horizontalAlignment = horizontal as NonNullable<CellStyle['horizontalAlignment']>;
    else unsupported.horizontal = horizontal;
  }
  const vertical = attrs.vertical;
  if (vertical !== undefined) {
    if (vertical === 'center') style.verticalAlignment = 'middle';
    else if (['top', 'bottom', 'justify', 'distributed'].includes(vertical)) style.verticalAlignment = vertical as NonNullable<CellStyle['verticalAlignment']>;
    else unsupported.vertical = vertical;
  }
  if (attrs.wrapText !== undefined) style.wrapText = strictXmlBoolean(attrs.wrapText, 'wrapText');
  if (attrs.shrinkToFit !== undefined) style.shrinkToFit = strictXmlBoolean(attrs.shrinkToFit, 'shrinkToFit');
  if (attrs.indent !== undefined) {
    const indent = Number(attrs.indent);
    if (!Number.isInteger(indent) || indent < 0 || indent > 250) throw new Error(`Invalid OOXML alignment indent: ${attrs.indent}`);
    style.indent = indent;
  }
  if (attrs.readingOrder !== undefined) {
    if (attrs.readingOrder === '0') style.readingOrder = 'context';
    else if (attrs.readingOrder === '1') style.readingOrder = 'ltr';
    else if (attrs.readingOrder === '2') style.readingOrder = 'rtl';
    else preserved.readingOrder = attrs.readingOrder;
  }
  if (attrs.textRotation !== undefined) {
    const rotation = Number(attrs.textRotation);
    if (!Number.isInteger(rotation) || rotation < 0 || rotation > 180 && rotation !== 255) throw new Error(`Invalid OOXML alignment textRotation: ${attrs.textRotation}`);
    if (rotation === 255) style.textOrientation = 'stacked';
    else if (rotation > 0) style.textRotate = rotation;
  }
  for (const [key, value] of Object.entries(attrs)) if (!known.has(key)) preserved[key] = value;
  if (Object.keys(preserved).length > 0) unsupported.attributes = preserved;
  if (unsupported.horizontal !== undefined || unsupported.vertical !== undefined || unsupported.attributes !== undefined) style.unsupportedAlignment = unsupported;
  return style;
}

function strictXmlBoolean(value: string, field: string): boolean {
  if (!/^(?:0|1|true|false)$/iu.test(value)) throw new Error(`Invalid OOXML alignment ${field}: ${value}`);
  return xmlBoolean(value);
}

function serializeAlignment(style: CellStyle): string {
  const attrs: Record<string, string> = { ...(style.unsupportedAlignment?.attributes ?? {}) };
  if (style.horizontalAlignment !== undefined) attrs.horizontal = style.horizontalAlignment;
  else if (style.unsupportedAlignment?.horizontal !== undefined) attrs.horizontal = style.unsupportedAlignment.horizontal;
  if (style.verticalAlignment !== undefined) attrs.vertical = style.verticalAlignment === 'middle' ? 'center' : style.verticalAlignment;
  else if (style.unsupportedAlignment?.vertical !== undefined) attrs.vertical = style.unsupportedAlignment.vertical;
  if (style.wrapText !== undefined) attrs.wrapText = style.wrapText ? '1' : '0';
  if (style.shrinkToFit !== undefined) attrs.shrinkToFit = style.shrinkToFit ? '1' : '0';
  if (style.indent !== undefined) attrs.indent = String(style.indent);
  if (style.readingOrder !== undefined) attrs.readingOrder = style.readingOrder === 'context' ? '0' : style.readingOrder === 'ltr' ? '1' : '2';
  if (style.textOrientation === 'stacked') attrs.textRotation = '255';
  else if (style.textOrientation === 'rotateUp') attrs.textRotation = '90';
  else if (style.textOrientation === 'rotateDown') attrs.textRotation = '180';
  else if (style.textRotate !== undefined) attrs.textRotation = String(style.textRotate);
  const serialized = Object.entries(attrs).map(([key, value]) => ` ${key}="${encodeXml(value)}"`).join('');
  return serialized ? `<alignment${serialized}/>` : '';
}

function descriptorsForSnapshot(snapshot: WorkbookSnapshot): SheetDescriptor[] {
  return snapshot.sheets.map((sheet, index) => ({ id: sheet.id, name: sheet.name, part: `xl/worksheets/sheet${index + 1}.xml`, hidden: Boolean(sheet.hidden) }));
}
