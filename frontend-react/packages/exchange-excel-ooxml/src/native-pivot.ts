import { strFromU8, strToU8 } from 'fflate';
import type {
  CellData,
  PivotAggregateFunction,
  PivotDefinition,
  PivotDateFilterOperator,
  PivotDynamicDateFilter,
  PivotErrorValue,
  PivotFieldDataType,
  PivotFieldPlacement,
  PivotGroup,
  PivotLabelFilterOperator,
  PivotLayout,
  PivotModel,
  PivotNativeCacheFlags,
  PivotNativeAutoSortMetadata,
  PivotNativeFilterMetadata,
  PivotRefreshPolicy,
  PivotScalar,
  PivotValueFilterOperator,
  PivotSlicerDrawingPayload,
  PivotTimelineDrawingPayload,
  PivotTimelineFilterType,
  PivotTimelineLevel,
  PivotSource,
  PivotValueField,
  RangeRef,
  SheetSnapshot,
  WorkbookSnapshot,
} from '@react-sheets/core-model';
import { canonicalizePivotDefinition, createPivotMemberKey, DEFAULT_PIVOT_COLLATION, DEFAULT_PIVOT_STYLE_OPTIONS, formatPivotMember, isPivotError, normalizePivotDisplayOptions, normalizePivotRefreshPolicy, pivotMemberKey, pivotTimelineInstant, refreshOnSaveForPivotMode } from '@react-sheets/core-model';
import { child, children, descendants, encodeXml, localName, parseXml, serializeXml, textContent, type XmlNode } from './xml';
import type {
  NativePivotCacheDefinition,
  NativePivotCacheField,
  NativePivotControlDefinition,
  NativePivotDataField,
  NativePivotFilter,
  NativePivotAutoSortScope,
  NativePivotFieldGroup,
  NativePivotFieldRange,
  NativePivotScalar,
  NativePivotGraph,
  NativePivotPackageUpdate,
  NativePivotSource,
  NativePivotTableDefinition,
  NativePivotTableField,
  XlsxRelationship,
} from './types';

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_DOC_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL_PIVOT_CACHE_DEFINITION = `${NS_DOC_REL}/pivotCacheDefinition`;
const REL_PIVOT_CACHE_RECORDS = `${NS_DOC_REL}/pivotCacheRecords`;
const REL_PIVOT_TABLE = `${NS_DOC_REL}/pivotTable`;
const NS_X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
const NS_X15 = 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main';
const NS_SLICER_DRAWING = 'http://schemas.microsoft.com/office/drawing/2010/slicer';
const NS_TIMELINE_DRAWING = 'http://schemas.microsoft.com/office/drawing/2012/timeslicer';
const REL_SLICER_CACHE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slicerCache';
const REL_SLICER_CACHE_MODERN = 'http://schemas.microsoft.com/office/2007/relationships/slicerCache';
const REL_SLICER = 'http://schemas.microsoft.com/office/2007/relationships/slicer';
const REL_TIMELINE_CACHE = 'http://schemas.microsoft.com/office/2010/relationships/timelineCache';
const REL_TIMELINE_CACHE_ALT = 'http://schemas.microsoft.com/office/2010/relationships/TimelineCache';
const REL_TIMELINE = 'http://schemas.microsoft.com/office/2011/relationships/timeline';
const REL_DRAWING = `${NS_DOC_REL}/drawing`;
const SLICER_CACHE_EXT_URI = '{BBE1A952-AA13-448E-AADC-164F8A28A991}';
const SLICER_LIST_EXT_URI = '{A8765BA9-456A-4DAB-B4F3-ACF838C121DE}';
const TIMELINE_CACHE_EXT_URI = '{D0CA8CA8-9F24-4464-BF8E-62219DCF47F9}';
const TIMELINE_REFS_EXT_URI = '{7E03D99C-DC04-49D9-9315-930204A7B6E9}';

export interface NativePivotReadInput {
  files: Record<string, Uint8Array>;
  relationships: Record<string, XlsxRelationship[]>;
  sheetPartById: Record<string, string>;
}

export interface NativePivotWriteInput {
  files: Record<string, Uint8Array>;
  relationships: Record<string, XlsxRelationship[]>;
  graph: NativePivotGraph;
  sheetNameByPart?: Record<string, string>;
}

export interface NativePivotPackageWriteInput {
  files: Record<string, Uint8Array>;
  relationships: Record<string, XlsxRelationship[]>;
  graph?: NativePivotGraph;
  snapshot: WorkbookSnapshot;
  sheetPartById: Record<string, string>;
}

/** Read all reachable Pivot caches, cache records, table definitions and sheet relations. */
export function readNativePivotGraph(input: NativePivotReadInput): NativePivotGraph {
  const workbookPart = 'xl/workbook.xml';
  const workbookBytes = input.files[workbookPart];
  if (!workbookBytes) throw new Error('Native Pivot reader requires xl/workbook.xml');
  const workbook = firstElement(parseXml(strFromU8(workbookBytes)), 'workbook');
  const workbookRels = input.relationships[workbookPart] ?? [];
  const sheetNames = readSheetNames(workbook, workbookRels, input.sheetPartById);
  const caches: NativePivotCacheDefinition[] = [];
  for (const cacheNode of children(child(workbook, 'pivotCaches'), 'pivotCache')) {
    const cacheId = requiredInteger(cacheNode.attrs.cacheId, 'pivotCache.cacheId');
    const relationId = cacheNode.attrs['r:id'] ?? cacheNode.attrs.id;
    if (!relationId) throw new Error(`Native Pivot cache ${cacheId} is missing r:id`);
    const relation = requireRelationship(workbookRels, relationId, REL_PIVOT_CACHE_DEFINITION, `pivot cache ${cacheId}`);
    const part = resolveTarget(workbookPart, relation.target);
    const definitionBytes = input.files[part];
    if (!definitionBytes) throw new Error(`Pivot cache definition relation points to missing part: ${part}`);
    const definition = firstElement(parseXml(strFromU8(definitionBytes)), 'pivotCacheDefinition');
    const cacheRels = input.relationships[part] ?? [];
    const recordsRelation = cacheRels.find((candidate) => candidate.type === REL_PIVOT_CACHE_RECORDS || candidate.type.endsWith('/pivotCacheRecords'));
    const recordsPart = recordsRelation ? resolveTarget(part, recordsRelation.target) : undefined;
    if (recordsPart && !input.files[recordsPart]) throw new Error(`Pivot cache records relation points to missing part: ${recordsPart}`);
    const recordsRoot = recordsPart ? firstElement(parseXml(strFromU8(input.files[recordsPart]!)), 'pivotCacheRecords') : undefined;
    caches.push({
      cacheId,
      part,
      ...(recordsPart ? { recordsPart } : {}),
      source: parseCacheSource(child(definition, 'cacheSource'), sheetNames),
      fields: parseCacheFields(child(definition, 'cacheFields')),
      ...(recordsRoot?.attrs.count !== undefined ? { recordCount: requiredInteger(recordsRoot.attrs.count, 'pivotCacheRecords.count') } : {}),
      ...optionalBoolean(definition.attrs.refreshOnLoad, 'refreshOnLoad'),
      ...optionalBoolean(definition.attrs.refreshOnSave, 'refreshOnSave'),
      ...optionalBoolean(definition.attrs.saveData, 'saveData'),
      ...optionalBoolean(definition.attrs.enableRefresh, 'enableRefresh'),
    });
  }
  const tables: NativePivotTableDefinition[] = [];
  for (const sheetPart of Object.values(input.sheetPartById)) {
    const bytes = input.files[sheetPart];
    if (!bytes) throw new Error(`Worksheet relation points to missing part: ${sheetPart}`);
    const root = firstElement(parseXml(strFromU8(bytes)), 'worksheet');
    const sheetRels = input.relationships[sheetPart] ?? [];
    for (const node of children(child(root, 'pivotTableParts'), 'pivotTablePart')) {
      const relationId = node.attrs['r:id'] ?? node.attrs.id;
      if (!relationId) throw new Error(`Worksheet ${sheetPart} contains a pivotTablePart without r:id`);
      const relation = requireRelationship(sheetRels, relationId, REL_PIVOT_TABLE, `worksheet ${sheetPart} PivotTable`);
      const part = resolveTarget(sheetPart, relation.target);
      const tableBytes = input.files[part];
      if (!tableBytes) throw new Error(`PivotTable relation points to missing part: ${part}`);
      const definition = firstElement(parseXml(strFromU8(tableBytes)), 'pivotTableDefinition');
      const style = child(definition, 'pivotTableStyleInfo');
      tables.push({
        name: definition.attrs.name ?? part,
        part,
        sheetPart,
        relationshipId: relationId,
        cacheId: requiredInteger(definition.attrs.cacheId, `PivotTable ${part}.cacheId`),
        ...(child(definition, 'location')?.attrs.ref ? { locationRef: child(definition, 'location')!.attrs.ref } : {}),
        fields: parsePivotFields(child(definition, 'pivotFields')),
        rowFields: parseFieldIndexes(child(definition, 'rowFields'), 'rowFields'),
        columnFields: parseFieldIndexes(child(definition, 'colFields'), 'colFields'),
        pageFields: parsePageFieldIndexes(child(definition, 'pageFields'), 'pageFields'),
        dataFields: parseDataFields(child(definition, 'dataFields')),
        ...(child(definition, 'pivotFilters') ? { pivotFilters: parsePivotFilters(child(definition, 'pivotFilters')) } : {}),
        ...optionalBoolean(definition.attrs.rowGrandTotals ?? definition.attrs.showRowGrandTotals, 'showRowGrandTotals'),
        ...optionalBoolean(definition.attrs.colGrandTotals ?? definition.attrs.showColumnGrandTotals, 'showColumnGrandTotals'),
        ...optionalBoolean(definition.attrs.compactData, 'compactData'),
        ...optionalBoolean(definition.attrs.multipleFieldFilters, 'multipleFieldFilters'),
        ...optionalBoolean(definition.attrs.repeatAllLabels, 'repeatLabels'),
        ...optionalBoolean(definition.attrs.showDrill, 'showButtons'),
        subtotalLocation: definition.attrs.showSubtotals === '0' ? 'off' : definition.attrs.subtotalTop === '1' || definition.attrs.subtotalTop === 'true' ? 'top' : 'bottom',
        ...(style?.attrs.name ? { styleName: style.attrs.name } : {}),
        ...(style ? { styleOptions: {
          showRowHeaders: style.attrs.showRowHeaders === undefined ? DEFAULT_PIVOT_STYLE_OPTIONS.showRowHeaders : style.attrs.showRowHeaders !== '0',
          showColumnHeaders: style.attrs.showColHeaders === undefined ? DEFAULT_PIVOT_STYLE_OPTIONS.showColumnHeaders : style.attrs.showColHeaders !== '0',
          showRowStripes: style.attrs.showRowStripes === undefined ? DEFAULT_PIVOT_STYLE_OPTIONS.showRowStripes : style.attrs.showRowStripes === '1',
          showColumnStripes: style.attrs.showColStripes === undefined ? DEFAULT_PIVOT_STYLE_OPTIONS.showColumnStripes : style.attrs.showColStripes === '1',
          showLastColumn: style.attrs.showLastColumn === undefined ? DEFAULT_PIVOT_STYLE_OPTIONS.showLastColumn : style.attrs.showLastColumn === '1',
        } } : {}),
        ...(definition.attrs.showHeaders !== undefined ? { showFieldHeaders: definition.attrs.showHeaders !== '0' && definition.attrs.showHeaders.toLowerCase() !== 'false' } : {}),
        ...(definition.attrs.showMissing !== undefined ? { fillEmptyCells: definition.attrs.showMissing === '1' || definition.attrs.showMissing.toLowerCase() === 'true' } : {}),
        ...(definition.attrs.missingCaption !== undefined ? { emptyCellText: definition.attrs.missingCaption } : {}),
        ...(definition.attrs.showError !== undefined ? { showErrorValues: definition.attrs.showError === '1' || definition.attrs.showError.toLowerCase() === 'true' } : {}),
        ...(definition.attrs.errorCaption !== undefined ? { errorCellText: definition.attrs.errorCaption } : {}),
        ...(definition.attrs.preserveFormatting !== undefined ? { preserveFormatting: definition.attrs.preserveFormatting !== '0' && definition.attrs.preserveFormatting.toLowerCase() !== 'false' } : {}),
      });
    }
  }
  const controls = readNativePivotControls(input, caches, tables);
  return { schema: 'NativePivotGraph', caches, tables, ...(controls.length ? { controls } : {}) };
}

interface NativeControlCachePart {
  kind: 'slicer' | 'timeline';
  part: string;
  relationId: string;
  name: string;
  sourceName?: string;
  pivotCacheId?: number;
  pivotTableNames: string[];
  selectedItemIndexes?: number[];
  selection?: { start?: string; end?: string };
  bounds?: { start?: string; end?: string };
  filterType?: PivotTimelineFilterType;
}

function readNativePivotControls(input: NativePivotReadInput, caches: NativePivotCacheDefinition[], tables: NativePivotTableDefinition[]): NativePivotControlDefinition[] {
  const workbookPart = 'xl/workbook.xml';
  const workbook = firstElement(parseXml(strFromU8(input.files[workbookPart]!)), 'workbook');
  const workbookRels = input.relationships[workbookPart] ?? [];
  const cacheParts: NativeControlCachePart[] = [];
  const addCache = (kind: 'slicer' | 'timeline', node: XmlNode): void => {
    const relationId = node.attrs['r:id'] ?? node.attrs.id;
    if (!relationId) return;
    const relation = workbookRels.find((candidate) => candidate.id === relationId && (kind === 'slicer' ? isSlicerCacheRelation(candidate) : isTimelineCacheRelation(candidate)));
    if (!relation) return;
    const part = resolveTarget(workbookPart, relation.target);
    const bytes = input.files[part];
    if (!bytes) return;
    const root = firstElement(parseXml(strFromU8(bytes)), kind === 'slicer' ? 'slicerCacheDefinition' : 'timelineCacheDefinition');
    const data = descendants(root, 'tabular')[0] ?? descendants(root, 'state')[0];
    const pivotCacheId = data?.attrs.pivotCacheId === undefined ? undefined : requiredInteger(data.attrs.pivotCacheId, `${kind} cache pivotCacheId`);
    const selection = descendants(root, 'selection')[0];
    const state = kind === 'timeline' ? descendants(root, 'state')[0] : undefined;
    const bounds = state ? readTimelineBounds(child(state, 'bounds'), `${kind} cache bounds`) : undefined;
    const filterType = state ? parseTimelineFilterType(state.attrs.filterType, `${kind} cache filterType`) : undefined;
    const selectedItemIndexes = kind === 'slicer' ? children(descendants(root, 'items')[0], 'i').flatMap((item) => item.attrs.s === '1' && item.attrs.x !== undefined ? [requiredInteger(item.attrs.x, `${kind} item.x`)] : []) : [];
    cacheParts.push({
      kind,
      part,
      relationId,
      name: root.attrs.name ?? part,
      ...(root.attrs.sourceName ? { sourceName: root.attrs.sourceName } : {}),
      ...(pivotCacheId === undefined ? {} : { pivotCacheId }),
      pivotTableNames: children(child(root, 'pivotTables'), 'pivotTable').flatMap((pivot) => pivot.attrs.name ? [pivot.attrs.name] : []),
      ...(selectedItemIndexes.length ? { selectedItemIndexes } : {}),
      ...(selection ? { selection: { ...(selection.attrs.startDate ? { start: selection.attrs.startDate } : {}), ...(selection.attrs.endDate ? { end: selection.attrs.endDate } : {}) } } : {}),
      ...(bounds ? { bounds } : {}),
      ...(filterType ? { filterType } : {}),
    });
  };
  for (const node of descendants(workbook, 'slicerCache')) addCache('slicer', node);
  for (const node of descendants(workbook, 'timelineCacheRef')) addCache('timeline', node);
  const controls: NativePivotControlDefinition[] = [];
  for (const sheetPart of Object.values(input.sheetPartById)) {
    const bytes = input.files[sheetPart];
    if (!bytes) continue;
    const root = firstElement(parseXml(strFromU8(bytes)), 'worksheet');
    const sheetRels = input.relationships[sheetPart] ?? [];
    for (const relation of sheetRels.filter((candidate) => isSlicerRelation(candidate))) {
      const part = resolveTarget(sheetPart, relation.target);
      const partBytes = input.files[part];
      if (!partBytes) continue;
      const slicers = firstElement(parseXml(strFromU8(partBytes)), 'slicers');
      const drawingRelation = sheetRels.find((candidate) => candidate.type === REL_DRAWING || candidate.type.endsWith('/drawing'));
      const drawingPart = drawingRelation ? resolveTarget(sheetPart, drawingRelation.target) : undefined;
      for (const node of children(slicers, 'slicer')) {
        const control = buildImportedControl('slicer', node, part, relation.id, sheetPart, cacheParts, caches, tables);
        if (drawingPart && input.files[drawingPart] && strFromU8(input.files[drawingPart]!).includes(`name="${encodeXml(control.name)}"`)) {
          control.drawingPart = drawingPart;
          control.drawingRelationshipId = drawingRelation!.id;
          const anchor = readControlDrawingAnchor(input.files[drawingPart]!, control.name);
          if (anchor) control.drawingAnchor = anchor;
        }
        controls.push(control);
      }
    }
    for (const relation of sheetRels.filter((candidate) => isTimelineRelation(candidate))) {
      const part = resolveTarget(sheetPart, relation.target);
      const partBytes = input.files[part];
      if (!partBytes) continue;
      const timelines = firstElement(parseXml(strFromU8(partBytes)), 'timelines');
      const drawingRelation = sheetRels.find((candidate) => candidate.type === REL_DRAWING || candidate.type.endsWith('/drawing'));
      const drawingPart = drawingRelation ? resolveTarget(sheetPart, drawingRelation.target) : undefined;
      for (const node of children(timelines, 'timeline')) {
        const control = buildImportedControl('timeline', node, part, relation.id, sheetPart, cacheParts, caches, tables);
        if (drawingPart && input.files[drawingPart] && strFromU8(input.files[drawingPart]!).includes(`name="${encodeXml(control.name)}"`)) {
          control.drawingPart = drawingPart;
          control.drawingRelationshipId = drawingRelation!.id;
          const anchor = readControlDrawingAnchor(input.files[drawingPart]!, control.name);
          if (anchor) control.drawingAnchor = anchor;
        }
        controls.push(control);
      }
    }
    void root;
  }
  return controls;
}

function buildImportedControl(
  kind: 'slicer' | 'timeline',
  node: XmlNode,
  part: string,
  relationshipId: string,
  sheetPart: string,
  cacheParts: NativeControlCachePart[],
  caches: NativePivotCacheDefinition[],
  tables: NativePivotTableDefinition[],
): NativePivotControlDefinition {
  const name = node.attrs.name ?? part;
  const cacheName = node.attrs.cache ?? '';
  const cachePart = cacheParts.find((candidate) => candidate.kind === kind && candidate.name === cacheName);
  const pivotCache = cachePart?.pivotCacheId === undefined ? undefined : caches.find((cache) => cache.cacheId === cachePart.pivotCacheId);
  const table = cachePart?.pivotTableNames.map((tableName) => tables.find((candidate) => candidate.name === tableName)).find((candidate): candidate is NativePivotTableDefinition => Boolean(candidate && (!pivotCache || candidate.cacheId === pivotCache.cacheId)));
  const field = pivotCache && cachePart?.sourceName ? pivotCache.fields.find((candidate) => candidate.name === cachePart.sourceName) : undefined;
  const connectedPivotIds = cachePart?.pivotTableNames.flatMap((tableName) => { const found = tables.find((candidate) => candidate.name === tableName); return found ? [nativePivotId(found)] : []; });
  const valid = Boolean(cachePart && pivotCache && table && field);
  const timelineLevel = kind === 'timeline' ? parseTimelineLevel(node.attrs.level, `${kind} level`) : undefined;
  const selectionLevel = kind === 'timeline' ? parseTimelineLevel(node.attrs.selectionLevel, `${kind} selectionLevel`) : undefined;
  const showHeader = kind === 'timeline' ? parseTimelineBoolean(node.attrs.showHeader, true, `${kind} showHeader`) : undefined;
  const showSelectionLabel = kind === 'timeline' ? parseTimelineBoolean(node.attrs.showSelectionLabel, true, `${kind} showSelectionLabel`) : undefined;
  const showTimeLevel = kind === 'timeline' ? parseTimelineBoolean(node.attrs.showTimeLevel, true, `${kind} showTimeLevel`) : undefined;
  const showHorizontalScrollbar = kind === 'timeline' ? parseTimelineBoolean(node.attrs.showHorizontalScrollbar, true, `${kind} showHorizontalScrollbar`) : undefined;
  const scrollPosition = kind === 'timeline' && node.attrs.scrollPosition !== undefined
    ? parseTimelineDate(node.attrs.scrollPosition, `${kind} scrollPosition`)
    : undefined;
  const selection = cachePart?.selection;
  if (kind === 'timeline') validateTimelinePeriod(selection, `${kind} selection`);
  return {
    kind,
    id: `native:${kind}:${part}:${name}`,
    name,
    sheetPart,
    part,
    cachePart: cachePart?.part ?? '',
    cacheName,
    relationshipId,
    cacheRelationshipId: cachePart?.relationId ?? '',
    ...(table ? { pivotId: nativePivotId(table) } : {}),
    ...(field ? { fieldId: nativeFieldId(pivotCache!.cacheId, field.index), fieldIndex: field.index } : {}),
    ...(pivotCache ? { pivotCacheId: pivotCache.cacheId } : {}),
    ...(connectedPivotIds?.length ? { connectedPivotIds } : {}),
    ...(selection ? { selection } : {}),
    ...(timelineLevel ? { level: timelineLevel } : {}),
    ...(selectionLevel ? { selectionLevel } : {}),
    ...(showHeader === undefined ? {} : { showHeader }),
    ...(showSelectionLabel === undefined ? {} : { showSelectionLabel }),
    ...(showTimeLevel === undefined ? {} : { showTimeLevel }),
    ...(showHorizontalScrollbar === undefined ? {} : { showHorizontalScrollbar }),
    ...(scrollPosition === undefined ? {} : { scrollPosition }),
    ...(cachePart?.bounds ? { bounds: cachePart.bounds } : {}),
    ...(cachePart?.filterType ? { filterType: cachePart.filterType } : {}),
    ...(cachePart?.selectedItemIndexes ? { selectedItemIndexes: cachePart.selectedItemIndexes } : {}),
    ...(node.attrs.style ? { styleName: node.attrs.style } : {}),
    ...(node.attrs.caption ? { caption: node.attrs.caption } : {}),
    valid,
    ...(valid ? {} : { reason: `Unable to validate ${kind} cache, PivotTable, or field binding` }),
  };
}

function readControlDrawingAnchor(bytes: Uint8Array, name: string): { row: number; column: number } | undefined {
  const root = firstElement(parseXml(strFromU8(bytes)), 'wsDr');
  for (const anchor of root.children.filter((node) => {
    const local = localName(node.name);
    return local === 'twoCellAnchor' || local === 'oneCellAnchor' || local === 'absoluteAnchor';
  })) {
    if (!descendants(anchor, 'cNvPr').some((node) => node.attrs.name === name)) continue;
    const from = child(anchor, 'from');
    const row = Number(textContent(child(from, 'row')));
    const column = Number(textContent(child(from, 'col')));
    if (Number.isSafeInteger(row) && row >= 0 && Number.isSafeInteger(column) && column >= 0) return { row, column };
  }
  return undefined;
}

function nativeCacheIdentity(cacheId: number): string { return `native-cache:${cacheId}`; }

const TIMELINE_LEVELS: readonly PivotTimelineLevel[] = ['years', 'quarters', 'months', 'days'];
const TIMELINE_LEVEL_TO_XML: Record<PivotTimelineLevel, number> = { years: 0, quarters: 1, months: 2, days: 3 };

function parseTimelineLevel(value: string | undefined, label: string): PivotTimelineLevel {
  if (value === undefined) throw new Error(`${label} is required`);
  const level = Number(value);
  if (!Number.isInteger(level) || level < 0 || level >= TIMELINE_LEVELS.length) throw new Error(`${label} must be one of 0, 1, 2, or 3`);
  return TIMELINE_LEVELS[level]!;
}

function timelineLevelXml(value: PivotTimelineLevel, label: string): string {
  const level = TIMELINE_LEVEL_TO_XML[value];
  if (level === undefined) throw new Error(`${label} is invalid: ${value}`);
  return String(level);
}

function parseTimelineBoolean(value: string | undefined, defaultValue: boolean, label: string): boolean {
  if (value === undefined) return defaultValue;
  if (value === '1' || value.toLowerCase() === 'true') return true;
  if (value === '0' || value.toLowerCase() === 'false') return false;
  throw new Error(`${label} must be true, false, 1, or 0`);
}

function timelineBooleanXml(value: boolean): string { return value ? '1' : '0'; }

function parseTimelineFilterType(value: string | undefined, label: string): PivotTimelineFilterType | undefined {
  if (value === undefined) return undefined;
  if (value === 'unknown' || value === 'dateBetween' || value === 'dateNotBetween') return value;
  throw new Error(`${label} is unsupported: ${value}`);
}

function parseTimelineDate(value: string, label: string): string {
  if (pivotTimelineInstant(value) === undefined) throw new Error(`${label} is not a valid xsd:dateTime: ${value}`);
  return value;
}

function validateTimelinePeriod(period: { start?: string; end?: string } | undefined, label: string): void {
  if (!period) return;
  if (period.start !== undefined) parseTimelineDate(period.start, `${label}.startDate`);
  if (period.end !== undefined) parseTimelineDate(period.end, `${label}.endDate`);
  const start = period.start === undefined ? undefined : pivotTimelineInstant(period.start);
  const end = period.end === undefined ? undefined : pivotTimelineInstant(period.end);
  if (start !== undefined && end !== undefined && start > end) throw new Error(`${label} startDate must not be after endDate`);
}

function readTimelineBounds(node: XmlNode | undefined, label: string): { start: string; end: string } | undefined {
  if (!node) return undefined;
  const start = node.attrs.startDate;
  const end = node.attrs.endDate;
  if (!start || !end) throw new Error(`${label} requires startDate and endDate`);
  const bounds = { start: parseTimelineDate(start, `${label}.startDate`), end: parseTimelineDate(end, `${label}.endDate`) };
  validateTimelinePeriod(bounds, label);
  return bounds;
}

function nativeRefreshMode(flags: PivotNativeCacheFlags | NativePivotCacheDefinition): PivotRefreshPolicy['mode'] {
  if (flags.refreshOnSave === true) return 'on-change';
  if (flags.refreshOnLoad === true) return 'on-open';
  return 'manual';
}

function cacheFlagsSnapshot(cache: NativePivotCacheDefinition): PivotNativeCacheFlags {
  return {
    ...(cache.refreshOnLoad === undefined ? {} : { refreshOnLoad: cache.refreshOnLoad }),
    ...(cache.refreshOnSave === undefined ? {} : { refreshOnSave: cache.refreshOnSave }),
    ...(cache.saveData === undefined ? {} : { saveData: cache.saveData }),
    ...(cache.enableRefresh === undefined ? {} : { enableRefresh: cache.enableRefresh }),
  };
}

function refreshPolicyForNativeCache(cache: NativePivotCacheDefinition): PivotRefreshPolicy {
  const mode = nativeRefreshMode(cache);
  return normalizePivotRefreshPolicy({ mode, preserveFormatting: true, refreshOnLoad: mode !== 'manual' });
}

function cacheRefreshPolicyChanged(pivot: PivotDefinition): boolean {
  const flags = pivot.nativeMetadata?.cacheFlags;
  if (!flags) return true;
  return normalizePivotRefreshPolicy(pivot.refreshPolicy).mode !== nativeRefreshMode(flags);
}

/**
 * Apply the canonical mode to a native cache only when a Pivot policy was
 * edited. This preserves omitted XML attributes on an import→export no-op,
 * while ensuring edits reach both new and reused cache definitions.
 */
function synchronizeCacheRefreshPolicy(cache: NativePivotCacheDefinition, entries: readonly PivotDefinition[]): void {
  const policy = normalizePivotRefreshPolicy(entries[0]!.refreshPolicy);
  if (entries.some(cacheRefreshPolicyChanged)) {
    cache.refreshOnLoad = policy.mode !== 'manual';
    cache.refreshOnSave = refreshOnSaveForPivotMode(policy.mode);
    if (policy.mode !== 'manual') cache.enableRefresh = true;
  }
}

function isSlicerCacheRelation(relation: XlsxRelationship): boolean { return relation.type === REL_SLICER_CACHE_MODERN || relation.type === REL_SLICER_CACHE || relation.type.endsWith('/slicerCache'); }
function isSlicerRelation(relation: XlsxRelationship): boolean { return relation.type === REL_SLICER || relation.type.endsWith('/slicer'); }
function isTimelineCacheRelation(relation: XlsxRelationship): boolean { return relation.type === REL_TIMELINE_CACHE || relation.type === REL_TIMELINE_CACHE_ALT || relation.type.endsWith('/timelineCache') || relation.type.endsWith('/TimelineCache'); }
function isTimelineRelation(relation: XlsxRelationship): boolean { return relation.type === REL_TIMELINE || relation.type.endsWith('/timeline'); }

function pruneRemovedControlDrawingAnchors(
  files: Record<string, Uint8Array>,
  existing: NativePivotControlDefinition[],
  current: NativePivotControlDefinition[],
): void {
  const oldNamesByPart = new Map<string, Set<string>>();
  for (const control of existing) {
    if (!control.drawingPart || !control.name) continue;
    const names = oldNamesByPart.get(control.drawingPart) ?? new Set<string>();
    names.add(control.name);
    oldNamesByPart.set(control.drawingPart, names);
  }
  const currentNamesByPart = new Map<string, Set<string>>();
  for (const control of current) {
    if (!control.drawingPart || !control.name) continue;
    const names = currentNamesByPart.get(control.drawingPart) ?? new Set<string>();
    names.add(control.name);
    currentNamesByPart.set(control.drawingPart, names);
  }
  for (const [drawingPart, oldNames] of oldNamesByPart) {
    const bytes = files[drawingPart];
    if (!bytes) continue;
    const currentNames = currentNamesByPart.get(drawingPart) ?? new Set<string>();
    if ([...oldNames].every((name) => currentNames.has(name))) continue;
    const root = firstElement(parseXml(strFromU8(bytes)), 'wsDr');
    root.children = root.children.filter((node) => {
      const name = localName(node.name);
      if (name !== 'twoCellAnchor' && name !== 'oneCellAnchor' && name !== 'absoluteAnchor') return true;
      const drawingNames = descendants(node, 'cNvPr').flatMap((candidate) => candidate.attrs.name ? [candidate.attrs.name] : []);
      return !drawingNames.some((candidate) => oldNames.has(candidate) && !currentNames.has(candidate));
    });
    files[drawingPart] = strToU8(withXmlDeclaration(serializeXml(root)));
  }
}

/** Convert a supported native table/cache pair to the canonical Pivot definition. */
export function mapNativePivotDefinition(
  table: NativePivotTableDefinition,
  cache: NativePivotCacheDefinition,
  snapshot: WorkbookSnapshot,
  sheetPartById: Record<string, string>,
): PivotDefinition | undefined {
  const target = snapshot.sheets.find((sheet) => sheetPartById[sheet.id] === table.sheetPart);
  const source = mapNativeSource(cache.source, snapshot, sheetPartById);
  const location = target && table.locationRef ? parseRange(table.locationRef, target.id) : undefined;
  if (!target || !source || !location) return undefined;
  const fields = cache.fields.map((field) => {
    const fieldId = nativeFieldId(cache.cacheId, field.index);
    return {
      fieldId,
      name: field.name,
      dataType: mapFieldType(field.dataType),
      ordinal: field.index,
      ...(field.sharedItems ? { values: structuredClone(field.sharedItems) } : {}),
    };
  });
  const fieldId = (index: number): string => fields[index]?.fieldId ?? nativeFieldId(cache.cacheId, index);
  const mappedFilters = mapNativePivotFilters(table.pivotFilters ?? [], fields, table.dataFields, new Set(table.pageFields));
  const manualItemFilters = table.fields.flatMap((nativeField) => {
    const hidden = nativeField.hiddenItemIndexes ?? [];
    if (hidden.length === 0) return [];
    const cacheField = cache.fields[nativeField.index];
    if (!cacheField) throw new Error(`Pivot field ${nativeField.index} is outside cache field bounds`);
    const sharedItems = cacheField.sharedItems;
    if (!sharedItems) throw new Error(`Pivot field ${nativeField.index} hides items without cache sharedItems`);
    const memberKeys = hidden.map((itemIndex) => {
      if (itemIndex < 0 || itemIndex >= sharedItems.length) throw new Error(`Pivot field ${nativeField.index} hidden item ${itemIndex} is outside sharedItems bounds`);
      const value = sharedItems[itemIndex];
      if (value === undefined) throw new Error(`Pivot field ${nativeField.index} hidden item ${itemIndex} is missing`);
      return createPivotMemberKey(value);
    });
    return [{
      kind: 'manual' as const,
      family: 'manual' as const,
      fieldId: fieldId(nativeField.index),
      scope: table.pageFields.includes(nativeField.index)
        || (!table.rowFields.includes(nativeField.index) && !table.columnFields.includes(nativeField.index))
        ? 'report' as const
        : 'field' as const,
      mode: 'exclude' as const,
      memberKeys,
    }];
  });
  const pageFilters = table.pageFields.map((index) => {
    if (index < 0 || index >= fields.length) throw new Error(`Pivot page field ${index} is outside cache field bounds`);
    if (manualItemFilters.some((filter) => filter.fieldId === fieldId(index) && filter.scope === 'report')) return undefined;
    if (mappedFilters.filters.some((filter) => filter.fieldId === fieldId(index) && (filter.scope ?? 'report') === 'report')) return undefined;
    return { kind: 'manual' as const, family: 'manual' as const, fieldId: fieldId(index), scope: 'report' as const, mode: 'all' as const, memberKeys: [] };
  }).filter((filter): filter is NonNullable<typeof filter> => filter !== undefined);
  const preservedAutoSortScopes = table.fields.flatMap((field) => {
    if (!field.autoSortScope || nativePivotSort(field, table.dataFields, fieldId) || (field.sortType === undefined && field.nonAutoSortDefault === undefined)) return [];
    return [{
      fieldIndex: field.index,
      ...(field.sortType ? { sortType: field.sortType } : {}),
      ...(field.nonAutoSortDefault !== undefined ? { nonAutoSortDefault: field.nonAutoSortDefault } : {}),
      attributes: { ...field.autoSortScope.attributes },
      references: field.autoSortScope.references.map((reference) => ({ ...reference, ...(reference.itemIndexes ? { itemIndexes: [...reference.itemIndexes] } : {}) })),
    } satisfies PivotNativeAutoSortMetadata];
  });
  const placement = (index: number): PivotFieldPlacement => {
    const group = nativePivotGroup(cache.fields[index], cache, fieldId(index));
    const subtotal = nativePivotSubtotal(table.fields[index]);
    const sort = nativePivotSort(table.fields[index], table.dataFields, fieldId);
    return { fieldId: fieldId(index), ...(sort ? { sort } : {}), ...(group ? { group } : {}), ...(subtotal ? { subtotal } : {}) };
  };
  const layout: PivotLayout = {
    rows: table.rowFields.map(placement),
    columns: table.columnFields.map(placement),
    filters: [...pageFilters, ...manualItemFilters, ...mappedFilters.filters],
    allowMultipleFiltersPerField: table.multipleFieldFilters ?? true,
    collation: { ...DEFAULT_PIVOT_COLLATION },
    values: table.dataFields.map((data) => ({ fieldId: fieldId(data.field), summarizeBy: mapAggregate(data.subtotal), ...(data.name ? { displayName: data.name } : {}), ...(data.showDataAs ? { showAs: mapShowAs(data.showDataAs) } : {}) })),
    subtotalLocation: table.subtotalLocation ?? 'bottom',
    showGrandTotals: (table.showRowGrandTotals ?? true) || (table.showColumnGrandTotals ?? true),
    compact: table.compactData ?? table.fields.some((field) => field.compact === true),
    repeatLabels: table.repeatLabels ?? false,
    expansion: {
      expandedNodeIds: [],
      collapsedNodeIds: table.fields.flatMap((field) => field.axis === 'row'
        ? (field.collapsedItemIndexes ?? []).flatMap((index) => {
          const value = cache.fields[field.index]?.sharedItems?.[index];
          return value === undefined ? [] : [`${fieldId(field.index)}=${pivotMemberKey(createPivotMemberKey(value))}`];
        })
        : []),
      showButtons: table.showButtons ?? true,
    },
  };
  const fieldBindings: Record<string, { cacheFieldIndex: number; sourceName?: string }> = {};
  for (const field of fields) fieldBindings[field.fieldId] = { cacheFieldIndex: field.ordinal, sourceName: field.name };
  const refreshPolicy = { ...refreshPolicyForNativeCache(cache), ...(table.preserveFormatting === undefined ? {} : { preserveFormatting: table.preserveFormatting }) };
  const displayOptions = table.showFieldHeaders === undefined && table.fillEmptyCells === undefined && table.emptyCellText === undefined
    && table.showErrorValues === undefined && table.errorCellText === undefined
    ? undefined
    : normalizePivotDisplayOptions({
      ...(table.showFieldHeaders === undefined ? {} : { showFieldHeaders: table.showFieldHeaders }),
      ...(table.fillEmptyCells === undefined ? {} : { fillEmptyCells: table.fillEmptyCells }),
      ...(table.emptyCellText === undefined ? {} : { emptyCellText: table.emptyCellText }),
      ...(table.showErrorValues === undefined ? {} : { showErrorValues: table.showErrorValues }),
      ...(table.errorCellText === undefined ? {} : { errorCellText: table.errorCellText }),
    });
  return {
    schema: 'PivotDefinition',
    id: nativePivotId(table),
    source,
    target: { sheetId: target.id, anchor: { row: location.startRow, column: location.startColumn } },
    fieldCatalog: { schema: 'PivotFieldCatalog', fields },
    layout,
    refreshPolicy,
    ...(table.styleName || table.styleOptions ? {
      presentation: {
        ...(table.styleName ? { styleName: table.styleName } : {}),
        styleOptions: { ...DEFAULT_PIVOT_STYLE_OPTIONS, ...(table.styleOptions ?? {}) },
        ...(displayOptions ? { displayOptions } : {}),
      },
    } : {}),
    nativeMetadata: {
      cacheKey: nativeCacheIdentity(cache.cacheId),
      cacheId: cache.cacheId,
      cacheDefinitionPart: cache.part,
      ...(cache.recordsPart ? { cacheRecordsPart: cache.recordsPart } : {}),
      pivotTablePart: table.part,
      cacheFlags: cacheFlagsSnapshot(cache),
      fieldBindings,
      ...(mappedFilters.preserved.length ? { preservedPivotFilters: mappedFilters.preserved } : {}),
      ...(preservedAutoSortScopes.length ? { preservedAutoSortScopes } : {}),
    },
  };
}

export interface NativePivotFeatureStatus {
  pivot: boolean;
  slicer: boolean;
  timeline: boolean;
}

/** Report capabilities of the canonical shapes accepted by the native writer. */
export function nativePivotFeatureStatus(snapshot: WorkbookSnapshot, graph?: NativePivotGraph): NativePivotFeatureStatus {
  const pivots = snapshot.sheets.flatMap((sheet) => sheet.pivots);
  const pivot = Boolean(graph) && pivots.some((candidate) => {
    const nativeTable = graph?.tables.some((table) => table.pivotId === candidate.id || nativePivotId(table) === candidate.id);
    return nativeTable && (candidate.source.kind === 'worksheet-range' || candidate.source.kind === 'table');
  });
  const controls = snapshot.sheets.flatMap((sheet) => (sheet.drawings ?? []).flatMap((drawing) => {
    const payload = sheet.drawingPayloads[drawing.payloadId];
    return payload && (payload.kind === 'slicer' || payload.kind === 'timeline') ? [payload] : [];
  }));
  const exportable = (payload: { kind: 'slicer' | 'timeline'; pivotId: string; fieldId: string }): boolean => {
    const target = pivots.find((candidate) => candidate.id === payload.pivotId);
    const field = target?.fieldCatalog.fields.find((candidate) => candidate.fieldId === payload.fieldId);
    return Boolean(target && (target.source.kind === 'worksheet-range' || target.source.kind === 'table') && field && (payload.kind === 'slicer' || field.dataType === 'date'));
  };
  const slicers = controls.filter((control) => control.kind === 'slicer');
  const timelines = controls.filter((control) => control.kind === 'timeline');
  const graphControls = graph?.controls ?? [];
  return {
    pivot,
    slicer: slicers.length > 0 && slicers.every(exportable) && (graph ? graphControls.filter((control) => control.kind === 'slicer' && control.valid).length >= slicers.length : true),
    timeline: timelines.length > 0 && timelines.every(exportable) && (graph ? graphControls.filter((control) => control.kind === 'timeline' && control.valid).length >= timelines.length : true),
  };
}

/** Rebuild reachable native parts and relationships from current canonical pivots. */
export function synchronizeNativePivotPackage(input: NativePivotPackageWriteInput): NativePivotPackageUpdate {
  const files = cloneFiles(input.files);
  const baseRelationships = cloneRelationships(input.relationships);
  const existing = input.graph ?? { schema: 'NativePivotGraph' as const, caches: [], tables: [] };
  const existingCaches = new Map(existing.caches.map((cache) => [cache.cacheId, cache]));
  const existingTables = new Map(existing.tables.map((table) => [table.part, table]));
  const usedTables = new Set<string>();
  const usedCaches = new Set<number>();
  const caches: NativePivotCacheDefinition[] = [];
  const tables: NativePivotTableDefinition[] = [];
  const displayCellsBySheetPart: NativePivotPackageUpdate['displayCellsBySheetPart'] = {};
  const partNumbers = nextPartNumbers(files);
  let nextCacheId = Math.max(0, ...existing.caches.map((cache) => cache.cacheId)) + 1;
  const cacheBySource = new Map<string, NativePivotCacheDefinition>();
  const cacheByIdentity = new Map<string, NativePivotCacheDefinition>();
  const pivotEntries = input.snapshot.sheets.flatMap((sheet) => sheet.pivots.map((pivot) => ({ sheet, pivot })));
  type PlannedPivotEntry = {
    targetSheet: typeof pivotEntries[number]['sheet'];
    pivot: PivotDefinition;
    targetPart: string;
    sourceInfo: { key: string; source: NativePivotSource; sheet: SheetSnapshot; range: RangeRef; tableName?: string };
    oldTable?: NativePivotTableDefinition;
    cache: NativePivotCacheDefinition;
  };
  const plannedEntries: PlannedPivotEntry[] = [];

  // First resolve every cache assignment. Policy validation happens before any
  // OOXML part is rebuilt, so a shared-cache conflict fails atomically.
  for (const { sheet: targetSheet, pivot: inputPivot } of pivotEntries) {
    const pivot = normalizePivot(inputPivot);
    if (!pivot) continue;
    const targetPart = input.sheetPartById[pivot.target.sheetId] ?? input.sheetPartById[targetSheet.id];
    const sourceInfo = targetPart ? resolveCanonicalSource(pivot.source, input.snapshot, input.sheetPartById) : undefined;
    const oldTable = pivot.nativeMetadata?.pivotTablePart ? existingTables.get(pivot.nativeMetadata.pivotTablePart) : existing.tables.find((table) => nativePivotId(table) === pivot.id);
    if (!sourceInfo) {
      // External/OLAP/consolidation and unresolved sources stay opaque only if
      // they already have a reachable part; no empty native Pivot is emitted.
      if (oldTable) {
        tables.push(structuredClone(oldTable));
        usedTables.add(oldTable.part);
        usedCaches.add(oldTable.cacheId);
      }
      continue;
    }
    const cacheKey = `${sourceInfo.key}|${pivotGroupingKey(pivot)}`;
    const requestedIdentity = pivot.nativeMetadata?.cacheKey ?? cacheKey;
    let cache = pivot.nativeMetadata?.cacheId === undefined ? undefined : existingCaches.get(pivot.nativeMetadata.cacheId);
    if (!cache || cache.source.kind === 'unsupported' || nativeSourceKey(cache.source) !== sourceInfo.key) cache = cacheByIdentity.get(requestedIdentity);
    if (!cache || cache.source.kind === 'unsupported' || nativeSourceKey(cache.source) !== sourceInfo.key) cache = cacheBySource.get(cacheKey);
    if (!cache) {
      const policy = normalizePivotRefreshPolicy(pivot.refreshPolicy);
      cache = {
        cacheId: nextCacheId++,
        part: `xl/pivotCache/pivotCacheDefinition${partNumbers.cacheDefinition++}.xml`,
        recordsPart: `xl/pivotCache/pivotCacheRecords${partNumbers.records++}.xml`,
        source: sourceInfo.source,
        fields: [],
        refreshOnLoad: policy.mode !== 'manual',
        refreshOnSave: refreshOnSaveForPivotMode(policy.mode),
        saveData: true,
        enableRefresh: true,
      };
    }
    cacheBySource.set(cacheKey, cache);
    cacheByIdentity.set(requestedIdentity, cache);
    cacheByIdentity.set(cacheKey, cache);
    plannedEntries.push({ targetSheet, pivot, targetPart: targetPart!, sourceInfo, ...(oldTable ? { oldTable } : {}), cache });
  }

  const policyByCache = new Map<number, { mode: PivotRefreshPolicy['mode']; pivots: PivotDefinition[] }>();
  for (const entry of plannedEntries) {
    const policy = normalizePivotRefreshPolicy(entry.pivot.refreshPolicy);
    const current = policyByCache.get(entry.cache.cacheId);
    if (current && current.mode !== policy.mode) {
      throw new Error(`Pivot cache ${entry.cache.cacheId} has conflicting refresh policies: ${current.mode} and ${policy.mode}`);
    }
    if (current) current.pivots.push(entry.pivot);
    else policyByCache.set(entry.cache.cacheId, { mode: policy.mode, pivots: [entry.pivot] });
  }
  for (const entry of plannedEntries) {
    const policy = policyByCache.get(entry.cache.cacheId);
    if (policy) synchronizeCacheRefreshPolicy(entry.cache, policy.pivots);
  }

  for (const { targetSheet, pivot, targetPart, sourceInfo, oldTable, cache } of plannedEntries) {
    usedCaches.add(cache.cacheId);
    const sourceRows = readSourceRows(sourceInfo.sheet, sourceInfo.range, pivot, sourceInfo.tableName);
    cache.source = sourceInfo.source;
    const previousFields = cache.fields;
    const groupedOrdinals = new Set([...pivot.layout.rows, ...pivot.layout.columns].flatMap((placement) => {
      if (!placement.group) return [];
      const field = pivot.fieldCatalog.fields.find((candidate) => candidate.fieldId === placement.fieldId || candidate.name === placement.fieldId);
      return field ? [field.ordinal] : [];
    }));
    cache.fields = sourceRows.fields.map((field, index) => ({
      index,
      name: field.name,
      dataType: field.dataType,
      sharedItems: uniqueScalars(sourceRows.rows.map((row) => row[index] ?? null)),
      ...(previousFields[index]?.fieldGroup && (!nativePivotGroup(previousFields[index], cache, nativeFieldId(cache.cacheId, index)) || groupedOrdinals.has(index))
        ? { fieldGroup: structuredClone(previousFields[index]!.fieldGroup) } : {}),
    }));
    applyCanonicalPivotGroups(cache, pivot);
    cache.recordCount = sourceRows.rows.length;
    cache.recordsPart ??= `xl/pivotCache/pivotCacheRecords${partNumbers.records++}.xml`;
    files[cache.part] = strToU8(buildCacheDefinitionXml(cache));
    files[cache.recordsPart] = strToU8(buildCacheRecordsXml(cache, sourceRows.rows));
    if (!caches.some((candidate) => candidate.cacheId === cache!.cacheId)) caches.push(cache);
    const tablePart = oldTable?.part ?? `xl/pivotTables/pivotTable${partNumbers.table++}.xml`;
    const table = buildNativeTable(pivot, cache, tablePart, targetPart!, oldTable, sourceRows);
    tables.push(table);
    if (oldTable) usedTables.add(oldTable.part);
    files[tablePart] = strToU8(buildPivotTableXml(table));
    const sheetCells = displayCellsBySheetPart[targetPart!] ??= {};
    mergeDisplayCells(sheetCells, buildDisplayCells(pivot, table, sourceRows), targetSheet);
  }

  // Keep only unsupported old tables/caches. A supported table missing from
  // the canonical snapshot is a deletion and all unreachable parts are pruned.
  for (const oldTable of existing.tables) {
    if (usedTables.has(oldTable.part)) continue;
    const cache = existingCaches.get(oldTable.cacheId);
    if (cache?.source.kind !== 'unsupported') continue;
    tables.push(structuredClone(oldTable));
    usedTables.add(oldTable.part);
    usedCaches.add(cache.cacheId);
    if (!caches.some((candidate) => candidate.cacheId === cache.cacheId)) caches.push(structuredClone(cache));
  }
  for (const cache of existing.caches) if (usedCaches.has(cache.cacheId) && !caches.some((candidate) => candidate.cacheId === cache.cacheId)) caches.push(structuredClone(cache));
  const reachable = new Set([...caches.flatMap((cache) => [cache.part, ...(cache.recordsPart ? [cache.recordsPart] : [])]), ...tables.map((table) => table.part)]);
  for (const name of Object.keys(files)) if (isNativePivotPart(name) && !reachable.has(name)) delete files[name];
  const pivotRelationships = rebuildRelationships(baseRelationships, caches, tables, files);
  const controlUpdate = synchronizeNativeControls({
    files,
    relationships: pivotRelationships,
    existing: existing.controls ?? [],
    snapshot: input.snapshot,
    sheetPartById: input.sheetPartById,
    caches,
    tables,
  });
  const relationships = controlUpdate.relationships;
  for (const cache of caches) {
    if (!files[cache.part]) throw new Error(`Native Pivot cache part is missing: ${cache.part}`);
    if (cache.recordsPart && !files[cache.recordsPart]) throw new Error(`Native Pivot records part is missing: ${cache.recordsPart}`);
  }
  for (const table of tables) if (!files[table.part]) throw new Error(`Native PivotTable part is missing: ${table.part}`);
  return { graph: { schema: 'NativePivotGraph', caches, tables, ...(controlUpdate.controls.length ? { controls: controlUpdate.controls } : {}) }, files: controlUpdate.files, relationships, displayCellsBySheetPart };
}

interface NativeControlSyncInput {
  files: Record<string, Uint8Array>;
  relationships: Record<string, XlsxRelationship[]>;
  existing: NativePivotControlDefinition[];
  snapshot: WorkbookSnapshot;
  sheetPartById: Record<string, string>;
  caches: NativePivotCacheDefinition[];
  tables: NativePivotTableDefinition[];
}

interface NativeControlSyncResult {
  files: Record<string, Uint8Array>;
  relationships: Record<string, XlsxRelationship[]>;
  controls: NativePivotControlDefinition[];
}

function synchronizeNativeControls(input: NativeControlSyncInput): NativeControlSyncResult {
  const files = input.files;
  const originalRelationships = cloneRelationships(input.relationships);
  const relationships = cloneRelationships(input.relationships);
  const controls: NativePivotControlDefinition[] = [];
  const partNumbers = nextControlPartNumbers(files);
  const entries = input.snapshot.sheets.flatMap((sheet) => (sheet.drawings ?? []).flatMap((drawing) => {
    const payload = sheet.drawingPayloads[drawing.payloadId];
    return payload && (payload.kind === 'slicer' || payload.kind === 'timeline') ? [{ sheet, drawing, payload }] : [];
  }));
  const oldById = new Map(input.existing.map((control) => [control.id, control]));
  type NativeControlDrawingEntry = { control: NativePivotControlDefinition; drawing: SheetSnapshot['drawings'][number] };
  const drawingEntries = new Map<string, NativeControlDrawingEntry[]>();
  for (const entry of entries) {
    const pivot = input.snapshot.sheets.flatMap((sheet) => sheet.pivots).find((candidate) => candidate.id === entry.payload.pivotId);
    const table = pivot ? input.tables.find((candidate) => candidate.pivotId === pivot.id || candidate.part === pivot.nativeMetadata?.pivotTablePart) : undefined;
    const field = pivot?.fieldCatalog.fields.find((candidate) => candidate.fieldId === entry.payload.fieldId);
    const cacheFieldIndex = pivot?.nativeMetadata?.fieldBindings?.[entry.payload.fieldId]?.cacheFieldIndex ?? field?.ordinal;
    const cache = table ? input.caches.find((candidate) => candidate.cacheId === table.cacheId) : undefined;
    if (!pivot || !table || !cache || !field || cacheFieldIndex === undefined || !input.sheetPartById[entry.sheet.id] || (entry.payload.kind === 'timeline' && field.dataType !== 'date')) continue;
    const old = oldById.get(entry.drawing.id);
    const kind = entry.payload.kind;
    const name = old?.name ?? safeControlName(entry.drawing.id, kind);
    const cacheName = old?.cacheName ?? `${kind === 'slicer' ? 'Slicer' : 'NativeTimeline'}_${safeControlName(entry.drawing.id, kind)}`;
    const cachePart = old?.cachePart || (kind === 'slicer' ? `xl/slicerCaches/slicerCache${partNumbers.slicerCache++}.xml` : `xl/timelineCaches/timelineCache${partNumbers.timelineCache++}.xml`);
    const part = old?.part || (kind === 'slicer' ? `xl/slicers/slicer${partNumbers.slicer++}.xml` : `xl/timelines/timeline${partNumbers.timeline++}.xml`);
    const connectedPivotIds = [...new Set([entry.payload.pivotId, ...(entry.payload.connectedPivotIds ?? [])])];
    const control: NativePivotControlDefinition = {
      kind, id: entry.drawing.id, name, sheetPart: input.sheetPartById[entry.sheet.id]!, part, cachePart, cacheName,
      relationshipId: old?.relationshipId ?? '', cacheRelationshipId: old?.cacheRelationshipId ?? '',
      ...(old?.drawingPart ? { drawingPart: old.drawingPart } : {}), ...(old?.drawingRelationshipId ? { drawingRelationshipId: old.drawingRelationshipId } : {}),
      pivotId: entry.payload.pivotId, fieldId: entry.payload.fieldId, fieldIndex: cacheFieldIndex, pivotCacheId: table.cacheId, connectedPivotIds,
      valid: true,
    };
    if (entry.payload.kind === 'timeline') {
      control.level = entry.payload.level;
      control.selectionLevel = entry.payload.selectionLevel;
      control.showHeader = entry.payload.showHeader;
      control.showSelectionLabel = entry.payload.showSelectionLabel;
      control.showTimeLevel = entry.payload.showTimeLevel;
      control.showHorizontalScrollbar = entry.payload.showHorizontalScrollbar;
      if (entry.payload.scrollPosition !== undefined) control.scrollPosition = parseTimelineDate(entry.payload.scrollPosition, `Timeline ${control.id} scrollPosition`);
      control.bounds = entry.payload.bounds;
      control.filterType = entry.payload.filterType;
      if (entry.payload.caption !== undefined) control.caption = entry.payload.caption;
      if (entry.payload.styleName !== undefined) control.styleName = entry.payload.styleName;
    }
    if (entry.payload.kind === 'slicer') {
      files[cachePart] = strToU8(buildSlicerCacheXml(control, entry.payload, pivot, cache, input.tables));
      files[part] = strToU8(buildSlicerXml(control, entry.payload));
    } else {
      files[cachePart] = strToU8(buildTimelineCacheXml(control, entry.payload, pivot, input.tables, entry.sheet, old?.cachePart ? input.files[old.cachePart] : undefined));
      files[part] = strToU8(buildTimelineXml(control, entry.payload, old?.part ? input.files[old.part] : undefined));
    }
    controls.push(control);
    const list = drawingEntries.get(control.sheetPart) ?? [];
    list.push({ control, drawing: entry.drawing });
    drawingEntries.set(control.sheetPart, list);
  }
  // Invalid imported controls stay byte-preserved and explicitly marked; valid
  // controls absent from canonical drawings are deletions.
  for (const old of input.existing) {
    if (controls.some((control) => control.id === old.id)) continue;
    const oldTable = old.pivotId ? input.tables.find((table) => nativePivotId(table) === old.pivotId) : undefined;
    const oldCache = oldTable ? input.caches.find((cache) => cache.cacheId === oldTable.cacheId) : undefined;
    if (old.part && old.cachePart && (!old.valid || oldCache?.source.kind === 'unsupported')) controls.push(structuredClone(old));
  }
  for (const name of Object.keys(files)) {
    if (!isNativeControlPart(name)) continue;
    if (!controls.some((control) => control.part === name || control.cachePart === name)) delete files[name];
  }
  for (const [sheetPart, entriesForSheet] of drawingEntries) {
    const current = relationships[sheetPart] ?? [];
    let drawingRelation = current.find((relation) => relation.type === REL_DRAWING || relation.type.endsWith('/drawing'));
    let drawingPart = drawingRelation ? resolveTarget(sheetPart, drawingRelation.target) : `xl/drawings/drawing${partNumbers.drawing++}.xml`;
    if (!drawingRelation) {
      drawingRelation = { id: allocateId(current), type: REL_DRAWING, target: relativeTarget(sheetPart, drawingPart) };
      relationships[sheetPart] = [...current, drawingRelation];
    }
    const oldDrawing = files[drawingPart];
    const newEntries = oldDrawing ? entriesForSheet.filter((entry) => !strFromU8(oldDrawing).includes(`name="${encodeXml(entry.control.name)}"`)) : entriesForSheet;
    if (newEntries.length) {
      const drawingXml = buildControlDrawingXml(newEntries);
      files[drawingPart] = oldDrawing ? appendControlDrawingXml(oldDrawing, drawingXml) : strToU8(drawingXml);
    }
    for (const entry of entriesForSheet) {
      entry.control.drawingPart = drawingPart;
      entry.control.drawingRelationshipId = drawingRelation.id;
    }
  }
  pruneRemovedControlDrawingAnchors(files, input.existing, controls);
  for (const [source, list] of Object.entries(relationships)) {
    relationships[source] = list.filter((relation) => source === 'xl/workbook.xml' ? !isSlicerCacheRelation(relation) && !isTimelineCacheRelation(relation) : !isSlicerRelation(relation) && !isTimelineRelation(relation));
  }
  for (const control of controls) {
    const workbookList = relationships['xl/workbook.xml'] ?? [];
    const oldCache = originalRelationships['xl/workbook.xml']?.find((relation) => (isSlicerCacheRelation(relation) || isTimelineCacheRelation(relation)) && resolveTarget('xl/workbook.xml', relation.target) === control.cachePart);
    control.cacheRelationshipId = oldCache?.id ?? allocateId(workbookList);
    relationships['xl/workbook.xml'] = [...workbookList, { id: control.cacheRelationshipId, type: control.kind === 'slicer' ? REL_SLICER_CACHE_MODERN : REL_TIMELINE_CACHE, target: relativeTarget('xl/workbook.xml', control.cachePart) }];
    const sheetList = relationships[control.sheetPart] ?? [];
    const oldPart = originalRelationships[control.sheetPart]?.find((relation) => (isSlicerRelation(relation) || isTimelineRelation(relation)) && resolveTarget(control.sheetPart, relation.target) === control.part);
    control.relationshipId = oldPart?.id ?? allocateId(sheetList);
    relationships[control.sheetPart] = [...sheetList, { id: control.relationshipId, type: control.kind === 'slicer' ? REL_SLICER : REL_TIMELINE, target: relativeTarget(control.sheetPart, control.part) }];
  }
  return { files, relationships, controls };
}

function buildSlicerCacheXml(control: NativePivotControlDefinition, payload: PivotSlicerDrawingPayload, pivot: PivotModel, cache: NativePivotCacheDefinition, tables: NativePivotTableDefinition[]): string {
  const field = cache.fields[control.fieldIndex ?? -1];
  const values = field?.sharedItems ?? pivot.fieldCatalog.fields.find((candidate) => candidate.fieldId === payload.fieldId)?.values ?? [];
  const selected = new Set(payload.filter.memberKeys.map((value) => `${value.type}:${JSON.stringify(value.value)}`));
  const items = values.map((value, index) => {
    const key = createPivotMemberKey(value);
    const keyValue = `${key.type}:${JSON.stringify(key.value)}`;
    const checked = payload.filter.mode === 'include' ? selected.has(keyValue) : payload.filter.mode === 'exclude' ? !selected.has(keyValue) : false;
    return `<i x="${index}"${checked ? ' s="1"' : ''}/>`;
  }).join('');
  const connected = [...new Set([payload.pivotId, ...(payload.connectedPivotIds ?? [])])].flatMap((pivotId) => { const table = tables.find((candidate) => candidate.pivotId === pivotId || candidate.name === pivotId); return table ? [`<pivotTable tabId="1" name="${encodeXml(table.name)}"/>`] : []; }).join('');
  return withXmlDeclaration(`<slicerCacheDefinition xmlns="${NS_X14}" xmlns:x="${NS_MAIN}" name="${encodeXml(control.cacheName)}" sourceName="${encodeXml(field?.name ?? payload.fieldId)}"><pivotTables>${connected}</pivotTables><data><tabular pivotCacheId="${control.pivotCacheId ?? cache.cacheId}"><items count="${values.length}">${items}</items></tabular></data></slicerCacheDefinition>`);
}

function buildSlicerXml(control: NativePivotControlDefinition, payload: { kind: 'slicer'; fieldId: string },): string {
  void payload;
  return withXmlDeclaration(`<slicers xmlns="${NS_X14}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="${NS_MAIN}" mc:Ignorable="x"><slicer name="${encodeXml(control.name)}" cache="${encodeXml(control.cacheName)}" caption="${encodeXml(control.name)}" rowHeight="228600"/></slicers>`);
}

function buildTimelineCacheXml(
  control: NativePivotControlDefinition,
  payload: PivotTimelineDrawingPayload,
  pivot: PivotModel,
  tables: NativePivotTableDefinition[],
  sheet: SheetSnapshot,
  originalBytes?: Uint8Array,
): string {
  const field = pivot.fieldCatalog.fields.find((candidate) => candidate.fieldId === payload.fieldId);
  const connected = [...new Set([payload.pivotId, ...(payload.connectedPivotIds ?? [])])].flatMap((pivotId) => { const table = tables.find((candidate) => candidate.pivotId === pivotId || candidate.name === pivotId); return table ? [`<pivotTable tabId="1" name="${encodeXml(table.name)}"/>`] : []; }).join('');
  const bounds = resolveTimelineBounds(payload.bounds, pivot, payload.fieldId, sheet);
  validateTimelinePeriod(payload.period, `Timeline ${control.id} selection`);
  const filterType = payload.filterType;
  const root = originalBytes
    ? firstElement(parseXml(strFromU8(originalBytes)), 'timelineCacheDefinition')
    : firstElement(parseXml(`<timelineCacheDefinition xmlns="${NS_X15}" xmlns:x15="${NS_X15}" name="${encodeXml(control.cacheName)}" sourceName="${encodeXml(field?.name ?? payload.fieldId)}"><pivotTables/><state/></timelineCacheDefinition>`), 'timelineCacheDefinition');
  root.attrs.name = control.cacheName;
  root.attrs.sourceName = field?.name ?? payload.fieldId;
  const pivotTables = child(root, 'pivotTables') ?? appendTimelineChild(root, 'pivotTables');
  pivotTables.children = connected ? childrenFromXml(`<pivotTables>${connected}</pivotTables>`, 'pivotTables').children : [];
  const state = child(root, 'state') ?? appendTimelineChild(root, 'state');
  state.attrs.pivotCacheId = String(control.pivotCacheId ?? 0);
  state.attrs.filterType = filterType;
  const selection = child(state, 'selection') ?? appendTimelineChild(state, 'selection');
  delete selection.attrs.startDate;
  delete selection.attrs.endDate;
  if (payload.period.start !== undefined) selection.attrs.startDate = parseTimelineDate(payload.period.start, `Timeline ${control.id} selection.startDate`);
  if (payload.period.end !== undefined) selection.attrs.endDate = parseTimelineDate(payload.period.end, `Timeline ${control.id} selection.endDate`);
  const boundsNode = child(state, 'bounds') ?? appendTimelineChild(state, 'bounds');
  boundsNode.attrs.startDate = bounds.start;
  boundsNode.attrs.endDate = bounds.end;
  return withXmlDeclaration(serializeXml(root));
}

function buildTimelineXml(control: NativePivotControlDefinition, payload: PivotTimelineDrawingPayload, originalBytes?: Uint8Array): string {
  const root = originalBytes
    ? firstElement(parseXml(strFromU8(originalBytes)), 'timelines')
    : firstElement(parseXml(`<timelines xmlns="${NS_X15}" xmlns:x15="${NS_X15}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="${NS_MAIN}" mc:Ignorable="x"><timeline/></timelines>`), 'timelines');
  const timeline = children(root, 'timeline').find((candidate) => candidate.attrs.name === control.name) ?? children(root, 'timeline')[0] ?? appendTimelineChild(root, 'timeline');
  timeline.attrs.name = control.name;
  timeline.attrs.cache = control.cacheName;
  if (payload.caption === undefined) delete timeline.attrs.caption;
  else timeline.attrs.caption = payload.caption;
  timeline.attrs.showHeader = timelineBooleanXml(payload.showHeader);
  timeline.attrs.showSelectionLabel = timelineBooleanXml(payload.showSelectionLabel);
  timeline.attrs.showTimeLevel = timelineBooleanXml(payload.showTimeLevel);
  timeline.attrs.showHorizontalScrollbar = timelineBooleanXml(payload.showHorizontalScrollbar);
  timeline.attrs.level = timelineLevelXml(payload.level, `Timeline ${control.id} level`);
  timeline.attrs.selectionLevel = timelineLevelXml(payload.selectionLevel, `Timeline ${control.id} selectionLevel`);
  if (payload.scrollPosition === undefined) delete timeline.attrs.scrollPosition;
  else timeline.attrs.scrollPosition = parseTimelineDate(payload.scrollPosition, `Timeline ${control.id} scrollPosition`);
  if (payload.styleName === undefined) delete timeline.attrs.style;
  else timeline.attrs.style = payload.styleName;
  return withXmlDeclaration(serializeXml(root));
}

function appendTimelineChild(parent: XmlNode, name: string): XmlNode {
  const node: XmlNode = { name, attrs: {}, children: [], text: '' };
  parent.children.push(node);
  return node;
}

function childrenFromXml(source: string, name: string): XmlNode {
  return firstElement(parseXml(source), name);
}

function resolveTimelineBounds(
  configured: { start?: string; end?: string },
  pivot: PivotModel,
  fieldId: string,
  sheet: SheetSnapshot,
): { start: string; end: string } {
  if ((configured.start === undefined) !== (configured.end === undefined)) throw new Error(`Timeline bounds for ${fieldId} must contain both start and end`);
  if (configured.start !== undefined && configured.end !== undefined) {
    const start = parseTimelineDate(configured.start, `Timeline ${fieldId} bounds.startDate`);
    const end = parseTimelineDate(configured.end, `Timeline ${fieldId} bounds.endDate`);
    validateTimelinePeriod({ start, end }, `Timeline ${fieldId} bounds`);
    return { start, end };
  }
  const field = pivot.fieldCatalog.fields.find((candidate) => candidate.fieldId === fieldId);
  const values = field?.values?.length ? [...field.values] : timelineSourceValues(pivot, field?.ordinal ?? -1, sheet);
  const instants = values.map((value) => pivotTimelineInstant(value)).filter((value): value is number => value !== undefined);
  if (!instants.length) throw new Error(`Timeline ${fieldId} bounds cannot be derived from a typed date-member domain`);
  const start = new Date(Math.min(...instants)).toISOString();
  const end = new Date(Math.max(...instants)).toISOString();
  return { start, end };
}

function timelineSourceValues(pivot: PivotModel, fieldOrdinal: number, sheet: SheetSnapshot): PivotScalar[] {
  if (fieldOrdinal < 0) return [];
  let range: RangeRef | undefined;
  const source = pivot.source;
  if (source.kind === 'worksheet-range' && source.range.sheetId === sheet.id) range = source.range;
  if (source.kind === 'table') {
    const table = sheet.sheetTables?.find((candidate) => candidate.id === source.tableId || candidate.name === source.tableId);
    range = table?.range;
  }
  if (!range) return [];
  const values: PivotScalar[] = [];
  const column = range.startColumn + fieldOrdinal;
  for (let row = range.startRow + 1; row <= range.endRow; row += 1) {
    const cell = sheet.cells[String(row)]?.[String(column)];
    const value = cell?.formulaValue ?? cell?.value;
    if (typeof value === 'string' || typeof value === 'number') values.push(value);
  }
  return values;
}

function buildControlDrawingXml(entries: Array<{ control: NativePivotControlDefinition; drawing: { anchor: unknown; transform: { width: number; height: number } } }>): string {
  const anchors = entries.map((entry, index) => buildControlDrawingAnchor(entry.control, entry.drawing, index)).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:r="${NS_DOC_REL}">${anchors}</xdr:wsDr>`;
}

function buildControlDrawingAnchor(control: NativePivotControlDefinition, drawing: { anchor: unknown; transform: { width: number; height: number } }, index: number): string {
  const anchor = drawing.anchor as { kind?: string; row?: number; column?: number };
  const row = anchor.row ?? index * 8;
  const column = anchor.column ?? 0;
  const width = Math.max(1, Math.round(drawing.transform.width / 80));
  const height = Math.max(2, Math.round(drawing.transform.height / 20));
  const requires = control.kind === 'slicer' ? 'x14' : 'tsle';
  const namespace = control.kind === 'slicer' ? NS_SLICER_DRAWING : NS_TIMELINE_DRAWING;
  const prefix = control.kind === 'slicer' ? 'sle' : 'tsle';
  return `<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>${column}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>${column + width}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${row + height}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><mc:AlternateContent><mc:Choice Requires="${requires}" xmlns:${requires}="${namespace}"><xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="${index + 1}" name="${encodeXml(control.name)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="${namespace}"><${prefix}:${control.kind === 'slicer' ? 'slicer' : 'timeline'} xmlns:${prefix}="${namespace}" name="${encodeXml(control.name)}"/></a:graphicData></a:graphic></xdr:graphicFrame></mc:Choice><mc:Fallback><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="${index + 1}" name="${encodeXml(control.name)}"/><xdr:cNvSpPr><a:spLocks noTextEdit="1"/></xdr:cNvSpPr></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:sp></mc:Fallback></mc:AlternateContent><xdr:clientData/></xdr:twoCellAnchor>`;
}

function appendControlDrawingXml(existing: Uint8Array, generated: string): Uint8Array {
  const root = firstElement(parseXml(strFromU8(existing)), 'wsDr');
  const generatedRoot = firstElement(parseXml(generated), 'wsDr');
  root.children.push(...generatedRoot.children);
  return strToU8(withXmlDeclaration(serializeXml(root)));
}

function safeControlName(value: string, kind: 'slicer' | 'timeline'): string {
  const normalized = value.replace(/[^A-Za-z0-9_]/g, '_').slice(0, 200) || `${kind}_control`;
  return normalized;
}

function nextControlPartNumbers(files: Record<string, Uint8Array>): { slicerCache: number; timelineCache: number; slicer: number; timeline: number; drawing: number } {
  const max = (pattern: RegExp): number => Object.keys(files).reduce((result, name) => Math.max(result, Number(name.match(pattern)?.[1] ?? 0)), 0);
  return { slicerCache: max(/slicerCache(\d+)\.xml$/i) + 1, timelineCache: max(/timelineCache(\d+)\.xml$/i) + 1, slicer: max(/slicer(\d+)\.xml$/i) + 1, timeline: max(/timeline(\d+)\.xml$/i) + 1, drawing: max(/drawing(\d+)\.xml$/i) + 1 };
}

/** Validate and patch source sheet names for callers that only have a graph. */
export function synchronizeNativePivotGraph(input: NativePivotWriteInput): Record<string, Uint8Array> {
  const files = cloneFiles(input.files);
  for (const cache of input.graph.caches) {
    if (cache.source.kind !== 'worksheet-range' || !cache.source.sheetPart || !input.sheetNameByPart) continue;
    const root = firstElement(parseXml(strFromU8(files[cache.part]!)), 'pivotCacheDefinition');
    const source = child(child(root, 'cacheSource'), 'worksheetSource');
    if (!source) throw new Error(`Pivot cache ${cache.part} has no worksheetSource`);
    const name = input.sheetNameByPart[cache.source.sheetPart];
    if (!name) throw new Error(`Pivot cache source worksheet part is missing from current workbook: ${cache.source.sheetPart}`);
    source.attrs.sheet = name;
    files[cache.part] = strToU8(withXmlDeclaration(serializeXml(root)));
  }
  for (const cache of input.graph.caches) {
    if (!files[cache.part]) throw new Error(`Native Pivot cache part is missing: ${cache.part}`);
    if (cache.recordsPart && !files[cache.recordsPart]) throw new Error(`Native Pivot records part is missing: ${cache.recordsPart}`);
  }
  for (const table of input.graph.tables) if (!files[table.part]) throw new Error(`Native PivotTable part is missing: ${table.part}`);
  return files;
}

export function serializeNativePivotCaches(graph: NativePivotGraph, relationships: XlsxRelationship[]): string {
  const items = graph.caches.map((cache) => {
    const relation = relationships.find((candidate) => candidate.type.endsWith('/pivotCacheDefinition') && resolveTarget('xl/workbook.xml', candidate.target) === cache.part);
    if (!relation) throw new Error(`Native Pivot cache relation is missing for ${cache.part}`);
    return `<pivotCache cacheId="${cache.cacheId}" r:id="${encodeXml(relation.id)}"/>`;
  }).join('');
  return `<pivotCaches count="${graph.caches.length}" xmlns:r="${NS_DOC_REL}">${items}</pivotCaches>`;
}

function mapNativeSource(source: NativePivotSource | { kind: 'unsupported'; reason: string }, snapshot: WorkbookSnapshot, sheetPartById: Record<string, string>): PivotSource | undefined {
  if (source.kind === 'unsupported') return undefined;
  if (source.kind === 'worksheet-range') {
    const sheet = snapshot.sheets.find((candidate) => candidate.name === source.sheetName || sheetPartById[candidate.id] === source.sheetPart);
    const range = sheet ? parseRange(source.ref, sheet.id) : undefined;
    return sheet && range ? { kind: 'worksheet-range', range } : undefined;
  }
  const table = snapshot.sheets.flatMap((sheet) => (sheet.sheetTables ?? []).map((candidate) => ({ sheet, candidate }))).find(({ candidate }) => candidate.name === source.tableName);
  return table ? { kind: 'table', tableId: table.candidate.id } : undefined;
}

function resolveCanonicalSource(source: PivotSource, snapshot: WorkbookSnapshot, parts: Record<string, string>): { key: string; source: NativePivotSource; sheet: SheetSnapshot; range: RangeRef; tableName?: string } | undefined {
  if (source.kind === 'worksheet-range') {
    const sheet = snapshot.sheets.find((candidate) => candidate.id === source.range.sheetId);
    const sheetPart = sheet ? parts[sheet.id] : undefined;
    return sheet && sheetPart ? { key: `range:${sheetPart}:${rangeToA1(source.range)}`, source: { kind: 'worksheet-range', sheetName: sheet.name, sheetPart, ref: rangeToA1(source.range) }, sheet, range: source.range } : undefined;
  }
  if (source.kind === 'table') {
    const found = snapshot.sheets.flatMap((sheet) => (sheet.sheetTables ?? []).map((table) => ({ sheet, table }))).find(({ table }) => table.id === source.tableId || table.name === source.tableId);
    const part = found ? parts[found.sheet.id] : undefined;
    return found && part ? { key: `table:${found.table.name}`, source: { kind: 'table', tableName: found.table.name, sheetName: found.sheet.name, sheetPart: part }, sheet: found.sheet, range: found.table.range, tableName: found.table.name } : undefined;
  }
  return undefined;
}

function nativeSourceKey(source: NativePivotSource | { kind: 'unsupported'; reason: string }): string {
  if (source.kind === 'worksheet-range') return `range:${source.sheetPart ?? source.sheetName}:${source.ref}`;
  if (source.kind === 'table') return `table:${source.tableName}`;
  return `unsupported:${source.reason}`;
}

function normalizePivot(input: PivotModel): PivotDefinition | undefined {
  if (input.schema !== 'PivotDefinition') return undefined;
  return canonicalizePivotDefinition(structuredClone(input));
}

function pivotGroupingKey(pivot: PivotDefinition): string {
  return JSON.stringify({ rows: pivot.layout.rows.map((placement) => placement.group ?? null), columns: pivot.layout.columns.map((placement) => placement.group ?? null) });
}

function readSourceRows(sheet: SheetSnapshot, range: RangeRef, pivot: PivotDefinition, tableName?: string): { fields: Array<{ name: string; dataType: NativePivotCacheField['dataType'] }>; rows: PivotScalar[][] } {
  const fields: Array<{ name: string; dataType: NativePivotCacheField['dataType'] }> = [];
  for (let column = range.startColumn; column <= range.endColumn; column += 1) {
    const cell = sheet.cells[String(range.startRow)]?.[String(column)];
    const value = cell?.formulaValue ?? cell?.value;
    const catalog = pivot.fieldCatalog.fields[column - range.startColumn];
    fields.push({ name: catalog?.name ?? (typeof value === 'string' && value ? value : `${tableName ?? 'Field'}${column - range.startColumn + 1}`), dataType: nativeDataType(catalog?.dataType ?? inferDataType(sheet, range, column)) });
  }
  const rows: PivotScalar[][] = [];
  for (let row = range.startRow + 1; row <= range.endRow; row += 1) {
    const values = fields.map((_, offset) => {
      const cell = sheet.cells[String(row)]?.[String(range.startColumn + offset)];
      const value = cell?.formulaValue ?? cell?.value;
      return isScalar(value) ? value : null;
    });
    if (values.some((value) => value !== null)) rows.push(values);
  }
  return { fields, rows };
}

function applyCanonicalPivotGroups(cache: NativePivotCacheDefinition, pivot: PivotDefinition): void {
  const placements = [...pivot.layout.rows, ...pivot.layout.columns];
  for (const placement of placements) {
    if (!placement.group) continue;
    const field = pivot.fieldCatalog.fields.find((candidate) => candidate.fieldId === placement.fieldId || candidate.name === placement.fieldId);
    if (!field || !cache.fields[field.ordinal]) continue;
    cache.fields[field.ordinal]!.fieldGroup = buildNativeFieldGroup(placement.group, field.ordinal, cache.fields[field.ordinal]!.sharedItems ?? []);
  }
}

function buildNativeFieldGroup(group: PivotGroup, base: number, sharedItems: NativePivotScalar[]): NativePivotFieldGroup {
  if (group.kind === 'date') {
    if (group.units && group.units.length > 1) throw new Error('Native Pivot export cannot represent multi-level date grouping without a native hierarchy');
    const range: NativePivotFieldRange = {
      groupBy: `${group.units?.[0] ?? group.unit}s`,
      ...(group.start === undefined ? {} : { start: group.start }),
      ...(group.end === undefined ? {} : { end: group.end }),
      ...(group.autoStart === undefined ? {} : { autoStart: group.autoStart }),
      ...(group.autoEnd === undefined ? {} : { autoEnd: group.autoEnd }),
    };
    return { base, range };
  }
  if (group.kind === 'number') {
    return {
      base,
      range: {
        groupBy: 'range', interval: group.interval,
        ...(group.start === undefined ? {} : { start: group.start }),
        ...(group.end === undefined ? {} : { end: group.end }),
        ...(group.autoStart === undefined ? {} : { autoStart: group.autoStart }),
        ...(group.autoEnd === undefined ? {} : { autoEnd: group.autoEnd }),
      },
    };
  }
  const groupItems = group.groups.map((candidate) => candidate.name);
  const groupByMember = new Map<string, number>();
  group.groups.forEach((candidate, groupIndex) => candidate.items.forEach((item) => groupByMember.set(pivotMemberKey(item), groupIndex)));
  const ungrouped = groupItems.length;
  const discreteIndexes = sharedItems.map((value) => groupByMember.get(pivotMemberKey(createPivotMemberKey(value))) ?? ungrouped);
  return { base, discreteIndexes, groupItems };
}

function buildNativeTable(pivot: PivotDefinition, cache: NativePivotCacheDefinition, part: string, sheetPart: string, old: NativePivotTableDefinition | undefined, source: { fields: Array<{ name: string; dataType: NativePivotCacheField['dataType'] }>; rows: PivotScalar[][] }): NativePivotTableDefinition {
  const fieldIndex = (placement: PivotFieldPlacement | PivotValueField): number => {
    const id = placement.fieldId;
    return pivot.fieldCatalog.fields.find((field) => field.fieldId === id || field.name === id)?.ordinal ?? source.fields.findIndex((field) => field.name === id);
  };
  for (const filter of pivot.layout.filters) {
    if (fieldIndex(filter) < 0) throw new Error(`Pivot filter references missing field ${filter.fieldId}`);
  }
  const rows = pivot.layout.rows.map(fieldIndex).filter((index) => index >= 0);
  const columns = pivot.layout.columns.map(fieldIndex).filter((index) => index >= 0);
  const pages = [...new Set(pivot.layout.filters
    .filter((filter) => (filter.scope ?? 'report') === 'report')
    .map((filter) => fieldIndex(filter))
    .filter((index) => index >= 0))];
  const subtotalForField = (index: number) => pivot.layout.rows.find((placement) => fieldIndex(placement) === index)?.subtotal
    ?? pivot.layout.columns.find((placement) => fieldIndex(placement) === index)?.subtotal;
  const dataFields = pivot.layout.values.map((value) => ({ field: fieldIndex(value), ...(value.displayName ? { name: value.displayName } : {}), subtotal: value.summarizeBy, ...(value.showAs && value.showAs.kind !== 'normal' ? { showDataAs: nativeShowAs(value.showAs.kind) } : {}) })).filter((value) => value.field >= 0);
  const placementForField = (index: number): PivotFieldPlacement | undefined => [...pivot.layout.rows, ...pivot.layout.columns].find((placement) => fieldIndex(placement) === index);
  const nativeSortForField = (index: number): Pick<NativePivotTableField, 'sortType' | 'nonAutoSortDefault' | 'autoSortScope'> => {
    const sort = placementForField(index)?.sort;
    if (!sort) {
      const preserved = pivot.nativeMetadata?.preservedAutoSortScopes?.find((candidate) => candidate.fieldIndex === index);
      return preserved ? {
        ...(preserved.sortType ? { sortType: preserved.sortType } : {}),
        ...(preserved.nonAutoSortDefault === undefined ? {} : { nonAutoSortDefault: preserved.nonAutoSortDefault }),
        autoSortScope: {
          attributes: { ...preserved.attributes },
          references: preserved.references.map((reference) => ({ ...reference, ...(reference.itemIndexes ? { itemIndexes: [...reference.itemIndexes] } : {}) })),
        },
      } : {};
    }
    const preserved = pivot.nativeMetadata?.preservedAutoSortScopes?.find((candidate) => candidate.fieldIndex === index);
    const valueFieldOrdinal = sort.valueFieldId === undefined ? -1 : pivot.fieldCatalog.fields.find((field) => field.fieldId === sort.valueFieldId || field.name === sort.valueFieldId)?.ordinal ?? -1;
    const valueField = valueFieldOrdinal < 0 ? undefined : dataFields.find((value) => value.field === valueFieldOrdinal);
    return {
      sortType: sort.direction,
      ...(preserved?.nonAutoSortDefault === undefined ? {} : { nonAutoSortDefault: preserved.nonAutoSortDefault }),
      autoSortScope: {
        ...(sort.by === 'value' ? { dataOnly: true } : { dataOnly: false }),
        fieldPosition: 0,
        attributes: { dataOnly: sort.by === 'value' ? '1' : '0', fieldPosition: '0' },
        references: valueField ? [{ field: valueField.field, selected: false, itemIndexes: [] }] : [],
      },
    };
  };
  const oldAnchor = old?.locationRef ? parseA1(old.locationRef.split(':')[0] ?? '') : undefined;
  const locationRef = oldAnchor
    && oldAnchor.row === pivot.target.anchor.row
    && oldAnchor.column === pivot.target.anchor.column
    ? old?.locationRef
    : deriveLocation(pivot, source.rows.length, rows.length, columns.length, dataFields.length);
  const styleName = pivot.presentation?.styleName ?? old?.styleName;
  const styleOptions = pivot.presentation?.styleOptions ?? old?.styleOptions;
  const collapsedItemIndexes = new Map<number, number[]>();
  for (const nodeId of pivot.layout.expansion?.collapsedNodeIds ?? []) {
    const separator = nodeId.indexOf('=');
    if (separator <= 0) continue;
    const fieldIdValue = nodeId.slice(0, separator);
    const member = nodeId.slice(separator + 1);
    const fieldIndexValue = pivot.fieldCatalog.fields.find((field) => field.fieldId === fieldIdValue || field.name === fieldIdValue)?.ordinal ?? -1;
    if (fieldIndexValue < 0) continue;
    const sharedItems = cache.fields[fieldIndexValue]?.sharedItems ?? [];
    const index = sharedItems.findIndex((value) => pivotMemberKey(createPivotMemberKey(value)) === member);
    if (index >= 0) collapsedItemIndexes.set(fieldIndexValue, [...(collapsedItemIndexes.get(fieldIndexValue) ?? []), index]);
  }
  const hiddenItemIndexes = new Map<number, number[]>();
  const manualFiltersByField = new Map<number, Extract<PivotLayout['filters'][number], { kind: 'manual' }>[] >();
  for (const filter of pivot.layout.filters) {
    if (filter.kind !== 'manual' || filter.mode === 'all') continue;
    const index = fieldIndex(filter);
    if (index < 0) throw new Error(`Pivot manual filter references missing field ${filter.fieldId}`);
    const current = manualFiltersByField.get(index) ?? [];
    current.push(filter);
    manualFiltersByField.set(index, current);
  }
  for (const [index, filters] of manualFiltersByField) {
    if (filters.length > 1) throw new Error(`Pivot manual filters cannot target field ${index} in multiple scopes`);
    const values = cache.fields[index]?.sharedItems;
    if (!values) throw new Error(`Pivot manual filter field ${index} has no cache sharedItems`);
    const keys = new Set(values.map((value) => pivotMemberKey(createPivotMemberKey(value))));
    const requested = new Set(filters[0]!.memberKeys.map((value) => pivotMemberKey(value)));
    const unknown = [...requested].find((key) => !keys.has(key));
    if (unknown) throw new Error(`Pivot manual filter field ${index} references an unknown member ${unknown}`);
    const hidden = values.flatMap((value, valueIndex) => {
      const selected = requested.has(pivotMemberKey(createPivotMemberKey(value)));
      const shouldHide = filters[0]!.mode === 'exclude' ? selected : !selected;
      return shouldHide ? [valueIndex] : [];
    });
    if (hidden.length) hiddenItemIndexes.set(index, hidden);
  }
  return {
    name: old?.name ?? (pivot.id.replace(/[^A-Za-z0-9_]/g, '_').slice(0, 255) || 'PivotTable'),
    part, sheetPart, relationshipId: old?.relationshipId ?? '', cacheId: cache.cacheId, pivotId: pivot.id,
    locationRef,
    fields: source.fields.map((_, index) => ({
      index,
      ...(rows.includes(index) ? { axis: 'row' as const } : columns.includes(index) ? { axis: 'column' as const } : pages.includes(index) ? { axis: 'page' as const } : {}),
      ...nativeSortForField(index),
      ...(pivot.layout.compact ? { compact: true } : {}),
      ...((collapsedItemIndexes.get(index)?.length ?? 0) > 0 ? { collapsedItemIndexes: [...new Set(collapsedItemIndexes.get(index))] } : {}),
      ...((hiddenItemIndexes.get(index)?.length ?? 0) > 0 ? { hiddenItemIndexes: hiddenItemIndexes.get(index) } : {}),
      ...(subtotalForField(index) ? { subtotal: structuredClone(subtotalForField(index)) } : {}),
    })),
    rowFields: rows, columnFields: columns, pageFields: pages, dataFields,
    pivotFilters: buildNativePivotFilters(pivot, dataFields),
    showRowGrandTotals: pivot.layout.showGrandTotals, showColumnGrandTotals: pivot.layout.showGrandTotals, subtotalLocation: pivot.layout.subtotalLocation, repeatLabels: pivot.layout.repeatLabels, compactData: pivot.layout.compact, multipleFieldFilters: pivot.layout.allowMultipleFiltersPerField,
    ...(styleName ? { styleName } : {}),
    ...(styleOptions ? { styleOptions: structuredClone(styleOptions) } : {}),
    showButtons: pivot.layout.expansion?.showButtons ?? old?.showButtons ?? true,
    ...(pivot.presentation?.displayOptions?.showFieldHeaders === undefined ? {} : { showFieldHeaders: pivot.presentation.displayOptions.showFieldHeaders }),
    ...(pivot.presentation?.displayOptions?.fillEmptyCells === undefined ? {} : { fillEmptyCells: pivot.presentation.displayOptions.fillEmptyCells }),
    ...(pivot.presentation?.displayOptions?.emptyCellText === undefined ? {} : { emptyCellText: pivot.presentation.displayOptions.emptyCellText }),
    ...(pivot.presentation?.displayOptions?.showErrorValues === undefined ? {} : { showErrorValues: pivot.presentation.displayOptions.showErrorValues }),
    ...(pivot.presentation?.displayOptions?.errorCellText === undefined ? {} : { errorCellText: pivot.presentation.displayOptions.errorCellText }),
    preserveFormatting: pivot.refreshPolicy.preserveFormatting,
  };
}

function buildCacheDefinitionXml(cache: NativePivotCacheDefinition): string {
  const source = cache.source.kind === 'worksheet-range' ? `<worksheetSource ref="${encodeXml(cache.source.ref)}" sheet="${encodeXml(cache.source.sheetName)}"/>` : cache.source.kind === 'table' ? `<worksheetSource name="${encodeXml(cache.source.tableName)}"${cache.source.sheetName ? ` sheet="${encodeXml(cache.source.sheetName)}"` : ''}/>` : '';
  const fields = cache.fields.map((field) => {
    const values = field.sharedItems ?? [];
    const contains = field.dataType === 'string' ? ' containsString="1"' : field.dataType === 'date' ? ' containsDate="1"' : field.dataType === 'number' ? ' containsNumber="1"' : field.dataType === 'boolean' ? ' containsBoolean="1"' : field.dataType === 'error' ? ' containsError="1"' : '';
    const shared = values.map((value) => nativePivotScalarXml(value)).join('');
    const fieldGroup = field.fieldGroup ? buildNativeFieldGroupXml(field.fieldGroup) : '';
    return `<cacheField name="${encodeXml(field.name)}"><sharedItems count="${values.length}"${contains}>${shared}</sharedItems>${fieldGroup}</cacheField>`;
  }).join('');
  const attrs = [...(cache.recordCount === undefined ? [] : [`recordCount="${cache.recordCount}"`]), ...boolAttr(cache.refreshOnLoad, 'refreshOnLoad'), ...boolAttr(cache.refreshOnSave, 'refreshOnSave'), ...boolAttr(cache.saveData, 'saveData'), ...boolAttr(cache.enableRefresh, 'enableRefresh')];
  return withXmlDeclaration(`<pivotCacheDefinition xmlns="${NS_MAIN}" xmlns:r="${NS_DOC_REL}"${attrs.length ? ` ${attrs.join(' ')}` : ''}><cacheSource type="worksheet">${source}</cacheSource><cacheFields count="${cache.fields.length}">${fields}</cacheFields></pivotCacheDefinition>`);
}

function buildNativeFieldGroupXml(group: NativePivotFieldGroup): string {
  const attrs = [
    ...(group.base === undefined ? [] : [`base="${group.base}"`]),
    ...(group.parent === undefined ? [] : [`par="${group.parent}"`]),
  ];
  const range = group.range ? buildNativeFieldRangeXml(group.range) : '';
  const discrete = group.discreteIndexes ? `<discretePr count="${group.discreteIndexes.length}">${group.discreteIndexes.map((index) => `<x v="${index}"/>`).join('')}</discretePr>` : '';
  const items = group.groupItems ? `<groupItems count="${group.groupItems.length}">${group.groupItems.map(nativePivotScalarXml).join('')}</groupItems>` : '';
  return `<fieldGroup${attrs.length ? ` ${attrs.join(' ')}` : ''}>${range}${discrete}${items}</fieldGroup>`;
}

function buildNativeFieldRangeXml(range: NativePivotFieldRange): string {
  const attrs = [
    ...(range.groupBy === undefined ? [] : [`groupBy="${encodeXml(range.groupBy)}"`]),
    ...(range.interval === undefined ? [] : [`groupInterval="${range.interval}"`]),
    ...rangeBoundAttr(range.start, 'start'),
    ...rangeBoundAttr(range.end, 'end'),
    ...boolAttr(range.autoStart, 'autoStart'),
    ...boolAttr(range.autoEnd, 'autoEnd'),
  ];
  return `<rangePr${attrs.length ? ` ${attrs.join(' ')}` : ''}/>`;
}

function rangeBoundAttr(value: NativePivotScalar | undefined, name: 'start' | 'end'): string[] {
  if (value === undefined || value === null) return [];
  if (typeof value === 'number') return [`${name}Num="${value}"`];
  return [`${name}Date="${encodeXml(String(value))}"`];
}

function nativePivotScalarXml(value: NativePivotScalar): string {
  if (value === null) return '<m/>';
  if (isPivotError(value)) return `<e v="${encodeXml(value.code)}"/>`;
  if (typeof value === 'string') return `<s v="${encodeXml(value)}"/>`;
  if (typeof value === 'boolean') return `<b v="${value ? '1' : '0'}"/>`;
  return `<n v="${encodeXml(String(value))}"/>`;
}

function buildCacheRecordsXml(cache: NativePivotCacheDefinition, rows: PivotScalar[][]): string {
  const indexes = cache.fields.map((field) => new Map((field.sharedItems ?? []).map((value, index) => [scalarKey(value), index])));
  const records = rows.map((row) => `<r>${cache.fields.map((_, index) => {
    const value = row[index] ?? null;
    if (value === null) return '<m/>';
    if (isPivotError(value)) return `<e v="${encodeXml(value.code)}"/>`;
    if (typeof value === 'string') return `<s v="${indexes[index]?.get(scalarKey(value)) ?? 0}"/>`;
    if (typeof value === 'boolean') return `<b v="${value ? '1' : '0'}"/>`;
    return `<n v="${encodeXml(String(value))}"/>`;
  }).join('')}</r>`).join('');
  return withXmlDeclaration(`<pivotCacheRecords xmlns="${NS_MAIN}" count="${rows.length}">${records}</pivotCacheRecords>`);
}

function buildPivotTableXml(table: NativePivotTableDefinition): string {
  const fields = table.fields.map((field) => {
    const itemIndexes = [...new Set([...(field.collapsedItemIndexes ?? []), ...(field.hiddenItemIndexes ?? [])])].sort((left, right) => left - right);
    const hiddenItems = new Set(field.hiddenItemIndexes ?? []);
    const items = itemIndexes.length ? `<items count="${itemIndexes.length}">${itemIndexes.map((index) => `<item x="${index}"${hiddenItems.has(index) ? ' h="1"' : ''}${field.collapsedItemIndexes?.includes(index) ? ' sd="0"' : ''}/>`).join('')}</items>` : '';
    const subtotal = nativePivotSubtotalAttrs(field.subtotal);
    const sortType = field.sortType && field.sortType !== 'manual' ? ` sortType="${field.sortType}"` : field.sortType === 'manual' ? ' sortType="manual"' : '';
    const nonAutoSortDefault = field.nonAutoSortDefault === undefined ? '' : ` nonAutoSortDefault="${field.nonAutoSortDefault ? '1' : '0'}"`;
    const autoSortScope = field.autoSortScope ? buildAutoSortScopeXml(field.autoSortScope) : '';
    return `<pivotField${field.axis === 'row' ? ' axis="axisRow"' : field.axis === 'column' ? ' axis="axisCol"' : field.axis === 'page' ? ' axis="axisPage"' : ''}${field.compact === undefined ? '' : ` compact="${field.compact ? '1' : '0'}"`}${sortType}${nonAutoSortDefault}${subtotal}>${items}${autoSortScope}</pivotField>`;
  }).join('');
  const rows = table.rowFields.map((field) => `<field x="${field}"/>`).join('');
  const columns = table.columnFields.map((field) => `<field x="${field}"/>`).join('');
  const pages = table.pageFields.map((field) => `<pageField fld="${field}"/>`).join('');
  const data = table.dataFields.map((field) => `<dataField fld="${field.field}"${field.name ? ` name="${encodeXml(field.name)}"` : ''} subtotal="${encodeXml(nativeAggregate(field.subtotal))}"${field.showDataAs ? ` showDataAs="${encodeXml(field.showDataAs)}"` : ''}/>`).join('');
  const filters = table.pivotFilters?.length ? `<pivotFilters count="${table.pivotFilters.length}">${table.pivotFilters.map(buildPivotFilterXml).join('')}</pivotFilters>` : '';
  const styleOptions = { ...DEFAULT_PIVOT_STYLE_OPTIONS, ...(table.styleOptions ?? {}) };
  const style = table.styleName ? `<pivotTableStyleInfo name="${encodeXml(table.styleName)}" showRowHeaders="${styleOptions.showRowHeaders ? '1' : '0'}" showColHeaders="${styleOptions.showColumnHeaders ? '1' : '0'}" showRowStripes="${styleOptions.showRowStripes ? '1' : '0'}" showColStripes="${styleOptions.showColumnStripes ? '1' : '0'}" showLastColumn="${styleOptions.showLastColumn ? '1' : '0'}"/>` : '';
  const displayAttrs = [
    ...boolAttr(table.showFieldHeaders, 'showHeaders'),
    ...boolAttr(table.fillEmptyCells, 'showMissing'),
    ...(table.emptyCellText === undefined ? [] : [`missingCaption="${encodeXml(table.emptyCellText)}"`]),
    ...boolAttr(table.showErrorValues, 'showError'),
    ...(table.errorCellText === undefined ? [] : [`errorCaption="${encodeXml(table.errorCellText)}"`]),
    ...boolAttr(table.preserveFormatting, 'preserveFormatting'),
  ];
  const subtotalAttrs = table.subtotalLocation === 'off' ? ' showSubtotals="0"' : table.subtotalLocation === 'top' ? ' subtotalTop="1"' : ' subtotalTop="0"';
  return withXmlDeclaration(`<pivotTableDefinition xmlns="${NS_MAIN}" xmlns:r="${NS_DOC_REL}" name="${encodeXml(table.name)}" cacheId="${table.cacheId}" rowGrandTotals="${table.showRowGrandTotals === false ? '0' : '1'}" colGrandTotals="${table.showColumnGrandTotals === false ? '0' : '1'}" compactData="${table.compactData === false ? '0' : '1'}" multipleFieldFilters="${table.multipleFieldFilters === false ? '0' : '1'}" showDrill="${table.showButtons === false ? '0' : '1'}"${displayAttrs.length ? ` ${displayAttrs.join(' ')}` : ''}${subtotalAttrs}><location ref="${encodeXml(table.locationRef ?? 'A1')}" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/><pivotFields count="${table.fields.length}">${fields}</pivotFields><rowFields count="${table.rowFields.length}">${rows}</rowFields><colFields count="${table.columnFields.length}">${columns}</colFields>${table.pageFields.length ? `<pageFields count="${table.pageFields.length}">${pages}</pageFields>` : '<pageFields count="0"/>'}<dataFields count="${table.dataFields.length}">${data}</dataFields>${filters}${style}</pivotTableDefinition>`);
}

function buildAutoSortScopeXml(scope: NativePivotAutoSortScope): string {
  const areaAttributes = {
    ...scope.attributes,
    ...(scope.dataOnly === undefined ? {} : { dataOnly: scope.dataOnly ? '1' : '0' }),
    ...(scope.labelOnly === undefined ? {} : { labelOnly: scope.labelOnly ? '1' : '0' }),
    ...(scope.outline === undefined ? {} : { outline: scope.outline ? '1' : '0' }),
    ...(scope.fieldPosition === undefined ? {} : { fieldPosition: String(scope.fieldPosition) }),
  };
  const attrs = Object.entries(areaAttributes).map(([name, value]) => `${name}="${encodeXml(value)}"`).join(' ');
  const references = scope.references.map((reference) => `<reference field="${reference.field}"${reference.selected === undefined ? '' : ` selected="${reference.selected ? '1' : '0'}"`}${reference.itemIndexes?.length ? ` count="${reference.itemIndexes.length}"` : ''}>${(reference.itemIndexes ?? []).map((index) => `<x v="${index}"/>`).join('')}</reference>`).join('');
  return `<autoSortScope><pivotArea${attrs ? ` ${attrs}` : ''}><references count="${scope.references.length}">${references}</references></pivotArea></autoSortScope>`;
}

function buildPivotFilterXml(filter: NativePivotFilter): string {
  const attrs: Record<string, string> = { ...filter.attributes, fld: String(filter.field), type: filter.type };
  if (filter.measureField !== undefined) attrs.iMeasureFld = String(filter.measureField);
  if (filter.secondMeasureField !== undefined) attrs.iMeasureFld2 = String(filter.secondMeasureField);
  if (filter.evalOrder !== undefined) attrs.evalOrder = String(filter.evalOrder);
  if (filter.id !== undefined) attrs.id = String(filter.id);
  if (filter.stringValue1 !== undefined) attrs.stringValue1 = filter.stringValue1;
  if (filter.stringValue2 !== undefined) attrs.stringValue2 = filter.stringValue2;
  if (filter.value1 !== undefined) attrs.val = String(filter.value1);
  if (filter.value2 !== undefined) attrs.val2 = String(filter.value2);
  if (filter.wholeDay !== undefined) attrs.wholeDay = filter.wholeDay ? '1' : '0';
  if (filter.top !== undefined) attrs.top = filter.top ? '1' : '0';
  if (filter.percent !== undefined) attrs.percent = filter.percent ? '1' : '0';
  return `<filter ${Object.entries(attrs).map(([name, value]) => `${name}="${encodeXml(value)}"`).join(' ')}/>`;
}

function buildNativePivotFilters(pivot: PivotDefinition, dataFields: NativePivotDataField[]): NativePivotFilter[] {
  const fieldIndex = (id: string): number => pivot.fieldCatalog.fields.find((field) => field.fieldId === id || field.name === id)?.ordinal ?? -1;
  const dataFieldIndex = (id: string): number => dataFields.findIndex((dataField) => dataField.field === fieldIndex(id));
  const filters = pivot.layout.filters.flatMap((filter): NativePivotFilter[] => {
    if (filter.kind === 'manual') return [];
    const field = fieldIndex(filter.fieldId);
    if (field < 0) throw new Error(`Pivot filter references missing field ${filter.fieldId}`);
    if (filter.kind === 'top-items') {
      const measure = dataFieldIndex(filter.valueFieldId);
      if (measure < 0) throw new Error(`Pivot top-items filter references missing value field ${filter.valueFieldId}`);
      return [{ field, type: filter.direction === 'top' ? 'valueTop10' : 'valueBottom10', measureField: measure, value1: filter.count, top: filter.direction === 'top', attributes: { fld: String(field), type: filter.direction === 'top' ? 'valueTop10' : 'valueBottom10', iMeasureFld: String(measure), val: String(filter.count), top: filter.direction === 'top' ? '1' : '0' } }];
    }
    if (filter.family === 'date' && filter.dynamic) {
      const dynamicTypes: Record<string, string> = { today: 'dateToday', yesterday: 'dateYesterday', tomorrow: 'dateTomorrow', 'this-week': 'dateThisWeek', 'last-week': 'dateLastWeek', 'next-week': 'dateNextWeek', 'this-month': 'dateThisMonth', 'last-month': 'dateLastMonth', 'next-month': 'dateNextMonth', 'this-quarter': 'dateThisQuarter', 'last-quarter': 'dateLastQuarter', 'next-quarter': 'dateNextQuarter', 'this-year': 'dateThisYear', 'last-year': 'dateLastYear', 'next-year': 'dateNextYear', 'year-to-date': 'dateYearToDate' };
      const type = dynamicTypes[filter.dynamic];
      if (!type) throw new Error(`Pivot dynamic date filter ${filter.dynamic} cannot be represented in native OOXML`);
      return [{ field, type, attributes: { fld: String(field), type }, ...(filter.wholeDay === undefined ? {} : { wholeDay: filter.wholeDay }) }];
    }
    const isValueFilter = filter.family === 'value';
    const measure = isValueFilter && filter.valueFieldId !== undefined ? dataFieldIndex(filter.valueFieldId) : undefined;
    const captionTypes: Record<string, string | undefined> = { equals: 'captionEqual', 'not-equals': 'captionNotEqual', 'begins-with': 'captionBeginsWith', 'not-begins-with': 'captionNotBeginsWith', 'ends-with': 'captionEndsWith', 'not-ends-with': 'captionNotEndsWith', contains: 'captionContains', 'not-contains': 'captionNotContains', between: 'captionBetween', 'not-between': 'captionNotBetween', 'greater-than': 'captionGreaterThan', 'greater-or-equal': 'captionGreaterThanOrEqual', 'less-than': 'captionLessThan', 'less-or-equal': 'captionLessThanOrEqual' };
    const dateTypes: Record<string, string | undefined> = { equals: 'dateEqual', 'not-equals': 'dateNotEqual', before: 'dateOlderThan', after: 'dateNewerThan', between: 'dateBetween', 'not-between': 'dateNotBetween' };
    const valueTypes: Record<string, string | undefined> = { equals: 'valueEqual', 'not-equals': 'valueNotEqual', 'greater-than': 'valueGreaterThan', 'greater-or-equal': 'valueGreaterThanOrEqual', 'less-than': 'valueLessThan', 'less-or-equal': 'valueLessThanOrEqual', between: 'valueBetween', 'not-between': 'valueNotBetween' };
    const type = isValueFilter ? valueTypes[filter.operator] : filter.family === 'date' ? dateTypes[filter.operator] : captionTypes[filter.operator];
    if (!type) throw new Error(`Pivot filter operator ${filter.operator} cannot be represented in native OOXML`);
    if (isValueFilter && (filter.valueFieldId === undefined || measure === undefined || measure < 0)) throw new Error(`Pivot value filter references missing value field ${filter.valueFieldId ?? filter.fieldId}`);
    return [{ field, type, ...(measure === undefined ? {} : { measureField: measure }), ...(typeof filter.value === 'string' && !isValueFilter ? { stringValue1: filter.value } : { value1: filter.value }), ...(filter.value2 === undefined ? {} : typeof filter.value2 === 'string' && !isValueFilter ? { stringValue2: filter.value2 } : { value2: filter.value2 }), ...(filter.family === 'date' && filter.wholeDay !== undefined ? { wholeDay: filter.wholeDay } : {}), attributes: { fld: String(field), type, ...(measure === undefined ? {} : { iMeasureFld: String(measure) }) } }];
  });
  const preserved = pivot.nativeMetadata?.preservedPivotFilters?.map((filter) => ({ field: filter.fieldIndex, type: filter.type, attributes: { ...filter.attributes } })) ?? [];
  return [...filters, ...preserved];
}

function nativePivotSubtotalAttrs(subtotal: NativePivotTableField['subtotal']): string {
  if (!subtotal || subtotal.mode === 'automatic') return subtotal ? ' defaultSubtotal="1"' : '';
  if (subtotal.mode === 'none') return ' defaultSubtotal="0"';
  const attrs: Record<string, string> = { defaultSubtotal: '0' };
  const names: Record<string, string> = { sum: 'sumSubtotal', count: 'countSubtotal', 'count-numbers': 'countASubtotal', average: 'averageSubtotal', max: 'maxSubtotal', min: 'minSubtotal', product: 'productSubtotal', stdev: 'stdDevSubtotal', stdevp: 'stdDevpSubtotal', var: 'varSubtotal', varp: 'varPSubtotal', 'distinct-count': 'distinctCountSubtotal' };
  for (const fn of subtotal.functions ?? []) { const name = names[fn]; if (name) attrs[name] = '1'; }
  return Object.entries(attrs).map(([name, value]) => ` ${name}="${value}"`).join('');
}

function buildDisplayCells(pivot: PivotDefinition, table: NativePivotTableDefinition, source: { fields: Array<{ name: string; dataType: NativePivotCacheField['dataType'] }>; rows: PivotScalar[][] }): Record<string, Record<string, CellData>> {
  const start = parseA1(table.locationRef?.split(':')[0] ?? 'A1');
  if (!start) return {};
  const output: Record<string, Record<string, CellData>> = {};
  const rowGroups = uniqueTuples(source.rows, table.rowFields);
  const columnGroups = uniqueTuples(source.rows, table.columnFields);
  const rows = rowGroups.length ? rowGroups : [[]];
  const columns = columnGroups.length ? columnGroups : [[]];
  const put = (row: number, column: number, value: PivotScalar): void => { if (value === null) return; output[String(row)] ??= {}; output[String(row)]![String(column)] = { value: isPivotError(value) ? value.code : value }; };
  put(start.row, start.column, table.name);
  table.rowFields.forEach((index, position) => put(start.row + 1, start.column + position, source.fields[index]?.name ?? `Field${index + 1}`));
  const headerRow = start.row + 1 + Math.max(1, table.rowFields.length);
  columns.forEach((tuple, index) => put(headerRow, start.column + table.rowFields.length + index, tuple.map((value) => value === null ? '' : formatPivotMember(value)).join(' / ') || 'Values'));
  rows.forEach((tuple, rowIndex) => {
    tuple.forEach((value, position) => put(headerRow + 1 + rowIndex, start.column + position, value));
    columns.forEach((columnTuple, columnIndex) => table.dataFields.forEach((field, valueIndex) => {
      const values = source.rows.filter((row) => matchesTuple(row, table.rowFields, tuple) && matchesTuple(row, table.columnFields, columnTuple)).map((row) => row[field.field] ?? null);
      put(headerRow + 1 + rowIndex, start.column + table.rowFields.length + columnIndex * Math.max(1, table.dataFields.length) + valueIndex, aggregate(values, field.subtotal));
    }));
  });
  void pivot;
  return output;
}

function mergeDisplayCells(target: Record<string, Record<string, CellData>>, display: Record<string, Record<string, CellData>>, sheet: SheetSnapshot): void {
  for (const [row, columns] of Object.entries(display)) for (const [column, cell] of Object.entries(columns)) { if (sheet.cells[row]?.[column]) continue; target[row] ??= {}; target[row]![column] ??= cell; }
}

function rebuildRelationships(input: Record<string, XlsxRelationship[]>, caches: NativePivotCacheDefinition[], tables: NativePivotTableDefinition[], files: Record<string, Uint8Array>): Record<string, XlsxRelationship[]> {
  const result = cloneRelationships(input);
  const workbook = result['xl/workbook.xml'] ?? [];
  const oldCacheRelations = input['xl/workbook.xml']?.filter((relation) => relation.type.endsWith('/pivotCacheDefinition')) ?? [];
  result['xl/workbook.xml'] = workbook.filter((relation) => !relation.type.endsWith('/pivotCacheDefinition'));
  for (const [source, list] of Object.entries(result)) {
    if (source.startsWith('xl/worksheets/')) result[source] = list.filter((relation) => !relation.type.endsWith('/pivotTable'));
    if (caches.some((cache) => cache.part === source)) result[source] = result[source]!.filter((relation) => !relation.type.endsWith('/pivotCacheRecords'));
  }
  for (const cache of caches) {
    const old = oldCacheRelations.find((relation) => resolveTarget('xl/workbook.xml', relation.target) === cache.part);
    result['xl/workbook.xml']!.push({ id: old?.id ?? allocateId(result['xl/workbook.xml']!), type: REL_PIVOT_CACHE_DEFINITION, target: relativeTarget('xl/workbook.xml', cache.part) });
    if (cache.recordsPart) { const list = result[cache.part] ?? []; const oldRecord = input[cache.part]?.find((relation) => relation.type.endsWith('/pivotCacheRecords')); result[cache.part] = [...list, { id: oldRecord?.id ?? allocateId(list), type: REL_PIVOT_CACHE_RECORDS, target: relativeTarget(cache.part, cache.recordsPart) }]; }
  }
  for (const table of tables) { const list = result[table.sheetPart] ?? []; const old = input[table.sheetPart]?.find((relation) => relation.type.endsWith('/pivotTable') && resolveTarget(table.sheetPart, relation.target) === table.part); result[table.sheetPart] = [...list, { id: (old?.id ?? table.relationshipId) || allocateId(list), type: REL_PIVOT_TABLE, target: relativeTarget(table.sheetPart, table.part) }]; }
  for (const [source, list] of Object.entries(result)) result[source] = list.filter((relation) => !isPivotRelation(relation) || files[resolveTarget(source, relation.target)] !== undefined);
  return result;
}

function parseCacheSource(node: XmlNode | undefined, sheets: Map<string, { name: string; part: string }>): NativePivotSource | { kind: 'unsupported'; reason: string } {
  const worksheet = child(node, 'worksheetSource');
  if (worksheet?.attrs.sheet && worksheet.attrs.ref) { const match = [...sheets.values()].find((sheet) => sheet.name === worksheet.attrs.sheet); return { kind: 'worksheet-range', sheetName: worksheet.attrs.sheet, ref: worksheet.attrs.ref, ...(match ? { sheetPart: match.part } : {}) }; }
  if (worksheet?.attrs.name) return { kind: 'table', tableName: worksheet.attrs.name, ...(worksheet.attrs.sheet ? { sheetName: worksheet.attrs.sheet } : {}) };
  if (child(node, 'external')) return { kind: 'unsupported', reason: 'external Pivot sources are not editable' };
  if (child(node, 'consolidation')) return { kind: 'unsupported', reason: 'consolidation sources are not editable' };
  return { kind: 'unsupported', reason: 'OLAP or unsupported Pivot source is not editable' };
}

function parseCacheFields(node: XmlNode | undefined): NativePivotCacheField[] {
  return children(node, 'cacheField').map((field, index) => {
    const shared = child(field, 'sharedItems');
    const values = parsePivotScalars(shared);
    const dataType = inferCacheType(shared);
    const fieldGroup = parseFieldGroup(child(field, 'fieldGroup'), `cacheField[${index}].fieldGroup`);
    return {
      index,
      name: field.attrs.name ?? `Field${index + 1}`,
      ...(dataType ? { dataType } : {}),
      ...(values.length ? { sharedItems: values } : {}),
      ...(fieldGroup ? { fieldGroup } : {}),
    };
  });
}

function parsePivotScalars(node: XmlNode | undefined): NativePivotScalar[] {
  if (!node) return [];
  return node.children.flatMap((item) => {
    switch (localName(item.name)) {
      case 's': return [item.attrs.v ?? null];
      case 'n': return [numberOrNull(item.attrs.v)];
      case 'd': return [item.attrs.v ?? null];
      case 'b': return [item.attrs.v === '1' || item.attrs.v?.toLowerCase() === 'true'];
      case 'e': return [nativePivotError(item.attrs.v)];
      case 'm': return [null];
      case 'groupItem': return [parseGroupItemValue(item)];
      default: return [];
    }
  });
}

function parseGroupItemValue(node: XmlNode): NativePivotScalar {
  const scalar = parsePivotScalars(node)[0];
  return scalar === undefined ? (node.attrs.v ?? null) : scalar;
}

function parseFieldGroup(node: XmlNode | undefined, label: string): NativePivotFieldGroup | undefined {
  if (!node) return undefined;
  const base = node.attrs.base === undefined ? undefined : requiredInteger(node.attrs.base, `${label}.base`);
  const parent = node.attrs.par === undefined ? undefined : requiredInteger(node.attrs.par, `${label}.par`);
  const rangeNode = child(node, 'rangePr');
  const range = rangeNode ? parseFieldRange(rangeNode, `${label}.rangePr`) : undefined;
  const discreteNode = child(node, 'discretePr');
  const discreteValues = discreteNode ? children(discreteNode, 'x').map((item, index) => requiredInteger(item.attrs.v, `${label}.discretePr.x[${index}]`)) : [];
  if (discreteNode?.attrs.count !== undefined && requiredInteger(discreteNode.attrs.count, `${label}.discretePr.count`) !== discreteValues.length) throw new Error(`${label}.discretePr.count does not match x elements`);
  const discreteIndexes = discreteNode ? discreteValues : undefined;
  const groupItemsNode = child(node, 'groupItems');
  const groupItems = groupItemsNode ? parsePivotScalars(groupItemsNode) : undefined;
  if (groupItemsNode?.attrs.count !== undefined && requiredInteger(groupItemsNode.attrs.count, `${label}.groupItems.count`) !== (groupItems?.length ?? 0)) throw new Error(`${label}.groupItems.count does not match group items`);
  return {
    ...(base === undefined ? {} : { base }),
    ...(parent === undefined ? {} : { parent }),
    ...(range ? { range } : {}),
    ...(discreteIndexes ? { discreteIndexes } : {}),
    ...(groupItems ? { groupItems } : {}),
  };
}

function parseFieldRange(node: XmlNode, label: string): NativePivotFieldRange {
  const groupBy = node.attrs.groupBy;
  const start = firstRangeValue(node.attrs, 'start');
  const end = firstRangeValue(node.attrs, 'end');
  const intervalValue = node.attrs.groupInterval ?? node.attrs.interval;
  const interval = intervalValue === undefined ? undefined : finiteNumber(intervalValue, `${label}.groupInterval`);
  if (interval !== undefined && interval <= 0) throw new Error(`${label}.groupInterval must be positive`);
  const autoStart = node.attrs.autoStart === undefined ? undefined : parseBoolean(node.attrs.autoStart, `${label}.autoStart`);
  const autoEnd = node.attrs.autoEnd === undefined ? undefined : parseBoolean(node.attrs.autoEnd, `${label}.autoEnd`);
  return {
    ...(groupBy === undefined ? {} : { groupBy }),
    ...(start === undefined ? {} : { start }),
    ...(end === undefined ? {} : { end }),
    ...(interval === undefined ? {} : { interval }),
    ...(autoStart === undefined ? {} : { autoStart }),
    ...(autoEnd === undefined ? {} : { autoEnd }),
  };
}

function firstRangeValue(attrs: Record<string, string>, name: 'start' | 'end'): NativePivotScalar | undefined {
  const date = attrs[`${name}Date`];
  if (date !== undefined) return date;
  const number = attrs[`${name}Num`];
  if (number !== undefined) {
    const parsed = Number(number);
    if (!Number.isFinite(parsed)) throw new Error(`Invalid ${name}Num range bound`);
    return parsed;
  }
  const generic = attrs[name];
  if (generic === undefined) return undefined;
  const numeric = Number(generic);
  return Number.isFinite(numeric) ? numeric : generic;
}

function finiteNumber(value: string, label: string): number {
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) throw new Error(`${label} must be finite`);
  return parsed;
}

function parseBoolean(value: string, label: string): boolean {
  if (value === '1' || value.toLowerCase() === 'true') return true;
  if (value === '0' || value.toLowerCase() === 'false') return false;
  throw new Error(`${label} must be boolean`);
}

function nativePivotSubtotal(field: NativePivotTableField | undefined): NonNullable<PivotFieldPlacement['subtotal']> | undefined {
  const subtotal = field?.subtotal;
  if (!subtotal) return undefined;
  return subtotal.mode === 'custom' ? { mode: 'custom', functions: (subtotal.functions ?? []).map(mapAggregate) } : { mode: subtotal.mode };
}

function nativePivotSort(field: NativePivotTableField | undefined, dataFields: NativePivotDataField[], fieldId: (index: number) => string): NonNullable<PivotFieldPlacement['sort']> | undefined {
  if (!field?.sortType || field.sortType === 'manual') return undefined;
  const scope = field.autoSortScope;
  const valueSort = scope?.dataOnly === true || scope?.labelOnly === false;
  if (!valueSort) return { direction: field.sortType, by: 'label' };
  const reference = scope?.references.find((candidate) => dataFields.some((dataField) => dataField.field === candidate.field));
  const valueField = reference ? dataFields.find((candidate) => candidate.field === reference.field) : undefined;
  if (!valueField) return undefined;
  return { direction: field.sortType, by: 'value', valueFieldId: fieldId(valueField.field) };
}

function mapNativePivotFilters(
  filters: NativePivotFilter[],
  fields: Array<{ fieldId: string }>,
  dataFields: NativePivotDataField[],
  pageFieldIndexes: ReadonlySet<number>,
): { filters: PivotLayout['filters']; preserved: PivotNativeFilterMetadata[] } {
  const mapped: PivotLayout['filters'] = [];
  const preserved: PivotNativeFilterMetadata[] = [];
  const fieldId = (index: number): string | undefined => fields[index]?.fieldId;
  const valueFieldId = (index: number | undefined): string | undefined => index === undefined ? undefined : dataFields[index]?.field !== undefined ? fields[dataFields[index]!.field]?.fieldId : undefined;
  for (const filter of filters) {
    const targetFieldId = fieldId(filter.field);
    const measureFieldId = valueFieldId(filter.measureField);
    const value = filter.value1 ?? filter.stringValue1;
    const type = filter.type.toLowerCase();
    const scope = pageFieldIndexes.has(filter.field) ? 'report' as const : 'field' as const;
    const condition = (operator: NonNullable<Extract<PivotLayout['filters'][number], { kind: 'condition' }>['operator']>, conditionValue: PivotScalar = value ?? null, withMeasure = false, upperValue: PivotScalar | undefined = filter.value2, dynamic: PivotDynamicDateFilter | undefined = undefined): void => {
      if (!targetFieldId || (withMeasure && !measureFieldId)) { preserved.push({ fieldIndex: filter.field, type: filter.type, attributes: { ...filter.attributes } }); return; }
      if (withMeasure) mapped.push({ kind: 'condition', family: 'value', fieldId: targetFieldId, ...(measureFieldId ? { valueFieldId: measureFieldId } : {}), operator: operator as PivotValueFilterOperator, value: conditionValue, ...(upperValue === undefined ? {} : { value2: upperValue }), scope });
      else if (type.startsWith('date')) mapped.push({ kind: 'condition', family: 'date', fieldId: targetFieldId, operator: operator as PivotDateFilterOperator, value: conditionValue, ...(upperValue === undefined ? {} : { value2: upperValue }), ...(dynamic === undefined ? {} : { dynamic }), scope, ...(filter.wholeDay === undefined ? {} : { wholeDay: filter.wholeDay }) });
      else mapped.push({ kind: 'condition', family: 'label', fieldId: targetFieldId, operator: operator as PivotLabelFilterOperator, value: conditionValue, ...(upperValue === undefined ? {} : { value2: upperValue }), scope });
    };
    if (type === 'captionequal' || type === 'dateequal') condition('equals');
    else if (type === 'captionnotequal' || type === 'datenotequal') condition('not-equals');
    else if (type === 'captioncontains') condition('contains');
    else if (type === 'captionnotcontains') condition('not-contains');
    else if (type === 'captionbeginswith') condition('begins-with');
    else if (type === 'captionnotbeginswith') condition('not-begins-with');
    else if (type === 'captionendswith') condition('ends-with');
    else if (type === 'captionnotendswith') condition('not-ends-with');
    else if (type === 'captionbetween') condition('between');
    else if (type === 'captionnotbetween') condition('not-between');
    else if (type === 'captiongreaterthan') condition('greater-than');
    else if (type === 'captiongreaterthanorequal') condition('greater-or-equal');
    else if (type === 'captionlessthan') condition('less-than');
    else if (type === 'captionlessthanorequal') condition('less-or-equal');
    else if (type === 'datenewerthan') condition('after');
    else if (type === 'dateolderthan') condition('before');
    else if (type === 'datebetween') condition('between');
    else if (type === 'datenotbetween') condition('not-between');
    else if (type === 'datetoday' || type === 'dateyesterday' || type === 'datetomorrow' || type === 'datethisweek' || type === 'datelastweek' || type === 'datenextweek' || type === 'datethismonth' || type === 'datelastmonth' || type === 'datenextmonth' || type === 'datethisquarter' || type === 'datelastquarter' || type === 'datenextquarter' || type === 'datethisyear' || type === 'datelastyear' || type === 'datenextyear' || type === 'dateyeartodate') {
      const dynamic: Record<string, PivotDynamicDateFilter> = { datetoday: 'today', dateyesterday: 'yesterday', datetomorrow: 'tomorrow', datethisweek: 'this-week', datelastweek: 'last-week', datenextweek: 'next-week', datethismonth: 'this-month', datelastmonth: 'last-month', datenextmonth: 'next-month', datethisquarter: 'this-quarter', datelastquarter: 'last-quarter', datenextquarter: 'next-quarter', datethisyear: 'this-year', datelastyear: 'last-year', datenextyear: 'next-year', dateyeartodate: 'year-to-date' };
      condition('between', null, false, undefined, dynamic[type]);
    }
    else if (type === 'valueequal') condition('equals', value ?? null, true);
    else if (type === 'valuenotequal') condition('not-equals', value ?? null, true);
    else if (type === 'valuegreaterthan') condition('greater-than', value ?? null, true);
    else if (type === 'valuegreaterthanorequal') condition('greater-or-equal', value ?? null, true);
    else if (type === 'valuelessthan') condition('less-than', value ?? null, true);
    else if (type === 'valuelessthanorequal') condition('less-or-equal', value ?? null, true);
    else if (type === 'valuebetween') condition('between', value ?? null, true);
    else if (type === 'valuenotbetween') condition('not-between', value ?? null, true);
    else if (type === 'valuetop10' || type === 'top10' || type === 'valuebottom10' || type === 'bottom10') {
      const count = typeof value === 'number' && Number.isSafeInteger(value) ? value : undefined;
      if (!targetFieldId || !measureFieldId || !count || count < 1 || filter.percent) preserved.push({ fieldIndex: filter.field, type: filter.type, attributes: { ...filter.attributes } });
      else mapped.push({ kind: 'top-items', family: 'top-items', fieldId: targetFieldId, valueFieldId: measureFieldId, count, direction: filter.top === false || type === 'valuebottom10' || type === 'bottom10' ? 'bottom' : 'top', scope });
    } else preserved.push({ fieldIndex: filter.field, type: filter.type, attributes: { ...filter.attributes } });
  }
  return { filters: mapped, preserved };
}

function parsePivotFieldSubtotal(field: XmlNode, index: number): NativePivotTableField['subtotal'] | undefined {
  const attrs = field.attrs;
  const enabled: Array<[string, string]> = [
    ['sumSubtotal', 'sum'], ['countSubtotal', 'count'], ['countASubtotal', 'count-numbers'], ['averageSubtotal', 'average'],
    ['maxSubtotal', 'max'], ['minSubtotal', 'min'], ['productSubtotal', 'product'], ['stdDevSubtotal', 'stdev'], ['stdDevpSubtotal', 'stdevp'],
    ['varSubtotal', 'var'], ['varPSubtotal', 'varp'], ['distinctCountSubtotal', 'distinct-count'],
  ];
  const functions = enabled.flatMap(([name, value]) => attrs[name] === '1' || attrs[name] === 'true' ? [value] : []);
  if (functions.length) return { mode: 'custom', functions };
  if (attrs.defaultSubtotal === '0' || attrs.subtotal === '0') return { mode: 'none' };
  if (attrs.defaultSubtotal === '1' || attrs.subtotal === '1') return { mode: 'automatic' };
  void index;
  return undefined;
}

function parsePivotFields(node: XmlNode | undefined): NativePivotTableField[] {
  return children(node, 'pivotField').map((field, index) => {
    const subtotal = parsePivotFieldSubtotal(field, index);
    const autoSortScope = parseAutoSortScope(child(field, 'autoSortScope'), `pivotField[${index}].autoSortScope`);
    const collapsedItemIndexes = children(child(field, 'items'), 'item').flatMap((item) => item.attrs.sd === '0' && item.attrs.x !== undefined
      ? [requiredInteger(item.attrs.x, `pivotField[${index}].items.item.x`)] : []);
    const hiddenItemIndexes = children(child(field, 'items'), 'item').flatMap((item) => {
      if (item.attrs.h === undefined) return [];
      const hidden = parseBoolean(item.attrs.h, `pivotField[${index}].items.item.h`);
      if (!hidden) return [];
      if (item.attrs.x === undefined) throw new Error(`pivotField[${index}].items.item.x is required for a hidden item`);
      return [requiredInteger(item.attrs.x, `pivotField[${index}].items.item.x`)];
    });
    if (new Set(hiddenItemIndexes).size !== hiddenItemIndexes.length) throw new Error(`pivotField[${index}].items contains duplicate hidden item indexes`);
    const sortType = field.attrs.sortType;
    if (sortType !== undefined && !['manual', 'ascending', 'descending'].includes(sortType)) {
      throw new Error(`pivotField[${index}].sortType is unsupported: ${sortType}`);
    }
    return {
      index,
      ...(field.attrs.axis === 'axisRow' ? { axis: 'row' as const } : field.attrs.axis === 'axisCol' ? { axis: 'column' as const } : field.attrs.axis === 'axisPage' ? { axis: 'page' as const } : {}),
      ...(field.attrs.compact === '0' ? { compact: false } : field.attrs.compact === '1' ? { compact: true } : {}),
      ...(field.attrs.outline === '0' ? { outline: false } : field.attrs.outline === '1' ? { outline: true } : {}),
      ...(sortType ? { sortType: sortType as NativePivotTableField['sortType'] } : {}),
      ...optionalBoolean(field.attrs.nonAutoSortDefault, 'nonAutoSortDefault'),
      ...(autoSortScope ? { autoSortScope } : {}),
      ...(subtotal ? { subtotal } : {}),
      ...(collapsedItemIndexes.length ? { collapsedItemIndexes } : {}),
      ...(hiddenItemIndexes.length ? { hiddenItemIndexes } : {}),
    };
  });
}

function parseAutoSortScope(node: XmlNode | undefined, label: string): NativePivotAutoSortScope | undefined {
  if (!node) return undefined;
  const area = child(node, 'pivotArea');
  if (!area) throw new Error(`${label}.pivotArea is missing`);
  const referencesNode = child(area, 'references');
  const references = children(referencesNode, 'reference').map((reference, index) => ({
    field: requiredInteger(reference.attrs.field, `${label}.pivotArea.references.reference[${index}].field`),
    ...optionalBoolean(reference.attrs.selected, 'selected'),
    ...(() => {
      const itemIndexes = children(reference, 'x').map((item) => requiredInteger(item.attrs.v, `${label}.pivotArea.references.reference[${index}].x.v`));
      return itemIndexes.length ? { itemIndexes } : {};
    })(),
  }));
  return {
    ...(area ? optionalBoolean(area.attrs.dataOnly, 'dataOnly') : {}),
    ...(area ? optionalBoolean(area.attrs.labelOnly, 'labelOnly') : {}),
    ...(area ? optionalBoolean(area.attrs.outline, 'outline') : {}),
    ...(area?.attrs.fieldPosition !== undefined ? { fieldPosition: requiredInteger(area.attrs.fieldPosition, `${label}.pivotArea.fieldPosition`) } : {}),
    attributes: { ...(area?.attrs ?? {}) },
    references,
  };
}

function parsePivotFilters(node: XmlNode | undefined): NativePivotFilter[] {
  return children(node, 'filter').map((filter, index) => {
    const field = requiredInteger(filter.attrs.fld, `pivotFilters.filter[${index}].fld`);
    const type = filter.attrs.type;
    if (!type) throw new Error(`pivotFilters.filter[${index}].type is missing`);
    return {
      field,
      type,
      ...(optionalInteger(filter.attrs.iMeasureFld, `pivotFilters.filter[${index}].iMeasureFld`) ?? {}),
      ...(optionalInteger(filter.attrs.iMeasureFld2, `pivotFilters.filter[${index}].iMeasureFld2`) ?? {}),
      ...(optionalInteger(filter.attrs.evalOrder, `pivotFilters.filter[${index}].evalOrder`) ?? {}),
      ...(optionalInteger(filter.attrs.id, `pivotFilters.filter[${index}].id`) ?? {}),
      ...(filter.attrs.stringValue1 !== undefined ? { stringValue1: filter.attrs.stringValue1 } : {}),
      ...(filter.attrs.stringValue2 !== undefined ? { stringValue2: filter.attrs.stringValue2 } : {}),
      ...(filter.attrs.val !== undefined || filter.attrs.value1 !== undefined ? { value1: nativeFilterScalar(filter.attrs.val ?? filter.attrs.value1!) } : {}),
      ...(filter.attrs.val2 !== undefined || filter.attrs.value2 !== undefined ? { value2: nativeFilterScalar(filter.attrs.val2 ?? filter.attrs.value2!) } : {}),
      ...(filter.attrs.wholeDay !== undefined ? optionalBoolean(filter.attrs.wholeDay, 'wholeDay') : {}),
      ...(filter.attrs.top !== undefined ? optionalBoolean(filter.attrs.top, 'top') : {}),
      ...(filter.attrs.percent !== undefined ? optionalBoolean(filter.attrs.percent, 'percent') : {}),
      attributes: { ...filter.attrs },
    };
  });
}

function optionalInteger(value: string | undefined, label: string): Record<string, number> | undefined {
  return value === undefined ? undefined : { [label.endsWith('iMeasureFld') ? 'measureField' : label.endsWith('iMeasureFld2') ? 'secondMeasureField' : label.endsWith('evalOrder') ? 'evalOrder' : 'id']: requiredInteger(value, label) };
}

function nativeFilterScalar(value: string): NativePivotScalar {
  if (value === '') return '';
  if (value.startsWith('#')) return nativePivotError(value);
  if (value === '0' || value === '1') return Number(value);
  const number = Number(value);
  return Number.isFinite(number) && value.trim() !== '' ? number : value;
}
function parseFieldIndexes(node: XmlNode | undefined, label: string): number[] { return children(node, 'field').map((field) => requiredInteger(field.attrs.x, `${label}.field.x`)); }
function parsePageFieldIndexes(node: XmlNode | undefined, label: string): number[] { return children(node, 'pageField').map((field) => requiredInteger(field.attrs.fld, `${label}.pageField.fld`)); }
function parseDataFields(node: XmlNode | undefined): NativePivotDataField[] { return children(node, 'dataField').map((field) => ({ field: requiredInteger(field.attrs.fld, 'dataField.fld'), ...(field.attrs.name ? { name: field.attrs.name } : {}), ...(field.attrs.subtotal ? { subtotal: field.attrs.subtotal } : {}), ...(field.attrs.showDataAs ? { showDataAs: field.attrs.showDataAs } : {}) })); }
function readSheetNames(workbook: XmlNode, rels: XlsxRelationship[], parts: Record<string, string>): Map<string, { name: string; part: string }> { const result = new Map<string, { name: string; part: string }>(); for (const [id, part] of Object.entries(parts)) { const node = children(child(workbook, 'sheets'), 'sheet').find((candidate) => `sheet-${candidate.attrs.sheetId}` === id); if (node?.attrs.name) result.set(id, { name: node.attrs.name, part }); } for (const node of children(child(workbook, 'sheets'), 'sheet')) { const relation = rels.find((candidate) => candidate.id === (node.attrs['r:id'] ?? node.attrs.id)); if (relation && node.attrs.name) result.set(`sheet-${node.attrs.sheetId ?? result.size + 1}`, { name: node.attrs.name, part: resolveTarget('xl/workbook.xml', relation.target) }); } return result; }
function requireRelationship(rels: XlsxRelationship[], id: string, type: string, context: string): XlsxRelationship { const relation = rels.find((candidate) => candidate.id === id); if (!relation) throw new Error(`${context} relation ${id} is missing`); if (relation.type !== type && !relation.type.endsWith(`/${type.split('/').pop()!}`)) throw new Error(`${context} relation ${id} has unexpected type ${relation.type}`); return relation; }
function firstElement(root: XmlNode, name: string): XmlNode { const found = localName(root.name) === name ? root : descendants(root, name)[0]; if (!found) throw new Error(`OOXML part is missing <${name}>`); return found; }
function resolveTarget(source: string, target: string): string { if (target.startsWith('/')) return normalizePartName(target.slice(1)); const base = source.includes('/') ? source.slice(0, source.lastIndexOf('/') + 1) : ''; return normalizePartName(`${base}${target}`); }
function normalizePartName(name: string): string { const pieces: string[] = []; for (const piece of name.replaceAll('\\', '/').split('/')) { if (!piece || piece === '.') continue; if (piece === '..') { if (!pieces.length) throw new Error(`Unsafe XLSX part target: ${name}`); pieces.pop(); } else pieces.push(piece); } const result = pieces.join('/'); if (!result || result.includes('\0')) throw new Error(`Unsafe XLSX part target: ${name}`); return result; }
function cloneFiles(files: Record<string, Uint8Array>): Record<string, Uint8Array> { return Object.fromEntries(Object.entries(files).map(([name, bytes]) => [name, bytes.slice()])); }
function cloneRelationships(input: Record<string, XlsxRelationship[]>): Record<string, XlsxRelationship[]> { return Object.fromEntries(Object.entries(input).map(([source, list]) => [source, list.map((relation) => ({ ...relation }))])); }
function allocateId(list: XlsxRelationship[]): string { const used = new Set(list.map((relation) => relation.id)); let index = 1; while (used.has(`rId${index}`)) index += 1; return `rId${index}`; }
function relativeTarget(source: string, target: string): string { const left = source.slice(0, source.lastIndexOf('/') + 1).split('/').filter(Boolean); const right = target.split('/').filter(Boolean); while (left.length && right.length && left[0] === right[0]) { left.shift(); right.shift(); } return `${'../'.repeat(left.length)}${right.join('/')}`; }
function isPivotRelation(relation: XlsxRelationship): boolean { return relation.type.endsWith('/pivotCacheDefinition') || relation.type.endsWith('/pivotCacheRecords') || relation.type.endsWith('/pivotTable'); }
function isNativePivotPart(name: string): boolean { return /^xl\/(pivotTables\/|pivotCache\/)/i.test(name); }
function isNativeControlPart(name: string): boolean { return /^xl\/(slicers\/|slicerCaches\/|timelines\/|timelineCaches\/)/i.test(name); }
function nextPartNumbers(files: Record<string, Uint8Array>): { cacheDefinition: number; records: number; table: number } { const max = (pattern: RegExp): number => Object.keys(files).reduce((value, name) => Math.max(value, Number(name.match(pattern)?.[1] ?? 0)), 0); return { cacheDefinition: max(/pivotCacheDefinition(\d+)\.xml$/i) + 1, records: max(/pivotCacheRecords(\d+)\.xml$/i) + 1, table: max(/pivotTable(\d+)\.xml$/i) + 1 }; }
function nativeFieldId(cacheId: number, index: number): string { return `native:cache:${cacheId}:field:${index}`; }
function nativePivotId(table: NativePivotTableDefinition): string { return `native:pivot:${table.part}`; }
function mapFieldType(value: NativePivotCacheField['dataType']): PivotFieldDataType {
  return value === 'string' ? 'text' : value === 'number' ? 'number' : value === 'date' ? 'date' : value === 'boolean' ? 'boolean' : value === 'error' ? 'error' : 'mixed';
}
function nativeDataType(value: PivotFieldDataType): NativePivotCacheField['dataType'] {
  return value === 'text' ? 'string' : value === 'number' ? 'number' : value === 'date' ? 'date' : value === 'boolean' ? 'boolean' : value === 'error' ? 'error' : 'mixed';
}
function nativePivotGroup(field: NativePivotCacheField | undefined, cache: NativePivotCacheDefinition, fieldId: string): PivotGroup | undefined {
  const grouping = field?.fieldGroup;
  if (!grouping) return undefined;
  const baseField = grouping.base === undefined ? field : cache.fields[grouping.base] ?? field;
  const range = grouping.range;
  const groupBy = range?.groupBy?.toLowerCase();
  const dateUnits: Record<string, 'year' | 'quarter' | 'month' | 'week' | 'day'> = {
    year: 'year', years: 'year', quarter: 'quarter', quarters: 'quarter',
    month: 'month', months: 'month', week: 'week', weeks: 'week', day: 'day', days: 'day',
  };
  if (groupBy && dateUnits[groupBy]) {
    return {
      kind: 'date', unit: dateUnits[groupBy],
      ...(range?.start !== undefined ? { start: range.start } : {}),
      ...(range?.end !== undefined ? { end: range.end } : {}),
      ...(range?.autoStart !== undefined ? { autoStart: range.autoStart } : {}),
      ...(range?.autoEnd !== undefined ? { autoEnd: range.autoEnd } : {}),
    };
  }
  if (groupBy === 'range' || groupBy === 'number' || groupBy === 'numbers' || (baseField?.dataType === 'number' && range?.interval !== undefined)) {
    const values = (grouping.groupItems ?? []).filter((value): value is number => typeof value === 'number');
    const derivedInterval = values.length > 1 ? Math.abs(values[1]! - values[0]!) : undefined;
    const interval = range?.interval ?? (derivedInterval && Number.isFinite(derivedInterval) && derivedInterval > 0 ? derivedInterval : undefined);
    if (interval === undefined || !Number.isFinite(interval) || interval <= 0) return undefined;
    return {
      kind: 'number', interval,
      ...(typeof range?.start === 'number' ? { start: range.start } : {}),
      ...(typeof range?.end === 'number' ? { end: range.end } : {}),
      ...(range?.autoStart !== undefined ? { autoStart: range.autoStart } : {}),
      ...(range?.autoEnd !== undefined ? { autoEnd: range.autoEnd } : {}),
    };
  }
  if (grouping.discreteIndexes && grouping.groupItems) {
    const groups = grouping.groupItems.map((name, groupIndex) => ({
      groupId: `${fieldId}:group:${groupIndex}`,
      name: name === null ? '' : String(name),
      items: (baseField?.sharedItems ?? []).flatMap((value, itemIndex) => grouping.discreteIndexes?.[itemIndex] === groupIndex ? [createPivotMemberKey(value)] : []),
    }));
    return { kind: 'manual', groups };
  }
  return undefined;
}
function inferDataType(sheet: SheetSnapshot, range: RangeRef, column: number): PivotFieldDataType {
  const values: unknown[] = [];
  for (let row = range.startRow + 1; row <= range.endRow; row += 1) {
    const cell = sheet.cells[String(row)]?.[String(column)];
    values.push(cell?.formulaValue ?? cell?.value);
  }
  const present = values.filter((value) => value !== null && value !== undefined);
  if (present.length && present.every(isPivotError)) return 'error';
  if (present.some(isPivotError)) return 'mixed';
  if (present.length && present.every((value) => typeof value === 'number')) return 'number';
  if (present.length && present.every((value) => typeof value === 'boolean')) return 'boolean';
  return 'text';
}
function mapAggregate(value: string | undefined): PivotAggregateFunction { switch ((value ?? 'sum').toLowerCase()) { case 'count': return 'count'; case 'countnums': case 'count-numbers': return 'count-numbers'; case 'average': case 'avg': return 'average'; case 'min': return 'min'; case 'max': return 'max'; case 'product': return 'product'; case 'stddev': case 'stdev': return 'stdev'; case 'stddevp': case 'stdevp': return 'stdevp'; case 'var': return 'var'; case 'varp': return 'varp'; case 'distinctcount': case 'distinct-count': return 'distinct-count'; default: return 'sum'; } }
function nativeAggregate(value: string | undefined): string { switch (value) { case 'count-numbers': return 'countNums'; case 'stdev': return 'stdDev'; case 'stdevp': return 'stdDevp'; case 'distinct-count': return 'distinctCount'; default: return value ?? 'sum'; } }
function mapShowAs(value: string): NonNullable<PivotValueField['showAs']> { return value === 'percentOfTotal' ? { kind: 'grand-percentage' } : value === 'percentOfRow' ? { kind: 'row-percentage' } : value === 'percentOfCol' ? { kind: 'column-percentage' } : value === 'runningTotal' ? { kind: 'running-total', axis: 'row' } : { kind: 'normal' }; }
function nativeShowAs(value: NonNullable<PivotValueField['showAs']>['kind']): string | undefined { return value === 'grand-percentage' ? 'percentOfTotal' : value === 'row-percentage' ? 'percentOfRow' : value === 'column-percentage' ? 'percentOfCol' : value === 'running-total' ? 'runningTotal' : undefined; }
function parseRange(value: string, sheetId: string): RangeRef | undefined { const parts = value.split(':'); const start = parseA1(parts[0] ?? ''); const end = parseA1(parts[1] ?? parts[0] ?? ''); return start && end ? { sheetId, startRow: start.row, endRow: end.row, startColumn: start.column, endColumn: end.column } : undefined; }
function rangeToA1(range: RangeRef): string { const start = a1(range.startRow, range.startColumn); const end = a1(range.endRow, range.endColumn); return start === end ? start : `${start}:${end}`; }
function deriveLocation(pivot: PivotDefinition, rows: number, rowFields: number, columnFields: number, values: number): string { return `${a1(pivot.target.anchor.row, pivot.target.anchor.column)}:${a1(pivot.target.anchor.row + Math.max(4, rows + rowFields + 2) - 1, pivot.target.anchor.column + Math.max(2, rowFields + columnFields * Math.max(1, values) + 1) - 1)}`; }
function parseA1(value: string): { row: number; column: number } | undefined { const match = /^\$?([A-Za-z]+)\$?(\d+)$/.exec(value.trim()); return match ? { row: Number(match[2]) - 1, column: columnFromLetter(match[1]!) } : undefined; }
function a1(row: number, column: number): string { return `${columnToLetter(column)}${row + 1}`; }
function columnFromLetter(value: string): number { let result = 0; for (const character of value.toUpperCase()) result = result * 26 + character.charCodeAt(0) - 64; return result - 1; }
function columnToLetter(index: number): string { let value = index + 1; let result = ''; while (value > 0) { const remainder = (value - 1) % 26; result = String.fromCharCode(65 + remainder) + result; value = Math.floor((value - 1) / 26); } return result; }
function isScalar(value: unknown): value is PivotScalar { return value === null || typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean' || isPivotError(value); }
function nativePivotError(code: string | undefined): PivotErrorValue {
  const allowed = new Set(['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#CALC!', '#BLOCKED!', '#SPILL!', '#PARSE!', '#CYCLE!']);
  if (!code || !allowed.has(code)) throw new Error(`Unsupported native Pivot error value: ${code ?? '<missing>'}`);
  return { kind: 'error', code: code as PivotErrorValue['code'] };
}
function scalarKey(value: PivotScalar): string { return pivotMemberKey(createPivotMemberKey(value)); }
function uniqueScalars(values: PivotScalar[]): PivotScalar[] { const result: PivotScalar[] = []; const seen = new Set<string>(); for (const value of values) { const key = scalarKey(value); if (!seen.has(key)) { seen.add(key); result.push(value); } } return result; }
function uniqueTuples(rows: PivotScalar[][], indexes: number[]): PivotScalar[][] { if (!indexes.length) return [[]]; const seen = new Set<string>(); const result: PivotScalar[][] = []; for (const row of rows) { const tuple = indexes.map((index) => row[index] ?? null); const key = tuple.map(scalarKey).join('|'); if (!seen.has(key)) { seen.add(key); result.push(tuple); } } return result; }
function matchesTuple(row: PivotScalar[], indexes: number[], tuple: PivotScalar[]): boolean { return indexes.every((index, position) => scalarKey(row[index] ?? null) === scalarKey(tuple[position] ?? null)); }
function aggregate(values: PivotScalar[], operation: string | undefined): PivotScalar {
  const numbers = values.filter((value): value is number => typeof value === 'number' && Number.isFinite(value));
  const firstError = values.find(isPivotError);
  switch (mapAggregate(operation)) {
    case 'count': return values.filter((value) => value !== null).length;
    case 'count-numbers': return numbers.length;
    case 'average': return firstError ?? (numbers.length ? numbers.reduce((sum, value) => sum + value, 0) / numbers.length : null);
    case 'min': return firstError ?? (numbers.length ? Math.min(...numbers) : null);
    case 'max': return firstError ?? (numbers.length ? Math.max(...numbers) : null);
    case 'product': return firstError ?? (numbers.length ? numbers.reduce((product, value) => product * value, 1) : null);
    case 'distinct-count': return new Set(values.filter((value) => value !== null).map(scalarKey)).size;
    default: return firstError ?? (numbers.length ? numbers.reduce((sum, value) => sum + value, 0) : values.filter((value) => value !== null).length || null);
  }
}
function inferCacheType(node: XmlNode | undefined): NativePivotCacheField['dataType'] | undefined { if (!node) return undefined; if (node.attrs.containsString === '1' || node.attrs.containsString === 'true') return 'string'; if (node.attrs.containsDate === '1' || node.attrs.containsDate === 'true') return 'date'; if (node.attrs.containsNumber === '1' || node.attrs.containsNumber === 'true') return 'number'; if (node.attrs.containsBoolean === '1' || node.attrs.containsBoolean === 'true') return 'boolean'; if (node.attrs.containsError === '1' || node.attrs.containsError === 'true') return 'error'; return undefined; }
function numberOrNull(value: string | undefined): number | null { const number = Number(value); return Number.isFinite(number) ? number : null; }
function optionalBoolean(value: string | undefined, key: string): Record<string, boolean> { return value === undefined ? {} : { [key]: value === '1' || value.toLowerCase() === 'true' }; }
function boolAttr(value: boolean | undefined, name: string): string[] { return value === undefined ? [] : [`${name}="${value ? '1' : '0'}"`]; }
function requiredInteger(value: string | undefined, label: string): number { const number = Number(value); if (!Number.isSafeInteger(number) || number < 0) throw new Error(`${label} must be a non-negative integer`); return number; }
function withXmlDeclaration(xml: string): string { return xml.startsWith('<?xml') ? xml : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${xml}`; }
