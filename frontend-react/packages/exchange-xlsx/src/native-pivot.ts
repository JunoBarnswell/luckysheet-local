import { strFromU8, strToU8 } from 'fflate';
import type {
  CellData,
  PivotAggregateFunction,
  PivotDefinition,
  PivotFieldDataType,
  PivotFieldPlacement,
  PivotLayout,
  PivotModel,
  PivotScalar,
  PivotSource,
  PivotValueField,
  RangeRef,
  SheetSnapshot,
  WorkbookSnapshot,
} from '@react-sheets/core-model';
import { child, children, descendants, encodeXml, localName, parseXml, serializeXml, type XmlNode } from './xml';
import type {
  NativePivotCacheDefinition,
  NativePivotCacheField,
  NativePivotControlDefinition,
  NativePivotDataField,
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
const REL_TIMELINE_CACHE = 'http://schemas.microsoft.com/office/2010/relationships/TimelineCache';
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
        pageFields: parseFieldIndexes(child(definition, 'pageFields'), 'pageFields'),
        dataFields: parseDataFields(child(definition, 'dataFields')),
        ...optionalBoolean(definition.attrs.rowGrandTotals ?? definition.attrs.showRowGrandTotals, 'showRowGrandTotals'),
        ...optionalBoolean(definition.attrs.colGrandTotals ?? definition.attrs.showColumnGrandTotals, 'showColumnGrandTotals'),
        ...optionalBoolean(definition.attrs.compactData, 'compactData'),
        ...optionalBoolean(definition.attrs.repeatAllLabels, 'repeatLabels'),
        ...(style?.attrs.name ? { styleName: style.attrs.name } : {}),
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
        if (drawingPart && input.files[drawingPart] && strFromU8(input.files[drawingPart]!).includes(`name="${encodeXml(control.name)}"`)) { control.drawingPart = drawingPart; control.drawingRelationshipId = drawingRelation!.id; }
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
        if (drawingPart && input.files[drawingPart] && strFromU8(input.files[drawingPart]!).includes(`name="${encodeXml(control.name)}"`)) { control.drawingPart = drawingPart; control.drawingRelationshipId = drawingRelation!.id; }
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
    ...(cachePart?.selection ? { selection: cachePart.selection } : {}),
    ...(cachePart?.selectedItemIndexes ? { selectedItemIndexes: cachePart.selectedItemIndexes } : {}),
    ...(node.attrs.style ? { styleName: node.attrs.style } : {}),
    ...(node.attrs.caption ? { caption: node.attrs.caption } : {}),
    valid,
    ...(valid ? {} : { reason: `Unable to validate ${kind} cache, PivotTable, or field binding` }),
  };
}

function isSlicerCacheRelation(relation: XlsxRelationship): boolean { return relation.type === REL_SLICER_CACHE_MODERN || relation.type === REL_SLICER_CACHE || relation.type.endsWith('/slicerCache'); }
function isSlicerRelation(relation: XlsxRelationship): boolean { return relation.type === REL_SLICER || relation.type.endsWith('/slicer'); }
function isTimelineCacheRelation(relation: XlsxRelationship): boolean { return relation.type === REL_TIMELINE_CACHE || relation.type.endsWith('/TimelineCache'); }
function isTimelineRelation(relation: XlsxRelationship): boolean { return relation.type === REL_TIMELINE || relation.type.endsWith('/timeline'); }

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
  const layout: PivotLayout = {
    rows: table.rowFields.map((index) => ({ fieldId: fieldId(index) })),
    columns: table.columnFields.map((index) => ({ fieldId: fieldId(index) })),
    filters: table.pageFields.map((index) => ({ kind: 'manual' as const, fieldId: fieldId(index), mode: 'all' as const, memberKeys: [] })),
    values: table.dataFields.map((data) => ({ fieldId: fieldId(data.field), summarizeBy: mapAggregate(data.subtotal), ...(data.name ? { displayName: data.name } : {}), ...(data.showDataAs ? { showAs: mapShowAs(data.showDataAs) } : {}) })),
    showSubtotals: table.showSubtotals ?? true,
    showGrandTotals: (table.showRowGrandTotals ?? true) || (table.showColumnGrandTotals ?? true),
    compact: table.compactData ?? table.fields.some((field) => field.compact === true),
    repeatLabels: table.repeatLabels ?? false,
    expansion: { expandedNodeIds: [], collapsedNodeIds: [], showButtons: true },
  };
  const fieldBindings: Record<string, { cacheFieldIndex: number; sourceName?: string }> = {};
  for (const field of fields) fieldBindings[field.fieldId] = { cacheFieldIndex: field.ordinal, sourceName: field.name };
  return {
    schema: 'PivotDefinition',
    id: nativePivotId(table),
    source,
    target: { sheetId: target.id, anchor: { row: location.startRow, column: location.startColumn } },
    fieldCatalog: { schema: 'PivotFieldCatalog', fields },
    layout,
    refreshPolicy: { mode: cache.refreshOnLoad ? 'on-open' : 'manual', preserveFormatting: true, refreshOnLoad: cache.refreshOnLoad ?? true },
    nativeMetadata: { cacheId: cache.cacheId, cacheDefinitionPart: cache.part, ...(cache.recordsPart ? { cacheRecordsPart: cache.recordsPart } : {}), pivotTablePart: table.part, fieldBindings },
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
  const pivot = pivots.some((candidate) => candidate.source.kind === 'worksheet-range' || candidate.source.kind === 'table');
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
  const pivotEntries = input.snapshot.sheets.flatMap((sheet) => sheet.pivots.map((pivot) => ({ sheet, pivot })));

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
    let cache = pivot.nativeMetadata?.cacheId === undefined ? undefined : existingCaches.get(pivot.nativeMetadata.cacheId);
    if (!cache || cache.source.kind === 'unsupported' || nativeSourceKey(cache.source) !== sourceInfo.key) cache = cacheBySource.get(sourceInfo.key);
    if (!cache) {
      cache = { cacheId: nextCacheId++, part: `xl/pivotCache/pivotCacheDefinition${partNumbers.cacheDefinition++}.xml`, recordsPart: `xl/pivotCache/pivotCacheRecords${partNumbers.records++}.xml`, source: sourceInfo.source, fields: [], refreshOnLoad: pivot.refreshPolicy.refreshOnLoad, refreshOnSave: pivot.refreshPolicy.mode === 'on-change', saveData: true, enableRefresh: true };
    }
    usedCaches.add(cache.cacheId);
    cacheBySource.set(sourceInfo.key, cache);
    const sourceRows = readSourceRows(sourceInfo.sheet, sourceInfo.range, pivot, sourceInfo.tableName);
    cache.source = sourceInfo.source;
    cache.fields = sourceRows.fields.map((field, index) => ({ index, name: field.name, dataType: field.dataType, sharedItems: uniqueScalars(sourceRows.rows.map((row) => row[index] ?? null)) }));
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
    if (entry.payload.kind === 'slicer') {
      files[cachePart] = strToU8(buildSlicerCacheXml(control, entry.payload, pivot, cache, input.tables));
      files[part] = strToU8(buildSlicerXml(control, entry.payload));
    } else {
      files[cachePart] = strToU8(buildTimelineCacheXml(control, entry.payload, pivot, input.tables));
      files[part] = strToU8(buildTimelineXml(control, entry.payload));
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

function buildSlicerCacheXml(control: NativePivotControlDefinition, payload: { kind: 'slicer'; pivotId: string; fieldId: string; filter: { mode: 'all' | 'include' | 'exclude'; memberKeys: Array<{ type: 'text' | 'number' | 'boolean' | 'blank'; value: string | number | boolean | null }> }; connectedPivotIds?: string[] }, pivot: PivotModel, cache: NativePivotCacheDefinition, tables: NativePivotTableDefinition[]): string {
  const field = cache.fields[control.fieldIndex ?? -1];
  const values = field?.sharedItems ?? pivot.fieldCatalog.fields.find((candidate) => candidate.fieldId === payload.fieldId)?.values ?? [];
  const selected = new Set(payload.filter.memberKeys.map((value) => `${value.type}:${JSON.stringify(value.value)}`));
  const items = values.map((value, index) => {
    const key = nativeMemberKey(value);
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

function buildTimelineCacheXml(control: NativePivotControlDefinition, payload: { kind: 'timeline'; pivotId: string; fieldId: string; period: { start?: string; end?: string }; connectedPivotIds?: string[] }, pivot: PivotModel, tables: NativePivotTableDefinition[]): string {
  const field = pivot.fieldCatalog.fields.find((candidate) => candidate.fieldId === payload.fieldId);
  const connected = [...new Set([payload.pivotId, ...(payload.connectedPivotIds ?? [])])].flatMap((pivotId) => { const table = tables.find((candidate) => candidate.pivotId === pivotId || candidate.name === pivotId); return table ? [`<pivotTable tabId="1" name="${encodeXml(table.name)}"/>`] : []; }).join('');
  const start = payload.period.start ?? '2000-01-01T00:00:00Z';
  const end = payload.period.end ?? '2100-01-01T00:00:00Z';
  const filterType = payload.period.start && payload.period.end ? 'dateBetween' : 'unknown';
  return withXmlDeclaration(`<timelineCacheDefinition xmlns="${NS_X15}" xmlns:x15="${NS_X15}" name="${encodeXml(control.cacheName)}" sourceName="${encodeXml(field?.name ?? payload.fieldId)}"><pivotTables>${connected}</pivotTables><state minimalRefreshVersion="6" lastRefreshVersion="6" pivotCacheId="${control.pivotCacheId ?? 0}" filterType="${filterType}"><selection startDate="${encodeXml(start)}" endDate="${encodeXml(end)}"/><bounds startDate="2000-01-01T00:00:00Z" endDate="2100-01-01T00:00:00Z"/></state></timelineCacheDefinition>`);
}

function buildTimelineXml(control: NativePivotControlDefinition, payload: { kind: 'timeline'; period: { start?: string; end?: string } }): string {
  return withXmlDeclaration(`<timelines xmlns="${NS_X15}" xmlns:x15="${NS_X15}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="${NS_MAIN}" mc:Ignorable="x"><timeline name="${encodeXml(control.name)}" cache="${encodeXml(control.cacheName)}" caption="${encodeXml(control.name)}" showHeader="1" showSelectionLabel="1" showTimeLevel="1" showHorizontalScrollbar="1" level="2" selectionLevel="2"${payload.period.start ? ` scrollPosition="${encodeXml(payload.period.start)}"` : ''} style="TimelineStyleLight2"/></timelines>`);
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
  return structuredClone(input);
}

function readSourceRows(sheet: SheetSnapshot, range: RangeRef, pivot: PivotDefinition, tableName?: string): { fields: Array<{ name: string; dataType: NativePivotCacheField['dataType'] }>; rows: PivotScalar[][] } {
  const fields: Array<{ name: string; dataType: NativePivotCacheField['dataType'] }> = [];
  for (let column = range.startColumn; column <= range.endColumn; column += 1) {
    const value = sheet.cells[String(range.startRow)]?.[String(column)]?.value;
    const catalog = pivot.fieldCatalog.fields[column - range.startColumn];
    fields.push({ name: catalog?.name ?? (typeof value === 'string' && value ? value : `${tableName ?? 'Field'}${column - range.startColumn + 1}`), dataType: nativeDataType(catalog?.dataType ?? inferDataType(sheet, range, column)) });
  }
  const rows: PivotScalar[][] = [];
  for (let row = range.startRow + 1; row <= range.endRow; row += 1) {
    const values = fields.map((_, offset) => { const value = sheet.cells[String(row)]?.[String(range.startColumn + offset)]?.value; return isScalar(value) ? value : null; });
    if (values.some((value) => value !== null)) rows.push(values);
  }
  return { fields, rows };
}

function buildNativeTable(pivot: PivotDefinition, cache: NativePivotCacheDefinition, part: string, sheetPart: string, old: NativePivotTableDefinition | undefined, source: { fields: Array<{ name: string; dataType: NativePivotCacheField['dataType'] }>; rows: PivotScalar[][] }): NativePivotTableDefinition {
  const fieldIndex = (placement: PivotFieldPlacement | PivotValueField): number => {
    const id = placement.fieldId;
    return pivot.fieldCatalog.fields.find((field) => field.fieldId === id || field.name === id)?.ordinal ?? source.fields.findIndex((field) => field.name === id);
  };
  const rows = pivot.layout.rows.map(fieldIndex).filter((index) => index >= 0);
  const columns = pivot.layout.columns.map(fieldIndex).filter((index) => index >= 0);
  const pages = pivot.layout.filters.map(fieldIndex).filter((index) => index >= 0);
  const dataFields = pivot.layout.values.map((value) => ({ field: fieldIndex(value), ...(value.displayName ? { name: value.displayName } : {}), subtotal: value.summarizeBy, ...(value.showAs && value.showAs.kind !== 'normal' ? { showDataAs: nativeShowAs(value.showAs.kind) } : {}) })).filter((value) => value.field >= 0);
  return {
    name: old?.name ?? (pivot.id.replace(/[^A-Za-z0-9_]/g, '_').slice(0, 255) || 'PivotTable'),
    part, sheetPart, relationshipId: old?.relationshipId ?? '', cacheId: cache.cacheId, pivotId: pivot.id,
    locationRef: old?.locationRef ?? deriveLocation(pivot, source.rows.length, rows.length, columns.length, dataFields.length),
    fields: source.fields.map((_, index) => ({ index, ...(rows.includes(index) ? { axis: 'row' as const } : columns.includes(index) ? { axis: 'column' as const } : pages.includes(index) ? { axis: 'page' as const } : {}), ...(pivot.layout.compact ? { compact: true } : {}) })),
    rowFields: rows, columnFields: columns, pageFields: pages, dataFields,
    showRowGrandTotals: pivot.layout.showGrandTotals, showColumnGrandTotals: pivot.layout.showGrandTotals, showSubtotals: pivot.layout.showSubtotals, repeatLabels: pivot.layout.repeatLabels, compactData: pivot.layout.compact,
    ...(old?.styleName ? { styleName: old.styleName } : {}),
  };
}

function buildCacheDefinitionXml(cache: NativePivotCacheDefinition): string {
  const source = cache.source.kind === 'worksheet-range' ? `<worksheetSource ref="${encodeXml(cache.source.ref)}" sheet="${encodeXml(cache.source.sheetName)}"/>` : cache.source.kind === 'table' ? `<worksheetSource name="${encodeXml(cache.source.tableName)}"${cache.source.sheetName ? ` sheet="${encodeXml(cache.source.sheetName)}"` : ''}/>` : '';
  const fields = cache.fields.map((field) => {
    const values = field.sharedItems ?? [];
    const contains = field.dataType === 'string' ? ' containsString="1"' : field.dataType === 'date' ? ' containsDate="1"' : field.dataType === 'number' ? ' containsNumber="1"' : field.dataType === 'boolean' ? ' containsBoolean="1"' : '';
    const shared = values.map((value) => value === null ? '<m/>' : typeof value === 'string' ? `<s v="${encodeXml(value)}"/>` : typeof value === 'boolean' ? `<b v="${value ? '1' : '0'}"/>` : field.dataType === 'date' ? `<d v="${value}"/>` : `<n v="${value}"/>`).join('');
    return `<cacheField name="${encodeXml(field.name)}"><sharedItems count="${values.length}"${contains}>${shared}</sharedItems></cacheField>`;
  }).join('');
  const attrs = [`cacheSource="worksheet"`, ...boolAttr(cache.refreshOnLoad, 'refreshOnLoad'), ...boolAttr(cache.refreshOnSave, 'refreshOnSave'), ...boolAttr(cache.saveData, 'saveData'), ...boolAttr(cache.enableRefresh, 'enableRefresh')];
  return withXmlDeclaration(`<pivotCacheDefinition xmlns="${NS_MAIN}" xmlns:r="${NS_DOC_REL}" ${attrs.join(' ')}><cacheSource>${source}</cacheSource><cacheFields count="${cache.fields.length}">${fields}</cacheFields></pivotCacheDefinition>`);
}

function buildCacheRecordsXml(cache: NativePivotCacheDefinition, rows: PivotScalar[][]): string {
  const indexes = cache.fields.map((field) => new Map((field.sharedItems ?? []).map((value, index) => [scalarKey(value), index])));
  const records = rows.map((row) => `<r>${cache.fields.map((_, index) => { const value = row[index] ?? null; if (value === null) return '<m/>'; if (typeof value === 'string') return `<s v="${indexes[index]?.get(scalarKey(value)) ?? 0}"/>`; if (typeof value === 'boolean') return `<b v="${value ? '1' : '0'}"/>`; return `<n v="${encodeXml(String(value))}"/>`; }).join('')}</r>`).join('');
  return withXmlDeclaration(`<pivotCacheRecords xmlns="${NS_MAIN}" count="${rows.length}">${records}</pivotCacheRecords>`);
}

function buildPivotTableXml(table: NativePivotTableDefinition): string {
  const fields = table.fields.map((field) => `<pivotField${field.axis === 'row' ? ' axis="axisRow"' : field.axis === 'column' ? ' axis="axisCol"' : field.axis === 'page' ? ' axis="axisPage"' : ''}${field.compact === undefined ? '' : ` compact="${field.compact ? '1' : '0'}"`}/>`).join('');
  const rows = table.rowFields.map((field) => `<field x="${field}"/>`).join('');
  const columns = table.columnFields.map((field) => `<field x="${field}"/>`).join('');
  const pages = table.pageFields.map((field) => `<pageField fld="${field}"/>`).join('');
  const data = table.dataFields.map((field) => `<dataField fld="${field.field}"${field.name ? ` name="${encodeXml(field.name)}"` : ''} subtotal="${encodeXml(nativeAggregate(field.subtotal))}"${field.showDataAs ? ` showDataAs="${encodeXml(field.showDataAs)}"` : ''}/>`).join('');
  const style = table.styleName ? `<pivotTableStyleInfo name="${encodeXml(table.styleName)}" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0"/>` : '';
  return withXmlDeclaration(`<pivotTableDefinition xmlns="${NS_MAIN}" xmlns:r="${NS_DOC_REL}" name="${encodeXml(table.name)}" cacheId="${table.cacheId}" rowGrandTotals="${table.showRowGrandTotals === false ? '0' : '1'}" colGrandTotals="${table.showColumnGrandTotals === false ? '0' : '1'}" compactData="${table.compactData === false ? '0' : '1'}"><location ref="${encodeXml(table.locationRef ?? 'A1')}" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/><pivotFields count="${table.fields.length}">${fields}</pivotFields><rowFields count="${table.rowFields.length}">${rows}</rowFields><colFields count="${table.columnFields.length}">${columns}</colFields>${table.pageFields.length ? `<pageFields count="${table.pageFields.length}">${pages}</pageFields>` : '<pageFields count="0"/>'}<dataFields count="${table.dataFields.length}">${data}</dataFields>${style}</pivotTableDefinition>`);
}

function buildDisplayCells(pivot: PivotDefinition, table: NativePivotTableDefinition, source: { fields: Array<{ name: string; dataType: NativePivotCacheField['dataType'] }>; rows: PivotScalar[][] }): Record<string, Record<string, CellData>> {
  const start = parseA1(table.locationRef?.split(':')[0] ?? 'A1');
  if (!start) return {};
  const output: Record<string, Record<string, CellData>> = {};
  const rowGroups = uniqueTuples(source.rows, table.rowFields);
  const columnGroups = uniqueTuples(source.rows, table.columnFields);
  const rows = rowGroups.length ? rowGroups : [[]];
  const columns = columnGroups.length ? columnGroups : [[]];
  const put = (row: number, column: number, value: PivotScalar): void => { if (value === null) return; output[String(row)] ??= {}; output[String(row)]![String(column)] = { value }; };
  put(start.row, start.column, table.name);
  table.rowFields.forEach((index, position) => put(start.row + 1, start.column + position, source.fields[index]?.name ?? `Field${index + 1}`));
  const headerRow = start.row + 1 + Math.max(1, table.rowFields.length);
  columns.forEach((tuple, index) => put(headerRow, start.column + table.rowFields.length + index, tuple.map((value) => value === null ? '' : String(value)).join(' / ') || 'Values'));
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

function parseCacheFields(node: XmlNode | undefined): NativePivotCacheField[] { return children(node, 'cacheField').map((field, index) => { const shared = child(field, 'sharedItems'); const values: Array<string | number | boolean | null> = []; for (const item of children(shared, 's')) values.push(item.attrs.v ?? null); for (const item of children(shared, 'n')) values.push(numberOrNull(item.attrs.v)); for (const item of children(shared, 'd')) values.push(item.attrs.v ?? null); for (const item of children(shared, 'b')) values.push(item.attrs.v === '1' || item.attrs.v === 'true'); for (const item of children(shared, 'e')) values.push(item.attrs.v ?? null); const dataType = inferCacheType(shared); return { index, name: field.attrs.name ?? `Field${index + 1}`, ...(dataType ? { dataType } : {}), ...(values.length ? { sharedItems: values } : {}) }; }); }

function parsePivotFields(node: XmlNode | undefined): NativePivotTableField[] { return children(node, 'pivotField').map((field, index) => ({ index, ...(field.attrs.axis === 'axisRow' ? { axis: 'row' as const } : field.attrs.axis === 'axisCol' ? { axis: 'column' as const } : field.attrs.axis === 'axisPage' ? { axis: 'page' as const } : {}), ...(field.attrs.compact === '0' ? { compact: false } : field.attrs.compact === '1' ? { compact: true } : {}), ...(field.attrs.outline === '0' ? { outline: false } : field.attrs.outline === '1' ? { outline: true } : {}) })); }
function parseFieldIndexes(node: XmlNode | undefined, label: string): number[] { return children(node, 'field').map((field) => requiredInteger(field.attrs.x, `${label}.field.x`)); }
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
function mapFieldType(value: NativePivotCacheField['dataType']): PivotFieldDataType { return value === 'string' ? 'text' : value === 'number' ? 'number' : value === 'date' ? 'date' : value === 'boolean' ? 'boolean' : 'mixed'; }
function nativeDataType(value: PivotFieldDataType): NativePivotCacheField['dataType'] { return value === 'text' ? 'string' : value === 'number' ? 'number' : value === 'date' ? 'date' : value === 'boolean' ? 'boolean' : 'mixed'; }
function inferDataType(sheet: SheetSnapshot, range: RangeRef, column: number): PivotFieldDataType { const values: unknown[] = []; for (let row = range.startRow + 1; row <= range.endRow; row += 1) values.push(sheet.cells[String(row)]?.[String(column)]?.value); if (values.filter((value) => value !== null).every((value) => typeof value === 'number')) return 'number'; if (values.filter((value) => value !== null).every((value) => typeof value === 'boolean')) return 'boolean'; return 'text'; }
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
function isScalar(value: unknown): value is PivotScalar { return value === null || typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean'; }
function nativeMemberKey(value: PivotScalar): { type: 'text' | 'number' | 'boolean' | 'blank'; value: string | number | boolean | null } { if (value === null || value === '') return { type: 'blank', value: null }; if (typeof value === 'number') return { type: 'number', value }; if (typeof value === 'boolean') return { type: 'boolean', value }; return { type: 'text', value }; }
function scalarKey(value: PivotScalar): string { return `${value === null ? 'blank' : typeof value}:${JSON.stringify(value)}`; }
function uniqueScalars(values: PivotScalar[]): PivotScalar[] { const result: PivotScalar[] = []; const seen = new Set<string>(); for (const value of values) { const key = scalarKey(value); if (!seen.has(key)) { seen.add(key); result.push(value); } } return result; }
function uniqueTuples(rows: PivotScalar[][], indexes: number[]): PivotScalar[][] { if (!indexes.length) return [[]]; const seen = new Set<string>(); const result: PivotScalar[][] = []; for (const row of rows) { const tuple = indexes.map((index) => row[index] ?? null); const key = tuple.map(scalarKey).join('|'); if (!seen.has(key)) { seen.add(key); result.push(tuple); } } return result; }
function matchesTuple(row: PivotScalar[], indexes: number[], tuple: PivotScalar[]): boolean { return indexes.every((index, position) => scalarKey(row[index] ?? null) === scalarKey(tuple[position] ?? null)); }
function aggregate(values: PivotScalar[], operation: string | undefined): PivotScalar { const numbers = values.filter((value): value is number => typeof value === 'number' && Number.isFinite(value)); switch (mapAggregate(operation)) { case 'count': return values.filter((value) => value !== null).length; case 'count-numbers': return numbers.length; case 'average': return numbers.length ? numbers.reduce((sum, value) => sum + value, 0) / numbers.length : null; case 'min': return numbers.length ? Math.min(...numbers) : null; case 'max': return numbers.length ? Math.max(...numbers) : null; case 'product': return numbers.length ? numbers.reduce((product, value) => product * value, 1) : null; case 'distinct-count': return new Set(values.filter((value) => value !== null).map(scalarKey)).size; default: return numbers.length ? numbers.reduce((sum, value) => sum + value, 0) : values.filter((value) => value !== null).length || null; } }
function inferCacheType(node: XmlNode | undefined): NativePivotCacheField['dataType'] | undefined { if (!node) return undefined; if (node.attrs.containsString === '1' || node.attrs.containsString === 'true') return 'string'; if (node.attrs.containsDate === '1' || node.attrs.containsDate === 'true') return 'date'; if (node.attrs.containsNumber === '1' || node.attrs.containsNumber === 'true') return 'number'; if (node.attrs.containsBoolean === '1' || node.attrs.containsBoolean === 'true') return 'boolean'; if (node.attrs.containsError === '1' || node.attrs.containsError === 'true') return 'error'; return undefined; }
function numberOrNull(value: string | undefined): number | null { const number = Number(value); return Number.isFinite(number) ? number : null; }
function optionalBoolean(value: string | undefined, key: string): Record<string, boolean> { return value === undefined ? {} : { [key]: value === '1' || value.toLowerCase() === 'true' }; }
function boolAttr(value: boolean | undefined, name: string): string[] { return value === undefined ? [] : [`${name}="${value ? '1' : '0'}"`]; }
function requiredInteger(value: string | undefined, label: string): number { const number = Number(value); if (!Number.isSafeInteger(number) || number < 0) throw new Error(`${label} must be a non-negative integer`); return number; }
function withXmlDeclaration(xml: string): string { return xml.startsWith('<?xml') ? xml : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${xml}`; }
