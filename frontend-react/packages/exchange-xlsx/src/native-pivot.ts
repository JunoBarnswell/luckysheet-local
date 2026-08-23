import { strFromU8, strToU8 } from 'fflate';
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
import type {
  NativePivotCacheDefinition,
  NativePivotCacheField,
  NativePivotDataField,
  NativePivotGraph,
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

export interface NativePivotReadInput {
  files: Record<string, Uint8Array>;
  relationships: Record<string, XlsxRelationship[]>;
  sheetPartById: Record<string, string>;
}

export interface NativePivotWriteInput {
  files: Record<string, Uint8Array>;
  relationships: Record<string, XlsxRelationship[]>;
  graph: NativePivotGraph;
  /** Current sheet names keyed by worksheet part. */
  sheetNameByPart?: Record<string, string>;
}

/**
 * Read the complete reachable native Pivot relationship graph.  It follows
 * workbook pivotCaches, cache-definition rels, cache-record rels and each
 * worksheet pivotTableParts relation.  Missing required targets are errors;
 * silently treating a broken Pivot as an empty grid would lose user data.
 */
export function readNativePivotGraph(input: NativePivotReadInput): NativePivotGraph {
  const workbookPart = 'xl/workbook.xml';
  const workbookBytes = input.files[workbookPart];
  if (!workbookBytes) throw new Error('Native Pivot reader requires xl/workbook.xml');
  const workbook = firstElement(parseXml(strFromU8(workbookBytes)), 'workbook');
  const workbookRels = input.relationships[workbookPart] ?? [];
  const sheetNames = readSheetNames(workbook, workbookRels, input.sheetPartById);
  const caches: NativePivotCacheDefinition[] = [];

  for (const cacheNode of children(child(workbook, 'pivotCaches'), 'pivotCache')) {
    const cacheId = parseRequiredInteger(cacheNode.attrs.cacheId, 'pivotCache.cacheId');
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
    const source = parseCacheSource(child(definition, 'cacheSource'), sheetNames);
    caches.push({
      cacheId,
      part,
      ...(recordsPart ? { recordsPart } : {}),
      source,
      fields: parseCacheFields(child(definition, 'cacheFields')),
      ...(recordsRoot?.attrs.count !== undefined ? { recordCount: parseNonNegativeInteger(recordsRoot.attrs.count, 'pivotCacheRecords.count') } : {}),
    });
  }

  const tables: NativePivotTableDefinition[] = [];
  for (const [sheetId, sheetPart] of Object.entries(input.sheetPartById)) {
    void sheetId;
    const sheetBytes = input.files[sheetPart];
    if (!sheetBytes) throw new Error(`Worksheet relation points to missing part: ${sheetPart}`);
    const root = firstElement(parseXml(strFromU8(sheetBytes)), 'worksheet');
    const sheetRels = input.relationships[sheetPart] ?? [];
    for (const tablePartNode of children(child(root, 'pivotTableParts'), 'pivotTablePart')) {
      const relationId = tablePartNode.attrs['r:id'] ?? tablePartNode.attrs.id;
      if (!relationId) throw new Error(`Worksheet ${sheetPart} contains a pivotTablePart without r:id`);
      const relation = requireRelationship(sheetRels, relationId, REL_PIVOT_TABLE, `worksheet ${sheetPart} PivotTable`);
      const part = resolveTarget(sheetPart, relation.target);
      const tableBytes = input.files[part];
      if (!tableBytes) throw new Error(`PivotTable relation points to missing part: ${part}`);
      const definition = firstElement(parseXml(strFromU8(tableBytes)), 'pivotTableDefinition');
      const cacheId = parseRequiredInteger(definition.attrs.cacheId, `PivotTable ${part}.cacheId`);
      tables.push({
        name: definition.attrs.name ?? part,
        part,
        sheetPart,
        relationshipId: relationId,
        cacheId,
        ...(child(definition, 'location')?.attrs.ref ? { locationRef: child(definition, 'location')!.attrs.ref } : {}),
        fields: parsePivotFields(child(definition, 'pivotFields')),
        rowFields: parseFieldIndexes(child(definition, 'rowFields'), 'rowFields'),
        columnFields: parseFieldIndexes(child(definition, 'colFields'), 'colFields'),
        pageFields: parseFieldIndexes(child(definition, 'pageFields'), 'pageFields'),
        dataFields: parseDataFields(child(definition, 'dataFields')),
      });
    }
  }

  return { schema: 'NativePivotGraph', caches, tables };
}

/**
 * Validate and patch native cache source sheet names after an editable sheet
 * rename.  Parts and relationship IDs remain content-addressed by the source
 * package; only the validated worksheetSource attribute is rewritten.
 */
export function synchronizeNativePivotGraph(input: NativePivotWriteInput): Record<string, Uint8Array> {
  const files = cloneFiles(input.files);
  for (const cache of input.graph.caches) {
    if (cache.source.kind !== 'worksheet-range' || !cache.source.sheetPart || !input.sheetNameByPart) continue;
    const bytes = files[cache.part];
    if (!bytes) throw new Error(`Native Pivot cache part is missing: ${cache.part}`);
    const root = firstElement(parseXml(strFromU8(bytes)), 'pivotCacheDefinition');
    const source = child(child(root, 'cacheSource'), 'worksheetSource');
    if (!source) throw new Error(`Pivot cache ${cache.part} has no worksheetSource`);
    const sheetName = input.sheetNameByPart[cache.source.sheetPart];
    if (!sheetName) throw new Error(`Pivot cache source worksheet part is missing from current workbook: ${cache.source.sheetPart}`);
    source.attrs.sheet = sheetName;
    files[cache.part] = strToU8(withXmlDeclaration(serializeXml(root)));
  }
  // Validate every relationship target in the constrained graph before the
  // package writer emits it.  Unknown relationships remain opaque bytes.
  for (const cache of input.graph.caches) {
    if (!files[cache.part]) throw new Error(`Native Pivot cache part is missing: ${cache.part}`);
    if (cache.recordsPart && !files[cache.recordsPart]) throw new Error(`Native Pivot records part is missing: ${cache.recordsPart}`);
  }
  for (const table of input.graph.tables) {
    if (!files[table.part]) throw new Error(`Native PivotTable part is missing: ${table.part}`);
    if (!input.relationships[table.sheetPart]?.some((relation) => relation.id === table.relationshipId && relation.type === REL_PIVOT_TABLE)) {
      throw new Error(`Native PivotTable relationship ${table.relationshipId} is missing from ${table.sheetPart}`);
    }
  }
  return files;
}

/** Serialize a canonical workbook pivotCaches node from validated metadata. */
export function serializeNativePivotCaches(graph: NativePivotGraph, relationships: XlsxRelationship[]): string {
  const items = graph.caches.map((cache) => {
    const relation = relationships.find((candidate) => (candidate.type === REL_PIVOT_CACHE_DEFINITION || candidate.type.endsWith('/pivotCacheDefinition')) && resolveTarget('xl/workbook.xml', candidate.target) === cache.part);
    if (!relation) throw new Error(`Native Pivot cache relation is missing for ${cache.part}`);
    return `<pivotCache cacheId="${cache.cacheId}" r:id="${encodeXml(relation.id)}"/>`;
  }).join('');
  return `<pivotCaches count="${graph.caches.length}" xmlns:r="${NS_DOC_REL}">${items}</pivotCaches>`;
}

function parseCacheSource(node: XmlNode | undefined, sheetNames: Map<string, { name: string; part: string }>): NativePivotSource | { kind: 'unsupported'; reason: string } {
  const worksheetSource = child(node, 'worksheetSource');
  if (worksheetSource) {
    const sheetName = worksheetSource.attrs.sheet;
    const ref = worksheetSource.attrs.ref;
    if (sheetName && ref) {
      const sheet = [...sheetNames.values()].find((candidate) => candidate.name === sheetName);
      return { kind: 'worksheet-range', sheetName, ref, ...(sheet ? { sheetPart: sheet.part } : {}) };
    }
    const tableName = worksheetSource.attrs.name;
    if (tableName) return { kind: 'table', tableName, ...(sheetName ? { sheetName } : {}) };
    return { kind: 'unsupported', reason: 'worksheetSource is missing both sheet/ref and table name' };
  }
  if (child(node, 'consolidation')) return { kind: 'unsupported', reason: 'consolidation sources are not editable' };
  if (child(node, 'external')) return { kind: 'unsupported', reason: 'external Pivot sources are not editable' };
  return { kind: 'unsupported', reason: 'Pivot cache has no supported cacheSource' };
}

function parseCacheFields(node: XmlNode | undefined): NativePivotCacheField[] {
  return children(node, 'cacheField').map((field, index) => {
    const shared = child(field, 'sharedItems');
    const values: Array<string | number | boolean | null> = [];
    if (shared) {
      for (const item of shared.children) {
        switch (localName(item.name)) {
          case 's': values.push(textOrNull(item.attrs.v)); break;
          case 'n': values.push(numericValue(item.attrs.v)); break;
          case 'd': values.push(textOrNull(item.attrs.v)); break;
          case 'b': values.push(item.attrs.v === '1' || item.attrs.v === 'true'); break;
          case 'e': values.push(textOrNull(item.attrs.v)); break;
          default: break;
        }
      }
    }
    const dataType = inferCacheDataType(shared);
    return {
      index,
      name: field.attrs.name ?? `Field${index + 1}`,
      ...(dataType ? { dataType } : {}),
      ...(values.length ? { sharedItems: values } : {}),
    };
  });
}

function parsePivotFields(node: XmlNode | undefined): NativePivotTableField[] {
  return children(node, 'pivotField').map((field, index) => ({
    index,
    ...(field.attrs.axis === 'axisRow' ? { axis: 'row' as const } : field.attrs.axis === 'axisCol' ? { axis: 'column' as const } : field.attrs.axis === 'axisPage' ? { axis: 'page' as const } : {}),
    ...(field.attrs.compact === '0' ? { compact: false } : field.attrs.compact === '1' ? { compact: true } : {}),
    ...(field.attrs.outline === '0' ? { outline: false } : field.attrs.outline === '1' ? { outline: true } : {}),
  }));
}

function parseFieldIndexes(node: XmlNode | undefined, label: string): number[] {
  return children(node, 'field').map((field) => parseRequiredInteger(field.attrs.x, `${label}.field.x`));
}

function parseDataFields(node: XmlNode | undefined): NativePivotDataField[] {
  return children(node, 'dataField').map((field) => ({
    field: parseRequiredInteger(field.attrs.fld, 'dataField.fld'),
    ...(field.attrs.name ? { name: field.attrs.name } : {}),
    ...(field.attrs.subtotal ? { subtotal: field.attrs.subtotal } : {}),
    ...(field.attrs.showDataAs ? { showDataAs: field.attrs.showDataAs } : {}),
  }));
}

function readSheetNames(workbook: XmlNode, rels: XlsxRelationship[], parts: Record<string, string>): Map<string, { name: string; part: string }> {
  const result = new Map<string, { name: string; part: string }>();
  for (const [sheetId, part] of Object.entries(parts)) {
    const node = children(child(workbook, 'sheets'), 'sheet').find((candidate) => `sheet-${candidate.attrs.sheetId}` === sheetId);
    if (node?.attrs.name) result.set(sheetId, { name: node.attrs.name, part });
  }
  // A malformed sheetPartById map should not prevent valid relation-based
  // source resolution when a caller only supplied workbook relationships.
  for (const node of children(child(workbook, 'sheets'), 'sheet')) {
    const relationId = node.attrs['r:id'] ?? node.attrs.id;
    const relation = rels.find((candidate) => candidate.id === relationId);
    if (!relation || !node.attrs.name) continue;
    const part = resolveTarget('xl/workbook.xml', relation.target);
    const key = `sheet-${node.attrs.sheetId ?? result.size + 1}`;
    if (!result.has(key)) result.set(key, { name: node.attrs.name, part });
  }
  return result;
}

function requireRelationship(rels: XlsxRelationship[], id: string, type: string, context: string): XlsxRelationship {
  const relation = rels.find((candidate) => candidate.id === id);
  if (!relation) throw new Error(`${context} relation ${id} is missing`);
  if (relation.type !== type && !relation.type.endsWith(`/${type.split('/').pop()!}`)) {
    throw new Error(`${context} relation ${id} has unexpected type ${relation.type}`);
  }
  return relation;
}

function firstElement(root: XmlNode, name: string): XmlNode {
  const found = localName(root.name) === name ? root : descendants(root, name)[0];
  if (!found) throw new Error(`OOXML part is missing <${name}>`);
  return found;
}

function resolveTarget(source: string, target: string): string {
  if (target.startsWith('/')) return normalizePartName(target.slice(1));
  const base = source.includes('/') ? source.slice(0, source.lastIndexOf('/') + 1) : '';
  return normalizePartName(`${base}${target}`);
}

function normalizePartName(name: string): string {
  const pieces: string[] = [];
  for (const piece of name.replaceAll('\\', '/').split('/')) {
    if (!piece || piece === '.') continue;
    if (piece === '..') {
      if (!pieces.length) throw new Error(`Unsafe XLSX part target: ${name}`);
      pieces.pop();
    } else pieces.push(piece);
  }
  const normalized = pieces.join('/');
  if (!normalized || normalized.includes('\0') || normalized.startsWith('/') || /^[A-Za-z]:/.test(normalized)) throw new Error(`Unsafe XLSX part target: ${name}`);
  return normalized;
}

function cloneFiles(files: Record<string, Uint8Array>): Record<string, Uint8Array> {
  return Object.fromEntries(Object.entries(files).map(([name, data]) => [name, data.slice()]));
}

function parseRequiredInteger(value: string | undefined, label: string): number {
  const number = Number(value);
  if (!Number.isSafeInteger(number) || number < 0) throw new Error(`${label} must be a non-negative integer`);
  return number;
}

function parseNonNegativeInteger(value: string | undefined, label: string): number {
  return parseRequiredInteger(value, label);
}

function textOrNull(value: string | undefined): string | null {
  return value === undefined ? null : value;
}

function numericValue(value: string | undefined): number | null {
  if (value === undefined) return null;
  const number = Number(value);
  return Number.isFinite(number) ? number : null;
}

function inferCacheDataType(node: XmlNode | undefined): NativePivotCacheField['dataType'] | undefined {
  if (!node) return undefined;
  if (node.attrs.containsString === '1' || node.attrs.containsString === 'true') return 'string';
  if (node.attrs.containsNumber === '1' || node.attrs.containsNumber === 'true') return 'number';
  if (node.attrs.containsDate === '1' || node.attrs.containsDate === 'true') return 'date';
  if (node.attrs.containsBoolean === '1' || node.attrs.containsBoolean === 'true') return 'boolean';
  if (node.attrs.containsError === '1' || node.attrs.containsError === 'true') return 'error';
  return undefined;
}

function withXmlDeclaration(xml: string): string {
  return xml.startsWith('<?xml') ? xml : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${xml}`;
}
