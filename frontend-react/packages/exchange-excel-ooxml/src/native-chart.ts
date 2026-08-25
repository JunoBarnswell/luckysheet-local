import { strFromU8, strToU8 } from 'fflate';
import type { ChartDrawingPayload, DrawingObject, SheetSnapshot, WorkbookSnapshot } from '@react-sheets/core-model';
import { descendants, encodeXml, localName, parseXml, serializeXml, type XmlNode } from './xml';
import type { NativePivotGraph, NativePivotTableDefinition, XlsxRelationship } from './types';

const NS_CHART = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const NS_DRAWING = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const NS_DRAWING_MAIN = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_DOC_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL_DRAWING = `${NS_DOC_REL}/drawing`;
const REL_CHART = `${NS_DOC_REL}/chart`;
const GENERATED_CHART_PREFIX = 'xl/charts/react-pivot-chart-';
const GENERATED_DRAWING_NAME_PREFIX = 'react-pivot-chart:';

export interface NativePivotChartSyncInput {
  files: Record<string, Uint8Array>;
  relationships: Record<string, XlsxRelationship[]>;
  snapshot: WorkbookSnapshot;
  graph: NativePivotGraph;
  sheetPartById: Record<string, string>;
  displayCellsBySheetPart: Record<string, Record<string, Record<string, import('@react-sheets/core-model').CellData>>>;
}

export interface NativePivotChartSyncResult {
  files: Record<string, Uint8Array>;
  relationships: Record<string, XlsxRelationship[]>;
}

/**
 * Synchronize only canonical PivotCharts into native DrawingML chart parts.
 * Generic charts and opaque imported chart parts deliberately remain outside
 * this writer; they are preserved by the surrounding OPC overlay.
 */
export function synchronizeNativePivotCharts(input: NativePivotChartSyncInput): NativePivotChartSyncResult {
  const files = cloneFiles(input.files);
  const relationships = cloneRelationships(input.relationships);
  const entries = collectEntries(input.snapshot);
  const activeParts = new Set<string>();
  const entriesBySheet = new Map<string, ChartEntry[]>();

  for (const entry of entries) {
    const part = input.sheetPartById[entry.sheet.id];
    if (!part) throw new Error(`PivotChart ${entry.drawing.id} points to a worksheet that is not in the package`);
    const table = resolveTable(entry.payload.pivotId, input.graph);
    if (!table) throw new Error(`PivotChart ${entry.drawing.id} references PivotTable ${entry.payload.pivotId}, but no native PivotTable was generated`);
    const cache = input.graph.caches.find((candidate) => candidate.cacheId === table.cacheId);
    if (!cache) throw new Error(`PivotChart ${entry.drawing.id} references PivotTable ${table.name}, but its PivotCache ${table.cacheId} is missing`);
    if (cache.source.kind === 'unsupported') throw new Error(`PivotChart ${entry.drawing.id} references an unsupported PivotCache source; export is fail-closed`);
    const pivot = input.snapshot.sheets.flatMap((sheet) => sheet.pivots).find((candidate) => candidate.id === entry.payload.pivotId);
    if (!pivot) throw new Error(`PivotChart ${entry.drawing.id} references missing PivotTable ${entry.payload.pivotId}`);
    assertSupportedChartType(entry.payload.chartType, entry.drawing.id);
    const pivotSheet = input.snapshot.sheets.find((candidate) => input.sheetPartById[candidate.id] === table.sheetPart);
    if (!pivotSheet) throw new Error(`PivotChart ${entry.drawing.id} references PivotTable ${table.name} on a missing worksheet`);
    const chartPart = chartPartFor(entry.sheet.id, entry.drawing.id);
    const chart = buildPivotChartXml(entry.payload, entry.drawing.id, pivotSheet, table, input.displayCellsBySheetPart[table.sheetPart] ?? {});
    files[chartPart] = strToU8(chart);
    activeParts.add(chartPart);
    const list = entriesBySheet.get(part) ?? [];
    list.push({ ...entry, sheetPart: part, chartPart, table });
    entriesBySheet.set(part, list);
  }

  pruneGeneratedChartParts(files, activeParts);
  const generatedDrawingParts = pruneGeneratedDrawingAnchors(files, relationships);

  for (const [sheetPart, sheetEntries] of entriesBySheet) {
    const current = relationships[sheetPart] ?? [];
    let drawingRelation = current.find((relation) => relation.type === REL_DRAWING || relation.type.endsWith('/drawing'));
    let drawingPart = drawingRelation ? resolveTarget(sheetPart, drawingRelation.target) : nextDrawingPart(files);
    if (!drawingRelation) {
      drawingRelation = { id: allocateId(current), type: REL_DRAWING, target: relativeTarget(sheetPart, drawingPart) };
      relationships[sheetPart] = [...current, drawingRelation];
    }
    const drawingRoot = files[drawingPart]
      ? firstElement(parseXml(strFromU8(files[drawingPart]!)), 'wsDr')
      : createDrawingRoot();
    drawingRoot.attrs['xmlns:c'] ??= NS_CHART;
    const drawingRelationships = relationships[drawingPart] ?? [];
    const anchors = sheetEntries.map((entry, index) => {
      const relationshipId = drawingRelationships.find((relation) => resolveTarget(drawingPart, relation.target) === entry.chartPart)?.id ?? allocateId(drawingRelationships);
      const relation = { id: relationshipId, type: REL_CHART, target: relativeTarget(drawingPart, entry.chartPart) };
      if (!drawingRelationships.some((candidate) => candidate.id === relation.id)) drawingRelationships.push(relation);
      return buildChartAnchor(entry, relationshipId, nextDrawingObjectId(drawingRoot, index));
    });
    drawingRoot.children.push(...anchors);
    files[drawingPart] = strToU8(withXmlDeclaration(serializeXml(drawingRoot)));
    relationships[drawingPart] = drawingRelationships;
    files[relationshipPartName(drawingPart)] = strToU8(buildRelationshipsXml(drawingRelationships));
  }

  pruneEmptyGeneratedDrawings(files, relationships, generatedDrawingParts);
  validateGeneratedPackage(files, relationships, activeParts);
  return { files, relationships };
}

interface ChartEntry {
  sheet: SheetSnapshot;
  drawing: DrawingObject;
  payload: ChartDrawingPayload & { pivotId: string };
  sheetPart: string;
  chartPart: string;
  table: NativePivotTableDefinition;
}

function collectEntries(snapshot: WorkbookSnapshot): Array<{ sheet: SheetSnapshot; drawing: DrawingObject; payload: ChartDrawingPayload & { pivotId: string } }> {
  const entries: Array<{ sheet: SheetSnapshot; drawing: DrawingObject; payload: ChartDrawingPayload & { pivotId: string } }> = [];
  for (const sheet of snapshot.sheets) {
    for (const drawing of sheet.drawings ?? []) {
      const payload = sheet.drawingPayloads[drawing.payloadId];
      if (payload?.kind !== 'chart' || !payload.pivotId) continue;
      entries.push({ sheet, drawing, payload: payload as ChartDrawingPayload & { pivotId: string } });
    }
  }
  return entries;
}

function resolveTable(pivotId: string, graph: NativePivotGraph): NativePivotTableDefinition | undefined {
  return graph.tables.find((table) => table.pivotId === pivotId);
}

function assertSupportedChartType(type: ChartDrawingPayload['chartType'], drawingId: string): asserts type is 'column' | 'bar' | 'line' | 'area' {
  if (!['column', 'bar', 'line', 'area'].includes(type)) {
    throw new Error(`PivotChart ${drawingId} uses unsupported native chart type ${type}; export is fail-closed`);
  }
}

function buildPivotChartXml(
  payload: ChartDrawingPayload & { pivotId: string },
  drawingId: string,
  sheet: SheetSnapshot,
  table: NativePivotTableDefinition,
  displayCells: Record<string, Record<string, import('@react-sheets/core-model').CellData>>,
): string {
  const location = parseA1Range(table.locationRef ?? 'A1');
  if (!location) throw new Error(`PivotTable ${table.name} has no valid locationRef for PivotChart ${drawingId}`);
  const used = usedDisplayBounds(displayCells, location);
  if (!used) throw new Error(`PivotTable ${table.name} has no materialized display values for PivotChart ${drawingId}`);
  const dataStartRow = location.startRow + 2 + Math.max(1, table.rowFields.length);
  const dataStartColumn = location.startColumn + table.rowFields.length;
  const categoryColumn = location.startColumn + Math.max(0, table.rowFields.length - 1);
  const dataEndRow = Math.max(dataStartRow, used.endRow);
  const dataEndColumn = Math.max(dataStartColumn, used.endColumn);
  if (dataEndRow < dataStartRow || dataEndColumn < dataStartColumn) throw new Error(`PivotChart ${drawingId} has an empty PivotTable data band`);
  const seriesCount = dataEndColumn - dataStartColumn + 1;
  const categoryValue = displayCells[String(dataStartRow)]?.[String(categoryColumn)]?.value;
  const categoryReference = typeof categoryValue === 'number' ? 'numRef' : 'strRef';
  const chartType = payload.chartType;
  assertSupportedChartType(chartType, drawingId);
  const seriesXml = Array.from({ length: seriesCount }, (_, index) => {
    const column = dataStartColumn + index;
    const declared = payload.series?.[index];
    const header = displayCells[String(location.startRow + 1 + Math.max(1, table.rowFields.length))]?.[String(column)]?.value;
    const name = declared?.name ?? (header === null || header === undefined || header === '' ? table.dataFields[index % Math.max(1, table.dataFields.length)]?.name ?? `Series ${index + 1}` : String(header));
    return `<c:ser><c:idx val="${index}"/><c:order val="${index}"/><c:tx><c:v>${encodeXml(name)}</c:v></c:tx><c:cat><c:${categoryReference}><c:f>${formula(sheet.name, categoryColumn, dataStartRow, dataEndRow)}</c:f></c:${categoryReference}></c:cat><c:val><c:numRef><c:f>${formula(sheet.name, column, dataStartRow, dataEndRow)}</c:f></c:numRef></c:val>${declared?.color ? `<c:spPr><a:solidFill><a:srgbClr val="${encodeXml(declared.color.replace(/^#/, ''))}"/></a:solidFill></c:spPr>` : ''}</c:ser>`;
  }).join('');
  const chartBody = chartBodyFor(chartType, seriesXml);
  const title = payload.elements.title ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>${encodeXml(payload.elements.title)}</a:t></a:r></a:p></c:rich></c:tx><c:layout/><c:overlay val="0"/></c:title>` : '<c:autoTitleDeleted val="1"/>';
  const legend = payload.elements.legend?.visible === false ? '' : `<c:legend><c:legendPos val="${payload.elements.legend?.position === 'left' ? 'l' : payload.elements.legend?.position === 'right' ? 'r' : payload.elements.legend?.position === 'top' ? 't' : 'b'}"/><c:layout/><c:overlay val="0"/></c:legend>`;
  return withXmlDeclaration(`<c:chartSpace xmlns:c="${NS_CHART}" xmlns:a="${NS_DRAWING_MAIN}" xmlns:r="${NS_DOC_REL}"><c:date1904 val="0"/><c:lang val="en-US"/><c:pivotSource><c:name>${encodeXml(table.name)}</c:name><c:fmtId val="0"/></c:pivotSource><c:chart>${title}<c:plotArea><c:layout/>${chartBody}<c:catAx><c:axId val="-201"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="-202"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/></c:catAx><c:valAx><c:axId val="-202"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="-201"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>${legend}</c:plotArea><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart><c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/><c:pageSetup/></c:printSettings></c:chartSpace>`);
}

function chartBodyFor(type: 'column' | 'bar' | 'line' | 'area', seriesXml: string): string {
  if (type === 'line') return `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${seriesXml}<c:axId val="-201"/><c:axId val="-202"/></c:lineChart>`;
  if (type === 'area') return `<c:areaChart><c:grouping val="standard"/><c:varyColors val="0"/>${seriesXml}<c:axId val="-201"/><c:axId val="-202"/></c:areaChart>`;
  return `<c:barChart><c:barDir val="${type === 'bar' ? 'bar' : 'col'}"/><c:grouping val="clustered"/><c:varyColors val="0"/>${seriesXml}<c:axId val="-201"/><c:axId val="-202"/></c:barChart>`;
}

function buildChartAnchor(entry: ChartEntry, relationshipId: string, objectId: number): XmlNode {
  const anchor = entry.drawing.anchor;
  const row = anchor.row ?? 0;
  const column = anchor.column ?? 0;
  const width = Math.max(2, Math.round(entry.drawing.transform.width / 80));
  const height = Math.max(8, Math.round(entry.drawing.transform.height / 20));
  return firstElement(parseXml(`<xdr:twoCellAnchor xmlns:xdr="${NS_DRAWING}" xmlns:a="${NS_DRAWING_MAIN}" xmlns:c="${NS_CHART}" xmlns:r="${NS_DOC_REL}" editAs="oneCell"><xdr:from><xdr:col>${column}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>${column + width}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${row + height}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="${objectId}" name="${encodeXml(`${GENERATED_DRAWING_NAME_PREFIX}${entry.drawing.id}`)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="${NS_CHART}"><c:chart r:id="${encodeXml(relationshipId)}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`), 'twoCellAnchor');
}

function usedDisplayBounds(display: Record<string, Record<string, import('@react-sheets/core-model').CellData>>, location: A1Range): A1Range | undefined {
  const rows = Object.keys(display).map(Number).filter((row) => row >= location.startRow && row <= location.endRow);
  const columns = Object.values(display).flatMap((cells) => Object.keys(cells).map(Number)).filter((column) => column >= location.startColumn && column <= location.endColumn);
  if (!rows.length || !columns.length) return undefined;
  return { ...location, startRow: Math.min(...rows), endRow: Math.max(...rows), startColumn: Math.min(...columns), endColumn: Math.max(...columns) };
}

function formula(sheetName: string, column: number, startRow: number, endRow: number): string {
  return `'${sheetName.replaceAll("'", "''")}'!$${columnToLetter(column)}$${startRow + 1}:$${columnToLetter(column)}$${endRow + 1}`;
}

function chartPartFor(sheetId: string, drawingId: string): string {
  const safe = `${sheetId}-${drawingId}`.replace(/[^A-Za-z0-9_.-]/g, '_').slice(0, 160) || 'chart';
  return `${GENERATED_CHART_PREFIX}${safe}.xml`;
}

function pruneGeneratedChartParts(files: Record<string, Uint8Array>, activeParts: Set<string>): void {
  for (const name of Object.keys(files)) if (name.startsWith(GENERATED_CHART_PREFIX) && !activeParts.has(name)) delete files[name];
}

function pruneGeneratedDrawingAnchors(files: Record<string, Uint8Array>, relationships: Record<string, XlsxRelationship[]>): Set<string> {
  const generatedDrawingParts = new Set<string>();
  for (const [source, list] of Object.entries(relationships)) {
    if (!source.startsWith('xl/drawings/') || source.includes('/_rels/')) continue;
    if (list.some((relation) => isChartRelation(relation) && resolveTarget(source, relation.target).startsWith(GENERATED_CHART_PREFIX))) generatedDrawingParts.add(source);
    relationships[source] = list.filter((relation) => !isChartRelation(relation) || !resolveTarget(source, relation.target).startsWith(GENERATED_CHART_PREFIX));
    const bytes = files[source];
    if (!bytes) continue;
    const root = firstElement(parseXml(strFromU8(bytes)), 'wsDr');
    root.children = root.children.filter((node) => !isGeneratedChartAnchor(node));
    files[source] = strToU8(withXmlDeclaration(serializeXml(root)));
    files[relationshipPartName(source)] = strToU8(buildRelationshipsXml(relationships[source] ?? []));
  }
  return generatedDrawingParts;
}

function isGeneratedChartAnchor(node: XmlNode): boolean {
  const name = descendants(node, 'cNvPr').map((candidate) => candidate.attrs.name).find((candidate) => candidate?.startsWith(GENERATED_DRAWING_NAME_PREFIX));
  return Boolean(name);
}

function pruneEmptyGeneratedDrawings(files: Record<string, Uint8Array>, relationships: Record<string, XlsxRelationship[]>, generatedDrawingParts: Set<string>): void {
  for (const [sheetPart, list] of Object.entries(relationships)) {
    if (!sheetPart.startsWith('xl/worksheets/')) continue;
    const drawing = list.find((relation) => relation.type === REL_DRAWING || relation.type.endsWith('/drawing'));
    if (!drawing) continue;
    const drawingPart = resolveTarget(sheetPart, drawing.target);
    const bytes = files[drawingPart];
    if (!bytes || descendants(parseXml(strFromU8(bytes)), 'twoCellAnchor').length || descendants(parseXml(strFromU8(bytes)), 'oneCellAnchor').length || descendants(parseXml(strFromU8(bytes)), 'absoluteAnchor').length) continue;
    if (!generatedDrawingParts.has(drawingPart)) continue;
    if ((relationships[drawingPart] ?? []).length > 0) continue;
    delete files[drawingPart];
    delete files[relationshipPartName(drawingPart)];
    delete relationships[drawingPart];
    relationships[sheetPart] = list.filter((relation) => relation !== drawing);
  }
}

function validateGeneratedPackage(files: Record<string, Uint8Array>, relationships: Record<string, XlsxRelationship[]>, activeParts: Set<string>): void {
  for (const chartPart of activeParts) {
    if (!files[chartPart]) throw new Error(`Generated PivotChart part is missing: ${chartPart}`);
    const chartXml = strFromU8(files[chartPart]!);
    if (!chartXml.includes('<c:pivotSource>') || !chartXml.includes('<c:plotArea>')) throw new Error(`Generated PivotChart part is structurally incomplete: ${chartPart}`);
  }
  for (const [source, list] of Object.entries(relationships)) for (const relation of list) {
    if (!isChartRelation(relation) && !(source.startsWith('xl/worksheets/') && (relation.type === REL_DRAWING || relation.type.endsWith('/drawing')))) continue;
    const target = resolveTarget(source, relation.target);
    if (!files[target]) throw new Error(`OOXML drawing relationship ${source}!${relation.id} points to missing part ${target}`);
  }
}

function isChartRelation(relation: XlsxRelationship): boolean { return relation.type === REL_CHART || relation.type.endsWith('/chart'); }

function createDrawingRoot(): XmlNode { return firstElement(parseXml(`<xdr:wsDr xmlns:xdr="${NS_DRAWING}" xmlns:a="${NS_DRAWING_MAIN}" xmlns:c="${NS_CHART}" xmlns:r="${NS_DOC_REL}"/>`), 'wsDr'); }

function firstElement(root: XmlNode, name: string): XmlNode { const found = root.name === name || localName(root.name) === name ? root : descendants(root, name)[0]; if (!found) throw new Error(`OOXML part is missing <${name}>`); return found; }
function withXmlDeclaration(xml: string): string { return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${xml}`; }
function cloneFiles(files: Record<string, Uint8Array>): Record<string, Uint8Array> { return Object.fromEntries(Object.entries(files).map(([name, bytes]) => [name, bytes.slice()])); }
function cloneRelationships(input: Record<string, XlsxRelationship[]>): Record<string, XlsxRelationship[]> { return Object.fromEntries(Object.entries(input).map(([source, list]) => [source, list.map((relation) => ({ ...relation }))])); }
function allocateId(list: XlsxRelationship[]): string { const used = new Set(list.map((relation) => relation.id)); let index = 1; while (used.has(`rId${index}`)) index += 1; return `rId${index}`; }
function nextDrawingPart(files: Record<string, Uint8Array>): string { let index = 1; while (files[`xl/drawings/drawing${index}.xml`]) index += 1; return `xl/drawings/drawing${index}.xml`; }
function nextDrawingObjectId(root: XmlNode, offset: number): number { const ids = descendants(root, 'cNvPr').map((node) => Number(node.attrs.id)).filter(Number.isFinite); return Math.max(0, ...ids, offset) + 1; }
function resolveTarget(source: string, target: string): string { if (target.startsWith('/')) return target.slice(1); const base = source.includes('/') ? source.slice(0, source.lastIndexOf('/') + 1) : ''; const result: string[] = []; for (const piece of `${base}${target}`.split('/')) { if (!piece || piece === '.') continue; if (piece === '..') result.pop(); else result.push(piece); } return result.join('/'); }
function relativeTarget(source: string, target: string): string { const left = source.slice(0, source.lastIndexOf('/') + 1).split('/').filter(Boolean); const right = target.split('/').filter(Boolean); while (left.length && right.length && left[0] === right[0]) { left.shift(); right.shift(); } return `${'../'.repeat(left.length)}${right.join('/')}`; }
function relationshipPartName(source: string): string { const slash = source.lastIndexOf('/'); return slash < 0 ? `_rels/${source}.rels` : `${source.slice(0, slash)}/_rels/${source.slice(slash + 1)}.rels`; }
function buildRelationshipsXml(relationships: XlsxRelationship[]): string { return withXmlDeclaration(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships.map((relation) => `<Relationship Id="${encodeXml(relation.id)}" Type="${encodeXml(relation.type)}" Target="${encodeXml(relation.target)}"${relation.targetMode ? ` TargetMode="${encodeXml(relation.targetMode)}"` : ''}/>`).join('')}</Relationships>`); }
function columnToLetter(index: number): string { let value = index + 1; let result = ''; while (value > 0) { const remainder = (value - 1) % 26; result = String.fromCharCode(65 + remainder) + result; value = Math.floor((value - 1) / 26); } return result; }
interface A1Range { startRow: number; endRow: number; startColumn: number; endColumn: number }
function parseA1Range(value: string): A1Range | undefined { const parts = value.split(':'); const start = parseA1(parts[0] ?? ''); const end = parseA1(parts[1] ?? parts[0] ?? ''); return start && end ? { startRow: start.row, endRow: end.row, startColumn: start.column, endColumn: end.column } : undefined; }
function parseA1(value: string): { row: number; column: number } | undefined { const match = /^\$?([A-Z]+)\$?(\d+)$/.exec(value.trim().toUpperCase()); if (!match) return undefined; let column = 0; for (const char of match[1]!) column = column * 26 + char.charCodeAt(0) - 64; const row = Number(match[2]); return Number.isSafeInteger(row) && row > 0 ? { row: row - 1, column: column - 1 } : undefined; }
