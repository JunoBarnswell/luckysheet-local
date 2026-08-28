import { strFromU8, strToU8 } from 'fflate';
import type {
  ChartDrawingPayload,
  ChartAxisModel,
  ChartSeriesModel,
  DrawingObject,
  RangeRef,
  SheetSnapshot,
  WorkbookSnapshot,
} from '@react-sheets/core-model';
import { child, children, descendants, encodeXml, localName, parseXml, serializeXml, textContent, type XmlNode } from './xml';
import type { NativeChartDefinition, NativeChartGraph, NativePivotGraph, NativePivotTableDefinition, XlsxRelationship } from './types';

const NS_CHART = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const NS_DRAWING = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const NS_DRAWING_MAIN = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_DOC_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL_DRAWING = `${NS_DOC_REL}/drawing`;
const REL_CHART = `${NS_DOC_REL}/chart`;
const GENERATED_CHART_PREFIX = 'xl/charts/react-chart-';
const GENERATED_DRAWING_NAME_PREFIX = 'react-chart:';

export interface NativePivotChartSyncInput {
  files: Record<string, Uint8Array>;
  relationships: Record<string, XlsxRelationship[]>;
  snapshot: WorkbookSnapshot;
  graph: NativePivotGraph;
  nativeChartGraph?: NativeChartGraph;
  sheetPartById: Record<string, string>;
  displayCellsBySheetPart: Record<string, Record<string, Record<string, import('@react-sheets/core-model').CellData>>>;
}

export interface NativePivotChartSyncResult {
  files: Record<string, Uint8Array>;
  relationships: Record<string, XlsxRelationship[]>;
  nativeChartGraph: NativeChartGraph;
}

/** All canonical worksheet and Pivot charts share one native writer. */
export function synchronizeNativePivotCharts(input: NativePivotChartSyncInput): NativePivotChartSyncResult {
  return synchronizeNativeCharts(input);
}

export function synchronizeNativeCharts(input: NativePivotChartSyncInput): NativePivotChartSyncResult {
  const files = cloneFiles(input.files);
  const relationships = cloneRelationships(input.relationships);
  const nativeGraph = input.nativeChartGraph ?? readNativeChartGraph({ files, relationships, sheetPartById: input.sheetPartById }) ?? { schema: 'NativeChartGraph' as const, charts: [] };
  const entries = collectEntries(input.snapshot);
  const activeParts = new Set<string>();
  const entriesBySheet = new Map<string, ChartEntry[]>();

  for (const entry of entries) {
    const sheetPart = input.sheetPartById[entry.sheet.id];
    if (!sheetPart) throw new Error(`Chart ${entry.drawing.id} points to a worksheet that is not in the package`);
    const chartPart = chartPartFor(entry.sheet.id, entry.drawing.id);
    const previous = nativeGraph.charts.find((candidate) => candidate.drawingId === entry.drawing.id);
    const originalChart = previous ? files[previous.chartPart] : undefined;
    const chartXml = entry.payload.source.kind === 'pivot'
      ? buildPivotChartXml(entry.payload, entry.drawing.id, entry.sheet, input.graph, input.displayCellsBySheetPart, input.snapshot)
      : buildWorksheetChartXml(entry.payload, entry.drawing.id, entry.sheet, input.snapshot);
    files[chartPart] = strToU8(preserveChartExtensions(chartXml, originalChart ? strFromU8(originalChart) : undefined));
    activeParts.add(chartPart);
    const list = entriesBySheet.get(sheetPart) ?? [];
    list.push({ sheet: entry.sheet, drawing: entry.drawing, payload: entry.payload, sheetPart, chartPart });
    entriesBySheet.set(sheetPart, list);
  }

  removeOwnedChartAnchors(files, relationships, nativeGraph, entries);
  pruneGeneratedChartParts(files, activeParts);
  const generatedDrawingParts = pruneGeneratedDrawingAnchors(files, relationships);
  const emitted: NativeChartDefinition[] = [];

  for (const [sheetPart, sheetEntries] of entriesBySheet) {
    const current = relationships[sheetPart] ?? [];
    let drawingRelation = current.find((relation) => relation.type === REL_DRAWING || relation.type.endsWith('/drawing'));
    const drawingPart = drawingRelation ? resolveTarget(sheetPart, drawingRelation.target) : nextDrawingPart(files);
    if (!drawingRelation) {
      drawingRelation = { id: allocateId(current), type: REL_DRAWING, target: relativeTarget(sheetPart, drawingPart) };
      relationships[sheetPart] = [...current, drawingRelation];
    }
    const drawingRoot = files[drawingPart] ? firstElement(parseXml(strFromU8(files[drawingPart]!)), 'wsDr') : createDrawingRoot();
    drawingRoot.attrs['xmlns:c'] ??= NS_CHART;
    const drawingRelationships = relationships[drawingPart] ?? [];
    const occupiedIds = descendants(drawingRoot, 'cNvPr').map((node) => Number(node.attrs.id)).filter(Number.isFinite);
    let objectId = Math.max(0, ...occupiedIds) + 1;
    for (const entry of sheetEntries) {
      const relationshipId = drawingRelationships.find((relation) => resolveTarget(drawingPart, relation.target) === entry.chartPart)?.id ?? allocateId(drawingRelationships);
      if (!drawingRelationships.some((relation) => relation.id === relationshipId)) drawingRelationships.push({ id: relationshipId, type: REL_CHART, target: relativeTarget(drawingPart, entry.chartPart) });
      drawingRoot.children.push(buildChartAnchor(entry, relationshipId, objectId++));
      emitted.push({ chartPart: entry.chartPart, drawingPart, drawingRelationshipId: relationshipId, drawingId: entry.drawing.id, sheetPart, family: entry.payload.chartType, subtype: entry.payload.subtype, ...(entry.payload.nativeIdentity?.xlChartType === undefined ? {} : { xlChartType: entry.payload.nativeIdentity.xlChartType }), editable: true });
    }
    files[drawingPart] = strToU8(withXmlDeclaration(serializeXml(drawingRoot)));
    relationships[drawingPart] = drawingRelationships;
    files[relationshipPartName(drawingPart)] = strToU8(buildRelationshipsXml(drawingRelationships));
  }

  pruneEmptyGeneratedDrawings(files, relationships, generatedDrawingParts);
  validateGeneratedPackage(files, relationships, activeParts);
  return { files, relationships, nativeChartGraph: { schema: 'NativeChartGraph', charts: emitted } };
}

interface ChartEntry {
  sheet: SheetSnapshot;
  drawing: DrawingObject;
  payload: ChartDrawingPayload;
  sheetPart: string;
  chartPart: string;
}

function collectEntries(snapshot: WorkbookSnapshot): Array<{ sheet: SheetSnapshot; drawing: DrawingObject; payload: ChartDrawingPayload }> {
  const entries: Array<{ sheet: SheetSnapshot; drawing: DrawingObject; payload: ChartDrawingPayload }> = [];
  for (const sheet of snapshot.sheets) for (const drawing of sheet.drawings ?? []) {
    const payload = sheet.drawingPayloads[drawing.payloadId];
    if (payload?.kind === 'chart') entries.push({ sheet, drawing, payload });
  }
  return entries;
}

function buildWorksheetChartXml(payload: ChartDrawingPayload, drawingId: string, sheet: SheetSnapshot, snapshot: WorkbookSnapshot): string {
  if (payload.nativeIdentity?.status === 'preserved-native') throw new Error(`UNSUPPORTED_FEATURE: Preserved-native chart ${drawingId} cannot be rewritten`);
  const sourceRange = sourceRangeFor(payload, snapshot);
  if (!sourceRange) throw new Error(`INVALID_CHART_SOURCE: Chart ${drawingId} has no worksheet-backed range`);
  const sheetNameForId = (sheetId: string): string => snapshot.sheets.find((candidate) => candidate.id === sheetId)?.name ?? sheet.name;
  const declarations = payload.series?.length ? payload.series : deriveSeriesDeclarations(payload, sourceRange);
  const categoryRange = payload.categoryRange ?? (sourceRange.endRow > sourceRange.startRow ? { ...sourceRange, startRow: sourceRange.startRow + 1, startColumn: sourceRange.startColumn, endColumn: sourceRange.startColumn } : sourceRange);
  const seriesXml = declarations.map((series, index) => buildSeriesXml(payload, series, index, categoryRange, sheetNameForId)).join('');
  return buildChartSpace(payload, drawingId, sheet, seriesXml, snapshot, undefined);
}

function sourceRangeFor(payload: ChartDrawingPayload, snapshot: WorkbookSnapshot): RangeRef | undefined {
  if (payload.source.kind === 'worksheet-ranges') return payload.source.ranges[0];
  if (payload.source.kind === 'report-range') return payload.source.range;
  if (payload.source.kind === 'table') {
    const source = payload.source;
    return snapshot.dataModel.tables.find((table) => table.id === source.tableId)?.sourceRange;
  }
  return undefined;
}

function deriveSeriesDeclarations(payload: ChartDrawingPayload, source: RangeRef): ChartSeriesModel[] {
  if (source.endColumn <= source.startColumn) return [{ id: 'series:1', name: 'Series 1', range: source, chartType: payload.chartType === 'combo' ? 'column' : payload.chartType } as ChartSeriesModel];
  return Array.from({ length: source.endColumn - source.startColumn }, (_, index) => ({ id: `series:${index + 1}`, name: `Series ${index + 1}`, range: { ...source, startColumn: source.startColumn + index + 1, endColumn: source.startColumn + index + 1 }, chartType: payload.chartType === 'combo' ? 'column' : payload.chartType } as ChartSeriesModel));
}

function buildSeriesXml(payload: ChartDrawingPayload, series: ChartSeriesModel, index: number, categoryRange: RangeRef, sheetNameForId: (sheetId: string) => string): string {
  const type = series.chartType ?? (payload.chartType === 'combo' ? 'column' : payload.chartType);
  const valueRange = series.yRange ?? series.range;
  const category = rangeRefXml(categoryRange, sheetNameForId);
  const name = series.name || `Series ${index + 1}`;
  const common = `<c:ser><c:idx val="${index}"/><c:order val="${index}"/><c:tx><c:v>${encodeXml(name)}</c:v></c:tx>`;
  const style = series.color ? `<c:spPr><a:solidFill><a:srgbClr val="${encodeXml(series.color.replace(/^#/, ''))}"/></a:solidFill></c:spPr>` : '';
  if (type === 'scatter' || type === 'bubble') {
    if (!series.xRange || !series.yRange || (type === 'bubble' && !series.sizeRange)) throw new Error(`INVALID_CHART_SOURCE: ${type} series ${name} requires X/Y${type === 'bubble' ? '/Size' : ''} range bindings`);
    const size = series.sizeRange ? `<c:bubbleSize>${numRefXml(series.sizeRange, sheetNameForId)}</c:bubbleSize>` : '';
    return `${common}<c:xVal>${numRefXml(series.xRange, sheetNameForId)}</c:xVal><c:yVal>${numRefXml(series.yRange, sheetNameForId)}</c:yVal>${size}${style}${seriesXmlExtras(series, series.subtype ?? payload.subtype)}</c:ser>`;
  }
  if (type === 'stock') {
    const roles = series.stockRoles;
    if (!roles) throw new Error(`INVALID_CHART_SOURCE: Stock series ${name} requires explicit role bindings`);
    return `${common}${stockRoleXml(roles, sheetNameForId)}${style}${seriesXmlExtras(series, series.subtype ?? payload.subtype)}</c:ser>`;
  }
  return `${common}<c:cat>${category}</c:cat><c:val>${numOrStringRefXml(valueRange, sheetNameForId, false)}</c:val>${style}${seriesXmlExtras(series, series.subtype ?? payload.subtype)}</c:ser>`;
}

function stockRoleXml(roles: NonNullable<ChartSeriesModel['stockRoles']>, sheetNameForId: (sheetId: string) => string): string {
  return `${roles.open ? `<c:open>${numRefXml(roles.open, sheetNameForId)}</c:open>` : ''}<c:high>${numRefXml(roles.high, sheetNameForId)}</c:high><c:low>${numRefXml(roles.low, sheetNameForId)}</c:low><c:close>${numRefXml(roles.close, sheetNameForId)}</c:close>${roles.volume ? `<c:vol>${numRefXml(roles.volume, sheetNameForId)}</c:vol>` : ''}`;
}

function seriesXmlExtras(series: ChartSeriesModel, subtype: ChartDrawingPayload['subtype']): string {
  const marker = series.marker?.enabled || subtype.includes('markers') ? `<c:marker><c:symbol val="${series.marker?.shape ?? 'circle'}"/><c:size val="${Math.max(2, Math.min(72, Math.round(series.marker?.size ?? 6)))}"/></c:marker>` : '';
  const smooth = series.smooth || subtype.includes('smooth') ? '<c:smooth val="1"/>' : '';
  const trendlines = (series.trendlines ?? []).map((trendline) => `<c:trendline>${trendline.name ? `<c:name>${encodeXml(trendline.name)}</c:name>` : ''}<c:trendlineType val="${trendline.type === 'moving-average' ? 'movingAvg' : trendline.type === 'logarithmic' ? 'log' : trendline.type === 'polynomial' ? 'poly' : trendline.type === 'power' ? 'power' : trendline.type}"/>${trendline.order === undefined ? '' : `<c:order val="${trendline.order}"/>`}${trendline.period === undefined ? '' : `<c:period val="${trendline.period}"/>`}${trendline.forwardForecast === undefined ? '' : `<c:forward val="${trendline.forwardForecast}"/>`}${trendline.backwardForecast === undefined ? '' : `<c:backward val="${trendline.backwardForecast}"/>`}${trendline.intercept === undefined ? '' : `<c:intercept val="${trendline.intercept}"/>`}${trendline.displayEquation ? '<c:dispEq val="1"/>' : ''}${trendline.displayRSquared ? '<c:dispRSqr val="1"/>' : ''}</c:trendline>`).join('');
  const error = series.errorBars ? `<c:errBars><c:errDir val="${series.errorBars.direction === 'horizontal' ? 'x' : 'y'}"/><c:errBarType val="${series.errorBars.direction === 'both' ? 'both' : 'bothDir'}"/><c:errValType val="${series.errorBars.type === 'standard-deviation' ? 'stdDev' : series.errorBars.type === 'standard-error' ? 'stdErr' : series.errorBars.type === 'fixed' ? 'fixedVal' : series.errorBars.type === 'percentage' ? 'percentage' : 'cust'}"/>${series.errorBars.endStyle === 'no-cap' ? '<c:noEndCap val="1"/>' : ''}${series.errorBars.value === undefined ? '' : `<c:val val="${series.errorBars.value}"/>`}</c:errBars>` : '';
  return `${marker}${smooth}${trendlines}${error}`;
}

function buildPivotChartXml(payload: ChartDrawingPayload, drawingId: string, sheet: SheetSnapshot, graph: NativePivotGraph, displayCellsBySheetPart: NativePivotChartSyncInput['displayCellsBySheetPart'], snapshot: WorkbookSnapshot): string {
  const pivotSource = payload.source;
  if (pivotSource.kind !== 'pivot') throw new Error(`INVALID_CHART_SOURCE: PivotChart ${drawingId} requires a Pivot source`);
  if (!['column', 'bar', 'line', 'area'].includes(payload.chartType)) throw new Error(`PivotChart ${drawingId} uses unsupported native chart type ${payload.chartType}; export is fail-closed`);
  const table = graph.tables.find((candidate) => candidate.pivotId === pivotSource.pivotId);
  if (!table) throw new Error(`PivotChart ${drawingId} references PivotTable ${pivotSource.pivotId}, but no native PivotTable was generated`);
  const cache = graph.caches.find((candidate) => candidate.cacheId === table.cacheId);
  if (!cache) throw new Error(`PivotChart ${drawingId} references PivotTable ${table.name}, but its PivotCache is missing`);
  if (cache.source.kind === 'unsupported') throw new Error(`PivotChart ${drawingId} references an unsupported PivotCache source; export is fail-closed`);
  const location = parseA1Range(table.locationRef ?? 'A1');
  if (!location) throw new Error(`PivotTable ${table.name} has no valid locationRef for PivotChart ${drawingId}`);
  const display = displayCellsBySheetPart[table.sheetPart] ?? {};
  const used = usedDisplayBounds(display, location);
  if (!used) throw new Error(`PivotTable ${table.name} has no materialized display values for PivotChart ${drawingId}`);
  const dataStartRow = location.startRow + 2 + Math.max(1, table.rowFields.length);
  const dataStartColumn = location.startColumn + table.rowFields.length;
  const categoryColumn = location.startColumn + Math.max(0, table.rowFields.length - 1);
  const dataEndRow = Math.max(dataStartRow, used.endRow);
  const dataEndColumn = Math.max(dataStartColumn, used.endColumn);
  const seriesXml = Array.from({ length: dataEndColumn - dataStartColumn + 1 }, (_, index) => {
    const column = dataStartColumn + index;
    const declared = payload.series?.[index];
    const header = display[String(location.startRow + 1 + Math.max(1, table.rowFields.length))]?.[String(column)]?.value;
    const name = declared?.name ?? (header == null || header === '' ? table.dataFields[index % Math.max(1, table.dataFields.length)]?.name ?? `Series ${index + 1}` : String(header));
    const valueRange: RangeRef = { sheetId: snapshot.sheets.find((candidate) => candidate.name === sheet.name)?.id ?? sheet.id, startRow: dataStartRow, endRow: dataEndRow, startColumn: column, endColumn: column };
    const categoryRange: RangeRef = { ...valueRange, startColumn: categoryColumn, endColumn: categoryColumn };
    return buildSeriesXml(payload, { ...(declared ?? { range: valueRange }), name, range: valueRange }, index, categoryRange, (sheetId) => snapshot.sheets.find((candidate) => candidate.id === sheetId)?.name ?? sheet.name);
  }).join('');
  return buildChartSpace(payload, drawingId, sheet, seriesXml, snapshot, table);
}

function buildChartSpace(payload: ChartDrawingPayload, drawingId: string, sheet: SheetSnapshot, seriesXml: string, snapshot: WorkbookSnapshot, pivot?: NativePivotTableDefinition): string {
  const bodies = chartBodies(payload, seriesXml);
  const title = payload.elements.title ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${encodeXml(payload.elements.title)}</a:t></a:r></a:p></c:rich></c:tx><c:layout/><c:overlay val="0"/></c:title>` : '<c:autoTitleDeleted val="1"/>';
  const legend = payload.elements.legend?.visible === false ? '' : `<c:legend><c:legendPos val="${legendPosition(payload.elements.legend?.position)}"/><c:layout/><c:overlay val="${payload.elements.legend?.overlay ? 1 : 0}"/></c:legend>`;
  const axes = ['pie', 'doughnut', 'funnel', 'treemap', 'sunburst', 'histogram', 'box-whisker', 'waterfall', 'stock', 'surface', 'radar', 'map'].includes(payload.chartType) ? '' : axisXml(payload);
  const pivotSource = pivot ? `<c:pivotSource><c:name>${encodeXml(pivot.name)}</c:name><c:fmtId val="0"/></c:pivotSource>` : '';
  const blank = payload.elements.emptyCells === 'zero' ? 'zero' : payload.elements.emptyCells === 'connect' ? 'span' : 'gap';
  const dataLabels = payload.elements.dataLabels?.visible ? `<c:dLbls>${payload.elements.dataLabels.showValue === false ? '' : '<c:showVal val="1"/>'}${payload.elements.dataLabels.showCategoryName ? '<c:showCatName val="1"/>' : ''}${payload.elements.dataLabels.showSeriesName ? '<c:showSerName val="1"/>' : ''}${payload.elements.dataLabels.showPercentage ? '<c:showPercent val="1"/>' : ''}${payload.elements.dataLabels.showLegendKey ? '<c:showLegendKey val="1"/>' : ''}<c:showLeaderLines val="${payload.elements.dataLabels.leaderLines ? 1 : 0}"/></c:dLbls>` : '';
  return withXmlDeclaration(`<c:chartSpace xmlns:c="${NS_CHART}" xmlns:a="${NS_DRAWING_MAIN}" xmlns:r="${NS_DOC_REL}">${pivotSource}<c:chart>${title}<c:plotArea><c:layout/>${bodies}${axes}${legend}${dataLabels}${payload.elements.dataTable?.visible ? '<c:dTable><c:showHorzBorder val="1"/><c:showVertBorder val="1"/><c:showOutline val="1"/><c:showKeys val="1"/></c:dTable>' : ''}</c:plotArea><c:plotVisOnly val="1"/><c:dispBlanksAs val="${blank}"/></c:chart><c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/><c:pageSetup/></c:printSettings>${chartAreaXml(payload)}</c:chartSpace>`);
}

function chartBodies(payload: ChartDrawingPayload, seriesXml: string): string {
  const type = payload.chartType;
  const subtype = payload.subtype;
  if (type === 'combo') {
    const series = payload.series ?? [];
    const grouped = new Map<string, string>();
    const all = [...seriesXml.matchAll(/<c:ser>[\s\S]*?<\/c:ser>/g)].map((match) => match[0]);
    series.forEach((entry, index) => { const chartType = entry.chartType ?? 'column'; grouped.set(chartType, `${grouped.get(chartType) ?? ''}${all[index] ?? ''}`); });
    return [...grouped.entries()].map(([chartType, xml]) => chartBodyFor(chartType as Exclude<ChartDrawingPayload['chartType'], 'combo'>, subtype, xml, payload.waterfallOptions)).join('');
  }
  return chartBodyFor(type, subtype, seriesXml, payload.waterfallOptions);
}

function chartBodyFor(type: Exclude<ChartDrawingPayload['chartType'], 'combo'>, subtype: ChartDrawingPayload['subtype'], seriesXml: string, waterfallOptions?: ChartDrawingPayload['waterfallOptions']): string {
  const grouping = subtype.includes('percent') ? 'percentStacked' : subtype.includes('stacked') ? 'stacked' : 'clustered';
  const axes = '<c:axId val="-201"/><c:axId val="-202"/>';
  if (type === 'column' || type === 'bar') {
    const shape = ['cone', 'cylinder', 'pyramid'].find((value) => subtype.startsWith(value));
    const tag = subtype.startsWith('three-dimensional') || Boolean(shape) ? 'bar3DChart' : 'barChart';
    return `<c:${tag}><c:barDir val="${type === 'bar' ? 'bar' : 'col'}"/><c:grouping val="${grouping}"/><c:varyColors val="0"/>${seriesXml}${shape ? `<c:shape val="${shape}"/>` : ''}${axes}</c:${tag}>`;
  }
  if (type === 'line') { const tag = subtype === 'three-dimensional' ? 'line3DChart' : 'lineChart'; return `<c:${tag}><c:grouping val="${grouping}"/><c:varyColors val="0"/>${seriesXml}${axes}</c:${tag}>`; }
  if (type === 'area') { const tag = subtype === 'three-dimensional' ? 'area3DChart' : 'areaChart'; return `<c:${tag}><c:grouping val="${grouping}"/><c:varyColors val="0"/>${seriesXml}${axes}</c:${tag}>`; }
  if (type === 'pie') return subtype === 'pie-of-pie' || subtype === 'bar-of-pie' ? `<c:ofPieChart><c:ofPieType val="${subtype === 'bar-of-pie' ? 'bar' : 'pie'}"/>${seriesXml}<c:splitType val="auto"/><c:secondPieSize val="75"/></c:ofPieChart>` : `<c:${subtype.includes('three-dimensional') ? 'pie3DChart' : 'pieChart'}><c:varyColors val="1"/>${seriesXml}<c:firstSliceAng val="0"/></c:${subtype.includes('three-dimensional') ? 'pie3DChart' : 'pieChart'}>`;
  if (type === 'doughnut') return `<c:doughnutChart><c:varyColors val="1"/>${seriesXml}<c:firstSliceAng val="0"/><c:holeSize val="55"/></c:doughnutChart>`;
  if (type === 'scatter') { const style = subtype.includes('smooth') ? subtype.includes('markers') ? 'smoothMarker' : 'smooth' : subtype.includes('straight') ? subtype.includes('markers') ? 'lineMarker' : 'line' : 'marker'; return `<c:scatterChart><c:scatterStyle val="${style}"/><c:varyColors val="0"/>${seriesXml}${axes}</c:scatterChart>`; }
  if (type === 'bubble') return `<c:bubbleChart><c:varyColors val="0"/>${seriesXml}<c:bubble3D val="${subtype === 'bubble-three-dimensional' ? '1' : '0'}"/><c:bubbleScale val="100"/><c:showNegBubbles val="0"/><c:sizeRepresents val="area"/>${axes}</c:bubbleChart>`;
  if (type === 'radar') return `<c:radarChart><c:radarStyle val="${subtype === 'radar-filled' ? 'filled' : subtype === 'radar-markers' ? 'marker' : 'standard'}"/><c:varyColors val="0"/>${seriesXml}${axes}</c:radarChart>`;
  if (type === 'stock') return `<c:stockChart>${seriesXml}${axes}</c:stockChart>`;
  if (type === 'surface') return `<c:${subtype.includes('contour') ? 'surfaceChart' : 'surface3DChart'}><c:wireframe val="${subtype.includes('wireframe') ? '1' : '0'}"/>${seriesXml}${axes}</c:${subtype.includes('contour') ? 'surfaceChart' : 'surface3DChart'}>`;
  if (type === 'treemap') return `<c:treemapChart>${seriesXml}</c:treemapChart>`;
  if (type === 'sunburst') return `<c:sunburstChart>${seriesXml}</c:sunburstChart>`;
  if (type === 'histogram' || type === 'pareto') return `<c:histogramChart>${seriesXml}${type === 'pareto' ? '<c:paretoLine val="1"/>' : ''}${axes}</c:histogramChart>`;
  if (type === 'box-whisker') return `<c:boxWhiskerChart>${seriesXml}<c:quartileMethod val="${subtype}"/>${axes}</c:boxWhiskerChart>`;
  if (type === 'waterfall') return `<c:waterfallChart>${seriesXml}<c:connectorLines val="${waterfallOptions?.connectorLines === false ? '0' : '1'}"/><c:reverseOrder val="0"/>${axes}</c:waterfallChart>`;
  if (type === 'funnel') return `<c:funnelChart>${seriesXml}<c:gapWidth val="150"/></c:funnelChart>`;
  if (type === 'map') throw new Error('UNSUPPORTED_FEATURE: Geographic map export requires an authoritative geography provider');
  return `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${seriesXml}${axes}</c:barChart>`;
}

function axisXml(payload: ChartDrawingPayload): string {
  const category = payload.elements.categoryAxis ?? { id: 'category', position: 'bottom' as const, visible: true, axisType: 'category' as const };
  const value = payload.elements.valueAxis ?? { id: 'value', position: 'left' as const, visible: true, axisType: 'value' as const };
  const secondary = payload.elements.secondaryValueAxis;
  const axis = (model: ChartAxisModel | undefined, id: number, cross: number, pos: string, valueAxis: boolean): string => {
    if (!model) return '';
    const min = model.minimum === undefined ? '' : `<c:min val="${model.minimum}"/>`;
    const max = model.maximum === undefined ? '' : `<c:max val="${model.maximum}"/>`;
    const major = model.majorUnit === undefined ? '' : `<c:majorUnit val="${model.majorUnit}"/>`;
    const minor = model.minorUnit === undefined ? '' : `<c:minorUnit val="${model.minorUnit}"/>`;
    return `<c:${valueAxis ? 'valAx' : 'catAx'}><c:axId val="${id}"/><c:scaling><c:orientation val="${model.reverseOrder ? 'maxMin' : 'minMax'}"/>${model.scale === 'logarithmic' ? `<c:logBase val="${model.logBase ?? 10}"/>` : ''}${min}${max}</c:scaling><c:delete val="${model.visible === false ? 1 : 0}"/><c:axPos val="${pos}"/>${major}${minor}<c:majorTickMark val="${model.majorTickMark === 'none' ? 'none' : 'out'}"/><c:minorTickMark val="${model.minorTickMark === 'none' ? 'none' : 'out'}"/><c:tickLblPos val="${model.tickLabelPosition === 'none' ? 'none' : 'nextTo'}"/><c:crossAx val="${cross}"/><c:crosses val="${model.crosses === 'maximum' ? 'max' : 'autoZero'}"/>${model.crossesAt === undefined ? '' : `<c:crossesAt val="${model.crossesAt}"/>`}${valueAxis && model.crossBetween ? `<c:crossBetween val="${model.crossBetween === 'mid-category' ? 'midCat' : 'between'}/> ` : ''}</c:${valueAxis ? 'valAx' : 'catAx'}>`;
  };
  return `${axis(category, -201, -202, category?.position === 'top' ? 't' : 'b', false)}${axis(value, -202, -201, value?.position === 'right' ? 'r' : 'l', true)}${axis(secondary, -203, -201, 'r', true)}`;
}

function chartAreaXml(payload: ChartDrawingPayload): string {
  const fill = payload.elements.chartArea?.fill;
  const color = typeof fill === 'string' ? fill : fill?.color;
  return color ? `<c:spPr><a:solidFill><a:srgbClr val="${encodeXml(color.replace(/^#/, ''))}"/></a:solidFill></c:spPr>` : '';
}

function legendPosition(position: NonNullable<ChartDrawingPayload['elements']['legend']>['position'] | undefined): string { return position === 'left' ? 'l' : position === 'right' ? 'r' : position === 'top' || position === 'top-right' ? 't' : 'b'; }

function numOrStringRefXml(range: RangeRef, sheetNameForId: (sheetId: string) => string, stringRef: boolean): string { return stringRef ? `<c:strRef><c:f>${encodeXml(rangeRefFormula(range, sheetNameForId))}</c:f></c:strRef>` : numRefXml(range, sheetNameForId); }
function numRefXml(range: RangeRef, sheetNameForId: (sheetId: string) => string): string { return `<c:numRef><c:f>${encodeXml(rangeRefFormula(range, sheetNameForId))}</c:f></c:numRef>`; }
function rangeRefXml(range: RangeRef, sheetNameForId: (sheetId: string) => string): string { return `<c:strRef><c:f>${encodeXml(rangeRefFormula(range, sheetNameForId))}</c:f></c:strRef>`; }
function rangeRefFormula(range: RangeRef, sheetNameForId: (sheetId: string) => string): string { return `'${sheetNameForId(range.sheetId).replaceAll("'", "''")}'!$${columnToLetter(range.startColumn)}$${range.startRow + 1}:$${columnToLetter(range.endColumn)}$${range.endRow + 1}`; }

function buildChartAnchor(entry: ChartEntry, relationshipId: string, objectId: number): XmlNode {
  const anchor = entry.drawing.anchor;
  const row = anchor.row ?? 0;
  const column = anchor.column ?? 0;
  const width = Math.max(2, anchor.endColumn === undefined ? Math.round(entry.drawing.transform.width / 80) : anchor.endColumn - column + 1);
  const height = Math.max(8, anchor.endRow === undefined ? Math.round(entry.drawing.transform.height / 20) : anchor.endRow - row + 1);
  return firstElement(parseXml(`<xdr:twoCellAnchor xmlns:xdr="${NS_DRAWING}" xmlns:a="${NS_DRAWING_MAIN}" xmlns:c="${NS_CHART}" xmlns:r="${NS_DOC_REL}" editAs="oneCell"><xdr:from><xdr:col>${column}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>${column + width}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${row + height}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="${objectId}" name="${encodeXml(`${GENERATED_DRAWING_NAME_PREFIX}${entry.drawing.id}`)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="${NS_CHART}"><c:chart r:id="${encodeXml(relationshipId)}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`), 'twoCellAnchor');
}

function removeOwnedChartAnchors(files: Record<string, Uint8Array>, relationships: Record<string, XlsxRelationship[]>, graph: NativeChartGraph, entries: readonly { drawing: DrawingObject; payload: ChartDrawingPayload }[]): void {
  const activeIds = new Set(entries.map((entry) => entry.drawing.id));
  const activeParts = new Set(entries.map((entry) => entry.payload.nativeIdentity?.part).filter((part): part is string => Boolean(part)));
  const ownedParts = new Set(graph.charts.map((entry) => entry.chartPart));
  for (const definition of graph.charts) {
    const bytes = files[definition.drawingPart];
    if (!bytes) continue;
    const root = firstElement(parseXml(strFromU8(bytes)), 'wsDr');
    const list = relationships[definition.drawingPart] ?? [];
    const relationIds = new Set(list.filter((relation) => ownedParts.has(resolveTarget(definition.drawingPart, relation.target))).map((relation) => relation.id));
    root.children = root.children.filter((node) => {
      if (!isAnchor(node)) return true;
      const name = descendants(node, 'cNvPr').map((candidate) => candidate.attrs.name).find(Boolean) ?? '';
      const chartRelation = descendants(node, 'chart')[0]?.attrs['r:id'];
      const ownedByActivePayload = activeIds.has(definition.drawingId) || activeParts.has(definition.chartPart);
      return !(ownedByActivePayload && (name === `${GENERATED_DRAWING_NAME_PREFIX}${definition.drawingId}` || chartRelation && relationIds.has(chartRelation)));
    });
    const ownedByActivePayload = activeIds.has(definition.drawingId) || activeParts.has(definition.chartPart);
    const nextRelations = list.filter((relation) => !(ownedByActivePayload && ownedParts.has(resolveTarget(definition.drawingPart, relation.target))));
    relationships[definition.drawingPart] = nextRelations;
    files[definition.drawingPart] = strToU8(withXmlDeclaration(serializeXml(root)));
    files[relationshipPartName(definition.drawingPart)] = strToU8(buildRelationshipsXml(nextRelations));
  }
}

function pruneGeneratedChartParts(files: Record<string, Uint8Array>, activeParts: Set<string>): void { for (const name of Object.keys(files)) if (name.startsWith(GENERATED_CHART_PREFIX) && !activeParts.has(name)) delete files[name]; }

function pruneGeneratedDrawingAnchors(files: Record<string, Uint8Array>, relationships: Record<string, XlsxRelationship[]>): Set<string> {
  const generatedDrawingParts = new Set<string>();
  for (const [source, list] of Object.entries(relationships)) {
    if (!source.startsWith('xl/drawings/') || source.includes('/_rels/')) continue;
    if (list.some((relation) => isChartRelation(relation) && resolveTarget(source, relation.target).startsWith(GENERATED_CHART_PREFIX))) generatedDrawingParts.add(source);
    const next = list.filter((relation) => !isChartRelation(relation) || !resolveTarget(source, relation.target).startsWith(GENERATED_CHART_PREFIX));
    relationships[source] = next;
    const bytes = files[source];
    if (!bytes) continue;
    const root = firstElement(parseXml(strFromU8(bytes)), 'wsDr');
    root.children = root.children.filter((node) => !isGeneratedChartAnchor(node));
    files[source] = strToU8(withXmlDeclaration(serializeXml(root)));
    files[relationshipPartName(source)] = strToU8(buildRelationshipsXml(next));
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
    if (!generatedDrawingParts.has(drawingPart) || (relationships[drawingPart] ?? []).length > 0) continue;
    delete files[drawingPart]; delete files[relationshipPartName(drawingPart)]; delete relationships[drawingPart];
    relationships[sheetPart] = list.filter((relation) => relation !== drawing);
  }
}

function validateGeneratedPackage(files: Record<string, Uint8Array>, relationships: Record<string, XlsxRelationship[]>, activeParts: Set<string>): void {
  for (const chartPart of activeParts) {
    const xml = files[chartPart] ? strFromU8(files[chartPart]!) : '';
    if (!xml || !xml.includes('<c:chart>') || !xml.includes('<c:plotArea>')) throw new Error(`Generated Chart part is structurally incomplete: ${chartPart}`);
  }
  for (const [source, list] of Object.entries(relationships)) for (const relation of list) {
    if (!isChartRelation(relation) && !(source.startsWith('xl/worksheets/') && (relation.type === REL_DRAWING || relation.type.endsWith('/drawing')))) continue;
    const target = resolveTarget(source, relation.target);
    if (!files[target]) throw new Error(`OOXML drawing relationship ${source}!${relation.id} points to missing part ${target}`);
  }
}

function isChartRelation(relation: XlsxRelationship): boolean { return relation.type === REL_CHART || relation.type.endsWith('/chart'); }
function isAnchor(node: XmlNode): boolean { const name = localName(node.name); return name === 'twoCellAnchor' || name === 'oneCellAnchor' || name === 'absoluteAnchor'; }
function createDrawingRoot(): XmlNode { return firstElement(parseXml(`<xdr:wsDr xmlns:xdr="${NS_DRAWING}" xmlns:a="${NS_DRAWING_MAIN}" xmlns:c="${NS_CHART}" xmlns:r="${NS_DOC_REL}"/>`), 'wsDr'); }
function firstElement(root: XmlNode, name: string): XmlNode { const found = root.name === name || localName(root.name) === name ? root : descendants(root, name)[0]; if (!found) throw new Error(`OOXML part is missing <${name}>`); return found; }
function withXmlDeclaration(xml: string): string { return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${xml}`; }
function cloneFiles(files: Record<string, Uint8Array>): Record<string, Uint8Array> { return Object.fromEntries(Object.entries(files).map(([name, bytes]) => [name, bytes.slice()])); }
function cloneRelationships(input: Record<string, XlsxRelationship[]>): Record<string, XlsxRelationship[]> { return Object.fromEntries(Object.entries(input).map(([source, list]) => [source, list.map((relation) => ({ ...relation }))])); }
function allocateId(list: XlsxRelationship[]): string { const used = new Set(list.map((relation) => relation.id)); let index = 1; while (used.has(`rId${index}`)) index += 1; return `rId${index}`; }
function nextDrawingPart(files: Record<string, Uint8Array>): string { let index = 1; while (files[`xl/drawings/drawing${index}.xml`]) index += 1; return `xl/drawings/drawing${index}.xml`; }
function resolveTarget(source: string, target: string): string { if (target.startsWith('/')) return target.slice(1); const base = source.includes('/') ? source.slice(0, source.lastIndexOf('/') + 1) : ''; const result: string[] = []; for (const piece of `${base}${target}`.split('/')) { if (!piece || piece === '.') continue; if (piece === '..') result.pop(); else result.push(piece); } return result.join('/'); }
function relativeTarget(source: string, target: string): string { const left = source.slice(0, source.lastIndexOf('/') + 1).split('/').filter(Boolean); const right = target.split('/').filter(Boolean); while (left.length && right.length && left[0] === right[0]) { left.shift(); right.shift(); } return `${'../'.repeat(left.length)}${right.join('/')}`; }
function relationshipPartName(source: string): string { const slash = source.lastIndexOf('/'); return slash < 0 ? `_rels/${source}.rels` : `${source.slice(0, slash)}/_rels/${source.slice(slash + 1)}.rels`; }
function buildRelationshipsXml(relationships: XlsxRelationship[]): string { return withXmlDeclaration(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships.map((relation) => `<Relationship Id="${encodeXml(relation.id)}" Type="${encodeXml(relation.type)}" Target="${encodeXml(relation.target)}"${relation.targetMode ? ` TargetMode="${encodeXml(relation.targetMode)}"` : ''}/>`).join('')}</Relationships>`); }
function chartPartFor(sheetId: string, drawingId: string): string { const safe = `${sheetId}-${drawingId}`.replace(/[^A-Za-z0-9_.-]/g, '_').slice(0, 160) || 'chart'; return `${GENERATED_CHART_PREFIX}${safe}.xml`; }
function columnToLetter(index: number): string { let value = index + 1; let result = ''; while (value > 0) { const remainder = (value - 1) % 26; result = String.fromCharCode(65 + remainder) + result; value = Math.floor((value - 1) / 26); } return result; }
function parseA1Range(value: string): A1Range | undefined { const parts = value.split(':'); const start = parseA1(parts[0] ?? ''); const end = parseA1(parts[1] ?? parts[0] ?? ''); return start && end ? { startRow: start.row, endRow: end.row, startColumn: start.column, endColumn: end.column } : undefined; }
function parseA1(value: string): { row: number; column: number } | undefined { const match = /^\$?([A-Z]+)\$?(\d+)$/.exec(value.trim().toUpperCase()); if (!match) return undefined; let column = 0; for (const char of match[1]!) column = column * 26 + char.charCodeAt(0) - 64; const row = Number(match[2]); return Number.isSafeInteger(row) && row > 0 ? { row: row - 1, column: column - 1 } : undefined; }
function usedDisplayBounds(display: Record<string, Record<string, import('@react-sheets/core-model').CellData>>, location: A1Range): A1Range | undefined { const rows = Object.keys(display).map(Number).filter((row) => row >= location.startRow && row <= location.endRow); const columns = Object.values(display).flatMap((cells) => Object.keys(cells).map(Number)).filter((column) => column >= location.startColumn && column <= location.endColumn); return rows.length && columns.length ? { ...location, startRow: Math.min(...rows), endRow: Math.max(...rows), startColumn: Math.min(...columns), endColumn: Math.max(...columns) } : undefined; }

function preserveChartExtensions(output: string, original: string | undefined): string {
  if (!original) return output;
  const sourceRoot = firstElement(parseXml(original), 'chartSpace');
  const outputRoot = firstElement(parseXml(output), 'chartSpace');
  const owned = new Set(['chart', 'printSettings', 'spPr', 'extLst']);
  const outputNames = new Set(outputRoot.children.map((node) => localName(node.name)));
  for (const node of sourceRoot.children) {
    const name = localName(node.name);
    if (owned.has(name) || outputNames.has(name)) continue;
    outputRoot.children.push(structuredClone(node));
    outputNames.add(name);
  }
  const sourceExtensions = sourceRoot.children.find((node) => localName(node.name) === 'extLst');
  if (sourceExtensions && !outputNames.has('extLst')) outputRoot.children.push(structuredClone(sourceExtensions));
  return withXmlDeclaration(serializeXml(outputRoot));
}

interface A1Range { startRow: number; endRow: number; startColumn: number; endColumn: number }

export function readNativeChartGraph(input: { files: Record<string, Uint8Array>; relationships: Record<string, XlsxRelationship[]>; sheetPartById: Record<string, string> }): NativeChartGraph | undefined {
  const charts: NativeChartDefinition[] = [];
  for (const sheetPart of Object.values(input.sheetPartById)) {
    const sheetRelation = (input.relationships[sheetPart] ?? []).find((relation) => relation.type === REL_DRAWING || relation.type.endsWith('/drawing'));
    if (!sheetRelation) continue;
    const drawingPart = resolveTarget(sheetPart, sheetRelation.target);
    const drawingBytes = input.files[drawingPart];
    if (!drawingBytes) continue;
    const drawingRoot = firstElement(parseXml(strFromU8(drawingBytes)), 'wsDr');
    const drawingRelations = input.relationships[drawingPart] ?? [];
    for (const anchor of drawingRoot.children.filter(isAnchor)) {
      const chartNode = descendants(anchor, 'chart')[0];
      const relationId = chartNode?.attrs['r:id'];
      if (!relationId) continue;
      const relation = drawingRelations.find((candidate) => candidate.id === relationId && isChartRelation(candidate));
      if (!relation) continue;
      const chartPart = resolveTarget(drawingPart, relation.target);
      const chartBytes = input.files[chartPart];
      if (!chartBytes) continue;
      const identity = chartIdentity(strFromU8(chartBytes));
      const rawDrawingId = descendants(anchor, 'cNvPr')[0]?.attrs.name ?? `${sheetPart}:${relationId}`;
      const drawingId = rawDrawingId.startsWith(GENERATED_DRAWING_NAME_PREFIX) ? rawDrawingId.slice(GENERATED_DRAWING_NAME_PREFIX.length) : rawDrawingId;
      charts.push({ chartPart, drawingPart, drawingRelationshipId: relationId, drawingId, sheetPart, family: identity.family, subtype: identity.subtype, ...(identity.xlChartType === undefined ? {} : { xlChartType: identity.xlChartType }), editable: identity.editable, ...(identity.reason ? { reason: identity.reason } : {}) });
    }
  }
  return charts.length ? { schema: 'NativeChartGraph', charts } : undefined;
}

function chartIdentity(xml: string): { family: string; subtype: string; xlChartType?: number; editable: boolean; reason?: string } {
  const root = parseXml(xml);
  const plot = descendants(root, 'plotArea')[0];
  const chartNodes = plot?.children.filter((node) => localName(node.name).endsWith('Chart')) ?? [];
  if (!chartNodes.length) return { family: 'unknown', subtype: 'unknown', editable: false, reason: 'UNSUPPORTED_FEATURE: chart has no recognized plot area family' };
  if (chartNodes.length > 1) return { family: 'combo', subtype: 'custom-combo', editable: true };
  const node = chartNodes[0]!;
  const name = localName(node.name);
  const family = name === 'barChart' || name === 'bar3DChart' ? nodeValue(node, 'barDir') === 'bar' ? 'bar' : 'column'
    : name === 'lineChart' ? 'line'
      : name === 'line3DChart' ? 'line'
      : name === 'areaChart' ? 'area'
        : name === 'area3DChart' ? 'area'
        : name === 'pieChart' || name === 'pie3DChart' ? 'pie'
          : name === 'doughnutChart' ? 'doughnut'
            : name === 'scatterChart' ? 'scatter'
              : name === 'bubbleChart' ? 'bubble'
                : name === 'radarChart' ? 'radar'
                  : name === 'stockChart' ? 'stock'
                    : name === 'surface3DChart' || name === 'surfaceChart' ? 'surface'
                      : name === 'treemapChart' ? 'treemap'
                        : name === 'sunburstChart' ? 'sunburst'
                          : name === 'histogramChart' ? 'histogram'
                            : name === 'boxWhiskerChart' ? 'box-whisker'
                              : name === 'waterfallChart' ? 'waterfall'
                                : name === 'funnelChart' ? 'funnel'
                                  : name === 'mapChart' ? 'map'
                                    : name === 'ofPieChart' ? 'pie' : 'unknown';
  const grouping = nodeValue(node, 'grouping');
  const subtype = family === 'column' || family === 'bar' ? nodeValue(node, 'shape') === 'cone' ? grouping === 'percentStacked' ? 'cone-percent-stacked' : grouping === 'stacked' ? 'cone-stacked' : 'cone' : nodeValue(node, 'shape') === 'cylinder' ? grouping === 'percentStacked' ? 'cylinder-percent-stacked' : grouping === 'stacked' ? 'cylinder-stacked' : 'cylinder' : nodeValue(node, 'shape') === 'pyramid' ? grouping === 'percentStacked' ? 'pyramid-percent-stacked' : grouping === 'stacked' ? 'pyramid-stacked' : 'pyramid' : name === 'bar3DChart' ? grouping === 'percentStacked' ? 'three-dimensional-percent-stacked' : grouping === 'stacked' ? 'three-dimensional-stacked' : 'three-dimensional' : grouping === 'percentStacked' ? 'percent-stacked' : grouping === 'stacked' ? 'stacked' : 'clustered'
    : family === 'line' ? grouping === 'percentStacked' ? 'percent-stacked-markers' : grouping === 'stacked' ? descendants(node, 'marker').length ? 'stacked-markers' : 'stacked' : descendants(node, 'marker').length ? 'line-markers' : 'line'
      : family === 'area' ? grouping === 'percentStacked' ? 'percent-stacked' : grouping === 'stacked' ? 'stacked' : name === 'area3DChart' ? 'three-dimensional' : 'area'
        : family === 'scatter' ? nodeValue(node, 'scatterStyle') === 'smoothMarker' ? 'scatter-smooth-lines-markers' : nodeValue(node, 'scatterStyle') === 'smooth' ? 'scatter-smooth-lines' : nodeValue(node, 'scatterStyle') === 'lineMarker' ? 'scatter-straight-lines-markers' : nodeValue(node, 'scatterStyle') === 'line' ? 'scatter-straight-lines' : 'scatter-markers'
          : family === 'bubble' ? nodeValue(node, 'bubble3D') === '1' ? 'bubble-three-dimensional' : 'bubble'
            : family === 'doughnut' ? 'doughnut'
              : family === 'stock' ? descendants(node, 'vol')[0] ? descendants(node, 'open')[0] ? 'stock-volume-open-high-low-close' : 'stock-volume-high-low-close' : descendants(node, 'open')[0] ? 'stock-open-high-low-close' : 'stock-high-low-close'
                : family === 'surface' ? nodeValue(node, 'wireframe') === '1' ? 'surface-wireframe' : 'surface-three-dimensional'
                  : family === 'radar' ? nodeValue(node, 'radarStyle') === 'filled' ? 'radar-filled' : nodeValue(node, 'radarStyle') === 'marker' ? 'radar-markers' : 'radar'
                    : family === 'map' ? 'filled-map' : family;
  const editable = family !== 'unknown' && family !== 'map';
  return editable ? { family, subtype, editable: true } : { family, subtype, editable: false, reason: `UNSUPPORTED_FEATURE: ${family} native chart is preserved without canonical editor ownership` };
}

function nodeValue(node: XmlNode, name: string): string | undefined { return descendants(node, name)[0]?.attrs.val; }

export function projectNativeCharts(snapshot: WorkbookSnapshot, graph: NativeChartGraph | undefined, files: Record<string, Uint8Array>, relationships: Record<string, XlsxRelationship[]>, sheetPartById: Record<string, string>, pivotGraph?: NativePivotGraph): NativeChartGraph | undefined {
  if (!graph) return undefined;
  const sheetByPart = new Map(Object.entries(sheetPartById).map(([sheetId, part]) => [part, snapshot.sheets.find((sheet) => sheet.id === sheetId)] as const));
  for (const definition of graph.charts) {
    if (!definition.editable) continue;
    const sheet = sheetByPart.get(definition.sheetPart);
    const bytes = files[definition.chartPart];
    if (!sheet || !bytes || sheet.drawings.some((drawing) => drawing.id === definition.drawingId || drawing.name === definition.drawingId)) continue;
    const payload = parseNativeChartPayload(strFromU8(bytes), definition, snapshot, sheet, pivotGraph);
    if (!payload) { definition.editable = false; definition.reason = 'UNSUPPORTED_FEATURE: native chart formulas could not be projected into canonical ranges'; continue; }
    const anchor = nativeAnchor(definition, files, relationships);
    const drawingId = definition.drawingId || `native-chart-${definition.chartPart.replace(/[^A-Za-z0-9]/g, '-')}`;
    const payloadId = drawingId;
    const drawing: DrawingObject = { id: drawingId, sheetId: sheet.id, kind: 'chart', payloadId, anchor: { kind: 'two-cell', row: anchor.row, column: anchor.column, endRow: anchor.endRow, endColumn: anchor.endColumn }, transform: { x: 0, y: 0, width: Math.max(160, (anchor.endColumn - anchor.column + 1) * 80), height: Math.max(120, (anchor.endRow - anchor.row + 1) * 20), rotation: 0 }, zIndex: sheet.drawings.length };
    payload.chartId = payloadId;
    payload.nativeIdentity = { family: definition.family, subtype: definition.subtype, ...(definition.xlChartType === undefined ? {} : { xlChartType: definition.xlChartType }), part: definition.chartPart, status: 'owned' };
    sheet.drawings.push(drawing);
    sheet.drawingPayloads[payloadId] = payload;
  }
  return graph;
}

function nativeAnchor(definition: NativeChartDefinition, files: Record<string, Uint8Array>, relationships: Record<string, XlsxRelationship[]>): { row: number; column: number; endRow: number; endColumn: number } {
  const drawing = files[definition.drawingPart];
  if (!drawing) return { row: 0, column: 0, endRow: 10, endColumn: 5 };
  const root = firstElement(parseXml(strFromU8(drawing)), 'wsDr');
  for (const anchor of root.children.filter(isAnchor)) if (descendants(anchor, 'chart')[0]?.attrs['r:id'] === definition.drawingRelationshipId) {
    const from = descendants(anchor, 'from')[0];
    const to = descendants(anchor, 'to')[0];
    return { row: intNode(from, 'row', 0), column: intNode(from, 'col', 0), endRow: Math.max(intNode(from, 'row', 0) + 1, intNode(to, 'row', intNode(from, 'row', 0) + 10)), endColumn: Math.max(intNode(from, 'col', 0) + 1, intNode(to, 'col', intNode(from, 'col', 0) + 5)) };
  }
  void relationships;
  return { row: 0, column: 0, endRow: 10, endColumn: 5 };
}

function intNode(parent: XmlNode | undefined, name: string, fallback: number): number { const value = parent ? textContent(child(parent, name)) : ''; const parsed = Number(value); return Number.isSafeInteger(parsed) && parsed >= 0 ? parsed : fallback; }

function parseNativeChartPayload(xml: string, definition: NativeChartDefinition, snapshot: WorkbookSnapshot, sheet: SheetSnapshot, pivotGraph?: NativePivotGraph): ChartDrawingPayload | undefined {
  const root = parseXml(xml);
  const plot = descendants(root, 'plotArea')[0];
  const chartType = definition.family as ChartDrawingPayload['chartType'];
  if (!['column', 'bar', 'line', 'area', 'pie', 'doughnut', 'scatter', 'bubble', 'radar', 'stock', 'surface', 'treemap', 'sunburst', 'histogram', 'waterfall', 'funnel', 'combo'].includes(chartType)) return undefined;
  const chartNodes = plot?.children.filter((node) => localName(node.name).endsWith('Chart')) ?? [];
  const projectionNodes = chartType === 'combo' ? chartNodes : chartNodes.slice(0, 1);
  if (!projectionNodes.length) return undefined;
  const sheetIdByName = new Map(snapshot.sheets.map((candidate) => [candidate.name, candidate.id]));
  const series: ChartSeriesModel[] = [];
  for (const chartNode of projectionNodes) {
    const seriesType = chartType === 'combo' ? nativeSeriesType(localName(chartNode.name)) : chartType === 'column' || chartType === 'bar' || chartType === 'line' || chartType === 'area' || chartType === 'pie' || chartType === 'doughnut' || chartType === 'scatter' || chartType === 'bubble' || chartType === 'radar' || chartType === 'stock' || chartType === 'surface' || chartType === 'treemap' || chartType === 'sunburst' || chartType === 'histogram' || chartType === 'waterfall' || chartType === 'funnel' ? chartType : undefined;
    if (!seriesType) continue;
    for (const node of descendants(chartNode, 'ser')) {
    const referenceFormula = (parentName: string, referenceName: string): string => {
      const reference = child(child(node, parentName), referenceName);
      return textContent(reference?.children.find((candidate) => localName(candidate.name) === 'f')).trim();
    };
    const yFormula = referenceFormula('val', 'numRef');
    const xFormula = referenceFormula('xVal', 'numRef');
    const catFormula = referenceFormula('cat', 'strRef') || referenceFormula('cat', 'numRef');
    const sizeFormula = referenceFormula('bubbleSize', 'numRef');
    const stockOpen = parseNativeFormula(referenceFormula('open', 'numRef'), sheetIdByName);
    const stockHigh = parseNativeFormula(referenceFormula('high', 'numRef'), sheetIdByName);
    const stockLow = parseNativeFormula(referenceFormula('low', 'numRef'), sheetIdByName);
    const stockClose = parseNativeFormula(referenceFormula('close', 'numRef'), sheetIdByName);
    const stockVolume = parseNativeFormula(referenceFormula('vol', 'numRef'), sheetIdByName);
    const range = parseNativeFormula(yFormula || xFormula || referenceFormula('high', 'numRef') || referenceFormula('close', 'numRef'), sheetIdByName);
    if (!range) continue;
    const xRange = parseNativeFormula(xFormula, sheetIdByName);
    const categoryRange = parseNativeFormula(catFormula, sheetIdByName);
    const sizeRange = parseNativeFormula(sizeFormula, sheetIdByName);
    const textParent = child(node, 'tx') ?? node;
    const name = textContent(descendants(textParent, 't')[0]).trim() || `Series ${series.length + 1}`;
    series.push({ id: `series:${series.length + 1}`, name, range, ...(xRange ? { xRange } : {}), ...(xRange ? { yRange: range } : {}), ...(sizeRange ? { sizeRange } : {}), ...(chartType === 'combo' || seriesType === 'scatter' || seriesType === 'bubble' ? { chartType: seriesType } : {}), ...(seriesType === 'stock' && stockHigh && stockLow && stockClose ? { stockRoles: { ...(stockOpen ? { open: stockOpen } : {}), high: stockHigh, low: stockLow, close: stockClose, ...(stockVolume ? { volume: stockVolume } : {}) } } : {}) });
    if (categoryRange && series.length === 1) (series as Array<ChartSeriesModel & { categoryRange?: RangeRef }>)[0]!.categoryRange = categoryRange;
    }
  }
  if (!series.length) return undefined;
  const title = textContent(descendants(root, 'title')[0] ? descendants(descendants(root, 'title')[0]!, 't')[0] : undefined).trim();
  const categoryRange = (series[0] as ChartSeriesModel & { categoryRange?: RangeRef }).categoryRange;
  const firstRange = series[0]!.range;
  const pivotName = textContent(descendants(root, 'pivotSource')[0] ? descendants(descendants(root, 'pivotSource')[0]!, 'name')[0] : undefined).trim();
  const pivot = pivotGraph?.tables.find((candidate) => candidate.name === pivotName);
  const source: ChartDrawingPayload['source'] = pivot?.pivotId ? { kind: 'pivot', pivotId: pivot.pivotId } : { kind: 'worksheet-ranges', ranges: [firstRange] };
  return { kind: 'chart', chartId: definition.drawingId, chartType, subtype: definition.subtype as ChartDrawingPayload['subtype'], source, series, ...(categoryRange && source.kind === 'worksheet-ranges' ? { categoryRange } : {}), elements: { ...(title ? { title } : {}), legend: { visible: descendants(root, 'legend').length > 0, position: 'bottom' }, hiddenData: 'show' } };
}

function nativeSeriesType(name: string): Exclude<ChartDrawingPayload['chartType'], 'combo'> | undefined {
  if (name === 'barChart' || name === 'bar3DChart') return 'column';
  if (name === 'lineChart' || name === 'line3DChart') return 'line';
  if (name === 'areaChart' || name === 'area3DChart') return 'area';
  return undefined;
}

function parseNativeFormula(value: string, sheetIdByName: ReadonlyMap<string, string>): RangeRef | undefined {
  const match = /^'?((?:[^']|'')+)'?!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/i.exec(value.replace(/^=/, '').trim());
  if (!match) return undefined;
  const sheetName = match[1]!.replaceAll("''", "'");
  const sheetId = sheetIdByName.get(sheetName);
  if (!sheetId) return undefined;
  const start = parseA1(`${match[2]}${match[3]}`);
  const end = parseA1(`${match[4] ?? match[2]}${match[5] ?? match[3]}`);
  return start && end ? { sheetId, startRow: start.row, endRow: end.row, startColumn: start.column, endColumn: end.column } : undefined;
}
