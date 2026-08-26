import {
  createPivotMemberKey,
  formatPivotMember,
  pivotMemberKey,
  pivotScalarFromMemberKey,
  buildPivotTimelineTiles,
} from "@react-sheets/core-model";
import type {
  ChartDrawingPayload,
  ChartMarkerModel,
  DataChartDrawingPayload,
  CameraDrawingPayload,
  FormControlDrawingPayload,
  ConnectorDrawingPayload,
  DrawingArrowhead,
  DrawingObject,
  DrawingPayload,
  PivotMemberKey,
  PivotScalar,
  PivotResultTree,
  PivotSlicerItemProjection,
  PivotSlicerDrawingPayload,
  PivotTimelineDrawingPayload,
  PivotTimelineLevel,
  RangeRef,
  SparklineModel,
  WorkbookTableModel,
  AssetRef,
} from "@react-sheets/core-model";
import { isDrawingConnectorPayload } from "@react-sheets/core-model";
import type { CanvasSheetSnapshot } from "@react-sheets/spreadsheet-app";
import { buildPivotChartData } from "@react-sheets/spreadsheet-app";
import {
  DEFAULT_RENDER_THEME,
  SheetSkeleton,
  drawCellLayer,
  drawGridLayer,
  type CellProvider,
  type CellRenderData,
  type FloatingDrawable,
  type Rect,
  type RenderPane,
} from "@react-sheets/render-engine";

/** Semantic child actions emitted by a Pivot control drawable. */
export type PivotControlAction =
  | { kind: 'slicer-member'; memberKey: PivotMemberKey }
  | { kind: 'slicer-clear' }
  | { kind: 'timeline-period'; start: string; end: string }
  | { kind: 'timeline-handle'; edge: 'start' | 'end' }
  | { kind: 'timeline-scroll'; direction: -1 | 1 }
  | { kind: 'timeline-level'; level: PivotTimelineLevel };

const CHART_PALETTE = ["#2563eb", "#10b981", "#f59e0b", "#ef4444", "#8b5cf6", "#ec4899", "#06b6d4"];

interface CanvasChartSeries {
  name: string;
  values: number[];
  missing?: boolean[];
  color?: string;
  marker?: ChartMarkerModel;
  chartType?: Exclude<ChartDrawingPayload['chartType'], 'combo'>;
}

function numericCellValue(value: string): number | undefined {
  if (!value.trim()) return undefined;
  const numeric = Number(value.replace(/[$,%]/g, ""));
  return Number.isFinite(numeric) ? numeric : undefined;
}

function getChartSeries(
  payload: ChartDrawingPayload,
  getSheet: (sheetId: string) => CanvasSheetSnapshot | undefined,
  pivotResults: Record<string, PivotResultTree>,
  sheets: readonly CanvasSheetSnapshot[],
): { categories: string[]; series: CanvasChartSeries[]; brokenReference?: string } {
  const categories: string[] = [];
  const series: CanvasChartSeries[] = [];
  const pivot = payload.pivotId
    ? pivotResults[payload.pivotId] ?? sheets.map((candidate) => candidate.pivotResults[payload.pivotId!]).find(Boolean)
    : undefined;
  if (payload.pivotId && !pivot) return { categories, series, brokenReference: `Pivot reference unavailable: ${payload.pivotId}` };
  if (pivot) {
    const pivotDefinition = sheets.flatMap((candidate) => candidate.pivots).find((candidate) => candidate.id === pivot.pivotId);
    const projected = buildPivotChartData(pivot, pivotDefinition);
    categories.push(...projected.categories.map((category) => category.label));
    for (const [index, entry] of projected.series.entries()) {
      const declared = payload.series?.[index];
      const missing = entry.values.map((value) => typeof value !== 'number' || !Number.isFinite(value));
      series.push({
        name: declared?.name ?? entry.name,
        values: entry.values.map((value) => typeof value === 'number' && Number.isFinite(value) ? value : 0),
        missing,
        color: declared?.color,
        marker: declared?.marker,
        chartType: declared?.chartType,
      });
    }
    return { categories, series };
  }

  const readRange = (range: RangeRef): string[][] => {
    const sourceSheet = getSheet(range.sheetId);
    if (!sourceSheet) return [];
    const rows: string[][] = [];
    for (let row = range.startRow; row <= range.endRow; row += 1) {
      if (payload.elements.hiddenData === 'hideRows' && sourceSheet.hiddenRows.includes(row)) continue;
      const values: string[] = [];
      for (let column = range.startColumn; column <= range.endColumn; column += 1) {
        if (payload.elements.hiddenData === 'hideColumns' && sourceSheet.hiddenColumns.includes(column)) continue;
        values.push(sourceSheet.getCell(row, column)?.value ?? "");
      }
      rows.push(values);
    }
    return rows;
  };

  if (payload.series && payload.series.length > 0) {
    for (const entry of payload.series) {
      const values = readRange(entry.range).flat().map(numericCellValue).filter((value): value is number => value !== undefined);
      series.push({ name: entry.name, values, color: entry.color, marker: entry.marker, chartType: entry.chartType });
    }
  }
  const source = payload.sourceRanges[0];
  if (!source) return { categories, series };
  const matrix = readRange(source);
  const categoryMatrix = payload.categoryRange ? readRange(payload.categoryRange) : [];
  if (categoryMatrix.length > 0) {
    categories.push(...categoryMatrix.flat().filter((value) => value !== ""));
  }

  if (series.length === 0 && matrix.length > 1 && (matrix[0]?.length ?? 0) > 1) {
    const width = matrix[0]?.length ?? 0;
    if (categories.length === 0) {
      for (let row = 1; row < matrix.length; row += 1) categories.push(matrix[row]?.[0] ?? String(row));
    }
    for (let column = 1; column < width; column += 1) {
      const values = matrix.slice(1).map((row) => numericCellValue(row?.[column] ?? "") ?? 0);
      series.push({ name: matrix[0]?.[column] || payload.elements.title || `Series ${column}`, values });
    }
    return { categories, series };
  }

  if (series.length === 0) {
    const values = matrix.flat().map(numericCellValue).filter((value): value is number => value !== undefined);
    if (categories.length === 0) {
      categories.push(...matrix.flat().filter((value) => numericCellValue(value) === undefined && value !== ""));
    }
    if (values.length > 0) series.push({ name: payload.elements.title || "Series 1", values });
  }
  const targetLength = Math.max(categories.length, ...series.map((entry) => entry.values.length), 1);
  if (categories.length === 0) {
    for (let index = 0; index < targetLength; index += 1) categories.push(String(index + 1));
  }
  return { categories, series };
}

function drawCanonicalShapeOnCanvas(options: {
  context: CanvasRenderingContext2D;
  payload: Extract<DrawingPayload, { kind: "shape" | "textbox" }>;
  bounds: Rect;
  rotation?: number;
}): void {
  const { context, payload, bounds, rotation } = options;
  const { x, y, width, height } = bounds;
  context.save();
  context.translate(x + width / 2, y + height / 2);
  if (rotation) context.rotate((rotation * Math.PI) / 180);
  context.translate(-width / 2, -height / 2);
  if (payload.kind === "textbox") {
    const frame = payload.textFrame;
    context.fillStyle = frame.textColor;
    context.font = `${frame.italic ? "italic " : ""}${frame.bold ? "bold " : ""}${frame.fontSize}px ${frame.fontFamily}, sans-serif`;
    context.textAlign = frame.horizontalAlignment === "center" ? "center" : frame.horizontalAlignment === "right" ? "right" : "left";
    context.textBaseline = frame.verticalAlignment === "middle" ? "middle" : frame.verticalAlignment === "bottom" ? "bottom" : "top";
    const margin = frame.margin;
    const textX = frame.horizontalAlignment === "center" ? width / 2 : frame.horizontalAlignment === "right" ? width - margin.right : margin.left;
    const textY = frame.verticalAlignment === "middle" ? height / 2 : frame.verticalAlignment === "bottom" ? height - margin.bottom : margin.top;
    const rawLines = frame.wrap ? payload.text.split(/\r?\n/) : [payload.text.replace(/[\r\n]+/g, " ")];
    const maxWidth = Math.max(10, width - margin.left - margin.right);
    const lines: string[] = [];
    for (const rawLine of rawLines) {
      if (!frame.wrap || rawLine.length === 0) { lines.push(rawLine); continue; }
      let line = "";
      for (const word of rawLine.split(/\s+/)) {
        const candidate = line ? `${line} ${word}` : word;
        if (line && context.measureText(candidate).width > maxWidth) { lines.push(line); line = word; }
        else line = candidate;
      }
      lines.push(line);
    }
    const lineHeight = frame.fontSize * 1.2;
    lines.forEach((line, index) => context.fillText(line, textX, textY + (frame.verticalAlignment === "bottom" ? -((lines.length - 1 - index) * lineHeight) : index * lineHeight), maxWidth));
    context.restore();
    return;
  }
  if (payload.effects?.shadow) {
    context.shadowColor = payload.effects.shadow.color;
    context.shadowBlur = payload.effects.shadow.blur;
    context.shadowOffsetX = payload.effects.shadow.offsetX;
    context.shadowOffsetY = payload.effects.shadow.offsetY;
    context.globalAlpha = payload.effects.shadow.opacity;
  }
  if (payload.effects?.glow) {
    context.shadowColor = payload.effects.glow.color;
    context.shadowBlur = payload.effects.glow.radius;
    context.globalAlpha = payload.effects.glow.opacity;
  }
  context.fillStyle = payload.fill;
  context.strokeStyle = payload.stroke;
  context.lineWidth = payload.strokeWidth ?? 1.5;
  if (payload.type === "rectangle") {
    context.fillRect(0, 0, width, height);
    context.strokeRect(0, 0, width, height);
  } else if (payload.type === "rounded-rectangle") {
    const radius = Math.min(8, width / 4, height / 4);
    context.beginPath();
    context.roundRect(0, 0, width, height, radius);
    context.fill();
    context.stroke();
  } else if (payload.type === "ellipse") {
    context.beginPath();
    context.ellipse(width / 2, height / 2, width / 2, height / 2, 0, 0, Math.PI * 2);
    context.fill();
    context.stroke();
  } else if (payload.type === "line" || payload.type === "arrow") {
    context.beginPath();
    context.moveTo(0, height / 2);
    context.lineTo(width - (payload.type === "arrow" ? Math.min(16, width / 3) : 0), height / 2);
    context.stroke();
    if (payload.type === "arrow") {
      const head = Math.min(16, width / 3);
      context.fillStyle = payload.stroke;
      context.beginPath();
      context.moveTo(width, height / 2);
      context.lineTo(width - head, height / 2 - head / 2);
      context.lineTo(width - head, height / 2 + head / 2);
      context.closePath();
      context.fill();
    }
  } else if (payload.type === "star") {
    let angle = -Math.PI / 2;
    const outer = Math.min(width, height) / 2;
    const inner = outer / 2;
    context.beginPath();
    for (let index = 0; index < 10; index += 1) {
      const radius = index % 2 === 0 ? outer : inner;
      const pointX = width / 2 + Math.cos(angle) * radius;
      const pointY = height / 2 + Math.sin(angle) * radius;
      if (index === 0) context.moveTo(pointX, pointY);
      else context.lineTo(pointX, pointY);
      angle += Math.PI / 5;
    }
    context.closePath();
    context.fill();
    context.stroke();
  } else {
    const radius = 6;
    const bodyHeight = height - 12;
    context.beginPath();
    context.roundRect(0, 0, width, bodyHeight, radius);
    context.moveTo(16, bodyHeight);
    context.lineTo(12, height);
    context.lineTo(28, bodyHeight);
    context.closePath();
    context.fill();
    context.stroke();
  }
  if (payload.text) {
    context.globalAlpha = 1;
    context.fillStyle = payload.textColor ?? "#1e293b";
    context.font = `${payload.fontSize ?? 13}px Inter, sans-serif`;
    context.textAlign = payload.textAlignment ?? "center";
    context.textBaseline = payload.textVerticalAlignment === 'top' ? 'top' : payload.textVerticalAlignment === 'bottom' ? 'bottom' : 'middle';
    const textX = payload.textAlignment === 'left' ? 6 : payload.textAlignment === 'right' ? width - 6 : width / 2;
    const textY = payload.textVerticalAlignment === 'top' ? 6 : payload.textVerticalAlignment === 'bottom' ? height - 6 : height / 2;
    if (payload.textDirection === 'vertical') {
      const chars = [...payload.text];
      const lineHeight = Math.max(10, (payload.fontSize ?? 13) * 1.15);
      const startY = height / 2 - ((chars.length - 1) * lineHeight) / 2;
      chars.forEach((char, index) => context.fillText(char, textX, startY + index * lineHeight, Math.max(10, width - 8)));
    } else {
      context.fillText(payload.text, textX, textY, Math.max(10, width - 8));
    }
  }
  context.restore();
}

/**
 * Render a connector from the route owned by the canonical drawing payload.
 *
 * `route.points` are worksheet-content coordinates. The extensions layer has
 * already applied the PaneMap content-to-screen translation before invoking a
 * drawable, so subtracting the drawable bounds here would make frozen panes
 * drift. Keeping the route in content coordinates also means moving a bound
 * shape is visible as soon as the canonical planner publishes a new route.
 */
export function drawCanonicalConnectorOnCanvas(
  context: CanvasRenderingContext2D,
  payload: ConnectorDrawingPayload,
  bounds: Rect,
): void {
  if (!isDrawingConnectorPayload(payload)) {
    drawUnsupportedDrawingOnCanvas(context, bounds, 'Unsupported connector payload');
    return;
  }
  const points = payload.route.points;
  const start = points[0];
  const end = points[points.length - 1];
  if (!start || !end) {
    drawUnsupportedDrawingOnCanvas(context, bounds, 'Connector route is empty');
    return;
  }
  const strokeWidth = payload.strokeWidth ?? 1.5;
  context.save();
  context.strokeStyle = payload.stroke;
  context.lineWidth = strokeWidth;
  context.lineCap = 'round';
  context.lineJoin = 'round';
  context.beginPath();
  context.moveTo(start.x, start.y);
  if (payload.connectorType === 'straight' || points.length === 2) {
    context.lineTo(end.x, end.y);
  } else if (payload.connectorType === 'curved') {
    for (let index = 1; index < points.length - 1; index += 1) {
      const control = points[index]!;
      const next = points[index + 1]!;
      const midpoint = { x: (control.x + next.x) / 2, y: (control.y + next.y) / 2 };
      context.quadraticCurveTo(control.x, control.y, midpoint.x, midpoint.y);
    }
    context.lineTo(end.x, end.y);
  } else {
    // Elbow routes are already orthogonalized by the canonical planner. Do not
    // infer extra bends in the renderer; doing so would diverge on replay.
    for (let index = 1; index < points.length; index += 1) {
      const point = points[index]!;
      context.lineTo(point.x, point.y);
    }
  }
  context.stroke();
  drawConnectorArrowhead(context, start, directionFrom(start, points[1]!), payload.startArrowhead, payload.stroke, strokeWidth);
  drawConnectorArrowhead(context, end, directionFrom(end, points[points.length - 2]!), payload.endArrowhead, payload.stroke, strokeWidth);
  context.restore();
}

function directionFrom(point: { x: number; y: number }, toward: { x: number; y: number }): number {
  return Math.atan2(point.y - toward.y, point.x - toward.x);
}

function drawConnectorArrowhead(
  context: CanvasRenderingContext2D,
  point: { x: number; y: number },
  angle: number,
  arrowhead: DrawingArrowhead,
  color: string,
  strokeWidth: number,
): void {
  if (arrowhead === 'none') return;
  const size = Math.max(6, Math.min(14, 6 + strokeWidth * 2));
  context.save();
  context.translate(point.x, point.y);
  context.rotate(angle);
  context.fillStyle = color;
  context.strokeStyle = color;
  context.lineWidth = Math.max(1, strokeWidth / 2);
  context.beginPath();
  if (arrowhead === 'diamond') {
    context.moveTo(0, 0);
    context.lineTo(-size, size * 0.55);
    context.lineTo(-size * 2, 0);
    context.lineTo(-size, -size * 0.55);
    context.closePath();
  } else if (arrowhead === 'oval') {
    context.ellipse(-size * 0.75, 0, size * 0.75, size * 0.5, 0, 0, Math.PI * 2);
  } else {
    context.moveTo(0, 0);
    context.lineTo(-size * (arrowhead === 'stealth' ? 1.7 : 2), size * (arrowhead === 'stealth' ? 0.42 : 0.65));
    context.lineTo(-size * (arrowhead === 'stealth' ? 1.25 : 1.1), 0);
    context.lineTo(-size * (arrowhead === 'stealth' ? 1.7 : 2), -size * (arrowhead === 'stealth' ? 0.42 : 0.65));
    context.closePath();
  }
  context.fill();
  if (arrowhead === 'diamond' || arrowhead === 'oval') context.stroke();
  context.restore();
}

function drawUnsupportedDrawingOnCanvas(context: CanvasRenderingContext2D, bounds: Rect, reason: string): void {
  context.save();
  context.strokeStyle = '#b91c1c';
  context.fillStyle = '#b91c1c';
  context.lineWidth = 1.5;
  context.setLineDash([4, 3]);
  context.strokeRect(bounds.x + 1, bounds.y + 1, Math.max(0, bounds.width - 2), Math.max(0, bounds.height - 2));
  context.setLineDash([]);
  context.font = '11px Segoe UI, sans-serif';
  context.textAlign = 'center';
  context.textBaseline = 'middle';
  context.fillText(reason, bounds.x + bounds.width / 2, bounds.y + bounds.height / 2, Math.max(10, bounds.width - 8));
  context.restore();
}

/** Return an endpoint control hit in drawable-local coordinates. */
export function connectorEndpointHitTest(
  payload: ConnectorDrawingPayload,
  bounds: Rect,
  point: { x: number; y: number },
  tolerance = 8,
): { action: 'drawing.connector.endpoint'; data: { kind: 'connector-endpoint'; edge: 'start' | 'end'; endpoint: ConnectorDrawingPayload['start'] } } | null {
  if (!isDrawingConnectorPayload(payload)) return null;
  const points = payload.route.points;
  const start = points[0];
  const end = points[points.length - 1];
  if (!start || !end) return null;
  const absolute = { x: point.x + bounds.x, y: point.y + bounds.y };
  const distance = (candidate: { x: number; y: number }) => Math.hypot(absolute.x - candidate.x, absolute.y - candidate.y);
  if (distance(start) <= tolerance) return { action: 'drawing.connector.endpoint', data: { kind: 'connector-endpoint', edge: 'start', endpoint: structuredClone(payload.start) } };
  if (distance(end) <= tolerance) return { action: 'drawing.connector.endpoint', data: { kind: 'connector-endpoint', edge: 'end', endpoint: structuredClone(payload.end) } };
  return null;
}

function dataChartSeries(
  payload: DataChartDrawingPayload,
  tables: readonly WorkbookTableModel[],
  getSheet: (sheetId: string) => CanvasSheetSnapshot | undefined,
): { categories: string[]; series: CanvasChartSeries[] } {
  const source = payload.source;
  const table = source.kind === 'table' ? tables.find((entry) => entry.id === source.tableId) : undefined;
  const sourceRange = table?.sourceRange ?? (source.kind === 'report-sheet' ? source.range : undefined);
  if (!sourceRange) return { categories: [], series: [] };
  const sheet = getSheet(sourceRange.sheetId);
  if (!sheet) return { categories: [], series: [] };
  const fields = source.kind === 'table'
    ? (table?.fields ?? []).map((field) => ({ id: field.id, name: field.name, ordinal: field.ordinal }))
    : Array.from({ length: sourceRange.endColumn - sourceRange.startColumn + 1 }, (_, offset) => ({ id: `report-column-${offset}`, name: String(sheet.getCell(sourceRange.startRow, sourceRange.startColumn + offset)?.value ?? `Column ${offset + 1}`), ordinal: offset }));
  const fieldById = new Map(fields.map((field) => [field.id, field]));
  const isVisibleField = (field: { ordinal: number } | undefined): boolean => Boolean(field && (payload.inspector.showHiddenData || !sheet.hiddenColumns.includes(sourceRange.startColumn + field.ordinal)));
  const categoryBinding = payload.bindings.category[0];
  const categoryField = categoryBinding && isVisibleField(fieldById.get(categoryBinding.fieldId)) ? fieldById.get(categoryBinding.fieldId) : undefined;
  const valueBindings = payload.bindings.values.filter((binding) => isVisibleField(fieldById.get(binding.fieldId)));
  if (!valueBindings.length) return { categories: [], series: [] };
  const buckets = new Map<string, Map<string, number[]>>();
  for (let row = sourceRange.startRow + 1; row <= sourceRange.endRow; row += 1) {
    if (!payload.inspector.showHiddenData && sheet.hiddenRows.includes(row)) continue;
    const category = String(categoryField ? sheet.getCell(row, sourceRange.startColumn + categoryField.ordinal)?.value ?? '' : row - sourceRange.startRow);
    const byField = buckets.get(category) ?? new Map<string, number[]>();
    for (const binding of valueBindings) {
      const field = fieldById.get(binding.fieldId)!;
      const raw = sheet.getCell(row, sourceRange.startColumn + field.ordinal)?.value ?? '';
      const value = numericCellValue(String(raw));
      if (value !== undefined) byField.set(binding.fieldId, [...(byField.get(binding.fieldId) ?? []), value]);
    }
    buckets.set(category, byField);
  }
  const aggregate = (values: number[], mode: DataChartDrawingPayload['bindings']['values'][number]['aggregate']): number => {
    if (!values.length || mode === 'none') return values.length ? values[values.length - 1]! : 0;
    if (mode === 'count') return values.length;
    if (mode === 'min') return Math.min(...values);
    if (mode === 'max') return Math.max(...values);
    if (mode === 'average') return values.reduce((sum, value) => sum + value, 0) / values.length;
    return values.reduce((sum, value) => sum + value, 0);
  };
  const categories = [...buckets.keys()];
  const series = valueBindings.map((binding) => ({ name: fieldById.get(binding.fieldId)!.name, values: categories.map((category) => aggregate(buckets.get(category)?.get(binding.fieldId) ?? [], binding.aggregate)) }));
  const sortBinding = valueBindings.find((binding) => binding.sort);
  if (sortBinding) {
    const index = series.findIndex((entry) => entry.name === fieldById.get(sortBinding.fieldId)?.name);
    if (index >= 0) {
      const order = sortBinding.sort === 'desc' ? -1 : 1;
      const orderIndexes = categories.map((_category, categoryIndex) => categoryIndex).sort((left, right) => (series[index]!.values[left]! - series[index]!.values[right]!) * order);
      return { categories: orderIndexes.map((categoryIndex) => categories[categoryIndex]!), series: series.map((entry) => ({ ...entry, values: orderIndexes.map((categoryIndex) => entry.values[categoryIndex]!) })) };
    }
  }
  return { categories, series };
}

export interface CameraSourceGeometry {
  left: number;
  top: number;
  width: number;
  height: number;
  firstRow: number;
  lastRow: number;
  firstColumn: number;
  lastColumn: number;
}

const CAMERA_SURFACE_MAX_EDGE = 4096;
const cameraSurfaceCache = new WeakMap<object, Map<string, HTMLCanvasElement>>();

function cameraRangeKey(range: RangeRef): string {
  return `${range.sheetId}:${range.startRow}:${range.endRow}:${range.startColumn}:${range.endColumn}`;
}

export function resolveCameraSourceGeometry(source: CanvasSheetSnapshot, range: RangeRef): CameraSourceGeometry | null {
  if (range.sheetId !== source.id
    || !Number.isSafeInteger(range.startRow) || !Number.isSafeInteger(range.endRow)
    || !Number.isSafeInteger(range.startColumn) || !Number.isSafeInteger(range.endColumn)
    || range.startRow < 0 || range.startColumn < 0
    || range.endRow < range.startRow || range.endColumn < range.startColumn
    || range.endRow >= source.rowCount || range.endColumn >= source.columnCount) return null;
  const skeleton = new SheetSkeleton({
    rowCount: source.rowCount,
    columnCount: source.columnCount,
    defaultRowHeight: source.defaultRowHeightPx,
    defaultColumnWidth: source.defaultColumnWidthPx,
    rowHeights: new Map(Object.entries(source.rowHeightsPx).map(([key, value]) => [Number(key), value])),
    columnWidths: new Map(Object.entries(source.columnWidthsPx).map(([key, value]) => [Number(key), value])),
    hiddenRows: new Set(source.hiddenRows),
    hiddenColumns: new Set(source.hiddenColumns),
  });
  let firstRow = range.startRow;
  while (firstRow <= range.endRow && skeleton.isRowHidden(firstRow)) firstRow += 1;
  let firstColumn = range.startColumn;
  while (firstColumn <= range.endColumn && skeleton.isColumnHidden(firstColumn)) firstColumn += 1;
  let lastRow = range.endRow;
  while (lastRow >= range.startRow && skeleton.isRowHidden(lastRow)) lastRow -= 1;
  let lastColumn = range.endColumn;
  while (lastColumn >= range.startColumn && skeleton.isColumnHidden(lastColumn)) lastColumn -= 1;
  if (firstRow > lastRow || firstColumn > lastColumn) return null;
  const left = skeleton.getColumnLeft(firstColumn);
  const top = skeleton.getRowTop(firstRow);
  if (left < 0 || top < 0) return null;
  let width = 0;
  for (let column = range.startColumn; column <= range.endColumn; column += 1) width += skeleton.getColumnWidth(column);
  let height = 0;
  for (let row = range.startRow; row <= range.endRow; row += 1) height += skeleton.getRowHeight(row);
  return width > 0 && height > 0 ? { left, top, width, height, firstRow, lastRow, firstColumn, lastColumn } : null;
}

function cameraCellProvider(source: CanvasSheetSnapshot, range: RangeRef): CellProvider {
  const merges = source.merges.filter((merge) => merge.range.endRow >= range.startRow
    && merge.range.startRow <= range.endRow
    && merge.range.endColumn >= range.startColumn
    && merge.range.startColumn <= range.endColumn);
  const mergeAt = (row: number, column: number) => merges.find((merge) => row >= merge.range.startRow && row <= merge.range.endRow && column >= merge.range.startColumn && column <= merge.range.endColumn);
  return ({ row, column }): CellRenderData | undefined => {
    const cell = source.getCell(row, column);
    const merge = mergeAt(row, column);
    if (!cell && !merge) return undefined;
    const value: CellRenderData = {
      value: cell?.value,
      displayValue: cell?.displayValue,
      formula: cell?.formula,
      style: cell?.style,
      editor: cell?.editor,
      presentation: cell?.presentation,
      hasComment: cell?.hasComment,
      invalid: cell?.invalid,
      overlay: cell?.overlay,
    };
    if (merge) {
      value.merge = {
        startRow: merge.range.startRow,
        endRow: merge.range.endRow,
        startColumn: merge.range.startColumn,
        endColumn: merge.range.endColumn,
        isAnchor: merge.anchor.row === row && merge.anchor.column === column,
      };
    }
    return value;
  };
}

function cameraSurface(source: CanvasSheetSnapshot, range: RangeRef): HTMLCanvasElement | null {
  if (typeof document === 'undefined') return null;
  const geometry = resolveCameraSourceGeometry(source, range);
  if (!geometry) return null;
  const key = cameraRangeKey(range);
  let surfaces = cameraSurfaceCache.get(source);
  if (!surfaces) {
    surfaces = new Map();
    cameraSurfaceCache.set(source, surfaces);
  }
  const cached = surfaces.get(key);
  if (cached) return cached;
  const scale = Math.min(1, CAMERA_SURFACE_MAX_EDGE / geometry.width, CAMERA_SURFACE_MAX_EDGE / geometry.height);
  const canvas = document.createElement('canvas');
  canvas.width = Math.max(1, Math.ceil(geometry.width * scale));
  canvas.height = Math.max(1, Math.ceil(geometry.height * scale));
  const surfaceContext = canvas.getContext('2d');
  if (!surfaceContext) return null;
  const pane: RenderPane = {
    id: 'main',
    screenRect: { x: geometry.left, y: geometry.top, width: geometry.width, height: geometry.height },
    contentOrigin: { x: geometry.left, y: geometry.top },
    visibleRange: {
      startRow: range.startRow,
      endRow: range.endRow,
      startColumn: range.startColumn,
      endColumn: range.endColumn,
    },
  };
  surfaceContext.save();
  surfaceContext.scale(scale, scale);
  surfaceContext.beginPath();
  surfaceContext.rect(0, 0, geometry.width, geometry.height);
  surfaceContext.clip();
  surfaceContext.translate(-geometry.left, -geometry.top);
  const skeleton = new SheetSkeleton({
    rowCount: source.rowCount,
    columnCount: source.columnCount,
    defaultRowHeight: source.defaultRowHeightPx,
    defaultColumnWidth: source.defaultColumnWidthPx,
    rowHeights: new Map(Object.entries(source.rowHeightsPx).map(([entry, value]) => [Number(entry), value])),
    columnWidths: new Map(Object.entries(source.columnWidthsPx).map(([entry, value]) => [Number(entry), value])),
    hiddenRows: new Set(source.hiddenRows),
    hiddenColumns: new Set(source.hiddenColumns),
  });
  const cellProvider = cameraCellProvider(source, range);
  const options = { context: surfaceContext, skeleton, pane, visibleRange: pane.visibleRange, cellProvider, theme: DEFAULT_RENDER_THEME };
  drawGridLayer(options);
  drawCellLayer(options);
  surfaceContext.restore();
  surfaces.set(key, canvas);
  return canvas;
}

export function drawCameraOnCanvas(context: CanvasRenderingContext2D, payload: CameraDrawingPayload, bounds: Rect, getSheet: (sheetId: string) => CanvasSheetSnapshot | undefined): void {
  const source = getSheet(payload.sourceRange.sheetId);
  context.save();
  context.fillStyle = '#ffffff';
  context.fillRect(bounds.x, bounds.y, bounds.width, bounds.height);
  if (!source) {
    context.fillStyle = '#b91c1c';
    context.fillText('Missing source range', bounds.x + 8, bounds.y + 18);
    context.restore();
    return;
  }
  const surface = cameraSurface(source, payload.sourceRange);
  if (!surface) {
    context.fillStyle = '#b91c1c';
    context.fillText('Invalid camera source range', bounds.x + 8, bounds.y + 18);
    context.restore();
    return;
  }
  const scale = Math.min(bounds.width / surface.width, bounds.height / surface.height);
  const width = surface.width * scale;
  const height = surface.height * scale;
  const x = bounds.x + (bounds.width - width) / 2;
  const y = bounds.y + (bounds.height - height) / 2;
  context.drawImage(surface, 0, 0, surface.width, surface.height, x, y, width, height);
  context.restore();
}

function drawFormControlOnCanvas(context: CanvasRenderingContext2D, payload: FormControlDrawingPayload, bounds: Rect): void {
  context.save();
  context.fillStyle = payload.style.fill;
  context.strokeStyle = payload.enabled ? payload.style.border : '#b7bdc4';
  context.lineWidth = 1;
  context.fillRect(bounds.x, bounds.y, bounds.width, bounds.height);
  context.strokeRect(bounds.x + 0.5, bounds.y + 0.5, bounds.width - 1, bounds.height - 1);
  context.fillStyle = payload.enabled ? payload.style.textColor : '#8a9097';
  context.font = `${payload.style.fontSize ?? 12}px "Microsoft YaHei", "Segoe UI", sans-serif`;
  context.textAlign = payload.controlType === 'label' ? 'left' : 'center';
  context.textBaseline = 'middle';
  const prefix = payload.controlType === 'checkbox' ? (payload.value ? '☑ ' : '☐ ') : payload.controlType === 'option-button' ? (payload.value ? '◉ ' : '○ ') : '';
  const valueLabel = payload.controlType === 'spin-button' || payload.controlType === 'scrollbar'
    ? ` (${payload.value})`
    : payload.controlType === 'list-box' || payload.controlType === 'combo-box'
      ? (payload.value ? `: ${payload.value}` : '')
      : '';
  context.fillText(`${prefix}${payload.text ?? payload.controlType}${valueLabel}`, payload.controlType === 'label' ? bounds.x + 5 : bounds.x + bounds.width / 2, bounds.y + bounds.height / 2, Math.max(4, bounds.width - 10));
  context.restore();
}

function drawCanonicalChartOnCanvas(options: {
  context: CanvasRenderingContext2D;
  payload: ChartDrawingPayload;
  bounds: Rect;
  categories: string[];
  series: CanvasChartSeries[];
}): void {
  const { context, payload, bounds, categories, series } = options;
  const elements = payload.elements;
  const legendPosition = elements.legend?.visible ? elements.legend.position : 'none';
  const title = elements.title;
  const { x, y, width, height } = bounds;
  context.save();
  context.translate(x, y);
  context.fillStyle = elements.chartArea?.fill ?? "#ffffff";
  context.fillRect(0, 0, width, height);
  context.strokeStyle = elements.chartArea?.border ?? "#e2e8f0";
  context.lineWidth = elements.chartArea?.borderWidth ?? 1;
  context.strokeRect(0, 0, width, height);
  if (title) {
    context.fillStyle = "#1e293b";
    context.font = "bold 14px Segoe UI, sans-serif";
    context.textAlign = "left";
    context.textBaseline = "top";
    context.fillText(title, 16, 12);
  }
  if (payload.chartType === "pie" || payload.chartType === "doughnut") {
    drawCanonicalPieChart(context, series[0], width, height, payload.chartType === "doughnut");
    drawCanonicalLegend(context, series, width, height, legendPosition);
    context.restore();
    return;
  }
  const plotTop = title ? 40 : 20;
  const plotBottom = height - (legendPosition === "bottom" ? 40 : 24);
  const plotLeft = 48;
  const plotRight = width - (legendPosition === "right" ? 90 : 16);
  const plotWidth = Math.max(10, plotRight - plotLeft);
  const plotHeight = Math.max(10, plotBottom - plotTop);
  if (elements.plotArea?.fill) {
    context.fillStyle = elements.plotArea.fill;
    context.fillRect(plotLeft, plotTop, plotWidth, plotHeight);
  }
  const allValues = series.flatMap((entry) => entry.values.filter((_value, index) => !entry.missing?.[index]));
  const maxValue = Math.max(1, ...allValues.map((value) => Math.abs(value)));
  const minValue = Math.min(0, ...allValues);
  const maxAxis = payload.stacked === "percent" ? 100 : elements.valueAxis?.maximum ?? Math.max(1, maxValue * 1.1);
  const minAxis = payload.stacked === "percent" ? 0 : elements.valueAxis?.minimum ?? Math.min(0, minValue);
  const axisSpan = Math.max(1, maxAxis - minAxis);
  context.strokeStyle = elements.valueAxis?.majorGridlines?.color ?? "#f1f5f9";
  context.lineWidth = elements.valueAxis?.majorGridlines?.width ?? 1;
  context.fillStyle = "#64748b";
  context.font = "11px Segoe UI, sans-serif";
  context.textAlign = "right";
  context.textBaseline = "middle";
  if (elements.valueAxis?.majorGridlines?.visible !== false) for (let index = 0; index <= 4; index += 1) {
    const ratio = index / 4;
    const gridY = plotTop + ratio * plotHeight;
    const gridValue = maxAxis - ratio * axisSpan;
    context.beginPath();
    context.moveTo(plotLeft, gridY);
    context.lineTo(plotRight, gridY);
    context.stroke();
    context.fillText(String(Math.round(gridValue)), plotLeft - 8, gridY);
  }
  const stackMode = payload.stacked ?? "none";
  const stacked = stackMode !== "none";
  const valueAt = (seriesIndex: number, categoryIndex: number): number => series[seriesIndex]?.values[categoryIndex] ?? 0;
  const stackedValue = (seriesIndex: number, categoryIndex: number): { start: number; end: number } => {
    const raw = valueAt(seriesIndex, categoryIndex);
    if (!stacked) return { start: 0, end: raw };
    const values = series.map((entry) => entry.values[categoryIndex] ?? 0);
    if (stackMode === "percent") {
      const total = values.reduce((sum, value) => sum + Math.abs(value), 0) || 1;
      const normalized = (raw / total) * 100;
      const start = values.slice(0, seriesIndex).reduce((sum, value) => sum + (value / total) * 100, 0);
      return { start, end: start + normalized };
    }
    const start = values.slice(0, seriesIndex).reduce((sum, value) => sum + value, 0);
    return { start, end: start + raw };
  };
  const yFor = (value: number): number => plotTop + ((maxAxis - value) / axisSpan) * plotHeight;
  const categoryCount = Math.max(1, categories.length, ...series.map((entry) => entry.values.length));
  if (payload.chartType === "column" || payload.chartType === "bar" || payload.chartType === "combo") {
    const categorySize = (payload.chartType === "bar" ? plotHeight : plotWidth) / categoryCount;
    const columnSeries = payload.chartType === "combo" ? series.slice(1) : series;
    const seriesCount = Math.max(1, stacked ? 1 : columnSeries.length);
    for (let categoryIndex = 0; categoryIndex < categoryCount; categoryIndex += 1) {
      for (let seriesIndex = 0; seriesIndex < columnSeries.length; seriesIndex += 1) {
        const sourceIndex = payload.chartType === "combo" ? seriesIndex + 1 : seriesIndex;
        const range = stackedValue(sourceIndex, categoryIndex);
        const color = columnSeries[seriesIndex]?.color ?? CHART_PALETTE[sourceIndex % CHART_PALETTE.length]!;
        context.fillStyle = color;
        if (payload.chartType === "bar") {
          const barHeight = Math.max(4, categorySize * (stacked ? 0.64 : 0.62));
          const yPosition = plotTop + categoryIndex * categorySize + (categorySize - barHeight) / 2;
          const startX = plotLeft + (Math.min(range.start, range.end) - minAxis) / axisSpan * plotWidth;
          const endX = plotLeft + (Math.max(range.start, range.end) - minAxis) / axisSpan * plotWidth;
          context.fillRect(startX, yPosition, Math.max(1, endX - startX), barHeight);
        } else {
          const groupWidth = categorySize * 0.72;
          const barWidth = Math.max(4, groupWidth / seriesCount);
          const baseX = plotLeft + categoryIndex * categorySize + categorySize * 0.14;
          const left = stacked ? baseX : baseX + seriesIndex * barWidth;
          const topY = yFor(Math.max(range.start, range.end));
          const bottomY = yFor(Math.min(range.start, range.end));
          context.fillRect(left, topY, Math.max(2, stacked ? groupWidth : barWidth - 2), Math.max(1, bottomY - topY));
          if (elements.dataLabels?.visible) {
            context.fillStyle = "#475569";
            context.textAlign = "center";
            context.textBaseline = "bottom";
            context.fillText(String(Math.round(valueAt(sourceIndex, categoryIndex))), left + (stacked ? groupWidth : barWidth) / 2, topY - 2);
          }
        }
      }
      if (payload.chartType !== "bar") {
        context.fillStyle = "#64748b";
        context.textAlign = "center";
        context.textBaseline = "top";
        context.fillText(categories[categoryIndex] ?? String(categoryIndex + 1), plotLeft + categoryIndex * categorySize + categorySize / 2, plotBottom + 6);
      }
    }
    if (payload.chartType === "combo" && series.length > 0) {
      drawCanonicalLineSeries(context, series[0]!, categories, plotLeft, plotTop, plotWidth, plotHeight, yFor, "#2563eb", false);
    }
  } else if (payload.chartType !== "scatter") {
    for (let seriesIndex = 0; seriesIndex < series.length; seriesIndex += 1) {
      const current = series[seriesIndex]!;
      const lineSeries = stacked
        ? { ...current, values: current.values.map((_value, categoryIndex) => stackedValue(seriesIndex, categoryIndex).end) }
        : current;
      drawCanonicalLineSeries(context, lineSeries, categories, plotLeft, plotTop, plotWidth, plotHeight, yFor, current.color ?? CHART_PALETTE[seriesIndex % CHART_PALETTE.length]!, payload.chartType === "area");
    }
  }
  if (payload.chartType === "scatter") {
    for (let seriesIndex = 0; seriesIndex < series.length; seriesIndex += 1) {
      const current = series[seriesIndex]!;
      context.fillStyle = current.color ?? CHART_PALETTE[seriesIndex % CHART_PALETTE.length]!;
      for (let categoryIndex = 0; categoryIndex < current.values.length; categoryIndex += 1) {
        const pointX = plotLeft + (categoryIndex / Math.max(1, categoryCount - 1)) * plotWidth;
        const pointY = yFor(current.values[categoryIndex] ?? 0);
        context.beginPath();
        context.arc(pointX, pointY, 4, 0, Math.PI * 2);
        context.fill();
      }
    }
  }
  drawCanonicalLegend(context, series, width, height, legendPosition);
  context.restore();
}

function drawCanonicalLineSeries(
  context: CanvasRenderingContext2D,
  series: CanvasChartSeries,
  categories: string[],
  left: number,
  top: number,
  width: number,
  height: number,
  yFor: (value: number) => number,
  color: string,
  area: boolean,
): void {
  const points = series.values.map((value, index) => ({
    x: left + (index / Math.max(1, Math.max(categories.length, series.values.length) - 1)) * width,
    y: yFor(value),
  }));
  if (points.length === 0) return;
  if (area) {
    context.save();
    context.fillStyle = `${color}22`;
    context.beginPath();
    context.moveTo(points[0]!.x, top + height);
    for (const point of points) context.lineTo(point.x, point.y);
    context.lineTo(points[points.length - 1]!.x, top + height);
    context.closePath();
    context.fill();
    context.restore();
  }
  context.strokeStyle = color;
  context.lineWidth = 2.5;
  context.beginPath();
  points.forEach((point, index) => index === 0 ? context.moveTo(point.x, point.y) : context.lineTo(point.x, point.y));
  context.stroke();
  if (series.marker?.enabled) {
    const radius = Math.max(2, (series.marker.size ?? 6) / 2);
    context.fillStyle = series.marker.fill ?? color;
    context.strokeStyle = series.marker.border ?? color;
    context.lineWidth = 1;
    for (const point of points) {
      context.beginPath();
      context.arc(point.x, point.y, radius, 0, Math.PI * 2);
      context.fill();
      context.stroke();
    }
  }
}

function drawCanonicalPieChart(
  context: CanvasRenderingContext2D,
  series: CanvasChartSeries | undefined,
  width: number,
  height: number,
  doughnut: boolean,
): void {
  if (!series) return;
  const total = series.values.reduce((sum, value) => sum + Math.max(0, value), 0);
  if (total <= 0) return;
  const centerX = width / 2;
  const centerY = height / 2;
  const radius = Math.min(width, height) / 2 - 24;
  let angle = -Math.PI / 2;
  for (let index = 0; index < series.values.length; index += 1) {
    const value = Math.max(0, series.values[index] ?? 0);
    const sweep = (value / total) * Math.PI * 2;
    context.fillStyle = CHART_PALETTE[index % CHART_PALETTE.length]!;
    context.beginPath();
    context.moveTo(centerX, centerY);
    context.arc(centerX, centerY, radius, angle, angle + sweep);
    context.closePath();
    context.fill();
    angle += sweep;
  }
  if (doughnut) {
    context.fillStyle = "#ffffff";
    context.beginPath();
    context.arc(centerX, centerY, radius * 0.55, 0, Math.PI * 2);
    context.fill();
  }
}

function drawCanonicalLegend(
  context: CanvasRenderingContext2D,
  series: readonly CanvasChartSeries[],
  width: number,
  height: number,
  position: 'top' | 'bottom' | 'left' | 'right' | 'none',
): void {
  if (position === "none") return;
  context.font = "11px Segoe UI, sans-serif";
  context.textAlign = "left";
  context.textBaseline = "middle";
  let x = position === "left" ? 4 : 16;
  const y = position === "bottom" ? height - 16 : 24;
  for (const [index, entry] of series.entries()) {
    const color = entry.color ?? CHART_PALETTE[index % CHART_PALETTE.length]!;
    context.fillStyle = color;
    context.fillRect(x, y - 4, 10, 10);
    context.fillStyle = "#475569";
    context.fillText(entry.name, x + 14, y + 1);
    x += context.measureText(entry.name).width + 32;
    if (x > width - 40) break;
  }
}

function drawCanonicalSparklineOnCanvas(options: {
  context: CanvasRenderingContext2D;
  sparkline: SparklineModel;
  values: number[];
  rect: Rect;
}): void {
  const { context, sparkline, values, rect } = options;
  const { x, y, width, height } = rect;
  if (values.length === 0) return;
  const max = Math.max(...values, 0);
  const min = Math.min(...values, 0);
  const span = Math.max(1, max - min);
  context.save();
  context.translate(x, y);
  if (sparkline.type === "line") {
    context.strokeStyle = sparkline.color || "#2563eb";
    context.lineWidth = 1.5;
    context.beginPath();
    values.forEach((value, index) => {
      const px = (index / Math.max(1, values.length - 1)) * width;
      const py = height - ((value - min) / span) * height;
      if (index === 0) context.moveTo(px, py);
      else context.lineTo(px, py);
    });
    context.stroke();
  } else {
    const barWidth = width / Math.max(1, values.length);
    values.forEach((value, index) => {
      const barHeight = Math.max(1, ((value - min) / span) * height);
      context.fillStyle = value < 0 ? sparkline.negativeColor || "#ef4444" : sparkline.color || "#2563eb";
      context.fillRect(index * barWidth + 1, height - barHeight, Math.max(1, barWidth - 2), barHeight);
    });
  }
  if (sparkline.showAxis) {
    context.strokeStyle = "#cbd5e1";
    context.lineWidth = 1;
    const axisY = height - ((0 - min) / span) * height;
    context.beginPath();
    context.moveTo(0, axisY);
    context.lineTo(width, axisY);
    context.stroke();
  }
  if (sparkline.showMarkers || sparkline.highlightMax || sparkline.highlightMin || sparkline.highlightFirst || sparkline.highlightLast) {
    const marker = (index: number, color: string): void => {
      const px = (index / Math.max(1, values.length - 1)) * width;
      const py = height - ((values[index]! - min) / span) * height;
      context.fillStyle = color;
      context.beginPath();
      context.arc(px, py, 2, 0, Math.PI * 2);
      context.fill();
    };
    if (sparkline.showMarkers) values.forEach((_value, index) => marker(index, sparkline.color || "#2563eb"));
    if (sparkline.highlightMax) marker(values.indexOf(max), "#16a34a");
    if (sparkline.highlightMin) marker(values.indexOf(min), "#dc2626");
    if (sparkline.highlightFirst) marker(0, "#f59e0b");
    if (sparkline.highlightLast) marker(values.length - 1, "#f59e0b");
  }
  context.restore();
}

interface PivotControlMember {
  value: PivotScalar;
  key: PivotMemberKey;
  label: string;
  selected: boolean;
  hasData: boolean;
}

function pivotControlMembers(drawingId: string, payload: PivotSlicerDrawingPayload | PivotTimelineDrawingPayload, pivotResults: Record<string, PivotResultTree>): PivotControlMember[] {
  const tree = pivotResults[payload.pivotId];
  if (payload.kind === 'slicer') {
    const projected = tree?.slicerItems?.[drawingId];
    if (projected) return projected.map((item: PivotSlicerItemProjection) => ({ ...item }));
  }
  const field = tree?.fields.fields.find((entry) => entry.fieldId === payload.fieldId);
  const values = [...(field?.values ?? [])];
  if (values.length === 0 && tree) {
    const collect = (nodes: readonly PivotResultTree['rows'][number][]): void => {
      for (const node of nodes) {
        if (node.fieldId === payload.fieldId && node.memberKey) values.push(pivotScalarFromMemberKey(node.memberKey));
        collect(node.children);
      }
    };
    collect(tree.rows);
  }
  const members = new Map<string, PivotControlMember>();
  for (const value of values) {
    const key = createPivotMemberKey(value);
    const selected = payload.kind === 'slicer'
      ? payload.filter.mode === 'all'
        || (payload.filter.mode === 'include' ? payload.filter.memberKeys.some((candidate) => pivotMemberKey(candidate) === pivotMemberKey(key)) : !payload.filter.memberKeys.some((candidate) => pivotMemberKey(candidate) === pivotMemberKey(key)))
      : false;
    members.set(pivotMemberKey(key), { value, key, label: formatPivotMember(value), selected, hasData: true });
  }
  return [...members.values()];
}

function pivotTimelinePeriods(payload: PivotTimelineDrawingPayload, pivotResults: Record<string, PivotResultTree>): ReturnType<typeof buildPivotTimelineTiles> {
  const values = pivotControlMembers('', payload, pivotResults).map((entry) => entry.value);
  const tiles = buildPivotTimelineTiles(values, payload.level);
  const boundedTiles = tiles.filter((tile) => (!payload.bounds.start || tile.end >= payload.bounds.start) && (!payload.bounds.end || tile.start <= payload.bounds.end));
  const startIndex = payload.scrollPosition ? Math.max(0, boundedTiles.findIndex((tile) => tile.start >= payload.scrollPosition!)) : 0;
  return boundedTiles.slice(startIndex);
}

function pivotControlHitTest(
  payload: PivotSlicerDrawingPayload | PivotTimelineDrawingPayload,
  members: readonly PivotControlMember[],
  periods: readonly ReturnType<typeof buildPivotTimelineTiles>[number][],
  point: { x: number; y: number },
  bounds: Rect,
): import('@react-sheets/render-engine').FloatingControlHit | null {
  const headerHeight = Math.min(26, bounds.height);
  if (payload.kind === 'slicer') {
    if (payload.settings.showHeader && point.y <= headerHeight && point.x >= Math.max(0, bounds.width - 26)) return { action: 'pivot.slicer.clear', data: { kind: 'slicer-clear' } satisfies PivotControlAction };
    if (payload.settings.showHeader && point.y <= headerHeight) return null;
    const itemHeight = payload.settings.itemHeight;
    const columnCount = Math.max(1, payload.settings.columnCount);
    const visibleMembers = members.filter((member) => payload.settings.showNoDataItems || member.hasData);
    const itemY = point.y - (payload.settings.showHeader ? headerHeight : 0);
    const rowIndex = Math.floor(itemY / itemHeight);
    const columnWidth = bounds.width / columnCount;
    const columnIndex = Math.floor(point.x / Math.max(1, columnWidth));
    const index = rowIndex * columnCount + columnIndex;
    const member = index >= 0 && index < visibleMembers.length ? visibleMembers[index] : undefined;
    return member ? { action: 'pivot.slicer.member', data: { kind: 'slicer-member', memberKey: member.key } satisfies PivotControlAction } : null;
  }
  if (payload.showHeader && point.y <= headerHeight && point.x < 74) return { action: 'pivot.timeline.scroll', data: { kind: 'timeline-scroll', direction: -1 } satisfies PivotControlAction };
  if (payload.showHeader && point.y <= headerHeight && point.x >= Math.max(0, bounds.width - 74)) return { action: 'pivot.timeline.scroll', data: { kind: 'timeline-scroll', direction: 1 } satisfies PivotControlAction };
  if (payload.showHeader && payload.showTimeLevel && point.y <= headerHeight && point.x >= 86 && point.x < 180) {
    const levels: PivotTimelineLevel[] = ['years', 'quarters', 'months', 'days'];
    return { action: 'pivot.timeline.level', data: { kind: 'timeline-level', level: levels[(levels.indexOf(payload.level) + 1) % levels.length]! } satisfies PivotControlAction };
  }
  if (point.y > headerHeight && point.y < bounds.height - 8 && point.x <= 12) return { action: 'pivot.timeline.handle', data: { kind: 'timeline-handle', edge: 'start' } satisfies PivotControlAction };
  if (point.y > headerHeight && point.y < bounds.height - 8 && point.x >= bounds.width - 12) return { action: 'pivot.timeline.handle', data: { kind: 'timeline-handle', edge: 'end' } satisfies PivotControlAction };
  const trackWidth = Math.max(1, bounds.width - 24);
  const periodIndex = Math.floor(((point.x - 12) / trackWidth) * periods.length);
  const period = periodIndex >= 0 && periodIndex < periods.length ? periods[periodIndex] : undefined;
  return period ? { action: 'pivot.timeline.period', data: { kind: 'timeline-period', ...period } satisfies PivotControlAction } : null;
}

function drawPivotControlOnCanvas(options: {
  context: CanvasRenderingContext2D;
  payload: PivotSlicerDrawingPayload | PivotTimelineDrawingPayload;
  bounds: Rect;
  members: readonly PivotControlMember[];
  periods: readonly ReturnType<typeof buildPivotTimelineTiles>[number][];
}): void {
  const { context, payload, bounds, members, periods } = options;
  const style = payload.style;
  context.save();
  context.fillStyle = style.fill;
  context.strokeStyle = style.border;
  context.lineWidth = 1;
  context.fillRect(bounds.x, bounds.y, bounds.width, bounds.height);
  context.strokeRect(bounds.x, bounds.y, bounds.width, bounds.height);
  const headerHeight = payload.kind === 'timeline' ? (payload.showHeader ? Math.min(26, bounds.height) : 0) : Math.min(26, bounds.height);
  context.fillStyle = style.accentColor;
  const effectiveHeaderHeight = payload.kind === 'slicer' && !payload.settings.showHeader ? 0 : headerHeight;
  if (effectiveHeaderHeight > 0) context.fillRect(bounds.x, bounds.y, bounds.width, effectiveHeaderHeight);
  context.fillStyle = style.textColor;
  context.font = `600 ${style.fontSize}px Segoe UI, sans-serif`;
  context.textBaseline = "middle";
  if (payload.kind === 'slicer') context.fillText(`Slicer · ${payload.fieldId}`, bounds.x + 8, bounds.y + Math.min(13, bounds.height / 2), Math.max(10, bounds.width - 16));
  else if (headerHeight > 0) {
    context.fillText(`${payload.caption || 'Timeline'} · ${payload.fieldId}`, bounds.x + 8, bounds.y + Math.min(13, bounds.height / 2), Math.max(10, bounds.width - 16));
    if (payload.showTimeLevel) context.fillText(payload.level, bounds.x + 88, bounds.y + Math.min(13, bounds.height / 2), 82);
  }
  context.font = `${style.fontSize}px Segoe UI, sans-serif`;
  if (payload.kind === "slicer") {
    const visibleMembers = members.filter((member) => payload.settings.showNoDataItems || member.hasData);
    const columnCount = Math.max(1, payload.settings.columnCount);
    const columnWidth = bounds.width / columnCount;
    const rows = Math.max(0, Math.floor((bounds.height - headerHeight) / payload.settings.itemHeight));
    visibleMembers.slice(0, rows * columnCount).forEach((member, index) => {
      const row = Math.floor(index / columnCount);
      const column = index % columnCount;
      const left = bounds.x + column * columnWidth;
      const top = bounds.y + headerHeight + row * payload.settings.itemHeight;
      if (member.selected) {
        context.fillStyle = style.selectedFill ?? '#bfdbfe';
        context.fillRect(left + 2, top + 1, Math.max(0, columnWidth - 4), payload.settings.itemHeight - 2);
      }
      if (!member.hasData && payload.settings.showNoDataStyle) {
        context.fillStyle = '#94a3b8';
        context.fillRect(left + 2, top + 1, Math.max(0, columnWidth - 4), payload.settings.itemHeight - 2);
      }
      context.fillStyle = style.textColor;
      context.globalAlpha = member.hasData ? 1 : 0.55;
      context.fillText(member.label, left + 8, top + payload.settings.itemHeight / 2, Math.max(10, columnWidth - 16));
      context.globalAlpha = 1;
    });
    if (headerHeight > 0) {
      context.fillStyle = style.textColor;
      context.font = `600 ${style.fontSize}px Segoe UI, sans-serif`;
      if (payload.settings.multiSelect) context.fillText('☷', bounds.x + Math.max(0, bounds.width - 44), bounds.y + 13, 14);
      if (payload.filter.mode !== 'all') context.fillText('×', bounds.x + Math.max(0, bounds.width - 20), bounds.y + 13, 14);
    }
  } else {
    const detail = `${payload.period.start ?? "Start"} — ${payload.period.end ?? "End"}`;
    const contentTop = bounds.y + headerHeight;
    if (payload.showSelectionLabel) context.fillText(detail, bounds.x + 8, contentTop + Math.min(18, Math.max(12, bounds.height - headerHeight - 12)), Math.max(10, bounds.width - 16));
    const trackY = bounds.y + Math.max(headerHeight + 28, bounds.height - (payload.showHorizontalScrollbar ? 22 : 10));
    context.strokeStyle = style.border;
    context.beginPath();
    context.moveTo(bounds.x + 12, trackY);
    context.lineTo(bounds.x + Math.max(12, bounds.width - 12), trackY);
    context.stroke();
    const visibleCount = Math.min(periods.length, Math.max(1, Math.floor((bounds.width - 24) / 24)));
    periods.slice(0, visibleCount).forEach((period, index) => {
      const x = bounds.x + 12 + (index + 0.5) * ((bounds.width - 24) / Math.max(1, visibleCount));
      context.fillStyle = style.textColor;
      context.fillRect(x - 2, trackY - 3, 4, 6);
      if (period.start >= (payload.period.start ?? '') && period.start <= (payload.period.end ?? '\uffff')) {
        context.fillStyle = style.selectedFill ?? '#bfdbfe';
        context.fillRect(x - 10, trackY - 6, 20, 12);
      }
    });
    if (payload.showHorizontalScrollbar) {
      context.fillStyle = style.accentColor;
      context.fillRect(bounds.x + 8, trackY - 7, 4, 14);
      context.fillRect(bounds.x + Math.max(8, bounds.width - 12), trackY - 7, 4, 14);
    }
  }
  context.restore();
}

export interface CanvasFloatingRendererInput {
  drawings: readonly DrawingObject[];
  drawingPayloads: ReadonlyMap<string, DrawingPayload>;
  allSheets: readonly CanvasSheetSnapshot[];
  sheet: CanvasSheetSnapshot;
  pivotResults: Record<string, PivotResultTree>;
  sparklines: readonly SparklineModel[];
  skeleton: SheetSkeleton;
  imageCache: Map<string, HTMLImageElement>;
  requestRender: () => void;
  tables: readonly WorkbookTableModel[];
  resolveAssetUrl?: (asset: AssetRef) => Promise<string>;
  assetUrlCache?: Map<string, string>;
  assetUrlPending?: Set<string>;
  assetUrlErrors?: Map<string, string>;
}

/** Build the render-engine floating scene without coupling it to SheetCanvas state. */
export function createCanvasFloatingDrawables(input: CanvasFloatingRendererInput): FloatingDrawable[] {
  const { allSheets, drawingPayloads, drawings, imageCache, pivotResults, requestRender, sheet, skeleton, sparklines, tables, resolveAssetUrl, assetUrlCache, assetUrlPending, assetUrlErrors } = input;
  const drawables: FloatingDrawable[] = [];
  const sheets = allSheets.length > 0 ? allSheets : [sheet];
  const getSheet = (sheetId: string): CanvasSheetSnapshot | undefined =>
    sheets.find((candidate) => candidate.id === sheetId) ?? (sheet.id === sheetId ? sheet : undefined);
  for (const drawing of [...drawings].sort((left, right) => left.zIndex - right.zIndex)) {
    if (drawing.visible === false) continue;
    const payload = drawingPayloads.get(drawing.payloadId);
    if (!payload) continue;
    const bounds = drawing.transform;
    if (payload.kind === "chart") {
      const data = getChartSeries(payload, getSheet, pivotResults, sheets);
      if (data.brokenReference) {
        drawables.push({
          kind: 'shape',
          id: drawing.id,
          bounds,
          draw: (context, rect) => {
            context.save();
            context.fillStyle = '#b91c1c';
            context.strokeStyle = '#b91c1c';
            context.strokeRect(rect.x, rect.y, rect.width, rect.height);
            context.font = '12px Segoe UI, sans-serif';
            context.fillText(data.brokenReference!, rect.x + 8, rect.y + Math.min(rect.height / 2, 24), Math.max(10, rect.width - 16));
            context.restore();
          },
        });
        continue;
      }
      const series = data.series.map((entry, index) => ({ ...entry, color: entry.color ?? CHART_PALETTE[index % CHART_PALETTE.length]! }));
      drawables.push({
        kind: "chart",
        id: drawing.id,
        bounds,
        draw: (context, rect) => drawCanonicalChartOnCanvas({ context, payload, bounds: rect, categories: data.categories, series }),
      });
      continue;
    }
    if (payload.kind === 'data-chart') {
      const data = dataChartSeries(payload, tables, getSheet);
      const chartPayload: ChartDrawingPayload = {
        kind: 'chart', chartId: drawing.payloadId, chartType: payload.plotType === 'radar' || payload.plotType === 'treemap' || payload.plotType === 'funnel' ? 'column' : payload.plotType,
        sourceRanges: [], elements: { title: payload.inspector.title, legend: { visible: payload.inspector.legendPosition !== 'none', position: payload.inspector.legendPosition === 'none' ? 'bottom' : payload.inspector.legendPosition }, dataLabels: { visible: payload.inspector.showDataLabels }, hiddenData: 'show', chartArea: payload.inspector.chartArea, plotArea: payload.inspector.plotArea, valueAxis: { id: 'value', position: 'left', title: payload.inspector.axis.valueTitle, majorGridlines: { visible: payload.inspector.axis.showGridlines, color: '#e2e8f0', width: 1, dash: 'solid' } }, categoryAxis: { id: 'category', position: 'bottom', title: payload.inspector.axis.categoryTitle, majorGridlines: { visible: false } } },
      };
      drawables.push({ kind: 'chart', id: drawing.id, bounds, draw: (context, rect) => drawCanonicalChartOnCanvas({ context, payload: chartPayload, bounds: rect, categories: data.categories, series: data.series }) });
      continue;
    }
    if (payload.kind === 'camera') {
      drawables.push({ kind: 'shape', id: drawing.id, bounds, draw: (context, rect) => drawCameraOnCanvas(context, payload, rect, getSheet) });
      continue;
    }
    if (payload.kind === 'form-control') {
      drawables.push({ kind: 'shape', id: drawing.id, bounds, draw: (context, rect) => drawFormControlOnCanvas(context, payload, rect) });
      continue;
    }
    if (payload.kind === 'connector') {
      drawables.push({
        // Connectors remain shape-layer drawables so the existing PaneMap
        // clipping and floating selection chrome own their coordinates.
        kind: 'shape',
        id: drawing.id,
        bounds,
        draw: (context, rect) => drawCanonicalConnectorOnCanvas(context, payload, rect),
        hitTest: (point) => connectorEndpointHitTest(payload, bounds, point),
      });
      continue;
    }
    if (payload.kind === "shape" || payload.kind === "textbox") {
      drawables.push({
        kind: "shape",
        id: drawing.id,
        bounds,
        draw: (context, rect) => drawCanonicalShapeOnCanvas({ context, payload, bounds: rect, rotation: drawing.transform.rotation }),
      });
      continue;
    }
    if (payload.kind === "slicer" || payload.kind === "timeline") {
      const members = pivotControlMembers(drawing.id, payload, pivotResults);
      const periods = payload.kind === 'timeline' ? pivotTimelinePeriods(payload, pivotResults) : [];
      drawables.push({
        kind: "pivot-control",
        id: drawing.id,
        bounds,
        draw: (context, rect) => drawPivotControlOnCanvas({ context, payload, bounds: rect, members, periods }),
        hitTest: (point) => pivotControlHitTest(payload, members, periods, point, bounds),
      });
      continue;
    }
    if (payload.kind === "image") {
      drawables.push({
        kind: "image",
        id: drawing.id,
        bounds,
        draw: (context, rect) => {
          const assetId = payload.asset.assetId;
          const assetUrl = assetUrlCache?.get(assetId);
          if (!assetUrl) {
            if (resolveAssetUrl && assetUrlPending && !assetUrlPending.has(assetId) && !assetUrlErrors?.has(assetId)) {
              assetUrlPending.add(assetId);
              void resolveAssetUrl(payload.asset)
                .then((url) => assetUrlCache?.set(assetId, url))
                .catch((error) => assetUrlErrors?.set(assetId, error instanceof Error ? error.message : `ASSET_RESOLVE_FAILED: ${assetId}`))
                .finally(() => {
                  assetUrlPending.delete(assetId);
                  requestRender();
                });
            }
            context.fillStyle = assetUrlErrors?.has(assetId) ? "#fee2e2" : "#f1f5f9";
            context.fillRect(rect.x, rect.y, rect.width, rect.height);
            context.strokeStyle = assetUrlErrors?.has(assetId) ? "#dc2626" : "#94a3b8";
            context.strokeRect(rect.x + 0.5, rect.y + 0.5, rect.width - 1, rect.height - 1);
            return;
          }
          let img = imageCache.get(assetId);
          if (!img) {
            img = new Image();
            img.src = assetUrl;
            imageCache.set(assetId, img);
            img.onload = requestRender;
          }
          if (img.complete && img.naturalWidth > 0) {
            const crop = payload.crop ?? { left: 0, top: 0, right: 0, bottom: 0 };
            const sourceX = img.naturalWidth * crop.left;
            const sourceY = img.naturalHeight * crop.top;
            const sourceWidth = img.naturalWidth * (1 - crop.left - crop.right);
            const sourceHeight = img.naturalHeight * (1 - crop.top - crop.bottom);
            const effects = payload.effects;
            context.save();
            context.globalAlpha = 1 - (effects?.transparency ?? 0);
            context.filter = `brightness(${1 + (effects?.brightness ?? 0)}) contrast(${1 + (effects?.contrast ?? 0)})`;
            context.drawImage(img, sourceX, sourceY, sourceWidth, sourceHeight, rect.x, rect.y, rect.width, rect.height);
            context.restore();
            return;
          }
          context.fillStyle = "#f1f5f9";
          context.fillRect(rect.x, rect.y, rect.width, rect.height);
          context.strokeStyle = "#94a3b8";
          context.strokeRect(rect.x + 0.5, rect.y + 0.5, rect.width - 1, rect.height - 1);
        },
      });
    }
  }
  const getSparklineValues = (sparkline: SparklineModel): number[] => {
    const values: number[] = [];
    const source = sparkline.sourceRange;
    const sourceSheet = (allSheets.find((candidate) => candidate.id === source.sheetId) ?? sheet);
    for (let row = source.startRow; row <= source.endRow; row += 1) {
      for (let column = source.startColumn; column <= source.endColumn; column += 1) {
        const cell = sourceSheet.getCell(row, column);
        if (!cell) continue;
        const numeric = Number(cell.value.replace(/[$,%]/g, ""));
        if (Number.isFinite(numeric) && cell.value !== "") values.push(numeric);
      }
    }
    return values;
  };
  for (const sparkline of sparklines) {
    const rect = skeleton.getCellRect(sparkline.anchor.row, sparkline.anchor.column);
    if (!rect) continue;
    drawables.push({
      kind: "shape",
      id: sparkline.id,
      bounds: rect,
      draw: (context, target) => drawCanonicalSparklineOnCanvas({ context, sparkline, values: getSparklineValues(sparkline), rect: target }),
    });
  }
  return drawables;
}
