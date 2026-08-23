import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
  Box,
  ContextMenu,
  type ContextMenuItem,
  Panel,
  Stack,
  StatePanel,
  Text,
  Button,
  Inline,
} from "@react-sheets/ui-system";
import {
  CanvasRenderSurface,
  CanvasRenderEngine,
  SheetSkeleton,
  type CellRenderData,
  type ChromeState,
  type FloatingDrawable,
  type FloatingHit,
  type HeaderHit,
  type Rect,
  createEmptyChromeState,
} from "@react-sheets/render-engine";
import type {
  ChartDrawingPayload,
  DrawingObject,
  DrawingPayload,
  PivotGridProjection,
  PivotHitTest,
  PivotProjectionCell,
  PivotSlicerDrawingPayload,
  PivotSourceRowPath,
  PivotTimelineDrawingPayload,
  PivotResultTree,
  RangeRef,
  SparklineModel,
} from "@react-sheets/core-model";
import { CellEditor } from "./CellEditor";
import { FilterPopover } from "./FilterPopover";
import { createSpreadsheetShortcutRegistry, resolveContextHit, type PeerCursor, type ResolvedContextHit, type SelectionState, type CanvasSheetSnapshot, type AppPhase } from "@react-sheets/spreadsheet-app";
import type { CanvasCellSnapshot } from "@react-sheets/spreadsheet-app";
import type { CommandDescriptor } from "@react-sheets/command-runtime";

const CHART_PALETTE = ["#2563eb", "#10b981", "#f59e0b", "#ef4444", "#8b5cf6", "#ec4899", "#06b6d4"];

interface CanvasChartSeries {
  name: string;
  values: number[];
  color?: string;
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
): { categories: string[]; series: CanvasChartSeries[] } {
  const categories: string[] = [];
  const series: CanvasChartSeries[] = [];
  const pivot = payload.pivotId
    ? pivotResults[payload.pivotId] ?? sheets.map((candidate) => candidate.pivotResults[payload.pivotId!]).find(Boolean)
    : undefined;
  if (pivot) {
    const leaves: typeof pivot.rows = [];
    const collect = (nodes: typeof pivot.rows): void => {
      for (const node of nodes) {
        if (node.children.length > 0) collect(node.children);
        else leaves.push(node);
      }
    };
    collect(pivot.rows);
    const valueCount = pivot.rows[0]?.values[0]?.values.length ?? 0;
    for (let index = 0; index < valueCount; index += 1) {
      series.push({ name: payload.series?.[index]?.name ?? payload.title ?? `Value ${index + 1}`, values: [] });
    }
    for (const node of leaves) {
      categories.push(node.label);
      const cell = node.values[0];
      for (let index = 0; index < valueCount; index += 1) {
        const numeric = Number(cell?.values[index]);
        series[index]?.values.push(Number.isFinite(numeric) ? numeric : 0);
      }
    }
    return { categories, series };
  }

  const readRange = (range: RangeRef): string[][] => {
    const sourceSheet = getSheet(range.sheetId);
    if (!sourceSheet) return [];
    const rows: string[][] = [];
    for (let row = range.startRow; row <= range.endRow; row += 1) {
      const values: string[] = [];
      for (let column = range.startColumn; column <= range.endColumn; column += 1) {
        values.push(sourceSheet.getCell(row, column)?.value ?? "");
      }
      rows.push(values);
    }
    return rows;
  };

  if (payload.series && payload.series.length > 0) {
    for (const entry of payload.series) {
      const values = readRange(entry.range).flat().map(numericCellValue).filter((value): value is number => value !== undefined);
      series.push({ name: entry.name, values, color: entry.color });
    }
  }
  const source = payload.sourceRanges[0];
  if (!source) return { categories, series };
  const matrix = readRange(source);
  const categoryMatrix = payload.categoryRange ? readRange(payload.categoryRange) : [];
  if (categoryMatrix.length > 0) {
    categories.push(...categoryMatrix.flat().filter((value) => value !== ""));
  }

  // A canonical source range with a header row and a label column becomes a
  // multi-series chart without constructing a second chart model.
  if (series.length === 0 && matrix.length > 1 && (matrix[0]?.length ?? 0) > 1) {
    const width = matrix[0]?.length ?? 0;
    if (categories.length === 0) {
      for (let row = 1; row < matrix.length; row += 1) categories.push(matrix[row]?.[0] ?? String(row));
    }
    for (let column = 1; column < width; column += 1) {
      const values = matrix.slice(1).map((row) => numericCellValue(row?.[column] ?? "") ?? 0);
      series.push({ name: matrix[0]?.[column] || payload.title || `Series ${column}`, values });
    }
    return { categories, series };
  }

  if (series.length === 0) {
    const values = matrix.flat().map(numericCellValue).filter((value): value is number => value !== undefined);
    if (categories.length === 0) {
      categories.push(...matrix.flat().filter((value) => numericCellValue(value) === undefined && value !== ""));
    }
    if (values.length > 0) series.push({ name: payload.title || "Series 1", values });
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
    context.fillStyle = payload.textColor ?? "#1e293b";
    context.font = `${payload.fontSize ?? 13}px Inter, sans-serif`;
    context.textAlign = "center";
    context.textBaseline = "middle";
    context.fillText(payload.text, width / 2, height / 2, Math.max(10, width - 8));
    context.restore();
    return;
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
    context.fillStyle = payload.textColor ?? "#1e293b";
    context.font = `${payload.fontSize ?? 13}px Inter, sans-serif`;
    context.textAlign = "center";
    context.textBaseline = "middle";
    context.fillText(payload.text, width / 2, height / 2, Math.max(10, width - 8));
  }
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
  const { x, y, width, height } = bounds;
  context.save();
  context.translate(x, y);
  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, width, height);
  context.strokeStyle = "#e2e8f0";
  context.lineWidth = 1;
  context.strokeRect(0, 0, width, height);
  if (payload.title) {
    context.fillStyle = "#1e293b";
    context.font = "bold 14px Segoe UI, sans-serif";
    context.textAlign = "left";
    context.textBaseline = "top";
    context.fillText(payload.title, 16, 12);
  }
  if (payload.chartType === "pie" || payload.chartType === "doughnut") {
    drawCanonicalPieChart(context, series[0], width, height, payload.chartType === "doughnut");
    drawCanonicalLegend(context, series, width, height, payload.legendPosition ?? "bottom");
    context.restore();
    return;
  }
  const plotTop = payload.title ? 40 : 20;
  const plotBottom = height - (payload.legendPosition === "bottom" ? 40 : 24);
  const plotLeft = 48;
  const plotRight = width - (payload.legendPosition === "right" ? 90 : 16);
  const plotWidth = Math.max(10, plotRight - plotLeft);
  const plotHeight = Math.max(10, plotBottom - plotTop);
  const allValues = series.flatMap((entry) => entry.values);
  const maxValue = Math.max(1, ...allValues.map((value) => Math.abs(value)));
  const minValue = Math.min(0, ...allValues);
  const maxAxis = payload.stacked === "percent" ? 100 : Math.max(1, maxValue * 1.1);
  const minAxis = payload.stacked === "percent" ? 0 : Math.min(0, minValue);
  const axisSpan = Math.max(1, maxAxis - minAxis);
  const zeroY = plotTop + (maxAxis / axisSpan) * plotHeight;
  context.strokeStyle = "#f1f5f9";
  context.fillStyle = "#64748b";
  context.font = "11px Segoe UI, sans-serif";
  context.textAlign = "right";
  context.textBaseline = "middle";
  for (let index = 0; index <= 4; index += 1) {
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
          if (payload.showDataLabels) {
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
  drawCanonicalLegend(context, series, width, height, payload.legendPosition ?? "top");
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
  context.fillStyle = "#ffffff";
  context.strokeStyle = color;
  context.lineWidth = 2;
  for (const point of points) {
    context.beginPath();
    context.arc(point.x, point.y, 4, 0, Math.PI * 2);
    context.fill();
    context.stroke();
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
  position: NonNullable<ChartDrawingPayload["legendPosition"]>,
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

function drawPivotControlOnCanvas(options: {
  context: CanvasRenderingContext2D;
  payload: PivotSlicerDrawingPayload | PivotTimelineDrawingPayload;
  bounds: Rect;
}): void {
  const { context, payload, bounds } = options;
  const style = payload.style;
  context.save();
  context.fillStyle = style.fill;
  context.strokeStyle = style.border;
  context.lineWidth = 1;
  context.fillRect(bounds.x, bounds.y, bounds.width, bounds.height);
  context.strokeRect(bounds.x, bounds.y, bounds.width, bounds.height);
  context.fillStyle = style.accentColor;
  context.fillRect(bounds.x, bounds.y, bounds.width, Math.min(26, bounds.height));
  context.fillStyle = style.textColor;
  context.font = `600 ${style.fontSize}px Segoe UI, sans-serif`;
  context.textBaseline = 'middle';
  context.fillText(payload.kind === 'slicer' ? `Slicer · ${payload.fieldId}` : `Timeline · ${payload.fieldId}`, bounds.x + 8, bounds.y + Math.min(13, bounds.height / 2), Math.max(10, bounds.width - 16));
  context.font = `${style.fontSize}px Segoe UI, sans-serif`;
  const detail = payload.kind === 'slicer'
    ? payload.filter.mode === 'all' ? 'All items' : `${payload.filter.memberKeys.length} selected`
    : `${payload.period.start ?? 'Start'} — ${payload.period.end ?? 'End'}`;
  context.fillText(detail, bounds.x + 8, bounds.y + Math.min(bounds.height - 12, 44), Math.max(10, bounds.width - 16));
  context.restore();
}

export interface SheetCanvasProps {
  sheet: CanvasSheetSnapshot;
  sheetId: string;
  selection: SelectionState;
  activeCell: string;
  formulaDraft: string;
  editingCell: { row: number; column: number } | null;
  phase: AppPhase;
  zoom: number;
  peers: PeerCursor[];
  cellStyle?: {
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
  };
  drawings?: readonly DrawingObject[];
  drawingPayloads?: ReadonlyMap<string, DrawingPayload>;
  allSheets?: readonly CanvasSheetSnapshot[];
  pivotResults?: Record<string, PivotResultTree>;
  sparklines?: SparklineModel[];
  selectedFloatingId: string | null;
  showFormulas?: boolean;
  /** Notifies the host when a visible Pivot projection becomes/leaves the active context. */
  onPivotContextHit?: (hit: ResolvedContextHit | null) => void;
  /** Lets the host add/replace Pivot-specific right-click commands. */
  getPivotContextMenuItems?: (hit: ResolvedContextHit) => readonly ContextMenuItem[];
  /** Opens a real details-sheet flow for a Pivot value/double-click or menu action. */
  onPivotShowDetails?: (request: PivotShowDetailsRequest) => void;
  onSelectionChange: (selection: SelectionState) => void;
  onMovePrimary: (rowDelta: number, columnDelta: number, opts?: { extend?: boolean }) => void;
  onCommitCell: (value: string) => void;
  onBeginEdit: (initialText?: string) => void;
  onCancelEdit: () => void;
  onCommitEdit: (moveAfter?: "down" | "up" | "left" | "right" | "none") => void;
  onFormulaDraftChange: (value: string) => void;
  onAppendFormulaDraft?: (fragment: string) => void;
  onInsertRef: (refText: string) => void;
  onToggleAbsolute: () => void;
  onJumpEdge: (direction: "up" | "down" | "left" | "right", extend?: boolean) => void;
  onSelectAll: () => void;
  onExtendSelection?: (row: number, column: number) => void;
  onResizeRow: (row: number, heightPx: number) => void;
  onResizeColumn: (column: number, widthPx: number) => void;
  onFillRange: (target: { startRow: number; endRow: number; startColumn: number; endColumn: number }) => void;
  onFloatingSelect: (hit: FloatingHit | null) => void;
  onFloatingMove: (drawingId: string, bounds: Rect, rotation?: number) => void;
  onFloatingRemove: (drawingId: string) => void;
  onCommand: (descriptor: CommandDescriptor) => void;
  onClearSelection?: (mode: "contents" | "formats") => void;
  onCopy: () => void;
  onCut: () => void;
  onPaste: () => void;
  onUndo: () => void;
  onRedo: () => void;
  /** Host owns command/session execution after the shared registry resolves a shortcut. */
  onShortcut?: (id: string) => boolean;
  canRepeat?: boolean;
  onOpenInspector: () => void;
  onApplyFilter: (column: number, patch: { selectedValues?: string[] | null }) => void;
  onToggleOutline?: (groupId: string) => void;
  getValidationList: (row: number, column: number) => string[] | undefined;
  onRetry: () => void;
  onCreateSheet: () => void;
}

interface DragState {
  kind: "select" | "fill" | "col-resize" | "row-resize" | "floating-move" | "floating-resize";
  startRow: number;
  startColumn: number;
  anchorRow: number;
  anchorColumn: number;
  currentRow: number;
  currentColumn: number;
  additive: boolean;
  extend: boolean;
  resizeStartSize: number;
  resizeIndex: number;
    floating?: { id: string; kind: 'chart' | 'shape' | 'image'; handle?: string; rotation?: number; startBounds: Rect; startLocal: { x: number; y: number } };
}

function mergeAtCell(sheet: CanvasSheetSnapshot, cell: { row: number; column: number }) {
  return sheet.merges.find((merge) =>
    cell.row >= merge.range.startRow && cell.row <= merge.range.endRow
    && cell.column >= merge.range.startColumn && cell.column <= merge.range.endColumn);
}

function resolveMergedCell(sheet: CanvasSheetSnapshot, cell: { row: number; column: number }): { row: number; column: number } {
  const merge = mergeAtCell(sheet, cell);
  return merge ? { ...merge.anchor } : { ...cell };
}

function intersectsRange(
  first: { startRow: number; endRow: number; startColumn: number; endColumn: number },
  second: { startRow: number; endRow: number; startColumn: number; endColumn: number },
): boolean {
  return first.startRow <= second.endRow && second.startRow <= first.endRow
    && first.startColumn <= second.endColumn && second.startColumn <= first.endColumn;
}

function containsRange(
  outer: { startRow: number; endRow: number; startColumn: number; endColumn: number },
  inner: { startRow: number; endRow: number; startColumn: number; endColumn: number },
): boolean {
  return outer.startRow <= inner.startRow && outer.endRow >= inner.endRow
    && outer.startColumn <= inner.startColumn && outer.endColumn >= inner.endColumn;
}

function expandRangeForMerges(sheet: CanvasSheetSnapshot, range: RangeRef): RangeRef {
  let expanded = { ...range };
  let changed = true;
  while (changed) {
    changed = false;
    for (const merge of sheet.merges) {
      if (!intersectsRange(expanded, merge.range)) continue;
      const next = {
        startRow: Math.min(expanded.startRow, merge.range.startRow),
        endRow: Math.max(expanded.endRow, merge.range.endRow),
        startColumn: Math.min(expanded.startColumn, merge.range.startColumn),
        endColumn: Math.max(expanded.endColumn, merge.range.endColumn),
      };
      changed = next.startRow !== expanded.startRow || next.endRow !== expanded.endRow
        || next.startColumn !== expanded.startColumn || next.endColumn !== expanded.endColumn;
      expanded = { ...expanded, ...next };
    }
  }
  return expanded;
}

function resolveDragCell(
  engine: CanvasRenderEngine,
  sheet: CanvasSheetSnapshot,
  local: { x: number; y: number },
  drag: DragState,
): { row: number; column: number } | null {
  const headerHit = engine.headerHitAtLocal(local);
  const hitCell = engine.cellAtLocalPoint(local);
  const isRowDrag = drag.floating?.id === "row";
  const isColDrag = drag.floating?.id === "col";
  if (isRowDrag) {
    if (headerHit?.kind === "row") return { row: headerHit.index, column: 0 };
    return hitCell ? { row: hitCell.row, column: 0 } : null;
  }
  if (isColDrag) {
    if (headerHit?.kind === "col") return { row: 0, column: headerHit.index };
    return hitCell ? { row: 0, column: hitCell.column } : null;
  }
  return hitCell ? resolveMergedCell(sheet, hitCell) : null;
}

function toChromeSelection(selection: SelectionState): ChromeState['selection'] {
  return {
    ranges: selection.ranges.map((range) => ({
      startRow: range.startRow,
      endRow: range.endRow,
      startColumn: range.startColumn,
      endColumn: range.endColumn,
    })),
    primary: { row: selection.activeCell.row, column: selection.activeCell.column },
    primaryIndex: selection.primaryRangeIndex,
  };
}

/**
 * A projection stores cell coordinates relative to its target anchor. The
 * canvas, on the other hand, always asks for worksheet coordinates. Keeping
 * this conversion at the UI boundary prevents the derived projection from
 * leaking into ordinary worksheet cells or selection state.
 */
export function findPivotProjectionCell(
  sheet: CanvasSheetSnapshot,
  row: number,
  column: number,
): { projection: PivotGridProjection; cell: PivotProjectionCell } | null {
  for (const projection of Object.values(sheet.pivotProjections)) {
    if (projection.sheetId !== sheet.id || projection.collision.status !== "clear") continue;
    const relativeRow = row - projection.target.anchor.row;
    const relativeColumn = column - projection.target.anchor.column;
    const cell = projection.cells.find((candidate) => candidate.row === relativeRow && candidate.column === relativeColumn);
    if (cell) return { projection, cell };
  }
  return null;
}

function pivotHitKind(cell: PivotProjectionCell): PivotHitTest['kind'] {
  if (cell.kind === "expand-toggle") return "expand-toggle";
  if (cell.kind === "filter") return "filter";
  if (cell.kind === "title" || cell.kind === "column-header" || cell.kind === "row-header") return "header";
  return "cell";
}

/** Resolve an absolute worksheet coordinate through the canonical context resolver. */
export function resolvePivotProjectionHit(
  sheet: CanvasSheetSnapshot,
  row: number,
  column: number,
): ResolvedContextHit | null {
  const target = findPivotProjectionCell(sheet, row, column);
  if (!target) return null;
  const hit: PivotHitTest = {
    kind: pivotHitKind(target.cell),
    pivotId: target.projection.pivotId,
    cellId: target.cell.id,
    row,
    column,
    nodeId: target.cell.nodeId,
    sourceRowPaths: target.cell.sourceRowPaths,
  };
  return resolveContextHit({ sheetId: sheet.id, pivot: hit });
}

export function isPivotValueCell(cell: PivotProjectionCell): boolean {
  return cell.kind === "value" || cell.kind === "subtotal" || cell.kind === "grand-total";
}

/** Convert a derived projection cell to the render-engine cell contract. */
export function pivotProjectionCellRenderData(cell: PivotProjectionCell): CellRenderData {
  const text = cell.kind === "expand-toggle"
    ? `${cell.expanded ? "▾" : "▸"} ${cell.text}`
    : cell.text;
  const style: NonNullable<CellRenderData["style"]> = {
    background: cell.kind === "title"
      ? "#dbeafe"
      : cell.kind === "column-header" || cell.kind === "filter"
        ? "#f1f5f9"
        : cell.kind === "grand-total"
          ? "#eff6ff"
          : cell.kind === "subtotal"
            ? "#f8fafc"
            : "#ffffff",
    textColor: cell.kind === "error" ? "#b91c1c" : cell.kind === "loading" ? "#92400e" : "#1e293b",
    bold: cell.kind === "title" || cell.kind === "column-header" || cell.kind === "subtotal" || cell.kind === "grand-total",
    italic: cell.kind === "filter",
    horizontalAlignment: isPivotValueCell(cell) ? "right" : "left",
    verticalAlignment: "middle",
    wrapText: cell.kind === "loading" || cell.kind === "error",
    borders: {
      bottom: { color: "#cbd5e1", style: cell.kind === "grand-total" ? "double" : "thin" },
    },
  };
  return {
    value: cell.value,
    displayValue: text,
    style,
    error: cell.kind === "error" ? text : undefined,
    invalid: cell.kind === "error",
  };
}

export interface PivotShowDetailsRequest {
  pivotId: string;
  sourceRowPaths: readonly PivotSourceRowPath[];
  hit: ResolvedContextHit;
}

export function SheetCanvas({
  sheet,
  sheetId,
  selection,
  activeCell,
  formulaDraft,
  editingCell,
  phase,
  zoom,
  peers,
  cellStyle = {},
  drawings = sheet.drawings,
  drawingPayloads = sheet.drawingPayloads,
  allSheets = [],
  pivotResults = {},
  sparklines = [],
  selectedFloatingId,
  showFormulas = false,
  onPivotContextHit,
  getPivotContextMenuItems,
  onPivotShowDetails,
  onSelectionChange,
  onMovePrimary,
  onCommitCell,
  onBeginEdit,
  onCancelEdit,
  onCommitEdit,
  onFormulaDraftChange,
  onAppendFormulaDraft,
  onInsertRef,
  onToggleAbsolute,
  onJumpEdge,
  onSelectAll,
  onExtendSelection,
  onResizeRow,
  onResizeColumn,
  onFillRange,
  onFloatingSelect,
  onFloatingMove,
  onFloatingRemove,
  onCommand,
  onClearSelection,
  onCopy,
  onCut,
  onPaste,
  onUndo,
  onRedo,
  onShortcut,
  canRepeat = false,
  onOpenInspector,
  onApplyFilter,
  onToggleOutline,
  getValidationList,
  onRetry,
  onCreateSheet,
}: SheetCanvasProps) {
  const engineRef = useRef<CanvasRenderEngine | null>(null);
  const shortcutRegistryRef = useRef<ReturnType<typeof createSpreadsheetShortcutRegistry> | null>(null);
  if (!shortcutRegistryRef.current) shortcutRegistryRef.current = createSpreadsheetShortcutRegistry();
  const imageCacheRef = useRef(new Map<string, HTMLImageElement>());
  const containerRef = useRef<HTMLDivElement | null>(null);
  const dragRef = useRef<DragState | null>(null);
  const contextRangeRef = useRef<RangeRef | null>(null);
  const [contextMenu, setContextMenu] = useState({ x: 0, y: 0, open: false });
  const [contextHit, setContextHit] = useState<ResolvedContextHit | null>(null);
  const [filterPopover, setFilterPopover] = useState<{ column: number; x: number; y: number } | null>(null);
  const [validationDropdown, setValidationDropdown] = useState<{ row: number; column: number; options: string[] } | null>(null);
  const [fillPreview, setFillPreview] = useState<{ startRow: number; endRow: number; startColumn: number; endColumn: number } | null>(null);
  const [scrollTick, setScrollTick] = useState(0);
  const [engineReady, setEngineReady] = useState(false);
  const transientSelectionRef = useRef<SelectionState | null>(null);
  const editingActiveRef = useRef(false);
  const autoScrollFrameRef = useRef<number | null>(null);
  const autoScrollPointRef = useRef<{ x: number; y: number } | null>(null);

  useEffect(() => {
    editingActiveRef.current = Boolean(editingCell);
  }, [editingCell]);
  const transientSelectionFrameRef = useRef<number | null>(null);

  const zoomFactor = zoom / 100;

  const skeleton = useMemo(
    () =>
      new SheetSkeleton({
        rowCount: Math.max(sheet.rowCount, 200),
        columnCount: Math.max(sheet.columnCount, 26),
        defaultRowHeight: 28,
        defaultColumnWidth: 110,
        rowHeights: new Map(Object.entries(sheet.rowHeights).map(([key, value]) => [Number(key), value])),
        columnWidths: new Map(Object.entries(sheet.columnWidths).map(([key, value]) => [Number(key), value])),
        hiddenRows: new Set(sheet.hiddenRows),
        hiddenColumns: new Set(sheet.hiddenColumns ?? []),
        zoom: zoomFactor,
      }),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [sheet.rowCount, sheet.columnCount, sheet.rowHeights, sheet.columnWidths, sheet.hiddenRows, sheet.hiddenColumns, zoomFactor],
  );

  const cellProvider = useCallback(({ row, column }: { row: number; column: number }): CellRenderData | undefined => {
    const pivotCell = findPivotProjectionCell(sheet, row, column);
    if (pivotCell) return pivotProjectionCellRenderData(pivotCell.cell);

    const cell = sheet.getCell(row, column);
    const merge = sheet.merges.find((span) =>
      row >= span.range.startRow && row <= span.range.endRow
      && column >= span.range.startColumn && column <= span.range.endColumn);
    // Empty cells inside a merge still need a render record so the grid layer
    // can suppress the merge's internal boundaries. Returning undefined here
    // made blank merged areas look like ordinary cells.
    if (!cell) {
      if (!merge) return undefined;
      return {
        value: undefined,
        merge: {
          startRow: merge.range.startRow,
          endRow: merge.range.endRow,
          startColumn: merge.range.startColumn,
          endColumn: merge.range.endColumn,
          isAnchor: merge.anchor.row === row && merge.anchor.column === column,
        },
      };
    }
    const isAnchor = merge ? merge.anchor.row === row && merge.anchor.column === column : true;
    return {
      value: showFormulas && cell.formula ? cell.formula : parseCellValue(cell),
      formula: cell.formula,
      displayValue: cell.value,
      style: cell.style,
      overlay: cell.overlay
        ? {
            dataBar: cell.overlay.dataBar,
            colorScale: cell.overlay.colorScale,
            icon: cell.overlay.icon,
          }
        : undefined,
      hasComment: cell.hasComment,
      invalid: cell.invalid,
      merge: merge
        ? {
            startRow: merge.range.startRow,
            endRow: merge.range.endRow,
            startColumn: merge.range.startColumn,
            endColumn: merge.range.endColumn,
            isAnchor,
          }
        : undefined,
    };
  }, [sheet, showFormulas]);

  const pivotStatusProjections = useMemo(
    () => Object.values(sheet.pivotProjections).filter((projection) =>
      projection.collision.status === "collision"
      || projection.refresh.status === "error"
      || projection.refresh.status === "refreshing"
      || projection.refresh.status === "stale"),
    [sheet.pivotProjections],
  );

  // ---------- 浮动对象绘制器 ----------

  const floatables = useMemo<FloatingDrawable[]>(() => {
    const drawables: FloatingDrawable[] = [];
    const sheets = allSheets.length > 0 ? allSheets : [sheet];
    const getSheet = (sheetId: string): CanvasSheetSnapshot | undefined =>
      sheets.find((candidate) => candidate.id === sheetId) ?? (sheet.id === sheetId ? sheet : undefined);
    const canonicalPayloads = drawingPayloads;
    for (const drawing of [...drawings].sort((left, right) => left.zIndex - right.zIndex)) {
      const payload = canonicalPayloads.get(drawing.payloadId);
      if (!payload) continue;
      const bounds = drawing.transform;
      if (payload.kind === "chart") {
        const data = getChartSeries(payload, getSheet, pivotResults, sheets);
        const series = data.series.map((entry, index) => ({
          ...entry,
          color: entry.color ?? CHART_PALETTE[index % CHART_PALETTE.length]!,
        }));
        drawables.push({
          kind: "chart",
          id: drawing.id,
          bounds,
          draw: (context, rect) => drawCanonicalChartOnCanvas({ context, payload, bounds: rect, categories: data.categories, series }),
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
        drawables.push({
          kind: "shape",
          id: drawing.id,
          bounds,
          draw: (context, rect) => drawPivotControlOnCanvas({ context, payload, bounds: rect }),
        });
        continue;
      }
      if (payload.kind === "image") {
        drawables.push({
          kind: "image",
          id: drawing.id,
          bounds,
          draw: (context, rect) => {
            let img = imageCacheRef.current.get(payload.src);
            if (!img) {
              img = new Image();
              img.src = payload.src;
              imageCacheRef.current.set(payload.src, img);
              img.onload = () => engineRef.current?.requestRender();
            }
            if (img.complete && img.naturalWidth > 0) {
              context.drawImage(img, rect.x, rect.y, rect.width, rect.height);
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

    for (const sparkline of sparklines) {
      const rect = skeleton.getCellRect(sparkline.anchor.row, sparkline.anchor.column);
      if (!rect) continue;
      drawables.push({
        kind: "shape",
        id: sparkline.id,
        bounds: rect,
        draw: (context, target) =>
          drawCanonicalSparklineOnCanvas({ context, sparkline, values: getSparklineValues(sparkline), rect: target }),
      });
    }
    return drawables;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [allSheets, drawingPayloads, drawings, pivotResults, sparklines, skeleton, sheet, sheetId]);

  function getSparklineValues(sparkline: SparklineModel): number[] {
    const values: number[] = [];
    const source = sparkline.sourceRange;
    const sourceSheet = (allSheets.find((candidate) => candidate.id === source.sheetId) ?? sheet);
    for (let row = source.startRow; row <= source.endRow; row++) {
      for (let column = source.startColumn; column <= source.endColumn; column++) {
        const cell = sourceSheet.getCell(row, column);
        if (!cell) continue;
        const numeric = Number(cell.value.replace(/[$,%]/g, ""));
        if (Number.isFinite(numeric) && cell.value !== "") values.push(numeric);
      }
    }
    return values;
  }

  // ---------- 引擎生命周期与 chrome 同步 ----------

  const chromeState = useMemo<ChromeState>(() => {
    const state = createEmptyChromeState();
    state.selection = toChromeSelection(selection);
    state.editing = editingCell ? { row: editingCell.row, column: editingCell.column } : null;
    state.filterColumns = sheet.filterColumns;
    state.filterButtons = sheet.filterButtons;
    state.tableOutlines = sheet.sheetTables.map((table) => ({
      startRow: table.range.startRow,
      endRow: table.range.endRow,
      startColumn: table.range.startColumn,
      endColumn: table.range.endColumn,
    }));
    state.outlineControls = sheet.outlineControls;
    state.remoteCursors = peers.map((peer) => ({
      actorId: peer.actorId,
      color: peer.color,
      name: peer.name,
      row: peer.row,
      column: peer.column,
    }));
    state.selectedFloatingId = selectedFloatingId;
    return state;
  }, [editingCell, peers, selectedFloatingId, selection, sheet.filterButtons, sheet.filterColumns, sheet.outlineControls, sheet.sheetTables]);

  useEffect(() => {
    const engine = engineRef.current;
    if (!engine) return;
    engine.setCellProvider(cellProvider);
  }, [cellProvider]);

  useEffect(() => {
    const engine = engineRef.current;
    if (!engine) return;
    engine.setSkeleton(skeleton);
  }, [skeleton]);

  useEffect(() => {
    const engine = engineRef.current;
    if (!engine) return;
    engine.setFreeze(
      sheet.freeze.xSplit > 0 || sheet.freeze.ySplit > 0
        ? { xSplit: sheet.freeze.xSplit, ySplit: sheet.freeze.ySplit }
        : null,
    );
  }, [sheet.freeze.xSplit, sheet.freeze.ySplit]);

  useEffect(() => {
    const engine = engineRef.current;
    if (!engine) return;
    engine.setFloating(floatables, selectedFloatingId);
  }, [floatables, selectedFloatingId]);

  useEffect(() => {
    const engine = engineRef.current;
    if (!engine) return;
    engine.setChrome(chromeState);
  }, [chromeState]);

  const clearTransientSelection = useCallback(() => {
    transientSelectionRef.current = null;
    if (transientSelectionFrameRef.current !== null) {
      if (typeof window !== 'undefined') window.cancelAnimationFrame(transientSelectionFrameRef.current);
      transientSelectionFrameRef.current = null;
    }
    const engine = engineRef.current;
    if (engine) engine.setChrome(chromeState);
  }, [chromeState]);

  const queueTransientSelection = useCallback((nextSelection: SelectionState) => {
    transientSelectionRef.current = nextSelection;
    if (transientSelectionFrameRef.current !== null) return;
    const draw = () => {
      transientSelectionFrameRef.current = null;
      const preview = transientSelectionRef.current;
      const engine = engineRef.current;
      if (!preview || !engine) return;
      engine.setChrome({ ...chromeState, selection: toChromeSelection(preview) });
    };
    if (typeof window !== 'undefined' && typeof window.requestAnimationFrame === 'function') {
      transientSelectionFrameRef.current = window.requestAnimationFrame(draw);
    } else {
      transientSelectionFrameRef.current = setTimeout(draw, 0) as unknown as number;
    }
  }, [chromeState]);

  useEffect(() => {
    const engine = engineRef.current;
    if (!engine) return;
    const detach = engine.onViewportChanged(() => setScrollTick((tick) => tick + 1));
    return detach;
  }, []);

  // 选区变化 → 滚动至可见
  useEffect(() => {
    const engine = engineRef.current;
    if (!engine || dragRef.current) return;
    engine.ensureVisible(selection.activeCell);
  }, [selection.activeCell]);

  // ---------- 指针交互 ----------

  const localPointOf = useCallback((event: { clientX: number; clientY: number }) => {
    const host = containerRef.current;
    if (!host) return { x: 0, y: 0 };
    const bounds = host.getBoundingClientRect();
    return { x: event.clientX - bounds.left, y: event.clientY - bounds.top };
  }, []);

  const stopAutoScroll = useCallback(() => {
    autoScrollPointRef.current = null;
    if (autoScrollFrameRef.current === null) return;
    if (typeof window !== "undefined") window.cancelAnimationFrame(autoScrollFrameRef.current);
    autoScrollFrameRef.current = null;
  }, []);

  const updateAutoScroll = useCallback((local: { x: number; y: number }) => {
    const host = containerRef.current;
    const engine = engineRef.current;
    const drag = dragRef.current;
    if (!host || !engine || !drag || (drag.kind !== "select" && drag.kind !== "fill")) {
      stopAutoScroll();
      return;
    }
    autoScrollPointRef.current = local;
    const origin = engine.headerOffset;
    const threshold = 24;
    const edge = local.x <= origin.x + threshold || local.x >= host.clientWidth - threshold
      || local.y <= origin.y + threshold || local.y >= host.clientHeight - threshold;
    if (!edge || typeof window === "undefined") {
      stopAutoScroll();
      return;
    }
    if (autoScrollFrameRef.current !== null) return;
    const tick = () => {
      autoScrollFrameRef.current = null;
      const currentDrag = dragRef.current;
      const point = autoScrollPointRef.current;
      const currentEngine = engineRef.current;
      const currentHost = containerRef.current;
      if (!currentDrag || !point || !currentEngine || !currentHost || (currentDrag.kind !== "select" && currentDrag.kind !== "fill")) return;
      const currentOrigin = currentEngine.headerOffset;
      const left = point.x < currentOrigin.x + threshold;
      const right = point.x > currentHost.clientWidth - threshold;
      const top = point.y < currentOrigin.y + threshold;
      const bottom = point.y > currentHost.clientHeight - threshold;
      const speed = (distance: number) => Math.max(2, Math.min(24, Math.round(distance / threshold * 24)));
      const dx = left ? -speed(currentOrigin.x + threshold - point.x) : right ? speed(point.x - (currentHost.clientWidth - threshold)) : 0;
      const dy = top ? -speed(currentOrigin.y + threshold - point.y) : bottom ? speed(point.y - (currentHost.clientHeight - threshold)) : 0;
      if (dx !== 0 || dy !== 0) currentEngine.scrollBy(dx, dy);
      const queryPoint = {
        x: Math.max(currentOrigin.x + 1, Math.min(currentHost.clientWidth - 1, point.x)),
        y: Math.max(currentOrigin.y + 1, Math.min(currentHost.clientHeight - 1, point.y)),
      };
        const cell = resolveDragCell(currentEngine, sheet, queryPoint, currentDrag);
        if (cell) {
          currentDrag.currentRow = cell.row;
          currentDrag.currentColumn = cell.column;
          if (currentDrag.kind === "fill") {
            setFillPreview({
              startRow: Math.min(currentDrag.startRow, currentDrag.currentRow),
            endRow: Math.max(currentDrag.anchorRow, currentDrag.currentRow),
            startColumn: Math.min(currentDrag.startColumn, currentDrag.currentColumn),
            endColumn: Math.max(currentDrag.anchorColumn, currentDrag.currentColumn),
          });
        } else {
          const isRowDrag = currentDrag.floating?.id === "row";
          const isColDrag = currentDrag.floating?.id === "col";
          const baseRange: RangeRef = {
            sheetId,
            startRow: isRowDrag || isColDrag ? Math.min(currentDrag.startRow, currentDrag.currentRow) : Math.min(currentDrag.anchorRow, currentDrag.currentRow),
            endRow: isRowDrag ? Math.max(currentDrag.startRow, currentDrag.currentRow) : isColDrag ? Math.max(0, skeleton.rowCount - 1) : Math.max(currentDrag.anchorRow, currentDrag.currentRow),
            startColumn: isColDrag ? Math.min(currentDrag.startColumn, currentDrag.currentColumn) : Math.min(currentDrag.anchorColumn, currentDrag.currentColumn),
            endColumn: isRowDrag ? Math.max(0, skeleton.columnCount - 1) : isColDrag ? Math.max(currentDrag.startColumn, currentDrag.currentColumn) : Math.max(currentDrag.anchorColumn, currentDrag.currentColumn),
          };
          const range = expandRangeForMerges(sheet, baseRange);
          const nextSelection: SelectionState = {
            ranges: [range],
            primaryRangeIndex: 0,
            activeCell: { row: currentDrag.currentRow, column: currentDrag.currentColumn },
            anchorCell: { row: currentDrag.anchorRow, column: currentDrag.anchorColumn },
          };
          queueTransientSelection(currentDrag.additive
            ? { ...selection, ranges: [...selection.ranges, range], primaryRangeIndex: selection.ranges.length, activeCell: nextSelection.activeCell, anchorCell: nextSelection.anchorCell }
            : nextSelection);
        }
      }
      if (autoScrollPointRef.current && typeof window !== "undefined") {
        autoScrollFrameRef.current = window.requestAnimationFrame(tick);
      }
    };
    autoScrollFrameRef.current = window.requestAnimationFrame(tick);
  }, [queueTransientSelection, selection, sheet, sheetId, skeleton, stopAutoScroll]);

  useEffect(() => () => {
    stopAutoScroll();
    if (transientSelectionFrameRef.current === null) return;
    if (typeof window !== 'undefined') window.cancelAnimationFrame(transientSelectionFrameRef.current);
    transientSelectionFrameRef.current = null;
  }, [stopAutoScroll]);

  const handlePointerDown = useCallback(
    (event: React.PointerEvent) => {
      if (phase !== "ready") return;
      if (event.button === 2) return; // 右键交给 contextmenu
      if ((event.target as Element).closest('[aria-label="Cell editor"]')) return;
      stopAutoScroll();
      setFillPreview(null);
      const engine = engineRef.current;
      const host = containerRef.current;
      if (!engine || !host) return;
      host.focus();
      if (editingCell || editingActiveRef.current) {
        editingActiveRef.current = false;
        onCommitEdit("none");
      }
      setFilterPopover(null);
      setValidationDropdown(null);
      const local = localPointOf(event);
      onPivotContextHit?.(null);

      // 1) 浮动对象优先
      const floatingHit = engine.hitTestFloating(local);
      if (floatingHit) {
        const drawableBounds = floatables.find((item) => item.id === floatingHit.id)?.bounds;
        if (drawableBounds) {
          dragRef.current = {
            kind: floatingHit.handle ? "floating-resize" : "floating-move",
            startRow: 0,
            startColumn: 0,
            anchorRow: 0,
            anchorColumn: 0,
            currentRow: 0,
            currentColumn: 0,
            additive: false,
            extend: false,
            resizeStartSize: 0,
            resizeIndex: 0,
            floating: {
              id: floatingHit.id,
              kind: floatingHit.kind,
              handle: floatingHit.handle,
              rotation: drawings.find((drawing) => drawing.id === floatingHit.id)?.transform.rotation,
              startBounds: { ...drawableBounds },
              startLocal: local,
            },
          };
          onFloatingSelect(floatingHit);
          (event.target as Element).setPointerCapture?.(event.pointerId);
          return;
        }
      }
      onFloatingSelect(null);

      // 2) 表头区
      const headerHit = engine.headerHitAtLocal(local);
      if (headerHit) {
        if (headerHit.kind === "corner") {
          onSelectAll();
          return;
        }
        if (headerHit.resizeBoundaryPx !== undefined) {
          dragRef.current = {
            kind: headerHit.kind === "col" ? "col-resize" : "row-resize",
            startRow: 0,
            startColumn: 0,
            anchorRow: 0,
            anchorColumn: 0,
            currentRow: 0,
            currentColumn: 0,
            additive: false,
            extend: false,
            resizeStartSize: headerHit.kind === "col" ? skeleton.getColumnWidth(headerHit.index) : skeleton.getRowHeight(headerHit.index),
            resizeIndex: headerHit.index,
          };
          (event.target as Element).setPointerCapture?.(event.pointerId);
          return;
        }
        const additive = event.ctrlKey || event.metaKey;
        if (headerHit.kind === "row") {
          for (const control of sheet.outlineControls) {
            if (control.axis !== "row" || control.index !== headerHit.index) continue;
            const buttonLeft = 4 + (control.level - 1) * 10;
            if (local.x >= buttonLeft && local.x <= buttonLeft + 10) {
              onToggleOutline?.(control.groupId);
              return;
            }
          }
        }
        if (headerHit.kind === "col") {
          for (const control of sheet.outlineControls) {
            if (control.axis !== "column" || control.index !== headerHit.index) continue;
            const buttonTop = 2 + (control.level - 1) * 10;
            if (local.y >= buttonTop && local.y <= buttonTop + 10 && local.x >= 2 && local.x <= 12) {
              onToggleOutline?.(control.groupId);
              return;
            }
          }
          if (sheet.filterColumns.includes(headerHit.index)) {
            setFilterPopover({ column: headerHit.index, x: event.clientX, y: event.clientY });
          }
        }
        dragRef.current = {
          kind: "select",
          startRow: headerHit.kind === "row" ? headerHit.index : 0,
          startColumn: headerHit.kind === "col" ? headerHit.index : 0,
          anchorRow: headerHit.kind === "row" ? headerHit.index : 0,
          anchorColumn: headerHit.kind === "col" ? headerHit.index : 0,
          currentRow: headerHit.kind === "row" ? headerHit.index : 0,
          currentColumn: headerHit.kind === "col" ? headerHit.index : 0,
          additive,
          extend: false,
          resizeStartSize: 0,
          resizeIndex: 0,
          floating: { id: headerHit.kind, kind: 'shape', handle: undefined, startBounds: { x: 0, y: 0, width: 0, height: 0 }, startLocal: { x: 0, y: 0 } },
        };
        (event.target as Element).setPointerCapture?.(event.pointerId);
        return;
      }

      // 3) 填充柄
      const primaryRange = selection.ranges[selection.primaryRangeIndex] ?? selection.ranges[0];
      const activePivotContextHit = resolvePivotProjectionHit(sheet, selection.activeCell.row, selection.activeCell.column);
      if (primaryRange && !activePivotContextHit) {
        const rect = skeleton.getRangeRect({
          startRow: primaryRange.endRow,
          endRow: primaryRange.endRow,
          startColumn: primaryRange.endColumn,
          endColumn: primaryRange.endColumn,
        });
        if (rect) {
          const screen = engine.contentToMainScreen(
            { x: rect.x + rect.width, y: rect.y + rect.height },
            { row: primaryRange.endRow, column: primaryRange.endColumn },
          );
          const half = 5;
          if (Math.abs(local.x - screen.x) <= half && Math.abs(local.y - screen.y) <= half) {
            dragRef.current = {
              kind: "fill",
              startRow: primaryRange.startRow,
              startColumn: primaryRange.startColumn,
              anchorRow: primaryRange.endRow,
              anchorColumn: primaryRange.endColumn,
              currentRow: primaryRange.endRow,
              currentColumn: primaryRange.endColumn,
              additive: false,
              extend: false,
              resizeStartSize: 0,
              resizeIndex: 0,
            };
            (event.target as Element).setPointerCapture?.(event.pointerId);
            return;
          }
        }
      }

      // 4) Pivot overlay context (it still uses the normal selection model;
      // only editing is blocked below). The hit is resolved before ordinary
      // cell handling so the host can activate Pivot contextual UI.
      const hitCell = engine.cellAtLocalPoint(local);
      const pivotContextHit = hitCell ? resolvePivotProjectionHit(sheet, hitCell.row, hitCell.column) : null;
      if (pivotContextHit) onPivotContextHit?.(pivotContextHit);

      // 5) 普通单元格选择/拖选
      if (!hitCell) return;
      const cell = resolveMergedCell(sheet, hitCell);
      const filterButton = sheet.filterButtons.find((button) => button.row === hitCell.row && button.column === hitCell.column);
      if (filterButton) {
        const cellRect = skeleton.getCellRect(hitCell.row, hitCell.column);
        if (cellRect) {
          const content = engine.localToContent(local);
          if (content.x >= cellRect.x + cellRect.width - 18) {
            setFilterPopover({ column: hitCell.column, x: event.clientX, y: event.clientY });
            return;
          }
        }
      }
      const additive = event.ctrlKey || event.metaKey;
      const extend = event.shiftKey && !additive;
      dragRef.current = {
        kind: "select",
        startRow: cell.row,
        startColumn: cell.column,
        anchorRow: extend ? selection.anchorCell.row : cell.row,
        anchorColumn: extend ? selection.anchorCell.column : cell.column,
        currentRow: cell.row,
        currentColumn: cell.column,
        additive,
        extend,
        resizeStartSize: 0,
        resizeIndex: 0,
        floating: undefined,
      };
      if (!additive && !extend) {
        onSelectionChange({
          ranges: [{ sheetId, startRow: cell.row, endRow: cell.row, startColumn: cell.column, endColumn: cell.column }],
          primaryRangeIndex: 0,
          activeCell: { row: cell.row, column: cell.column },
          anchorCell: { row: cell.row, column: cell.column },
        });
      }
      (event.target as Element).setPointerCapture?.(event.pointerId);
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [editingCell, floatables, localPointOf, onCommitEdit, onFloatingSelect, onPivotContextHit, onSelectAll, onSelectionChange, phase, selection, sheet, sheetId, skeleton, stopAutoScroll],
  );

  const handlePointerMove = useCallback(
    (event: React.PointerEvent) => {
      const engine = engineRef.current;
      if (!engine) return;
      const local = localPointOf(event);
      const drag = dragRef.current;

      if (!drag) {
        stopAutoScroll();
        // 悬停光标提示
        const headerHit = engine.headerHitAtLocal(local);
        const host = containerRef.current;
        if (host) {
          host.style.cursor = headerHit?.resizeBoundaryPx !== undefined
            ? (headerHit.kind === "col" ? "col-resize" : "row-resize")
            : "default";
        }
        return;
      }

      if (drag.kind === "col-resize" || drag.kind === "row-resize") {
        setFillPreview(null);
        stopAutoScroll();
        const content = engine.localToContent(local);
        const boundary = drag.kind === "col-resize"
          ? skeleton.getColumnLeft(drag.resizeIndex)
          : skeleton.getRowTop(drag.resizeIndex);
        const size = Math.max(24, (drag.kind === "col-resize" ? content.x : content.y) - boundary);
        const nextChrome = createEmptyChromeState();
        nextChrome.resizePreview = { axis: drag.kind === "col-resize" ? "column" : "row", index: drag.resizeIndex, sizePx: size };
        engine.setChrome({ ...chromeState, resizePreview: nextChrome.resizePreview });
        return;
      }

      if (drag.kind === "floating-move" && drag.floating) {
        setFillPreview(null);
        stopAutoScroll();
        const deltaX = local.x - drag.floating.startLocal.x;
        const deltaY = local.y - drag.floating.startLocal.y;
        const content = engine.localToContent(local);
        void content;
        onFloatingMove(
          drag.floating.id,
          {
            x: drag.floating.startBounds.x + deltaX,
            y: drag.floating.startBounds.y + deltaY,
            width: drag.floating.startBounds.width,
            height: drag.floating.startBounds.height,
          },
          drag.floating.rotation,
        );
        return;
      }

      if (drag.kind === "floating-resize" && drag.floating?.handle) {
        setFillPreview(null);
        stopAutoScroll();
        const handle = drag.floating.handle;
        const start = drag.floating.startBounds;
        const deltaX = local.x - drag.floating.startLocal.x;
        const deltaY = local.y - drag.floating.startLocal.y;
        let x = start.x;
        let y = start.y;
        let width = start.width;
        let height = start.height;
        if (handle.includes("e")) width = Math.max(40, start.width + deltaX);
        if (handle.includes("s")) height = Math.max(30, start.height + deltaY);
        if (handle.includes("w")) { width = Math.max(40, start.width - deltaX); x = start.x + (start.width - width); }
        if (handle.includes("n")) { height = Math.max(30, start.height - deltaY); y = start.y + (start.height - height); }
        onFloatingMove(drag.floating.id, { x, y, width, height }, drag.floating.rotation);
        return;
      }

      // select / fill:更新当前行列。行列头拖动不能依赖 cellAtLocalPoint，
      // 因为 pointer 可能始终位于表头区域。
      updateAutoScroll(local);
      const headerHit = engine.headerHitAtLocal(local);
      const hitCell = engine.cellAtLocalPoint(local);
      const isRowDrag = drag.floating?.id === "row";
      const isColDrag = drag.floating?.id === "col";
      const cell = isRowDrag
        ? headerHit?.kind === "row"
          ? { row: headerHit.index, column: 0 }
          : hitCell
            ? { row: hitCell.row, column: 0 }
            : null
        : isColDrag
          ? headerHit?.kind === "col"
            ? { row: 0, column: headerHit.index }
            : hitCell
              ? { row: 0, column: hitCell.column }
              : null
          : hitCell ? resolveMergedCell(sheet, hitCell) : null;
      if (!cell) return;
      if (drag.kind === "fill") {
        drag.currentRow = cell.row;
        drag.currentColumn = cell.column;
        setFillPreview({
          startRow: Math.min(drag.startRow, drag.currentRow),
          endRow: Math.max(drag.anchorRow, drag.currentRow),
          startColumn: Math.min(drag.startColumn, drag.currentColumn),
          endColumn: Math.max(drag.anchorColumn, drag.currentColumn),
        });
        return;
      }
      drag.currentRow = cell.row;
      drag.currentColumn = cell.column;
      const startRow = Math.min(drag.anchorRow, cell.row);
      const endRow = Math.max(drag.anchorRow, cell.row);
      const startColumn = Math.min(drag.anchorColumn, cell.column);
      const endColumn = Math.max(drag.anchorColumn, cell.column);
      const baseRange: RangeRef = {
        sheetId,
        startRow: isRowDrag || isColDrag ? Math.min(drag.startRow, drag.currentRow) : startRow,
        endRow: isRowDrag ? Math.max(drag.startRow, drag.currentRow) : isColDrag ? Math.max(0, skeleton.rowCount - 1) : endRow,
        startColumn: isColDrag ? Math.min(drag.startColumn, drag.currentColumn) : startColumn,
        endColumn: isRowDrag ? Math.max(0, skeleton.columnCount - 1) : isColDrag ? Math.max(drag.startColumn, drag.currentColumn) : endColumn,
      };
      const range = expandRangeForMerges(sheet, baseRange);
      const nextSelection: SelectionState = {
        ranges: [range],
        activeCell: { row: cell.row, column: cell.column },
        primaryRangeIndex: 0,
        anchorCell: { row: drag.anchorRow, column: drag.anchorColumn },
      };
      const previewSelection = drag.additive
        ? { ...selection, ranges: [...selection.ranges, ...nextSelection.ranges], primaryRangeIndex: selection.ranges.length, activeCell: nextSelection.activeCell, anchorCell: nextSelection.anchorCell }
        : nextSelection;
      queueTransientSelection(previewSelection);
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [localPointOf, onFloatingMove, queueTransientSelection, selection, sheet, sheetId, skeleton, stopAutoScroll, updateAutoScroll],
  );

  const handlePointerUp = useCallback(
    (event: React.PointerEvent) => {
      const drag = dragRef.current;
      dragRef.current = null;
      const engine = engineRef.current;
      if (!drag || !engine) return;
      (event.target as Element).releasePointerCapture?.(event.pointerId);

      if (drag.kind === "select" || drag.kind === "fill") {
        const finalCell = resolveDragCell(engine, sheet, localPointOf(event), drag);
        if (finalCell) {
          drag.currentRow = finalCell.row;
          drag.currentColumn = finalCell.column;
          const pivotContextHit = resolvePivotProjectionHit(sheet, finalCell.row, finalCell.column);
          onPivotContextHit?.(pivotContextHit);
        }
      }
      stopAutoScroll();

      if (drag.kind === "col-resize") {
        const content = engine.localToContent(localPointOf(event));
        const width = Math.max(24, content.x - skeleton.getColumnLeft(drag.resizeIndex));
        onResizeColumn(drag.resizeIndex, Math.round(width / (zoom / 100)));
        return;
      }
      if (drag.kind === "row-resize") {
        const content = engine.localToContent(localPointOf(event));
        const height = Math.max(18, content.y - skeleton.getRowTop(drag.resizeIndex));
        onResizeRow(drag.resizeIndex, Math.round(height / (zoom / 100)));
        return;
      }
      if (drag.kind === "fill") {
        setFillPreview(null);
        const target: RangeRef = {
          sheetId,
          startRow: Math.min(drag.startRow, drag.currentRow),
          endRow: Math.max(drag.anchorRow, drag.currentRow),
          startColumn: Math.min(drag.startColumn, drag.currentColumn),
          endColumn: Math.max(drag.anchorColumn, drag.currentColumn),
        };
        const partialMerge = sheet.merges.some((merge) =>
          intersectsRange(target, merge.range) && !containsRange(target, merge.range));
        if (!partialMerge && (target.endRow !== drag.anchorRow || target.endColumn !== drag.anchorColumn)) {
          onFillRange(target);
          onSelectionChange({
            ranges: [target],
            primaryRangeIndex: 0,
            activeCell: { row: drag.currentRow, column: drag.currentColumn },
            anchorCell: { row: drag.startRow, column: drag.startColumn },
          });
        }
        return;
      }
      if (drag.kind === "select") {
        setFillPreview(null);
        if (drag.extend) {
          clearTransientSelection();
          onExtendSelection?.(drag.currentRow, drag.currentColumn);
          return;
        }
        const startRow = Math.min(drag.anchorRow, drag.currentRow);
        const endRow = Math.max(drag.anchorRow, drag.currentRow);
        const startColumn = Math.min(drag.anchorColumn, drag.currentColumn);
        const endColumn = Math.max(drag.anchorColumn, drag.currentColumn);
        const isRowDrag = drag.floating?.id === "row";
        const isColDrag = drag.floating?.id === "col";
        const baseRange: RangeRef = {
          sheetId,
          startRow: isRowDrag || isColDrag ? Math.min(drag.startRow, drag.currentRow) : startRow,
          endRow: isRowDrag ? Math.max(drag.startRow, drag.currentRow) : isColDrag ? Math.max(0, skeleton.rowCount - 1) : endRow,
          startColumn: isColDrag ? Math.min(drag.startColumn, drag.currentColumn) : startColumn,
          endColumn: isRowDrag ? Math.max(0, skeleton.columnCount - 1) : isColDrag ? Math.max(drag.startColumn, drag.currentColumn) : endColumn,
        };
        const range = expandRangeForMerges(sheet, baseRange);
        const nextSelection: SelectionState = drag.additive
          ? { ...selection, ranges: [...selection.ranges, range], primaryRangeIndex: selection.ranges.length, activeCell: { row: drag.currentRow, column: drag.currentColumn }, anchorCell: { row: drag.anchorRow, column: drag.anchorColumn } }
          : { ranges: [range], primaryRangeIndex: 0, activeCell: { row: drag.currentRow, column: drag.currentColumn }, anchorCell: { row: drag.anchorRow, column: drag.anchorColumn } };
        clearTransientSelection();
        onSelectionChange(nextSelection);
      }
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [clearTransientSelection, localPointOf, onExtendSelection, onFillRange, onPivotContextHit, onResizeColumn, onResizeRow, onSelectionChange, selection, sheet, sheetId, skeleton, stopAutoScroll, zoom],
  );

  const handleDoubleClick = useCallback(
    (event: React.PointerEvent | React.MouseEvent) => {
      const engine = engineRef.current;
      if (!engine) return;
      const local = localPointOf(event);
      onPivotContextHit?.(null);
      const headerHit = engine.headerHitAtLocal(local);
      if (headerHit?.resizeBoundaryPx !== undefined) {
        // 双击边界 = 自适应内容宽(近似:取列内最长显示文本)
        const column = headerHit.index;
        let maxWidth = 60;
        const context = engine.getCanvas("content")?.getContext("2d");
        if (context) {
          context.font = "13px Segoe UI, sans-serif";
          for (let row = 0; row < Math.min(sheet.rowCount, 200); row += 1) {
            const cell = sheet.getCell(row, column);
            if (cell?.value) maxWidth = Math.max(maxWidth, context.measureText(cell.value).width + 16);
          }
        }
        onResizeColumn(column, Math.round(maxWidth / (zoom / 100)));
        return;
      }
      const hitCell = engine.cellAtLocalPoint(local);
      if (!hitCell) return;
      const pivotContextHit = resolvePivotProjectionHit(sheet, hitCell.row, hitCell.column);
      if (pivotContextHit) {
        onPivotContextHit?.(pivotContextHit);
        const pivotTarget = findPivotProjectionCell(sheet, hitCell.row, hitCell.column);
        if (pivotTarget && isPivotValueCell(pivotTarget.cell) && pivotTarget.cell.sourceRowPaths && pivotTarget.cell.sourceRowPaths.length > 0) {
          onPivotShowDetails?.({
            pivotId: pivotTarget.projection.pivotId,
            sourceRowPaths: pivotTarget.cell.sourceRowPaths,
            hit: pivotContextHit,
          });
        }
        // A Pivot projection is derived and therefore never editable through
        // the ordinary CellEditor, including headers and loading/error cells.
        return;
      }
      const cell = resolveMergedCell(sheet, hitCell);
      const validationList = getValidationList(cell.row, cell.column);
      if (validationList && validationList.length > 0) {
        setValidationDropdown({ row: cell.row, column: cell.column, options: validationList });
        return;
      }
      onBeginEdit();
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [getValidationList, localPointOf, onBeginEdit, onPivotContextHit, onPivotShowDetails, onResizeColumn, sheet, zoom],
  );

  const handleWheel = useCallback(
    (event: React.WheelEvent) => {
      const engine = engineRef.current;
      if (!engine) return;
      event.preventDefault();
      engine.scrollBy(event.deltaX, event.deltaY);
    },
    [],
  );

  // ---------- 键盘 ----------

  const handleKeyDown = useCallback(
    (event: React.KeyboardEvent) => {
      if (phase !== "ready") return;
      const key = event.key;
      const ctrl = event.ctrlKey || event.metaKey;
      const isEditing = Boolean(editingCell) || editingActiveRef.current;
      const activePivotContextHit = resolvePivotProjectionHit(sheet, selection.activeCell.row, selection.activeCell.column);
      const editingPivotContextHit = editingCell
        ? resolvePivotProjectionHit(sheet, editingCell.row, editingCell.column)
        : null;

      // A stale host editing state must not expose the ordinary editor over a
      // derived Pivot cell. Cancel it and keep the Pivot context authoritative.
      if (editingPivotContextHit) {
        event.preventDefault();
        editingActiveRef.current = false;
        onPivotContextHit?.(editingPivotContextHit);
        onCancelEdit();
        return;
      }
      if (activePivotContextHit && isEditing) {
        event.preventDefault();
        editingActiveRef.current = false;
        onPivotContextHit?.(activePivotContextHit);
        onCancelEdit();
        return;
      }

      if (isEditing) {
        if (key === "Escape") {
          event.preventDefault();
          editingActiveRef.current = false;
          onCancelEdit();
          return;
        }
        if (key === "Enter" && !event.shiftKey) {
          event.preventDefault();
          editingActiveRef.current = false;
          onCommitEdit("down");
          return;
        }
        if (key === "Tab") {
          event.preventDefault();
          editingActiveRef.current = false;
          onCommitEdit(event.shiftKey ? "left" : "right");
          return;
        }
        if (key === "F4") {
          event.preventDefault();
          onToggleAbsolute();
          return;
        }
        if (key.length === 1 && !ctrl && !event.altKey) {
          event.preventDefault();
          if (onAppendFormulaDraft) onAppendFormulaDraft(key);
          else onFormulaDraftChange(formulaDraft + key);
          return;
        }
        return;
      }

      const activeScope = activePivotContextHit
        ? 'pivot' as const
        : selectedFloatingId
          ? 'drawing' as const
          : 'grid' as const;
      const shortcut = shortcutRegistryRef.current?.resolve(event, {
        scope: activeScope,
        canRepeat,
      }) ?? (activeScope === 'grid' ? undefined : shortcutRegistryRef.current?.resolve(event, {
        scope: 'grid',
        canRepeat,
      }));
      if (shortcut?.id === 'context.open') {
        event.preventDefault();
        const activeRange = selection.ranges[selection.primaryRangeIndex] ?? selection.ranges[0] ?? {
          sheetId,
          startRow: selection.activeCell.row,
          endRow: selection.activeCell.row,
          startColumn: selection.activeCell.column,
          endColumn: selection.activeCell.column,
        };
        contextRangeRef.current = activeRange;
        setContextHit(activePivotContextHit ?? resolveContextHit({
          sheetId,
          cell: { row: selection.activeCell.row, column: selection.activeCell.column, range: activeRange },
        }));
        if (activePivotContextHit) onPivotContextHit?.(activePivotContextHit);
        const bounds = containerRef.current?.getBoundingClientRect();
        setContextMenu({ x: (bounds?.left ?? 0) + 24, y: (bounds?.top ?? 0) + 24, open: true });
        return;
      }
      if (shortcut && onShortcut?.(shortcut.id)) {
        event.preventDefault();
        return;
      }

      if (ctrl && (key === "z" || key === "Z")) { event.preventDefault(); onUndo(); return; }
      if (ctrl && (key === "y" || key === "Y")) { event.preventDefault(); onRedo(); return; }
      if (ctrl && (key === "c" || key === "C")) { event.preventDefault(); onCopy(); return; }
      if (ctrl && (key === "x" || key === "X")) { event.preventDefault(); onCut(); return; }
      if (ctrl && (key === "v" || key === "V")) { event.preventDefault(); onPaste(); return; }
      if (ctrl && (key === "b" || key === "B")) { event.preventDefault(); onCommand({ commandId: "sheet.style.set", params: { style: { bold: !cellStyle.bold } } }); return; }
      if (ctrl && (key === "i" || key === "I")) { event.preventDefault(); onCommand({ commandId: "sheet.style.set", params: { style: { italic: !cellStyle.italic } } }); return; }
      if (ctrl && (key === "u" || key === "U")) { event.preventDefault(); onCommand({ commandId: "sheet.style.set", params: { style: { underline: !cellStyle.underline } } }); return; }
      if (key === "F2") {
        event.preventDefault();
        if (activePivotContextHit) onPivotContextHit?.(activePivotContextHit);
        else onBeginEdit();
        return;
      }
      if (key === "Delete" || key === "Backspace") { event.preventDefault(); onCommand({ commandId: "sheet.range.clear" }); return; }
      if (key === "Enter") {
        event.preventDefault();
        if (activePivotContextHit) onPivotContextHit?.(activePivotContextHit);
        else onBeginEdit();
        return;
      }

      const moves: Record<string, [number, number]> = {
        ArrowUp: [-1, 0],
        ArrowDown: [1, 0],
        ArrowLeft: [0, -1],
        ArrowRight: [0, 1],
        Tab: [0, event.shiftKey ? -1 : 1],
      };
      if (key in moves && !ctrl) {
        event.preventDefault();
        const [dr, dc] = moves[key]!;
        onMovePrimary(dr, dc, { extend: event.shiftKey });
        return;
      }
      if (key in moves && ctrl) {
        event.preventDefault();
        const direction = key === "ArrowUp" ? "up" : key === "ArrowDown" ? "down" : key === "ArrowLeft" ? "left" : "right";
        onJumpEdge(direction, event.shiftKey);
        return;
      }
      if (key === "Home") {
        event.preventDefault();
        onMovePrimary(0, -selection.activeCell.column, { extend: event.shiftKey });
        return;
      }
      if (key === "PageDown" || key === "PageUp") {
        event.preventDefault();
        const rowHeight = Math.max(1, skeleton.getRowHeight(selection.activeCell.row));
        const rows = Math.max(1, Math.floor((containerRef.current?.clientHeight ?? 600) / rowHeight) - 2);
        const delta = key === "PageDown" ? rows : -rows;
        onMovePrimary(delta, 0, { extend: event.shiftKey });
        return;
      }
      // 直接输入进入编辑
      if (key.length === 1 && !ctrl && !event.altKey) {
        event.preventDefault();
        if (activePivotContextHit) {
          onPivotContextHit?.(activePivotContextHit);
        } else if (editingCell || editingActiveRef.current) {
          onAppendFormulaDraft?.(key);
        } else {
          editingActiveRef.current = true;
          onBeginEdit(key);
        }
      }
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [canRepeat, cellStyle.bold, cellStyle.italic, cellStyle.underline, editingCell, formulaDraft, onAppendFormulaDraft, onCancelEdit, onCommitEdit, onFormulaDraftChange, onBeginEdit, onInsertRef, onJumpEdge, onMovePrimary, onCommand, onCopy, onCut, onPaste, onPivotContextHit, onRedo, onShortcut, onUndo, phase, selectedFloatingId, selection, sheet, sheetId, skeleton],
  );

  // ---------- 右键菜单 ----------

  const contextMenuItems = useMemo<ContextMenuItem[]>(() => {
    if (contextHit?.kind === "pivot" && contextHit.pivot) {
      const supplied = getPivotContextMenuItems?.(contextHit);
      if (supplied) return [...supplied];
      const sourceRowPaths = contextHit.pivot.sourceRowPaths ?? [];
      return [
        {
          id: "pivot-show-details",
          label: "Show Details",
          disabled: sourceRowPaths.length === 0 || !onPivotShowDetails,
          onSelect: () => onPivotShowDetails?.({
            pivotId: contextHit.pivot!.pivotId,
            sourceRowPaths,
            hit: contextHit,
          }),
        },
        {
          id: "pivot-refresh",
          label: "Refresh PivotTable",
          onSelect: () => onCommand({ commandId: "pivot.refresh", params: { sheetId, pivotId: contextHit.pivot!.pivotId } }),
        },
      ];
    }
    const getContextRange = () => contextRangeRef.current ?? selection.ranges[selection.primaryRangeIndex] ?? selection.ranges[0];
    const items: ContextMenuItem[] = [
      { id: "cut", label: "Cut", shortcut: "Ctrl+X", onSelect: onCut },
      { id: "copy", label: "Copy", shortcut: "Ctrl+C", onSelect: onCopy },
      { id: "paste", label: "Paste", shortcut: "Ctrl+V", onSelect: onPaste },
      { id: "sep-1", label: "", separator: true },
      { id: "insert-row", label: "Insert row above", onSelect: () => { const range = getContextRange(); if (range) onCommand({ commandId: "sheet.rows.insert", params: { sheetId, at: range.startRow, count: 1 } }); } },
      { id: "insert-column", label: "Insert column left", onSelect: () => { const range = getContextRange(); if (range) onCommand({ commandId: "sheet.columns.insert", params: { sheetId, at: range.startColumn, count: 1 } }); } },
      { id: "delete-row", label: "Delete row", danger: true, onSelect: () => { const range = getContextRange(); if (range) onCommand({ commandId: "sheet.rows.delete", params: { sheetId, at: range.startRow, count: range.endRow - range.startRow + 1 } }); } },
      { id: "delete-column", label: "Delete column", danger: true, onSelect: () => { const range = getContextRange(); if (range) onCommand({ commandId: "sheet.columns.delete", params: { sheetId, at: range.startColumn, count: range.endColumn - range.startColumn + 1 } }); } },
      { id: "sep-2", label: "", separator: true },
      { id: "hide-row", label: "Hide rows", onSelect: () => { const range = getContextRange(); if (range) onCommand({ commandId: "sheet.row.hide", params: { sheetId, index: range.startRow } }); } },
      { id: "hide-col", label: "Hide columns", onSelect: () => { const range = getContextRange(); if (range) onCommand({ commandId: "sheet.column.hide", params: { sheetId, index: range.startColumn } }); } },
      { id: "unhide-all", label: "Unhide all", onSelect: () => onCommand({ commandId: "sheet.rows.unhide.all", params: { sheetId } }) },
      { id: "sep-3", label: "", separator: true },
      { id: "clear", label: "Clear contents", onSelect: () => onClearSelection?.("contents") },
      { id: "clear-formats", label: "Clear formats", onSelect: () => onClearSelection?.("formats") },
      { id: "comment-add", label: "Add comment", onSelect: onOpenInspector },
    ];
    return items;
  }, [contextHit, getPivotContextMenuItems, onClearSelection, onCommand, onCopy, onCut, onOpenInspector, onPaste, onPivotShowDetails, selection, sheetId]);

  const handleContextMenu = useCallback((event: React.MouseEvent) => {
    event.preventDefault();
    setFillPreview(null);
    onPivotContextHit?.(null);
    const engine = engineRef.current;
    if (!engine || phase !== "ready") {
      setContextHit(null);
      setContextMenu({ x: event.clientX, y: event.clientY, open: true });
      return;
    }
    const local = localPointOf(event);
    const headerHit = engine.headerHitAtLocal(local);
    setContextHit(null);
    contextRangeRef.current = selection.ranges[selection.primaryRangeIndex] ?? selection.ranges[0] ?? null;
    if (headerHit?.kind === "corner") {
      contextRangeRef.current = { sheetId, startRow: 0, endRow: Math.max(0, skeleton.rowCount - 1), startColumn: 0, endColumn: Math.max(0, skeleton.columnCount - 1) };
      onSelectAll();
    } else if (headerHit?.kind === "row") {
      const row = headerHit.index;
      contextRangeRef.current = { sheetId, startRow: row, endRow: row, startColumn: 0, endColumn: Math.max(0, skeleton.columnCount - 1) };
      onSelectionChange({
        ranges: [{ sheetId, startRow: row, endRow: row, startColumn: 0, endColumn: Math.max(0, skeleton.columnCount - 1) }],
        primaryRangeIndex: 0,
        activeCell: { row, column: 0 },
        anchorCell: { row, column: 0 },
      });
    } else if (headerHit?.kind === "col") {
      const column = headerHit.index;
      contextRangeRef.current = { sheetId, startRow: 0, endRow: Math.max(0, skeleton.rowCount - 1), startColumn: column, endColumn: column };
      onSelectionChange({
        ranges: [{ sheetId, startRow: 0, endRow: Math.max(0, skeleton.rowCount - 1), startColumn: column, endColumn: column }],
        primaryRangeIndex: 0,
        activeCell: { row: 0, column },
        anchorCell: { row: 0, column },
      });
    } else {
      const hitCell = engine.cellAtLocalPoint(local);
      if (hitCell) {
        const pivotContextHit = resolvePivotProjectionHit(sheet, hitCell.row, hitCell.column);
        if (pivotContextHit) {
          setContextHit(pivotContextHit);
          onPivotContextHit?.(pivotContextHit);
          const targetRange: RangeRef = {
            sheetId,
            startRow: hitCell.row,
            endRow: hitCell.row,
            startColumn: hitCell.column,
            endColumn: hitCell.column,
          };
          const alreadySelected = selection.ranges.some((range) => intersectsRange(range, targetRange));
          if (!alreadySelected) {
            contextRangeRef.current = targetRange;
            onSelectionChange({
              ranges: [targetRange],
              primaryRangeIndex: 0,
              activeCell: { row: hitCell.row, column: hitCell.column },
              anchorCell: { row: hitCell.row, column: hitCell.column },
            });
          } else {
            contextRangeRef.current = selection.ranges[selection.primaryRangeIndex] ?? selection.ranges[0] ?? targetRange;
          }
        } else {
          setContextHit(null);
          const cell = resolveMergedCell(sheet, hitCell);
          const merge = mergeAtCell(sheet, hitCell);
          const targetRange: RangeRef = merge?.range ?? {
            sheetId,
            startRow: cell.row,
            endRow: cell.row,
            startColumn: cell.column,
            endColumn: cell.column,
          };
          const alreadySelected = selection.ranges.some((range) => intersectsRange(range, targetRange));
          if (!alreadySelected) {
            contextRangeRef.current = expandRangeForMerges(sheet, targetRange);
            onSelectionChange({
              ranges: [contextRangeRef.current],
              primaryRangeIndex: 0,
              activeCell: cell,
              anchorCell: cell,
            });
          } else {
            contextRangeRef.current = selection.ranges[selection.primaryRangeIndex] ?? selection.ranges[0] ?? targetRange;
          }
        }
      }
    }
    setContextMenu({ x: event.clientX, y: event.clientY, open: true });
  }, [localPointOf, onPivotContextHit, onSelectAll, onSelectionChange, phase, selection, sheet, sheetId, skeleton]);

  // ---------- 编辑器定位(随滚动更新) ----------

  const editingPivotContextHit = editingCell
    ? resolvePivotProjectionHit(sheet, editingCell.row, editingCell.column)
    : null;

  // 编辑器随滚动重定位:依赖 scrollTick 触发重算
  const editorRect = useMemo(() => {
    void scrollTick;
    const engine = engineRef.current;
    if (!engine || !editingCell || editingPivotContextHit) return null;
    // The engine owns the render geometry. Reading the local React skeleton
    // here can race with setSkeleton during a session refresh, and the main
    // pane is not the correct origin for frozen rows/columns.
    const rect = engine.skeleton.getCellRect(editingCell.row, editingCell.column);
    if (!rect) return null;
    const topLeft = engine.contentToMainScreen({ x: rect.x, y: rect.y }, editingCell);
    return {
      x: topLeft.x,
      y: topLeft.y,
      width: rect.width,
      height: rect.height,
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [editingCell, editingPivotContextHit, skeleton, scrollTick]);

  if (phase === "empty") {
    return (
      <Panel className="m-4 flex-1">
        <StatePanel
          kind="empty"
          description="Create a workbook to start editing cells."
          actionLabel="Create workbook"
          onAction={onCreateSheet}
          title="No workbook loaded"
        />
      </Panel>
    );
  }

  if (phase === "error") {
    return (
      <Panel className="m-4 flex-1">
        <StatePanel
          kind="error"
          description="The workbook engine failed to initialize. Retry to recover."
          actionLabel="Retry"
          onAction={onRetry}
          title="Engine error"
        />
      </Panel>
    );
  }

  return (
    <Panel className="h-full min-h-0 flex-1 overflow-hidden">
      <Stack gap="none" className="h-full">
        <Inline gap="xs" className="items-center justify-between border-b border-slate-100 px-3 py-1.5">
          <Inline gap="xs" className="items-center">
            <Text size="xs" tone="muted">Sheet</Text>
            <Text size="xs" weight="semibold">{sheet.name}</Text>
            {sheet.freeze.xSplit > 0 || sheet.freeze.ySplit > 0 ? (
              <Text size="xs" tone="subtle">frozen {sheet.freeze.xSplit}x{sheet.freeze.ySplit}</Text>
            ) : null}
          </Inline>
          <Text size="xs" tone="subtle">{activeCell}</Text>
        </Inline>
        <Box className="relative min-h-0 flex-1">
          <Box
            ref={containerRef}
            role="grid"
            aria-label="Spreadsheet canvas"
            data-testid="sheet-canvas"
            aria-rowcount={sheet.rowCount}
            aria-colcount={sheet.columnCount}
            aria-rowindex={selection.activeCell.row + 1}
            aria-colindex={selection.activeCell.column + 1}
            tabIndex={0}
            className="absolute inset-0 outline-none"
            onPointerDown={handlePointerDown}
            onPointerMove={handlePointerMove}
            onPointerUp={handlePointerUp}
            onPointerCancel={handlePointerUp}
            onDoubleClick={handleDoubleClick}
            onWheel={handleWheel}
            onKeyDown={handleKeyDown}
            onContextMenu={handleContextMenu}
          >
            <CanvasRenderSurface
              onReady={(engine) => {
                engineRef.current = engine;
                setEngineReady(true);
                engine.setCellProvider(cellProvider);
                engine.setSkeleton(skeleton);
                engine.setFloating(floatables, selectedFloatingId);
                engine.setChrome(chromeState);
              }}
              className="absolute inset-0"
            />
            {engineReady ? pivotStatusProjections.map((projection) => (
              <PivotProjectionStatusNotice
                key={`${projection.pivotId}:${projection.refresh.status}:${projection.refresh.error ?? ""}`}
                engine={engineRef.current}
                projection={projection}
                scrollTick={scrollTick}
              />
            )) : null}
          </Box>

          {editorRect && editingCell ? (
            <Box
              className="absolute z-20 border-2 border-blue-600 bg-white shadow-lg"
              style={{ left: editorRect.x - 1, top: editorRect.y - 1, minWidth: Math.max(editorRect.width + 2, 120) }}
            >
              <CellEditor
                initialText={formulaDraft}
                onCancel={onCancelEdit}
                onChange={onFormulaDraftChange}
                onCommit={onCommitEdit}
                onInsertRef={onInsertRef}
              />
            </Box>
          ) : null}

          {fillPreview ? (
            <FillPreviewOverlay engine={engineRef.current} preview={fillPreview} />
          ) : null}

          {validationDropdown ? (
            <ValidationDropdown
              options={validationDropdown.options}
              onPick={(value) => {
                onCommitCell(value);
                setValidationDropdown(null);
              }}
              onClose={() => setValidationDropdown(null)}
            />
          ) : null}

          {filterPopover ? (
            <FilterPopover
              column={filterPopover.column}
              sheet={sheet}
              onApply={(patch) => {
                onApplyFilter(filterPopover.column, patch);
                setFilterPopover(null);
              }}
              onClose={() => setFilterPopover(null)}
            />
          ) : null}
        </Box>
      </Stack>

      <ContextMenu
        x={contextMenu.x}
        y={contextMenu.y}
        open={contextMenu.open}
        items={contextMenuItems}
        onClose={() => setContextMenu((previous) => ({ ...previous, open: false }))}
      />
    </Panel>
  );
}

function parseCellValue(cell: CanvasCellSnapshot): string | number | boolean | null {
  const numeric = Number(cell.value.replace(/[$,]/g, ""));
  if (cell.value !== "" && Number.isFinite(numeric) && /\d/.test(cell.value)) return numeric;
  if (cell.value === "TRUE") return true;
  if (cell.value === "FALSE") return false;
  return cell.value;
}

function pivotProjectionStatusMessage(projection: PivotGridProjection): string | null {
  if (projection.collision.status === "collision") {
    const reason = projection.collision.reasons.length > 0
      ? projection.collision.reasons.join(", ")
      : "target range is occupied";
    return `PivotTable ${projection.pivotId} cannot expand: ${reason}`;
  }
  if (projection.refresh.status === "error") {
    return `PivotTable ${projection.pivotId} error: ${projection.refresh.error ?? "refresh failed"}`;
  }
  if (projection.refresh.status === "refreshing") return `PivotTable ${projection.pivotId}: Loading PivotTable…`;
  if (projection.refresh.status === "stale") return `PivotTable ${projection.pivotId}: Refresh required`;
  return null;
}

function PivotProjectionStatusNotice({
  engine,
  projection,
  scrollTick,
}: {
  engine: CanvasRenderEngine | null;
  projection: PivotGridProjection;
  scrollTick: number;
}): React.ReactElement | null {
  void scrollTick;
  const message = pivotProjectionStatusMessage(projection);
  if (!engine || !message) return null;
  const rect = engine.contentRangeToScreenRects(projection.occupiedRange)[0];
  if (!rect) return null;
  const collision = projection.collision.status === "collision";
  return (
    <Box
      role="status"
      aria-live="polite"
      className={`pointer-events-none absolute z-10 max-w-[min(32rem,calc(100%-1rem))] rounded border px-2 py-1 text-[11px] shadow-sm ${collision ? "border-amber-300 bg-amber-50 text-amber-800" : "border-red-200 bg-red-50 text-red-700"}`}
      style={{ left: rect.x + 2, top: rect.y + 2 }}
    >
      {message}
    </Box>
  );
}

function FillPreviewOverlay({
  engine,
  preview,
}: {
  engine: CanvasRenderEngine | null;
  preview: { startRow: number; endRow: number; startColumn: number; endColumn: number };
}): React.ReactElement | null {
  if (!engine) return null;
  const rects = engine.contentRangeToScreenRects(preview);
  if (rects.length === 0) return null;
  return (
    <>
      {rects.map((screen, index) => (
        <Box
          key={`${screen.x}:${screen.y}:${index}`}
          className="pointer-events-none absolute z-10 border-2 border-dashed border-blue-500 bg-blue-500/5"
          style={{ left: screen.x, top: screen.y, width: screen.width, height: screen.height }}
        />
      ))}
    </>
  );
}

function ValidationDropdown({
  options,
  onPick,
  onClose,
}: {
  options: string[];
  onPick: (value: string) => void;
  onClose: () => void;
}): React.ReactElement {
  return (
    <Box className="absolute left-1/2 top-1/2 z-30 -translate-x-1/2 -translate-y-1/2 rounded-lg border border-slate-200 bg-white p-1 shadow-xl">
      <Stack gap="none">
        {options.map((option) => (
          <Button
            key={option}
            size="sm"
            variant="ghost"
            className="justify-start"
            onClick={() => onPick(option)}
          >
            {option}
          </Button>
        ))}
        <Button size="sm" variant="ghost" className="justify-start text-slate-400" onClick={onClose}>
          Cancel
        </Button>
      </Stack>
    </Box>
  );
}
