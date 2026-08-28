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
  CameraDrawingPayload,
  ScreenshotDrawingPayload,
  FormControlDrawingPayload,
  IconDrawingPayload,
  Model3dDrawingPayload,
  SmartArtDrawingPayload,
  WordArtDrawingPayload,
  SignatureLineDrawingPayload,
  EmbeddedObjectDrawingPayload,
  EquationDrawingPayload,
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
import { buildChartLayout, resolveChartDataFromSources, resolveSparklineData } from "@react-sheets/spreadsheet-app";
import type { ChartLayout, ResolvedChartData } from "@react-sheets/spreadsheet-app";
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

function getChartSeries(
  payload: ChartDrawingPayload,
  getSheet: (sheetId: string) => CanvasSheetSnapshot | undefined,
  pivotResults: Record<string, PivotResultTree>,
  sheets: readonly CanvasSheetSnapshot[],
  tables: readonly WorkbookTableModel[],
): ResolvedChartData {
  const pivotSources = { ...pivotResults };
  for (const source of sheets) for (const [pivotId, result] of Object.entries(source.pivotResults)) pivotSources[pivotId] ??= result;
  const data = resolveChartDataFromSources(payload, (sheetId) => getSheet(sheetId), pivotSources, tables);
  return data;
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
  const geometry = drawShapeGeometry(context, payload.type, width, height);
  if (geometry.fill) geometry.fillRule ? context.fill(geometry.fillRule) : context.fill();
  context.stroke();
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

interface ShapeGeometryResult {
  fill: boolean;
  fillRule?: CanvasFillRule;
}

function drawShapeGeometry(context: CanvasRenderingContext2D, type: Extract<DrawingPayload, { kind: 'shape' }>['type'], width: number, height: number): ShapeGeometryResult {
  const polygon = (points: readonly [number, number][]): ShapeGeometryResult => {
    context.beginPath();
    points.forEach(([x, y], index) => index === 0 ? context.moveTo(x, y) : context.lineTo(x, y));
    context.closePath();
    return { fill: true };
  };
  const arrowHead = (x: number, y: number, direction: number, size: number): void => {
    context.save();
    context.translate(x, y);
    context.rotate(direction);
    context.moveTo(0, 0);
    context.lineTo(-size, -size * 0.55);
    context.lineTo(-size, size * 0.55);
    context.closePath();
    context.restore();
  };
  const line = (start: [number, number], end: [number, number], endArrow = false, startArrow = false): ShapeGeometryResult => {
    context.beginPath();
    context.moveTo(start[0], start[1]);
    context.lineTo(end[0], end[1]);
    if (endArrow || startArrow) {
      const angle = Math.atan2(end[1] - start[1], end[0] - start[0]);
      const size = Math.max(8, Math.min(18, Math.min(width, height) * 0.35));
      if (endArrow) arrowHead(end[0], end[1], angle, size);
      if (startArrow) arrowHead(start[0], start[1], angle + Math.PI, size);
    }
    return { fill: endArrow || startArrow };
  };
  const star = (points: number, innerRatio: number): ShapeGeometryResult => {
    const centerX = width / 2;
    const centerY = height / 2;
    const outer = Math.min(width, height) / 2;
    const total = points * 2;
    const vertices: [number, number][] = [];
    for (let index = 0; index < total; index += 1) {
      const angle = -Math.PI / 2 + index * Math.PI / points;
      const radius = index % 2 === 0 ? outer : outer * innerRatio;
      vertices.push([centerX + Math.cos(angle) * radius, centerY + Math.sin(angle) * radius]);
    }
    return polygon(vertices);
  };

  switch (type) {
    case 'rectangle':
      context.beginPath(); context.rect(0, 0, width, height); return { fill: true };
    case 'rounded-rectangle': {
      const radius = Math.min(12, width / 4, height / 4);
      context.beginPath(); context.roundRect(0, 0, width, height, radius); return { fill: true };
    }
    case 'ellipse':
      context.beginPath(); context.ellipse(width / 2, height / 2, width / 2, height / 2, 0, 0, Math.PI * 2); return { fill: true };
    case 'triangle': return polygon([[width / 2, 0], [width, height], [0, height]]);
    case 'right-triangle': return polygon([[0, 0], [width, height], [0, height]]);
    case 'diamond': return polygon([[width / 2, 0], [width, height / 2], [width / 2, height], [0, height / 2]]);
    case 'parallelogram': {
      const skew = Math.min(width * 0.24, Math.max(8, height * 0.5));
      return polygon([[skew, 0], [width, 0], [width - skew, height], [0, height]]);
    }
    case 'trapezoid': {
      const inset = Math.min(width * 0.24, Math.max(8, width / 4));
      return polygon([[inset, 0], [width - inset, 0], [width, height], [0, height]]);
    }
    case 'hexagon': {
      const inset = Math.min(width * 0.2, Math.max(8, height / 2));
      return polygon([[inset, 0], [width - inset, 0], [width, height / 2], [width - inset, height], [inset, height], [0, height / 2]]);
    }
    case 'octagon': {
      const inset = Math.min(width, height) * 0.28;
      return polygon([[inset, 0], [width - inset, 0], [width, inset], [width, height - inset], [width - inset, height], [inset, height], [0, height - inset], [0, inset]]);
    }
    case 'plus': {
      const arm = Math.min(width, height) * 0.32;
      return polygon([[arm, 0], [width - arm, 0], [width - arm, arm], [width, arm], [width, height - arm], [width - arm, height - arm], [width - arm, height], [arm, height], [arm, height - arm], [0, height - arm], [0, arm], [arm, arm]]);
    }
    case 'home-plate': return polygon([[0, 0], [width * 0.72, 0], [width, height / 2], [width * 0.72, height], [0, height]]);
    case 'cube': return polygon([[width * 0.22, height * 0.18], [width * 0.72, 0], [width, height * 0.2], [width * 0.5, height * 0.38], [width * 0.22, height * 0.18], [0, height * 0.38], [width * 0.5, height * 0.58], [width, height * 0.38], [width, height * 0.82], [width * 0.5, height], [width * 0.5, height * 0.58], [0, height * 0.38], [0, height * 0.82], [width * 0.5, height], [width, height * 0.82], [width, height * 0.38], [width * 0.5, height * 0.58]]);
    case 'cylinder': {
      const ry = Math.min(height * 0.18, Math.max(5, width * 0.12));
      context.beginPath();
      context.moveTo(0, ry);
      context.ellipse(width / 2, ry, width / 2, ry, 0, Math.PI, 0, true);
      context.lineTo(width, height - ry);
      context.ellipse(width / 2, height - ry, width / 2, ry, 0, 0, Math.PI, true);
      context.closePath();
      return { fill: true };
    }
    case 'sun': return star(16, 0.78);
    case 'moon': {
      context.beginPath();
      context.moveTo(width * 0.72, height * 0.1);
      context.bezierCurveTo(width * 0.28, height * 0.14, width * 0.2, height * 0.78, width * 0.7, height * 0.9);
      context.bezierCurveTo(width * 0.44, height * 0.65, width * 0.44, height * 0.34, width * 0.72, height * 0.1);
      context.closePath();
      return { fill: true };
    }
    case 'heart': {
      context.beginPath();
      context.moveTo(width / 2, height * 0.92);
      context.bezierCurveTo(width * 0.08, height * 0.64, 0, height * 0.32, width * 0.24, height * 0.16);
      context.bezierCurveTo(width * 0.38, height * 0.07, width * 0.48, height * 0.18, width / 2, height * 0.3);
      context.bezierCurveTo(width * 0.52, height * 0.18, width * 0.62, height * 0.07, width * 0.76, height * 0.16);
      context.bezierCurveTo(width, height * 0.32, width * 0.92, height * 0.64, width / 2, height * 0.92);
      context.closePath();
      return { fill: true };
    }
    case 'lightning': return polygon([[width * 0.58, 0], [width * 0.12, height * 0.56], [width * 0.45, height * 0.56], [width * 0.32, height], [width * 0.88, height * 0.32], [width * 0.55, height * 0.32]]);
    case 'cloud': {
      context.beginPath();
      context.moveTo(width * 0.15, height * 0.75);
      context.bezierCurveTo(width * 0.02, height * 0.56, width * 0.13, height * 0.34, width * 0.34, height * 0.36);
      context.bezierCurveTo(width * 0.38, height * 0.1, width * 0.72, height * 0.08, width * 0.78, height * 0.36);
      context.bezierCurveTo(width * 1.02, height * 0.28, width * 1.03, height * 0.72, width * 0.84, height * 0.78);
      context.lineTo(width * 0.22, height * 0.78);
      context.closePath();
      return { fill: true };
    }
    case 'frame':
      context.beginPath(); context.rect(0, 0, width, height); context.rect(width * 0.18, height * 0.18, width * 0.64, height * 0.64); return { fill: true, fillRule: 'evenodd' };
    case 'line': return line([0, height / 2], [width, height / 2]);
    case 'arrow': return line([0, height / 2], [width, height / 2], true);
    case 'left-right-arrow': return polygon([[0, height / 2], [width * 0.18, 0], [width * 0.18, height * 0.27], [width * 0.82, height * 0.27], [width * 0.82, 0], [width, height / 2], [width * 0.82, height], [width * 0.82, height * 0.73], [width * 0.18, height * 0.73], [width * 0.18, height]]);
    case 'up-down-arrow': return polygon([[width / 2, 0], [width, height * 0.2], [width * 0.64, height * 0.2], [width * 0.64, height * 0.8], [width, height * 0.8], [width / 2, height], [0, height * 0.8], [width * 0.36, height * 0.8], [width * 0.36, height * 0.2], [0, height * 0.2]]);
    case 'quad-arrow': return polygon([[width / 2, 0], [width * 0.68, height * 0.2], [width * 0.58, height * 0.2], [width * 0.58, height * 0.42], [width * 0.8, height * 0.42], [width * 0.8, height * 0.32], [width, height / 2], [width * 0.8, height * 0.68], [width * 0.8, height * 0.58], [width * 0.58, height * 0.58], [width * 0.58, height * 0.8], [width * 0.68, height * 0.8], [width / 2, height], [width * 0.32, height * 0.8], [width * 0.42, height * 0.8], [width * 0.42, height * 0.58], [width * 0.2, height * 0.58], [width * 0.2, height * 0.68], [0, height / 2], [width * 0.2, height * 0.32], [width * 0.2, height * 0.42], [width * 0.42, height * 0.42], [width * 0.42, height * 0.2], [width * 0.32, height * 0.2]]);
    case 'bent-arrow': {
      const head = Math.min(width, height) * 0.25;
      context.beginPath(); context.moveTo(width * 0.12, height * 0.82); context.lineTo(width * 0.12, height * 0.3); context.quadraticCurveTo(width * 0.12, height * 0.12, width * 0.3, height * 0.12); context.lineTo(width * 0.76, height * 0.12); context.lineTo(width * 0.76, height * 0.02); context.lineTo(width, height * 0.22); context.lineTo(width * 0.76, height * 0.42); context.lineTo(width * 0.76, height * 0.3); context.lineTo(width * 0.34, height * 0.3); context.lineTo(width * 0.34, height * 0.82); context.closePath(); void head; return { fill: true };
    }
    case 'u-turn-arrow': {
      context.beginPath(); context.moveTo(width * 0.78, height * 0.8); context.lineTo(width * 0.78, height * 0.3); context.quadraticCurveTo(width * 0.78, height * 0.14, width * 0.62, height * 0.14); context.lineTo(width * 0.32, height * 0.14); context.lineTo(width * 0.32, 0); context.lineTo(0, height * 0.27); context.lineTo(width * 0.32, height * 0.54); context.lineTo(width * 0.32, height * 0.38); context.lineTo(width * 0.55, height * 0.38); context.quadraticCurveTo(width * 0.6, height * 0.38, width * 0.6, height * 0.45); context.lineTo(width * 0.6, height * 0.8); context.closePath(); return { fill: true };
    }
    case 'left-brace': return brace(context, width, height, 'left');
    case 'right-brace': return brace(context, width, height, 'right');
    case 'left-right-brace': return brace(context, width, height, 'both');
    case 'left-bracket': return bracket(context, width, height, 'left');
    case 'right-bracket': return bracket(context, width, height, 'right');
    case 'left-right-bracket': return bracket(context, width, height, 'both');
    case 'callout': return callout(context, width, height, false, false);
    case 'cloud-callout': return callout(context, width, height, true, false);
    case 'wedge-rect-callout': return callout(context, width, height, false, true);
    case 'wedge-round-rect-callout': return callout(context, width, height, false, true);
    case 'star': return star(5, 0.48);
    case 'star4': return star(4, 0.42);
    case 'star5': return star(5, 0.42);
    case 'star6': return star(6, 0.46);
    case 'star8': return star(8, 0.5);
    case 'star16': return star(16, 0.62);
    case 'explosion1': return star(12, 0.34);
    case 'explosion2': return star(20, 0.48);
  }
  return { fill: false };
}

function brace(context: CanvasRenderingContext2D, width: number, height: number, side: 'left' | 'right' | 'both'): ShapeGeometryResult {
  const left = (x: number): void => { context.moveTo(x + width * 0.18, 0); context.bezierCurveTo(x, 0, x, height * 0.18, x + width * 0.18, height * 0.25); context.bezierCurveTo(x + width * 0.32, height * 0.34, x + width * 0.32, height * 0.42, x + width * 0.18, height * 0.5); context.bezierCurveTo(x + width * 0.32, height * 0.58, x + width * 0.32, height * 0.66, x + width * 0.18, height * 0.75); context.bezierCurveTo(x, height * 0.82, x, height, x + width * 0.18, height); };
  context.beginPath();
  if (side === 'left' || side === 'both') left(0);
  if (side === 'both') { context.moveTo(width - width * 0.18, 0); context.bezierCurveTo(width, 0, width, height * 0.18, width - width * 0.18, height * 0.25); context.bezierCurveTo(width - width * 0.32, height * 0.34, width - width * 0.32, height * 0.42, width - width * 0.18, height * 0.5); context.bezierCurveTo(width - width * 0.32, height * 0.58, width - width * 0.32, height * 0.66, width - width * 0.18, height * 0.75); context.bezierCurveTo(width, height * 0.82, width, height, width - width * 0.18, height); }
  if (side === 'right') { context.moveTo(width - width * 0.18, 0); context.bezierCurveTo(width, 0, width, height * 0.18, width - width * 0.18, height * 0.25); context.bezierCurveTo(width - width * 0.32, height * 0.34, width - width * 0.32, height * 0.42, width - width * 0.18, height * 0.5); context.bezierCurveTo(width - width * 0.32, height * 0.58, width - width * 0.32, height * 0.66, width - width * 0.18, height * 0.75); context.bezierCurveTo(width, height * 0.82, width, height, width - width * 0.18, height); }
  return { fill: false };
}

function bracket(context: CanvasRenderingContext2D, width: number, height: number, side: 'left' | 'right' | 'both'): ShapeGeometryResult {
  context.beginPath();
  const draw = (x: number, direction: number): void => { context.moveTo(x + direction * width * 0.2, 0); context.lineTo(x, 0); context.lineTo(x, height); context.lineTo(x + direction * width * 0.2, height); };
  if (side === 'left' || side === 'both') draw(2, 1);
  if (side === 'right' || side === 'both') draw(width - 2, -1);
  return { fill: false };
}

function callout(context: CanvasRenderingContext2D, width: number, height: number, cloud: boolean, wedge: boolean): ShapeGeometryResult {
  context.beginPath();
  if (cloud) {
    context.moveTo(width * 0.14, height * 0.66);
    context.bezierCurveTo(0, height * 0.5, width * 0.1, height * 0.2, width * 0.32, height * 0.3);
    context.bezierCurveTo(width * 0.38, height * 0.04, width * 0.72, height * 0.08, width * 0.74, height * 0.32);
    context.bezierCurveTo(width, height * 0.22, width * 1.02, height * 0.6, width * 0.82, height * 0.68);
    context.lineTo(width * 0.34, height * 0.68); context.lineTo(width * 0.18, height); context.lineTo(width * 0.24, height * 0.66); context.closePath();
  } else {
    const radius = wedge ? 4 : 8;
    context.roundRect(0, 0, width, height * 0.76, radius);
    context.moveTo(width * 0.26, height * 0.76); context.lineTo(width * 0.18, height); context.lineTo(width * 0.42, height * 0.76); context.closePath();
  }
  return { fill: true };
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

export function drawScreenshotOnCanvas(context: CanvasRenderingContext2D, payload: ScreenshotDrawingPayload, bounds: Rect, getSheet: (sheetId: string) => CanvasSheetSnapshot | undefined): void {
  drawCameraOnCanvas(context, { kind: 'camera', sourceRange: payload.sourceRange, refreshPolicy: 'live' }, bounds, getSheet);
}

function drawIconOnCanvas(context: CanvasRenderingContext2D, payload: IconDrawingPayload, bounds: Rect): void {
  context.save();
  context.fillStyle = '#ffffff';
  context.fillRect(bounds.x, bounds.y, bounds.width, bounds.height);
  context.strokeStyle = '#cbd5e1';
  context.strokeRect(bounds.x + 0.5, bounds.y + 0.5, bounds.width - 1, bounds.height - 1);
  // The SVG path is persisted with the object. Canvas uses a path fallback for
  // deterministic rendering while the DOM inspector can still expose the same path.
  const scale = Math.min(bounds.width, bounds.height) / 24;
  const path = new Path2D(payload.svgPath);
  context.save();
  context.translate(bounds.x + (bounds.width - 24 * scale) / 2, bounds.y + (bounds.height - 24 * scale) / 2);
  context.scale(scale, scale);
  context.fillStyle = payload.fill;
  context.fill(path);
  context.restore();
  context.restore();
}

function drawModel3dOnCanvas(context: CanvasRenderingContext2D, payload: Model3dDrawingPayload, bounds: Rect): void {
  const { vertices, faces } = payload.geometry;
  const angle = payload.rotation.y;
  const cos = Math.cos(angle);
  const sin = Math.sin(angle);
  const points = vertices.map((vertex) => {
    const x = vertex.x * cos - vertex.z * sin;
    const z = vertex.x * sin + vertex.z * cos;
    const perspective = 1 / Math.max(0.25, 2.8 - z * 0.25);
    return { x: bounds.x + bounds.width / 2 + x * payload.scale * bounds.width * 0.35 * perspective, y: bounds.y + bounds.height / 2 - vertex.y * payload.scale * bounds.height * 0.35 * perspective };
  });
  context.save();
  context.fillStyle = '#f8fafc'; context.fillRect(bounds.x, bounds.y, bounds.width, bounds.height);
  context.strokeStyle = '#64748b'; context.lineWidth = 1;
  for (const face of faces) {
    const [a, b, c] = face.map((index) => points[index]!);
    context.beginPath(); context.moveTo(a!.x, a!.y); context.lineTo(b!.x, b!.y); context.lineTo(c!.x, c!.y); context.closePath(); context.stroke();
  }
  context.strokeStyle = '#2563eb'; context.strokeRect(bounds.x + 0.5, bounds.y + 0.5, bounds.width - 1, bounds.height - 1);
  context.restore();
}

function drawSmartArtOnCanvas(context: CanvasRenderingContext2D, payload: SmartArtDrawingPayload, bounds: Rect): void {
  context.save();
  const gap = 12;
  const columns = payload.layout === 'process' || payload.layout === 'list' ? Math.min(4, payload.nodes.length) : 1;
  const rows = Math.max(1, Math.ceil(payload.nodes.length / columns));
  const nodeWidth = Math.max(48, (bounds.width - gap * (columns - 1) - 16) / columns);
  const nodeHeight = Math.max(28, (bounds.height - gap * (rows - 1) - 16) / rows);
  const positions = new Map<string, { x: number; y: number }>();
  payload.nodes.forEach((node, index) => { const column = index % columns; const row = Math.floor(index / columns); positions.set(node.id, { x: bounds.x + 8 + column * (nodeWidth + gap), y: bounds.y + 8 + row * (nodeHeight + gap) }); });
  context.strokeStyle = payload.stroke; context.lineWidth = 1.5;
  for (const edge of payload.edges) { const start = positions.get(edge.from); const end = positions.get(edge.to); if (!start || !end) continue; context.beginPath(); context.moveTo(start.x + nodeWidth / 2, start.y + nodeHeight / 2); context.lineTo(end.x + nodeWidth / 2, end.y + nodeHeight / 2); context.stroke(); }
  context.font = '12px Segoe UI, sans-serif'; context.textAlign = 'center'; context.textBaseline = 'middle';
  for (const node of payload.nodes) { const position = positions.get(node.id)!; context.fillStyle = payload.fill; context.fillRect(position.x, position.y, nodeWidth, nodeHeight); context.strokeStyle = payload.stroke; context.strokeRect(position.x + 0.5, position.y + 0.5, nodeWidth - 1, nodeHeight - 1); context.fillStyle = payload.textColor; context.fillText(node.text, position.x + nodeWidth / 2, position.y + nodeHeight / 2, nodeWidth - 10); }
  context.restore();
}

function drawWordArtOnCanvas(context: CanvasRenderingContext2D, payload: WordArtDrawingPayload, bounds: Rect, rotation = 0): void {
  context.save(); context.translate(bounds.x + bounds.width / 2, bounds.y + bounds.height / 2); context.rotate(rotation * Math.PI / 180);
  context.font = `${payload.italic ? 'italic ' : ''}${payload.bold ? 'bold ' : ''}${payload.fontSize}px ${payload.fontFamily}`; context.textAlign = 'center'; context.textBaseline = 'middle'; context.lineWidth = payload.outlineWidth; context.strokeStyle = payload.outline; context.fillStyle = payload.fill; context.strokeText(payload.text, 0, 0, Math.max(1, bounds.width - 8)); context.fillText(payload.text, 0, 0, Math.max(1, bounds.width - 8)); context.restore();
}

function drawSignatureLineOnCanvas(context: CanvasRenderingContext2D, payload: SignatureLineDrawingPayload, bounds: Rect): void {
  context.save(); context.fillStyle = '#fff'; context.fillRect(bounds.x, bounds.y, bounds.width, bounds.height); context.strokeStyle = '#334155'; context.beginPath(); context.moveTo(bounds.x + 8, bounds.y + bounds.height * 0.62); context.lineTo(bounds.x + bounds.width - 8, bounds.y + bounds.height * 0.62); context.stroke(); context.fillStyle = '#334155'; context.font = '12px Segoe UI, sans-serif'; context.textAlign = 'left'; context.textBaseline = 'top'; context.fillText(payload.signerName, bounds.x + 8, bounds.y + 6); if (payload.signerTitle) context.fillText(payload.signerTitle, bounds.x + 8, bounds.y + 21); context.textAlign = 'right'; context.fillStyle = payload.status === 'signed' ? '#15803d' : '#64748b'; context.fillText(payload.status === 'signed' ? `已签名 ${payload.signedBy ?? ''}` : '未签名', bounds.x + bounds.width - 8, bounds.y + 6); context.restore();
}

function drawEmbeddedObjectOnCanvas(context: CanvasRenderingContext2D, payload: EmbeddedObjectDrawingPayload, bounds: Rect): void {
  context.save(); context.fillStyle = '#f8fafc'; context.fillRect(bounds.x, bounds.y, bounds.width, bounds.height); context.strokeStyle = '#94a3b8'; context.strokeRect(bounds.x + 0.5, bounds.y + 0.5, bounds.width - 1, bounds.height - 1); context.fillStyle = '#334155'; context.font = '12px Segoe UI, sans-serif'; context.textAlign = 'center'; context.textBaseline = 'middle'; context.fillText(payload.fileName, bounds.x + bounds.width / 2, bounds.y + bounds.height / 2 - 8, Math.max(10, bounds.width - 16)); context.fillStyle = '#64748b'; context.font = '10px Segoe UI, sans-serif'; context.fillText(payload.relationship === 'embedded' ? '本地嵌入对象' : '本地链接对象', bounds.x + bounds.width / 2, bounds.y + bounds.height / 2 + 10); context.restore();
}

function drawEquationOnCanvas(context: CanvasRenderingContext2D, payload: EquationDrawingPayload, bounds: Rect): void {
  context.save(); context.fillStyle = '#fff'; context.fillRect(bounds.x, bounds.y, bounds.width, bounds.height); context.fillStyle = payload.textColor; context.font = `${payload.fontSize}px Cambria Math, Cambria, serif`; context.textAlign = 'center'; context.textBaseline = 'middle'; context.fillText(payload.tokens.map((token) => token.value).join(''), bounds.x + bounds.width / 2, bounds.y + bounds.height / 2, Math.max(10, bounds.width - 12)); context.restore();
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

function chartFillColor(fill: string | { kind: string; color?: string; secondaryColor?: string; transparency?: number } | undefined, fallback: string): string {
  if (!fill) return fallback;
  if (typeof fill === 'string') return fill;
  if (fill.kind === 'none') return 'transparent';
  return fill.color ?? fallback;
}

function paintChartFill(context: CanvasRenderingContext2D, fill: string | { kind: string; color?: string; secondaryColor?: string; transparency?: number; angle?: number; pattern?: string } | undefined, rect: { x: number; y: number; width: number; height: number }, fallback: string): boolean {
  if (!fill) { context.fillStyle = fallback; return true; }
  if (typeof fill === 'string') { context.fillStyle = fill; return true; }
  if (fill.kind === 'none') return false;
  if (fill.kind === 'gradient') {
    const angle = (fill.angle ?? 90) * Math.PI / 180;
    const x = Math.cos(angle) * rect.width;
    const y = Math.sin(angle) * rect.height;
    const gradient = context.createLinearGradient(rect.x, rect.y, rect.x + x, rect.y + y);
    gradient.addColorStop(0, fill.color ?? fallback);
    gradient.addColorStop(1, fill.secondaryColor ?? fill.color ?? fallback);
    context.fillStyle = gradient;
    return true;
  }
  context.fillStyle = fill.color ?? fallback;
  return true;
}

function drawChartText(context: CanvasRenderingContext2D, text: string, x: number, y: number, options?: { color?: string; size?: number; bold?: boolean; align?: CanvasTextAlign }): void {
  context.fillStyle = options?.color ?? '#334155';
  context.font = `${options?.bold ? '600 ' : ''}${options?.size ?? 11}px Segoe UI, sans-serif`;
  context.textAlign = options?.align ?? 'left';
  context.textBaseline = 'middle';
  context.fillText(text, x, y);
}

function chartScale(value: number, axis: NonNullable<ChartLayout['valueAxis']>): number {
  const axisModel = axis.model;
  if (axisModel.scale === 'logarithmic') {
    const base = axisModel.logBase ?? 10;
    const min = Math.log(Math.max(Number.MIN_VALUE, axis.minimum)) / Math.log(base);
    const max = Math.log(Math.max(Number.MIN_VALUE, axis.maximum)) / Math.log(base);
    return (Math.log(Math.max(Number.MIN_VALUE, value)) / Math.log(base) - min) / Math.max(Number.MIN_VALUE, max - min);
  }
  return (value - axis.minimum) / Math.max(Number.MIN_VALUE, axis.maximum - axis.minimum);
}

function drawChartMarker(context: CanvasRenderingContext2D, x: number, y: number, marker: ChartMarkerModel | undefined, color: string): void {
  if (!marker?.enabled) return;
  const radius = Math.max(2, (marker.size ?? 6) / 2);
  const shape = marker.shape ?? 'circle';
  context.save();
  context.fillStyle = marker.fill ?? color;
  context.strokeStyle = marker.border ?? color;
  context.lineWidth = 1;
  context.beginPath();
  if (shape === 'square') context.rect(x - radius, y - radius, radius * 2, radius * 2);
  else if (shape === 'diamond') { context.moveTo(x, y - radius); context.lineTo(x + radius, y); context.lineTo(x, y + radius); context.lineTo(x - radius, y); context.closePath(); }
  else if (shape === 'triangle') { context.moveTo(x, y - radius); context.lineTo(x + radius, y + radius); context.lineTo(x - radius, y + radius); context.closePath(); }
  else context.arc(x, y, radius, 0, Math.PI * 2);
  context.fill();
  context.stroke();
  context.restore();
}

function drawChartLegend(context: CanvasRenderingContext2D, layout: ChartLayout): void {
  if (!layout.legend.visible) return;
  const position = layout.legend.position;
  const entries = layout.series.filter((series) => series.visible);
  let x = position === 'left' ? 8 : position === 'right' ? layout.width - 92 : 16;
  let y = position === 'top' || position === 'top-right' ? 22 : layout.height - 16;
  if (position === 'top-right') { x = layout.width - 112; y = 22; }
  for (const [index, series] of entries.entries()) {
    context.fillStyle = series.color;
    context.fillRect(x, y - 5, 10, 10);
    drawChartText(context, series.name, x + 14, y, { color: '#475569', size: 10 });
    x += Math.max(54, context.measureText(series.name).width + 32);
    if ((position === 'left' || position === 'right') && y + 18 < layout.height - 4) y += 18;
    else if (x > layout.width - 36 && index < entries.length - 1) { x = 16; y += 16; }
  }
}

function drawChartAxes(context: CanvasRenderingContext2D, layout: ChartLayout, categories: readonly PivotScalar[]): void {
  const { plot } = layout;
  const valueAxis = layout.valueAxis;
  if (!valueAxis) return;
  const valueAxisModel = valueAxis.model;
  const grid = valueAxisModel.majorGridlines;
  for (const tick of valueAxis.ticks) {
    const y = plot.top + plot.height * (1 - chartScale(tick, valueAxis));
    if (grid?.visible !== false) {
      context.save();
      context.strokeStyle = grid?.color ?? '#e2e8f0';
      context.lineWidth = grid?.width ?? 1;
      if (grid?.dash === 'dash') context.setLineDash([4, 3]);
      if (grid?.dash === 'dot') context.setLineDash([1, 3]);
      context.beginPath();
      context.moveTo(plot.left, y);
      context.lineTo(plot.left + plot.width, y);
      context.stroke();
      context.restore();
    }
    drawChartText(context, String(Math.round(tick * 100000) / 100000), plot.left - 8, y, { color: '#64748b', size: 9, align: 'right' });
  }
  if (valueAxisModel.visible !== false) {
    context.strokeStyle = valueAxisModel.line?.color ?? '#94a3b8';
    context.lineWidth = valueAxisModel.line?.width ?? 1;
    context.beginPath();
    context.moveTo(plot.left, plot.top);
    context.lineTo(plot.left, plot.top + plot.height);
    context.stroke();
  }
  const categoryAxis = layout.categoryAxis;
  const categoryAxisModel = categoryAxis?.model;
  if (categoryAxisModel?.visible !== false) {
    context.strokeStyle = categoryAxisModel?.line?.color ?? '#94a3b8';
    context.lineWidth = categoryAxisModel?.line?.width ?? 1;
    context.beginPath();
    context.moveTo(plot.left, plot.top + plot.height);
    context.lineTo(plot.left + plot.width, plot.top + plot.height);
    context.stroke();
    const interval = Math.max(1, categoryAxisModel?.labelInterval ?? 1);
    const count = Math.max(1, categories.length);
    categories.forEach((category, index) => {
      if (index % interval !== 0) return;
      const x = plot.left + (index + 0.5) * plot.width / count;
      drawChartText(context, String(category ?? ''), x, plot.top + plot.height + 12, { color: '#64748b', size: 9, align: 'center' });
    });
  }
  if (valueAxisModel.title) drawChartText(context, valueAxisModel.title, 12, plot.top + plot.height / 2, { color: '#475569', size: 10 });
  if (categoryAxisModel?.title) drawChartText(context, categoryAxisModel.title, plot.left + plot.width / 2, layout.height - 6, { color: '#475569', size: 10, align: 'center' });
}

function drawChartLine(context: CanvasRenderingContext2D, series: ChartLayout['series'][number], area: boolean, smooth: boolean, baseline = 0, emptyCells: NonNullable<ChartDrawingPayload['elements']['emptyCells']> = 'gap', pixelsPerValue = 1): void {
  const points = series.points.filter((point) => point.visible);
  if (!points.length) return;
  const segments: typeof points[] = [];
  if (emptyCells === 'connect') segments.push(points);
  let current: typeof points = [];
  if (emptyCells !== 'connect') {
    for (const point of series.points) {
      if (!point.visible) { if (current.length) segments.push(current); current = []; }
      else current.push(point);
    }
    if (current.length) segments.push(current);
  }
  for (const segment of segments) {
    if (!segment.length) continue;
    context.save();
    context.strokeStyle = series.color;
    context.lineWidth = 2;
    if (area) {
      context.fillStyle = `${series.color}22`;
      context.beginPath();
      context.moveTo(segment[0]!.x, baseline);
      segment.forEach((point) => context.lineTo(point.x, point.y));
      context.lineTo(segment.at(-1)!.x, baseline);
      context.closePath();
      context.fill();
    }
    context.beginPath();
    context.moveTo(segment[0]!.x, segment[0]!.y);
    for (let index = 1; index < segment.length; index += 1) {
      const point = segment[index]!;
      const previous = segment[index - 1]!;
      if (smooth) {
        const middle = (previous.x + point.x) / 2;
        context.quadraticCurveTo(middle, previous.y, middle, (previous.y + point.y) / 2);
        context.quadraticCurveTo(middle, point.y, point.x, point.y);
      } else context.lineTo(point.x, point.y);
    }
    context.stroke();
    for (const point of segment) drawChartMarker(context, point.x, point.y, series.subtype?.includes('markers') ? { enabled: true } : undefined, series.color);
    for (const point of segment) if (point.errorPlus || point.errorMinus) {
      const plus = point.errorPlus ?? 0;
      const minus = point.errorMinus ?? 0;
      const up = plus * pixelsPerValue;
      const down = minus * pixelsPerValue;
      context.strokeStyle = series.color;
      context.beginPath();
      context.moveTo(point.x, point.y - up); context.lineTo(point.x, point.y + down);
      context.moveTo(point.x - 3, point.y - up); context.lineTo(point.x + 3, point.y - up);
      context.moveTo(point.x - 3, point.y + down); context.lineTo(point.x + 3, point.y + down);
      context.stroke();
    }
    context.restore();
  }
  for (const trendline of series.trendlines) {
    const trendlineModel = trendline.model;
    context.save();
    context.strokeStyle = trendlineModel.color ?? series.color;
    context.lineWidth = trendlineModel.width ?? 1.5;
    if (trendlineModel.dash === 'dash') context.setLineDash([5, 3]);
    if (trendlineModel.dash === 'dot') context.setLineDash([1, 3]);
    context.beginPath();
    trendline.points.forEach((point, index) => index === 0 ? context.moveTo(point.x, point.y) : context.lineTo(point.x, point.y));
    context.stroke();
    context.restore();
  }
}

function drawScatterSeries(context: CanvasRenderingContext2D, series: ChartLayout['series'][number], bubble: boolean): void {
  const points = series.points.filter((point) => point.visible && point.xValue !== null && (!bubble || point.sizeValue !== null));
  if (!points.length) return;
  const subtype = series.subtype ?? '';
  const hasLine = subtype.includes('lines');
  if (hasLine && points.length > 1) {
    context.save(); context.strokeStyle = series.color; context.lineWidth = 2; context.beginPath();
    points.forEach((point, index) => { if (index === 0) context.moveTo(point.x, point.y); else if (subtype.includes('smooth')) { const previous = points[index - 1]!; const middle = (previous.x + point.x) / 2; context.quadraticCurveTo(middle, previous.y, point.x, point.y); } else context.lineTo(point.x, point.y); });
    context.stroke(); context.restore();
  }
  const sizes = points.map((point) => Math.abs(point.sizeValue ?? 1));
  const maxSize = Math.max(1, ...sizes);
  for (const point of points) {
    const radius = bubble ? Math.max(3, Math.min(24, 3 + Math.sqrt(Math.abs(point.sizeValue ?? 0) / maxSize) * 18)) : 4;
    context.save(); context.globalAlpha = bubble ? 0.72 : 1; context.fillStyle = series.color; context.strokeStyle = series.color; context.beginPath(); context.arc(point.x, point.y, radius, 0, Math.PI * 2); context.fill(); context.stroke(); context.restore();
  }
}

function drawChartDataLabels(context: CanvasRenderingContext2D, payload: ChartDrawingPayload, series: ChartLayout['series'][number]): void {
  const chartLabels = payload.elements.dataLabels;
  for (const point of series.points) {
    const labels = payload.series?.find((entry) => entry.id === series.id || entry.name === series.name)?.dataLabels ?? chartLabels;
    if (!labels?.visible || point.value === null) continue;
    const parts: string[] = [];
    if (labels.showSeriesName) parts.push(series.name);
    if (labels.showCategoryName) parts.push(String(point.category));
    if (labels.showValue !== false) parts.push(String(point.value));
    if (labels.showPercentage) parts.push(`${Math.round(Math.abs(point.value) * 100) / 100}%`);
    if (!parts.length) parts.push(String(point.value));
    const position = labels.position === 'inside-base' ? { x: point.x, y: point.y + 10 } : labels.position === 'below' ? { x: point.x, y: point.y + 14 } : { x: point.x, y: point.y - 10 };
    drawChartText(context, parts.join(labels.separator ?? ', '), position.x, position.y, { color: '#334155', size: 9, align: 'center' });
  }
}

function drawChartSpecial(context: CanvasRenderingContext2D, payload: ChartDrawingPayload, layout: ChartLayout): void {
  const { plot } = layout;
  if (layout.kind === 'pie') {
    const centerX = plot.left + plot.width / 2;
    const centerY = plot.top + plot.height / 2;
    for (const slice of layout.pieSlices ?? []) {
      const mid = (slice.startAngle + slice.endAngle) / 2;
      const offsetX = Math.cos(mid) * slice.explosion;
      const offsetY = Math.sin(mid) * slice.explosion;
      context.save();
      context.translate(offsetX, offsetY);
      context.fillStyle = slice.color;
      context.beginPath();
      context.moveTo(centerX + Math.cos(slice.startAngle) * slice.innerRadius, centerY + Math.sin(slice.startAngle) * slice.innerRadius);
      context.arc(centerX, centerY, slice.outerRadius, slice.startAngle, slice.endAngle);
      if (slice.innerRadius > 0) context.arc(centerX, centerY, slice.innerRadius, slice.endAngle, slice.startAngle, true);
      else context.lineTo(centerX, centerY);
      context.closePath();
      context.fill();
      context.restore();
    }
    return;
  }
  if (layout.kind === 'histogram') {
    const bins = layout.histogramBins ?? [];
    const maximum = Math.max(1, ...bins.map((bin) => bin.count));
    const width = plot.width / Math.max(1, bins.length);
    bins.forEach((bin, index) => {
      const barHeight = bin.count / maximum * plot.height;
      context.fillStyle = '#2563eb';
      context.fillRect(plot.left + index * width, plot.top + plot.height - barHeight, Math.max(1, width - 1), barHeight);
      drawChartText(context, bin.label, plot.left + (index + 0.5) * width, plot.top + plot.height + 12, { size: 8, align: 'center' });
    });
    if (layout.paretoPoints?.length) {
      context.strokeStyle = '#dc2626';
      context.lineWidth = 2;
      context.beginPath();
      layout.paretoPoints.forEach((point, index) => index === 0 ? context.moveTo(point.x, point.y) : context.lineTo(point.x, point.y));
      context.stroke();
    }
    return;
  }
  if (layout.kind === 'box-whisker') {
    const boxes = layout.boxes ?? [];
    const all = boxes.flatMap((box) => [box.minimum, box.maximum, ...box.outliers]);
    const minimum = Math.min(0, ...all);
    const maximum = Math.max(1, ...all);
    const axis = { minimum, maximum, model: { scale: 'linear' } } as NonNullable<ChartLayout['valueAxis']>;
    const y = (value: number) => plot.top + plot.height * (1 - chartScale(value, axis));
    const slot = plot.width / Math.max(1, boxes.length);
    boxes.forEach((box, index) => {
      const x = plot.left + slot * (index + 0.5);
      context.strokeStyle = box.color;
      context.lineWidth = 1.5;
      context.beginPath();
      context.moveTo(x, y(box.lowerWhisker)); context.lineTo(x, y(box.upperWhisker));
      context.moveTo(x - slot * 0.12, y(box.lowerWhisker)); context.lineTo(x + slot * 0.12, y(box.lowerWhisker));
      context.moveTo(x - slot * 0.12, y(box.upperWhisker)); context.lineTo(x + slot * 0.12, y(box.upperWhisker));
      context.stroke();
      context.fillStyle = `${box.color}33`;
      context.fillRect(x - slot * 0.24, y(box.q3), slot * 0.48, Math.max(1, y(box.q1) - y(box.q3)));
      context.strokeRect(x - slot * 0.24, y(box.q3), slot * 0.48, Math.max(1, y(box.q1) - y(box.q3)));
      context.beginPath(); context.moveTo(x - slot * 0.24, y(box.median)); context.lineTo(x + slot * 0.24, y(box.median)); context.stroke();
      for (const outlier of box.outliers) { context.fillStyle = box.color; context.beginPath(); context.arc(x, y(outlier), 2, 0, Math.PI * 2); context.fill(); }
    });
    return;
  }
  if (layout.kind === 'waterfall') {
    const bars = layout.waterfallBars ?? [];
    const minimum = Math.min(0, ...bars.map((bar) => bar.start));
    const maximum = Math.max(1, ...bars.map((bar) => bar.end));
    const span = Math.max(1, maximum - minimum);
    const slot = plot.width / Math.max(1, bars.length);
    bars.forEach((bar) => {
      const x = plot.left + bar.index * slot + slot * 0.16;
      const top = plot.top + plot.height * (1 - (bar.end - minimum) / span);
      const bottom = plot.top + plot.height * (1 - (bar.start - minimum) / span);
      context.fillStyle = bar.color;
      context.fillRect(x, Math.min(top, bottom), slot * 0.68, Math.max(1, Math.abs(bottom - top)));
      if (payload.waterfallOptions?.connectorLines !== false && bar.index > 0) { context.strokeStyle = '#94a3b8'; context.beginPath(); context.moveTo(x - slot * 0.16, bottom); context.lineTo(x, bottom); context.stroke(); }
    });
    return;
  }
  if (layout.kind === 'funnel') {
    const stages = layout.funnelStages ?? [];
    const maximum = Math.max(1, ...stages.map((stage) => stage.value));
    const band = plot.height / Math.max(1, stages.length);
    stages.forEach((stage) => {
      const current = plot.width * stage.value / maximum;
      const next = plot.width * stage.nextValue / maximum;
      const center = plot.left + plot.width / 2;
      context.fillStyle = stage.color;
      context.beginPath();
      context.moveTo(center - current / 2, plot.top + stage.index * band);
      context.lineTo(center + current / 2, plot.top + stage.index * band);
      context.lineTo(center + next / 2, plot.top + (stage.index + 1) * band);
      context.lineTo(center - next / 2, plot.top + (stage.index + 1) * band);
      context.closePath(); context.fill();
      drawChartText(context, `${stage.label}: ${stage.value}`, center, plot.top + (stage.index + 0.5) * band, { color: '#fff', size: 9, align: 'center' });
    });
    return;
  }
  if (layout.kind === 'stock') {
    const points = layout.stockPoints ?? [];
    const values = points.flatMap((point) => [point.high, point.low, point.close, point.open ?? point.close]);
    const minimum = Math.min(...values, 0);
    const maximum = Math.max(...values, 1);
    const span = Math.max(1, maximum - minimum);
    const slot = plot.width / Math.max(1, points.length);
    for (const point of points) {
      const x = plot.left + (point.index + 0.5) * slot;
      const y = (value: number) => plot.top + plot.height * (1 - (value - minimum) / span);
      context.strokeStyle = point.color;
      context.beginPath(); context.moveTo(x, y(point.high)); context.lineTo(x, y(point.low)); context.stroke();
      if (point.open !== undefined) { const top = Math.min(y(point.open), y(point.close)); context.fillStyle = point.color; context.fillRect(x - slot * 0.2, top, Math.max(2, slot * 0.4), Math.max(1, Math.abs(y(point.open) - y(point.close)))); }
      else { context.fillStyle = point.color; context.fillRect(x - 2, y(point.close) - 2, 4, 4); }
    }
    return;
  }
  if (layout.kind === 'surface') {
    const cells = layout.surfaceCells ?? [];
    const rows = Math.max(1, ...cells.map((cell) => cell.row + 1));
    const columns = Math.max(1, ...cells.map((cell) => cell.column + 1));
    for (const cell of cells) { context.fillStyle = cell.color; context.fillRect(plot.left + cell.column * plot.width / columns, plot.top + cell.row * plot.height / rows, Math.ceil(plot.width / columns), Math.ceil(plot.height / rows)); }
    return;
  }
  if (layout.kind === 'radar') {
    const radar = layout.radar;
    if (!radar) return;
    const centerX = plot.left + plot.width / 2;
    const centerY = plot.top + plot.height / 2;
    const radius = Math.min(plot.width, plot.height) * 0.42;
    for (let ring = 1; ring <= 4; ring += 1) {
      context.strokeStyle = '#cbd5e1'; context.beginPath();
      for (let index = 0; index < radar.count; index += 1) { const angle = -Math.PI / 2 + Math.PI * 2 * index / radar.count; const r = radius * ring / 4; const x = centerX + Math.cos(angle) * r; const y = centerY + Math.sin(angle) * r; index === 0 ? context.moveTo(x, y) : context.lineTo(x, y); }
      context.closePath(); context.stroke();
    }
    for (const entry of radar.points) {
      context.strokeStyle = entry.color; context.fillStyle = `${entry.color}22`; context.beginPath();
      entry.values.forEach((value, index) => { const angle = -Math.PI / 2 + Math.PI * 2 * index / radar.count; const r = radius * Math.abs(value) / radar.maximum; const x = centerX + Math.cos(angle) * r; const y = centerY + Math.sin(angle) * r; index === 0 ? context.moveTo(x, y) : context.lineTo(x, y); });
      context.closePath(); context.fill(); context.stroke();
    }
    return;
  }
  if (layout.kind === 'treemap') {
    const values = layout.series.flatMap((series) => series.points.filter((point) => point.visible).map((point) => ({ value: Math.max(0, point.value ?? 0), label: String(point.category), color: series.color })));
    const total = values.reduce((sum, entry) => sum + entry.value, 0) || 1;
    let x = plot.left;
    for (const [index, entry] of values.entries()) { const width = plot.width * entry.value / total; context.fillStyle = entry.color ?? '#2563eb'; context.fillRect(x, plot.top, Math.max(1, width - 1), plot.height); drawChartText(context, entry.label, x + width / 2, plot.top + plot.height / 2, { color: '#fff', size: 9, align: 'center' }); x += width; void index; }
    return;
  }
  if (layout.kind === 'sunburst') {
    const values = layout.series.flatMap((series) => series.points.filter((point) => point.visible).map((point) => Math.max(0, point.value ?? 0)));
    const total = values.reduce((sum, value) => sum + value, 0) || 1;
    const centerX = plot.left + plot.width / 2;
    const centerY = plot.top + plot.height / 2;
    const radius = Math.min(plot.width, plot.height) * 0.42;
    let angle = -Math.PI / 2;
    values.forEach((value, index) => { const sweep = value / total * Math.PI * 2; context.fillStyle = CHART_PALETTE[index % CHART_PALETTE.length]!; context.beginPath(); context.moveTo(centerX, centerY); context.arc(centerX, centerY, radius, angle, angle + sweep); context.closePath(); context.fill(); angle += sweep; });
    context.fillStyle = '#fff'; context.beginPath(); context.arc(centerX, centerY, radius * 0.42, 0, Math.PI * 2); context.fill();
    return;
  }
}

function drawChartLayoutOnCanvas(options: { context: CanvasRenderingContext2D; payload: ChartDrawingPayload; bounds: Rect; layout: ChartLayout }): void {
  const { context, payload, bounds, layout } = options;
  context.save();
  context.translate(bounds.x, bounds.y);
  const chartArea = payload.elements.chartArea;
  const fill = chartArea?.fill;
  const chartHasFill = paintChartFill(context, fill, { x: 0, y: 0, width: bounds.width, height: bounds.height }, '#fff');
  if (chartHasFill) { context.globalAlpha = 1 - (chartArea?.transparency ?? 0); context.fillRect(0, 0, bounds.width, bounds.height); context.globalAlpha = 1; }
  context.strokeStyle = chartArea?.line?.color ?? chartArea?.border ?? '#cbd5e1';
  context.lineWidth = chartArea?.line?.width ?? chartArea?.borderWidth ?? 1;
  if (chartArea?.line?.dash === 'dash' || chartArea?.borderDash === 'dash') context.setLineDash([5, 3]);
  if (chartArea?.line?.dash === 'dot' || chartArea?.borderDash === 'dot') context.setLineDash([1, 3]);
  context.strokeRect(0.5, 0.5, Math.max(0, bounds.width - 1), Math.max(0, bounds.height - 1));
  context.setLineDash([]);
  if (layout.title) drawChartText(context, layout.title.text, layout.title.x, layout.title.y, { color: payload.elements.titleText?.color ?? '#1e293b', size: payload.elements.titleText?.fontSize ?? 14, bold: payload.elements.titleText?.bold !== false });
  if (layout.status.kind !== 'ready') {
    context.strokeStyle = '#dc2626'; context.setLineDash([4, 3]); context.strokeRect(6, Math.max(6, layout.plot.top), Math.max(0, bounds.width - 12), Math.max(12, bounds.height - layout.plot.top - 8)); context.setLineDash([]);
    drawChartText(context, layout.status.message ?? `${layout.status.code ?? 'UNSUPPORTED_FEATURE'}: chart data is unavailable`, 14, Math.min(bounds.height - 16, layout.plot.top + 18), { color: '#b91c1c', size: 11 });
    context.restore();
    return;
  }
  const plotArea = payload.elements.plotArea;
  const plotHasFill = paintChartFill(context, plotArea?.fill, { x: layout.plot.left, y: layout.plot.top, width: layout.plot.width, height: layout.plot.height }, 'transparent');
  if (plotHasFill) { context.globalAlpha = 1 - (plotArea?.transparency ?? 0); context.fillRect(layout.plot.left, layout.plot.top, layout.plot.width, layout.plot.height); context.globalAlpha = 1; }
  if (layout.kind === 'cartesian') {
    const categories = layout.series[0]?.points.map((point) => point.category) ?? [];
    drawChartAxes(context, layout, categories);
    for (const series of layout.series) {
      if (!series.visible) continue;
      if (series.bars.length) for (const bar of series.bars) { if (!bar.visible) continue; context.fillStyle = bar.color; context.fillRect(bar.x, bar.y, bar.width, bar.height); }
      if (series.chartType === 'line' || series.chartType === 'area') {
        const valueAxis = series.axis === 'secondary' ? layout.secondaryValueAxis ?? layout.valueAxis : layout.valueAxis;
        const pixelsPerValue = valueAxis ? layout.plot.height / Math.max(Number.EPSILON, valueAxis.maximum - valueAxis.minimum) : 1;
        drawChartLine(context, series, series.chartType === 'area' || payload.chartType === 'area', series.subtype?.includes('smooth') === true || series.smooth === true, layout.plot.top + layout.plot.height, payload.elements.emptyCells ?? 'gap', pixelsPerValue);
      }
      if (payload.chartType === 'scatter' || payload.chartType === 'bubble') drawScatterSeries(context, series, payload.chartType === 'bubble');
      drawChartDataLabels(context, payload, series);
    }
  } else drawChartSpecial(context, payload, layout);
  if (payload.elements.dataTable?.visible) {
    const y = Math.min(bounds.height - 8, layout.plot.top + layout.plot.height + 28);
    drawChartText(context, payload.elements.dataTable.showLegendKeys === false ? 'Chart Data Table' : 'Chart Data Table · Legend Keys', 8, y, { color: '#475569', size: 9 });
  }
  drawChartLegend(context, layout);
  context.restore();
}

function chartHitTest(layout: ChartLayout, point: { x: number; y: number }, dataTableVisible = false): { action: string; data: unknown } | null {
  if (layout.title && point.y <= layout.title.y + 18) return { action: 'chart.select-element', data: { kind: 'title' } };
  if (layout.legend.visible) {
    const legendBand = layout.legend.position === 'bottom' ? { left: 0, top: layout.height - 30, right: layout.width, bottom: layout.height }
      : layout.legend.position === 'top' || layout.legend.position === 'top-right' ? { left: 0, top: 0, right: layout.width, bottom: 34 }
        : layout.legend.position === 'left' ? { left: 0, top: 0, right: layout.plot.left - 8, bottom: layout.height }
          : { left: layout.plot.left + layout.plot.width + 8, top: 0, right: layout.width, bottom: layout.height };
    if (point.x >= legendBand.left && point.x <= legendBand.right && point.y >= legendBand.top && point.y <= legendBand.bottom) return { action: 'chart.select-element', data: { kind: 'legend' } };
  }
  if (dataTableVisible && point.y >= layout.plot.top + layout.plot.height + 18) return { action: 'chart.select-element', data: { kind: 'data-table' } };
  if (layout.kind === 'pie') {
    const centerX = layout.plot.left + layout.plot.width / 2;
    const centerY = layout.plot.top + layout.plot.height / 2;
    const radius = Math.hypot(point.x - centerX, point.y - centerY);
    const angle = Math.atan2(point.y - centerY, point.x - centerX);
    for (const slice of layout.pieSlices ?? []) {
      let normalized = angle;
      while (normalized < slice.startAngle) normalized += Math.PI * 2;
      if (normalized >= slice.startAngle && normalized <= slice.endAngle && radius >= slice.innerRadius && radius <= slice.outerRadius + slice.explosion) {
        const series = layout.series[slice.seriesIndex];
        if (series) return { action: 'chart.select-element', data: { kind: 'point', seriesId: series.id, pointIndex: slice.pointIndex } };
      }
    }
  }
  for (const series of layout.series) {
    for (const chartPoint of series.points) {
      if (!chartPoint.visible) continue;
      if (Math.hypot(point.x - chartPoint.x, point.y - chartPoint.y) <= 7) return { action: 'chart.select-element', data: { kind: 'point', seriesId: series.id, pointIndex: chartPoint.index } };
    }
    for (const bar of series.bars) if (bar.visible && point.x >= bar.x && point.x <= bar.x + bar.width && point.y >= bar.y && point.y <= bar.y + bar.height) return { action: 'chart.select-element', data: { kind: 'point', seriesId: series.id, pointIndex: bar.index } };
    for (const trendline of series.trendlines) {
      for (let index = 1; index < trendline.points.length; index += 1) {
        const start = trendline.points[index - 1]!;
        const end = trendline.points[index]!;
        const dx = end.x - start.x;
        const dy = end.y - start.y;
        const lengthSquared = dx * dx + dy * dy || 1;
        const ratio = Math.max(0, Math.min(1, ((point.x - start.x) * dx + (point.y - start.y) * dy) / lengthSquared));
        if (Math.hypot(point.x - (start.x + ratio * dx), point.y - (start.y + ratio * dy)) <= 5) return { action: 'chart.select-element', data: { kind: 'trendline', seriesId: series.id } };
      }
    }
    if (series.points.some((candidate) => candidate.visible && Math.hypot(point.x - candidate.x, point.y - candidate.y) <= 10)) return { action: 'chart.select-element', data: { kind: 'series', seriesId: series.id } };
  }
  if (layout.valueAxis && point.x >= layout.plot.left - 12 && point.x <= layout.plot.left + 12 && point.y >= layout.plot.top && point.y <= layout.plot.top + layout.plot.height) return { action: 'chart.select-element', data: { kind: 'axis' } };
  if (layout.categoryAxis && point.y >= layout.plot.top + layout.plot.height - 12 && point.y <= layout.plot.top + layout.plot.height + 18 && point.x >= layout.plot.left && point.x <= layout.plot.left + layout.plot.width) return { action: 'chart.select-element', data: { kind: 'axis' } };
  if (point.x >= layout.plot.left && point.x <= layout.plot.left + layout.plot.width && point.y >= layout.plot.top && point.y <= layout.plot.top + layout.plot.height) return { action: 'chart.select-element', data: { kind: 'plot-area' } };
  return { action: 'chart.select-element', data: { kind: 'chart-area' } };
}

function drawCanonicalSparklineOnCanvas(options: {
  context: CanvasRenderingContext2D;
  sparkline: SparklineModel;
  group?: import('@react-sheets/core-model').SparklineGroup;
  values: Array<number | null>;
  minimum?: number;
  maximum?: number;
  rect: Rect;
}): void {
  const { context, values, rect } = options;
  const group = options.group;
  const sparkline: SparklineModel = {
    ...options.sparkline,
    ...(group?.type ? { type: group.type } : {}),
    ...(group?.showAxis === undefined ? {} : { showAxis: group.showAxis }),
    ...(group?.showMarkers === undefined ? {} : { showMarkers: group.showMarkers }),
    ...(group?.lineWeight === undefined ? {} : { lineWeight: group.lineWeight }),
    ...(group?.emptyCells === undefined ? {} : { emptyCells: group.emptyCells }),
    ...(group?.axisColor === undefined ? {} : { axisColor: group.axisColor }),
    ...(group?.firstColor === undefined ? {} : { firstColor: group.firstColor }),
    ...(group?.lastColor === undefined ? {} : { lastColor: group.lastColor }),
    ...(group?.highColor === undefined ? {} : { highColor: group.highColor }),
    ...(group?.lowColor === undefined ? {} : { lowColor: group.lowColor }),
    ...(group?.negativeColor === undefined ? {} : { negativeColor: group.negativeColor }),
    ...(group?.markerColor === undefined ? {} : { markerColor: group.markerColor }),
  };
  const { x, y, width, height } = rect;
  if (values.every((value) => value === null)) return;
  const numbers = values.filter((value): value is number => value !== null);
  const max = options.maximum ?? Math.max(...numbers, 0);
  const min = options.minimum ?? Math.min(...numbers, 0);
  const span = Math.max(1, max - min);
  context.save();
  context.translate(x, y);
  if (sparkline.type === "line") {
    context.strokeStyle = sparkline.color || "#2563eb";
    context.lineWidth = sparkline.lineWeight ?? 1.5;
    const connect = sparkline.emptyCells === 'connect';
    let started = false;
    values.forEach((value, index) => {
      if (value === null && !connect) { started = false; return; }
      const previous = value === null ? values.slice(0, index).reverse().find((candidate): candidate is number => candidate !== null) : value;
      if (previous === undefined) return;
      const px = (index / Math.max(1, values.length - 1)) * width;
      const py = height - ((previous - min) / span) * height;
      if (!started) { context.beginPath(); context.moveTo(px, py); started = true; }
      else context.lineTo(px, py);
    });
    context.stroke();
  } else {
    const barWidth = width / Math.max(1, values.length);
    values.forEach((value, index) => {
      if (value === null) return;
      const zero = height - ((0 - min) / span) * height;
      const valueY = height - ((value - min) / span) * height;
      const top = sparkline.type === 'win-loss' ? Math.min(zero, value >= 0 ? zero - Math.max(1, barWidth * 0.55) : zero) : Math.min(zero, valueY);
      const bottom = sparkline.type === 'win-loss' ? Math.max(zero, value < 0 ? zero + Math.max(1, barWidth * 0.55) : zero) : Math.max(zero, valueY);
      context.fillStyle = value < 0 ? sparkline.negativeColor || "#ef4444" : sparkline.color || "#2563eb";
      context.fillRect(index * barWidth + 1, top, Math.max(1, barWidth - 2), Math.max(1, bottom - top));
    });
  }
  if (sparkline.showAxis) {
    context.strokeStyle = sparkline.axisColor || "#cbd5e1";
    context.lineWidth = 1;
    const axisY = height - ((0 - min) / span) * height;
    context.beginPath();
    context.moveTo(0, axisY);
    context.lineTo(width, axisY);
    context.stroke();
  }
  if (sparkline.showMarkers || sparkline.highlightMax || sparkline.highlightMin || sparkline.highlightFirst || sparkline.highlightLast || sparkline.highlightNegative) {
    const marker = (index: number, color: string): void => {
      if (index < 0 || index >= values.length || values[index] === null) return;
      const px = (index / Math.max(1, values.length - 1)) * width;
      const py = height - ((values[index]! - min) / span) * height;
      context.fillStyle = color;
      context.beginPath();
      context.arc(px, py, 2, 0, Math.PI * 2);
      context.fill();
    };
    if (sparkline.showMarkers) values.forEach((_value, index) => marker(index, sparkline.markerColor || sparkline.color || "#2563eb"));
    const extremumIndex = (direction: 'max' | 'min'): number => values.reduce<number>((best, value, index) => {
      if (value === null) return best;
      const previous = best >= 0 ? values[best] : null;
      if (previous === null || previous === undefined) return index;
      return direction === 'max' ? value > previous ? index : best : value < previous ? index : best;
    }, -1);
    if (sparkline.highlightMax) marker(extremumIndex('max'), sparkline.highColor || "#16a34a");
    if (sparkline.highlightMin) marker(extremumIndex('min'), sparkline.lowColor || "#dc2626");
    if (sparkline.highlightFirst) marker(values.findIndex((value) => value !== null), sparkline.firstColor || "#f59e0b");
    if (sparkline.highlightLast) marker(values.reduce<number>((last, value, index) => value === null ? last : index, -1), sparkline.lastColor || "#f59e0b");
    if (sparkline.highlightNegative) values.forEach((value, index) => { if (value !== null && value < 0) marker(index, sparkline.negativeColor || '#ef4444'); });
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
      const data = getChartSeries(payload, getSheet, pivotResults, sheets, tables);
      const layout = buildChartLayout(payload, data, bounds.width, bounds.height);
      if (layout.status.kind !== 'ready') {
        drawables.push({ kind: 'shape', id: drawing.id, bounds, draw: (context, rect) => drawUnsupportedDrawingOnCanvas(context, rect, layout.status.message ?? `${layout.status.code ?? 'UNSUPPORTED_FEATURE'}: chart data is unavailable`) });
        continue;
      }
      drawables.push({
        kind: "chart",
        id: drawing.id,
        bounds,
        draw: (context, rect) => drawChartLayoutOnCanvas({ context, payload, bounds: rect, layout }),
        hitTest: (point) => chartHitTest(layout, point, payload.elements.dataTable?.visible === true),
      });
      continue;
    }
    if (payload.kind === 'camera') {
      drawables.push({ kind: 'shape', id: drawing.id, bounds, draw: (context, rect) => drawCameraOnCanvas(context, payload, rect, getSheet) });
      continue;
    }
    if (payload.kind === 'screenshot') {
      drawables.push({ kind: 'shape', id: drawing.id, bounds, draw: (context, rect) => drawScreenshotOnCanvas(context, payload, rect, getSheet) });
      continue;
    }
    if (payload.kind === 'form-control') {
      drawables.push({ kind: 'shape', id: drawing.id, bounds, draw: (context, rect) => drawFormControlOnCanvas(context, payload, rect) });
      continue;
    }
    if (payload.kind === 'icon') {
      drawables.push({ kind: 'shape', id: drawing.id, bounds, draw: (context, rect) => drawIconOnCanvas(context, payload, rect) });
      continue;
    }
    if (payload.kind === 'model3d') {
      drawables.push({ kind: 'shape', id: drawing.id, bounds, draw: (context, rect) => drawModel3dOnCanvas(context, payload, rect) });
      continue;
    }
    if (payload.kind === 'smartart') {
      drawables.push({ kind: 'shape', id: drawing.id, bounds, draw: (context, rect) => drawSmartArtOnCanvas(context, payload, rect) });
      continue;
    }
    if (payload.kind === 'wordart') {
      drawables.push({ kind: 'shape', id: drawing.id, bounds, draw: (context, rect) => drawWordArtOnCanvas(context, payload, rect, drawing.transform.rotation) });
      continue;
    }
    if (payload.kind === 'signature-line') {
      drawables.push({ kind: 'shape', id: drawing.id, bounds, draw: (context, rect) => drawSignatureLineOnCanvas(context, payload, rect) });
      continue;
    }
    if (payload.kind === 'embedded-object') {
      drawables.push({ kind: 'shape', id: drawing.id, bounds, draw: (context, rect) => drawEmbeddedObjectOnCanvas(context, payload, rect) });
      continue;
    }
    if (payload.kind === 'equation') {
      drawables.push({ kind: 'shape', id: drawing.id, bounds, draw: (context, rect) => drawEquationOnCanvas(context, payload, rect) });
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
  const sparklineGroups = sheet.sparklineGroups ?? [];
  const groupBounds = new Map<string, { min: number; max: number }>();
  for (const group of sparklineGroups) {
    const values = group.sparklineIds.flatMap((id) => {
      const member = sparklines.find((entry) => entry.id === id);
      if (!member) return [];
      try { return resolveSparklineData(member, (sheetId) => getSheet(sheetId), group).values.filter((value): value is number => value !== null); } catch { return []; }
    });
    if (values.length) groupBounds.set(group.id, { min: Math.min(0, ...values), max: Math.max(0, ...values) });
  }
  for (const sparkline of sparklines) {
    const rect = skeleton.getCellRect(sparkline.anchor.row, sparkline.anchor.column);
    if (!rect) continue;
    const group = sparkline.groupId ? sparklineGroups.find((entry) => entry.id === sparkline.groupId) : undefined;
    drawables.push({
      kind: "shape",
      id: sparkline.id,
      bounds: rect,
      draw: (context, target) => {
        try {
          const resolved = resolveSparklineData(sparkline, (sheetId) => getSheet(sheetId), group);
          const bounds = group?.verticalAxis?.mode === 'same-group' && groupBounds.get(group.id) ? groupBounds.get(group.id)! : { min: resolved.min, max: resolved.max };
          drawCanonicalSparklineOnCanvas({ context, sparkline, group, values: resolved.values, minimum: group?.verticalAxis?.mode === 'custom' ? group.verticalAxis.minimum : bounds.min, maximum: group?.verticalAxis?.mode === 'custom' ? group.verticalAxis.maximum : bounds.max, rect: target });
        } catch (error) {
          drawUnsupportedDrawingOnCanvas(context, target, error instanceof Error ? error.message : 'INVALID_CHART_SOURCE: Sparkline data is unavailable');
        }
      },
    });
  }
  return drawables;
}
