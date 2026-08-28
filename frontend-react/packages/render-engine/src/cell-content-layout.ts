import type { CellRenderData, CellRenderStyle, Rect, RenderTheme } from './types';

const MIN_SHRINK_FONT_SIZE_PX = 8;
const TEXT_LAYOUT_CACHE_LIMIT = 4_000;
const layoutCaches = new WeakMap<object, Map<string, CellContentLayoutResult>>();

export type CellContentLayoutMode = 'display' | 'edit';
type CellHorizontalAlignment = NonNullable<NonNullable<CellRenderData['style']>['horizontalAlignment']>;

export interface CellLayoutNeighbor {
  column: number;
  widthPx: number;
  occupied: boolean;
}

export interface CellNeighborOccupancy {
  left: readonly CellLayoutNeighbor[];
  right: readonly CellLayoutNeighbor[];
}

export interface CellContentLayoutInput {
  context: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;
  cell: CellRenderData;
  theme: RenderTheme;
  text: string;
  cellRect: Rect;
  mode: CellContentLayoutMode;
  mergedRect?: Rect;
  alignmentSpan?: Rect;
  cellRange?: { startColumn: number; endColumn: number };
  neighborOccupancy?: CellNeighborOccupancy;
  viewportRect?: Rect;
  caret?: { start: number; end: number };
  reserveFilterButton?: boolean;
  /** Rectangles and font sizes are already CSS/model pixels; zoom is identity metadata, not an extra scale. */
  zoom?: number;
}

export interface CellCaretGeometry {
  x: number;
  y: number;
  width: number;
  height: number;
  lineIndex: number;
  visible: boolean;
}

export interface CellMeasuredTextRun {
  text: string;
  font: string;
  fontSizePx: number;
  widthPx: number;
}

export interface CellContentLayoutResult {
  font: string;
  fontSizePx: number;
  lineHeightPx: number;
  lines: readonly string[];
  fontRuns: readonly CellMeasuredTextRun[];
  rawTextWidthPx: number;
  widthPx: number;
  heightPx: number;
  contentWidthPx: number;
  contentHeightPx: number;
  horizontalAlignment: CellHorizontalAlignment;
  multiline: boolean;
  displayRect: Rect;
  editRect: Rect;
  contentRect: Rect;
  overflowWidthPx: number;
  overflowSpan?: { startColumn: number; endColumn: number };
  caretGeometry?: CellCaretGeometry;
  requiresInternalScroll: boolean;
}

/**
 * Canonical owner for static text, AutoFit, and in-cell edit geometry.
 * Callers provide rectangles in one CSS/model coordinate space. The domain
 * never applies a second zoom transform.
 */
export class CellContentLayoutDomain {
  static resolve(input: CellContentLayoutInput): CellContentLayoutResult {
    validateInput(input);
    const style = input.cell.style;
    const padding = style?.padding ?? input.theme.cellPadding;
    const indent = Math.max(0, Math.trunc(style?.indent ?? 0)) * 12;
    const borderWidth = (style?.borders?.left ? 1 : 0) + (style?.borders?.right ? 1 : 0);
    const borderHeight = (style?.borders?.top ? 1 : 0) + (style?.borders?.bottom ? 1 : 0);
    const reserve = input.reserveFilterButton ? 18 : 0;
    const baseRect = input.mergedRect ?? input.alignmentSpan ?? input.cellRect;
    const availableWidthPx = baseRect.width > 0 ? Math.max(0, baseRect.width - padding * 2 - indent - reserve - borderWidth) : undefined;
    const baseFontSizePx = style?.fontSizePx ?? 13;
    const baseFont = cellRenderFont(style, input.theme, baseFontSizePx);
    const cacheKey = JSON.stringify({
      text: input.text,
      font: baseFont,
      presentation: input.cell.presentation?.kind,
      richText: input.cell.richText,
      mode: input.mode,
      baseRect,
      cellRect: input.cellRect,
      alignmentSpan: input.alignmentSpan,
      cellRange: input.cellRange,
      neighbors: input.neighborOccupancy,
      viewport: input.viewportRect,
      caret: input.caret,
      padding,
      indent,
      borderWidth,
      borderHeight,
      reserve,
      wrapText: Boolean(style?.wrapText),
      shrinkToFit: Boolean(style?.shrinkToFit),
      textOrientation: style?.textOrientation,
      textRotate: style?.textRotate ?? 0,
      zoom: input.zoom ?? 1,
    });
    const cache = layoutCaches.get(input.context) ?? new Map<string, CellContentLayoutResult>();
    layoutCaches.set(input.context, cache);
    const cached = cache.get(cacheKey);
    if (cached) return cached;

    input.context.save();
    input.context.font = baseFont;
    let fontSizePx = baseFontSizePx;
    let font = baseFont;
    let fontRuns = measureTextRuns(input.context, input.cell, input.theme, input.text, baseFontSizePx);
    let rawTextWidthPx = widestLineFromRuns(input.context, input.text, fontRuns);
    if (style?.shrinkToFit && !style.wrapText && availableWidthPx !== undefined && availableWidthPx > 0 && rawTextWidthPx > availableWidthPx) {
      fontSizePx = Math.max(MIN_SHRINK_FONT_SIZE_PX, baseFontSizePx * availableWidthPx / rawTextWidthPx);
      font = cellRenderFont(style, input.theme, fontSizePx);
      input.context.font = font;
      fontRuns = measureTextRuns(input.context, input.cell, input.theme, input.text, fontSizePx);
      rawTextWidthPx = widestLineFromRuns(input.context, input.text, fontRuns);
    }
    const lineHeightPx = Math.max(fontSizePx * 1.25, ...fontRuns.map((run) => run.fontSizePx * 1.25), 16);
    input.context.font = font;
    const lines = style?.textOrientation === 'stacked'
      ? Array.from(input.text.replace(/\r?\n/g, ''))
      : style?.wrapText && availableWidthPx !== undefined
        ? wrapText(input.context, input.text, availableWidthPx)
        : input.text.split(/\r?\n/);
    const textWidthPx = measureLines(input.context, input.cell, input.theme, input.text, lines, fontSizePx, fontRuns);
    const contentHeightPx = Math.max(1, lines.length) * lineHeightPx;
    const naturalWidthPx = textWidthPx + padding * 2 + indent + reserve + borderWidth;
    const naturalHeightPx = contentHeightPx + padding * 2 + borderHeight;
    input.context.restore();

    const horizontalAlignment = resolveHorizontalAlignment(style?.horizontalAlignment, input.cell.value);
    const displayRect = input.mode === 'display'
      ? resolveStaticDisplayRect(input, baseRect, naturalWidthPx, horizontalAlignment)
      : { ...baseRect };
    const editResult = input.mode === 'edit'
      ? resolveEditRect(input, baseRect, naturalWidthPx, naturalHeightPx, horizontalAlignment)
      : { rect: { ...baseRect }, requiresInternalScroll: false };
    const activeRect = input.mode === 'edit' ? editResult.rect : displayRect;
    const overflowSpan = resolveOverflowSpan(input, baseRect, displayRect);
    const contentRect = {
      x: activeRect.x + padding + indent,
      y: activeRect.y + padding,
      width: Math.max(1, activeRect.width - padding * 2 - indent - reserve - borderWidth),
      height: Math.max(1, activeRect.height - padding * 2 - borderHeight),
    };
    const overflowWidthPx = Math.max(1, activeRect.width - padding * 2 - indent - reserve - borderWidth);
    const result: CellContentLayoutResult = {
      font,
      fontSizePx,
      lineHeightPx,
      lines,
      fontRuns,
      rawTextWidthPx,
      widthPx: Math.ceil(naturalWidthPx),
      heightPx: Math.ceil(naturalHeightPx),
      contentWidthPx: textWidthPx,
      contentHeightPx,
      horizontalAlignment,
      multiline: lines.length > 1 || Boolean(style?.wrapText),
      displayRect,
      editRect: editResult.rect,
      contentRect,
      overflowWidthPx,
      ...(overflowSpan ? { overflowSpan } : {}),
      ...(input.caret ? { caretGeometry: resolveCaretGeometry(input, { lines, lineHeightPx, font, contentRect, horizontalAlignment }) } : {}),
      requiresInternalScroll: editResult.requiresInternalScroll,
    };
    if (cache.size >= TEXT_LAYOUT_CACHE_LIMIT) cache.clear();
    cache.set(cacheKey, result);
    return result;
  }
}

export function resolveCellContentLayout(input: CellContentLayoutInput): CellContentLayoutResult {
  return CellContentLayoutDomain.resolve(input);
}

export function cellRenderFont(style: CellRenderData['style'], theme: RenderTheme, fontSizePx = style?.fontSizePx ?? 13): string {
  const family = style?.fontFamily ? `"${style.fontFamily}", sans-serif` : '"Microsoft YaHei", "Segoe UI", sans-serif';
  const weight = style?.bold ? '700' : '400';
  const slant = style?.italic ? ' italic' : '';
  return `${slant} ${weight} ${fontSizePx}px ${family}`;
}

function measureTextRuns(
  context: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  cell: CellRenderData,
  theme: RenderTheme,
  text: string,
  baseFontSizePx: number,
): CellMeasuredTextRun[] {
  const runs = cell.richText && cell.richText.length > 0 && cell.richText.map((run) => run.text).join('') === text
    ? cell.richText
    : [{ text, style: undefined }];
  const measured: CellMeasuredTextRun[] = [];
  for (const run of runs) {
    const scale = baseFontSizePx > 0 ? baseFontSizePx / (cell.style?.fontSizePx ?? baseFontSizePx) : 1;
    const fontSizePx = Math.max(1, (run.style?.fontSizePx ?? cell.style?.fontSizePx ?? baseFontSizePx) * scale);
    const runStyle: CellRenderStyle = {
      fontFamily: run.style?.fontFamily ?? cell.style?.fontFamily,
      fontSizePx,
      bold: run.style?.bold ?? cell.style?.bold,
      italic: run.style?.italic ?? cell.style?.italic,
      textColor: run.style?.textColor ?? cell.style?.textColor,
    };
    const font = cellRenderFont(runStyle, theme, fontSizePx);
    context.font = font;
    measured.push({ text: run.text, font, fontSizePx, widthPx: widestLine(context, run.text) });
  }
  context.font = cellRenderFont(cell.style, theme, baseFontSizePx);
  return measured;
}

function widestLineFromRuns(context: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D, text: string, runs: readonly CellMeasuredTextRun[]): number {
  if (runs.length === 0) return 0;
  const widths: number[] = [];
  let lineWidth = 0;
  let runIndex = 0;
  let runOffset = 0;
  for (const character of text) {
    if (character === '\n') {
      widths.push(lineWidth);
      lineWidth = 0;
      continue;
    }
    while (runIndex < runs.length - 1 && runOffset >= runs[runIndex]!.text.length) {
      runOffset -= runs[runIndex]!.text.length;
      runIndex += 1;
    }
    const run = runs[Math.min(runIndex, runs.length - 1)]!;
    context.font = run.font;
    lineWidth += context.measureText(character).width;
    runOffset += character.length;
  }
  widths.push(lineWidth);
  return Math.max(0, ...widths);
}

function measureLines(
  context: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  cell: CellRenderData,
  theme: RenderTheme,
  sourceText: string,
  lines: readonly string[],
  baseFontSizePx: number,
  measuredRuns: readonly CellMeasuredTextRun[],
): number {
  if (!cell.richText || cell.richText.length === 0) return Math.max(0, ...lines.map((line) => context.measureText(line).width));
  const richText = cell.richText.map((run) => run.text).join('');
  if (richText !== sourceText) return Math.max(0, ...lines.map((line) => context.measureText(line).width));
  let searchCursor = 0;
  const widths = lines.map((line) => {
    const found = sourceText.indexOf(line, searchCursor);
    const sourceStart = found >= 0 ? found : searchCursor;
    searchCursor = Math.min(sourceText.length, sourceStart + line.length);
    return measureRichTextSegment(context, line, sourceStart, measuredRuns);
  });
  return Math.max(0, ...widths);
}

function measureRichTextSegment(
  context: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  text: string,
  sourceStart: number,
  runs: readonly CellMeasuredTextRun[],
): number {
  let width = 0;
  let runIndex = 0;
  let sourceOffset = sourceStart;
  let runEnd = runs[0]?.text.length ?? 0;
  for (const character of text) {
    while (runIndex < runs.length - 1 && sourceOffset >= runEnd) {
      runIndex += 1;
      runEnd += runs[runIndex]!.text.length;
    }
    const run = runs[Math.min(runIndex, runs.length - 1)];
    if (!run) {
      width += context.measureText(character).width;
      sourceOffset += character.length;
      continue;
    }
    context.font = run.font;
    width += context.measureText(character).width;
    sourceOffset += character.length;
  }
  return width;
}

function validateInput(input: CellContentLayoutInput): void {
  if (!Number.isFinite(input.cellRect.x) || !Number.isFinite(input.cellRect.y) || !Number.isFinite(input.cellRect.width) || !Number.isFinite(input.cellRect.height) || input.cellRect.width < 0 || input.cellRect.height < 0) throw new Error('CellContentLayoutDomain received an invalid cell rectangle');
  if (input.zoom !== undefined && (!Number.isFinite(input.zoom) || input.zoom <= 0)) throw new Error('CellContentLayoutDomain zoom must be positive');
  if (input.alignmentSpan && input.alignmentSpan.width < input.cellRect.width) throw new Error('CellContentLayoutDomain alignment span cannot be narrower than the cell rectangle');
}

function widestLine(context: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D, text: string): number {
  return Math.max(0, ...text.split(/\r?\n/).map((line) => context.measureText(line).width));
}

function wrapText(context: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D, text: string, widthPx: number): string[] {
  if (widthPx <= 0) return text.split(/\r?\n/);
  return text.split(/\r?\n/).flatMap((paragraph) => wrapParagraph(context, paragraph, widthPx));
}

function wrapParagraph(context: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D, paragraph: string, widthPx: number): string[] {
  if (paragraph === '') return [''];
  const cjk = /[\u3040-\u30ff\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]/u.test(paragraph);
  const tokens = cjk ? Array.from(paragraph) : paragraph.match(/\S+|\s+/g) ?? [paragraph];
  const lines: string[] = [];
  let current = '';
  for (const token of tokens) {
    const candidate = current + token;
    if (current && context.measureText(candidate).width > widthPx) {
      lines.push(current.trimEnd());
      current = token.trimStart();
      while (current && context.measureText(current).width > widthPx) {
        const split = splitToken(context, current, widthPx);
        lines.push(split.head);
        current = split.tail;
      }
    } else current = candidate;
  }
  if (current || lines.length === 0) lines.push(current.trimEnd());
  return lines;
}

function splitToken(context: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D, token: string, widthPx: number): { head: string; tail: string } {
  const graphemes = Array.from(token);
  let head = '';
  let index = 0;
  while (index < graphemes.length && context.measureText(head + graphemes[index]).width <= widthPx) head += graphemes[index++];
  if (!head) head = graphemes[index++] ?? '';
  return { head, tail: graphemes.slice(index).join('') };
}

function resolveHorizontalAlignment(alignment: CellHorizontalAlignment | undefined, value: CellRenderData['value']): CellHorizontalAlignment {
  if (alignment !== undefined && alignment !== 'general') return alignment;
  return typeof value === 'number' || typeof value === 'boolean' ? 'right' : 'left';
}

function resolveStaticDisplayRect(input: CellContentLayoutInput, baseRect: Rect, naturalWidthPx: number, alignment: CellHorizontalAlignment): Rect {
  const style = input.cell.style;
  const canOverflow = typeof input.cell.value === 'string'
    && !style?.wrapText
    && !style?.shrinkToFit
    && (!style?.textOrientation || style.textOrientation === 'horizontal')
    && !input.cell.presentation
    && alignment !== 'fill'
    && alignment !== 'justify'
    && alignment !== 'distributed'
    && !input.alignmentSpan;
  if (!canOverflow || naturalWidthPx <= baseRect.width || !input.neighborOccupancy) return { ...baseRect };
  const required = naturalWidthPx - baseRect.width;
  const leftAndRight = alignment === 'right' ? consumeOneSide(input.neighborOccupancy.left, required, 'left') : alignment === 'center' ? consumeBothSides(input.neighborOccupancy, required) : consumeOneSide(input.neighborOccupancy.right, required, 'right');
  const result = { x: baseRect.x - leftAndRight.leftPx, y: baseRect.y, width: baseRect.width + leftAndRight.leftPx + leftAndRight.rightPx, height: baseRect.height };
  return result;
}

function resolveOverflowSpan(input: CellContentLayoutInput, baseRect: Rect, displayRect: Rect): { startColumn: number; endColumn: number } | undefined {
  if (sameRect(baseRect, displayRect) || !input.cellRange || !input.neighborOccupancy) return undefined;
  const start = input.cellRange.startColumn;
  const end = input.cellRange.endColumn;
  const leftPx = Math.max(0, input.cellRect.x - displayRect.x);
  const rightPx = Math.max(0, displayRect.x + displayRect.width - (input.cellRect.x + input.cellRect.width));
  const leftColumns = consumeWidth(input.neighborOccupancy.left, leftPx, true).columns;
  const rightColumns = consumeWidth(input.neighborOccupancy.right, rightPx, true).columns;
  return { startColumn: start - leftColumns, endColumn: end + rightColumns };
}

function consumeOneSide(neighbors: readonly CellLayoutNeighbor[], requiredPx: number, side: 'left' | 'right'): { leftPx: number; rightPx: number; leftColumns: number; rightColumns: number } {
  const consumed = consumeWidth(neighbors, requiredPx, true);
  return side === 'left'
    ? { leftPx: consumed.widthPx, rightPx: 0, leftColumns: consumed.columns, rightColumns: 0 }
    : { leftPx: 0, rightPx: consumed.widthPx, leftColumns: 0, rightColumns: consumed.columns };
}

function consumeBothSides(occupancy: CellNeighborOccupancy, requiredPx: number): { leftPx: number; rightPx: number; leftColumns: number; rightColumns: number } {
  const leftTarget = Math.ceil(requiredPx / 2);
  const rightTarget = Math.floor(requiredPx / 2);
  const left = consumeWidth(occupancy.left, leftTarget, true);
  const right = consumeWidth(occupancy.right, rightTarget, true);
  let remaining = Math.max(0, requiredPx - left.widthPx - right.widthPx);
  const extraRight = consumeWidth(occupancy.right.slice(right.columns), remaining, true);
  right.widthPx += extraRight.widthPx;
  right.columns += extraRight.columns;
  remaining -= extraRight.widthPx;
  const extraLeft = consumeWidth(occupancy.left.slice(left.columns), remaining, true);
  left.widthPx += extraLeft.widthPx;
  left.columns += extraLeft.columns;
  return { leftPx: left.widthPx, rightPx: right.widthPx, leftColumns: left.columns, rightColumns: right.columns };
}

function consumeWidth(neighbors: readonly CellLayoutNeighbor[], requiredPx: number, stopAtOccupied = false): { widthPx: number; columns: number } {
  let widthPx = 0;
  let columns = 0;
  for (const neighbor of neighbors) {
    if (stopAtOccupied && neighbor.occupied) break;
    if (widthPx >= requiredPx) break;
    widthPx += Math.max(0, neighbor.widthPx);
    columns += 1;
  }
  return { widthPx: Math.min(widthPx, requiredPx), columns };
}

function resolveEditRect(input: CellContentLayoutInput, baseRect: Rect, naturalWidthPx: number, naturalHeightPx: number, alignment: CellHorizontalAlignment): { rect: Rect; requiresInternalScroll: boolean } {
  const style = input.cell.style;
  const multiline = input.text.includes('\n') || Boolean(style?.wrapText);
  if (multiline) {
    const rect = { ...baseRect, height: Math.max(baseRect.height, naturalHeightPx) };
    return { rect: clampRectToViewport(rect, input.viewportRect), requiresInternalScroll: rect.height > (input.viewportRect?.height ?? rect.height) };
  }
  if (naturalWidthPx <= baseRect.width || !input.neighborOccupancy) {
    const rect = clampRectToViewport({ ...baseRect }, input.viewportRect);
    return { rect, requiresInternalScroll: naturalWidthPx > rect.width };
  }
  const required = naturalWidthPx - baseRect.width;
  const occupancy = input.neighborOccupancy;
  const leftAvailable = Math.max(0, baseRect.x - (input.viewportRect?.x ?? -Infinity));
  const rightAvailable = Math.max(0, (input.viewportRect ? input.viewportRect.x + input.viewportRect.width : Infinity) - (baseRect.x + baseRect.width));
  const leftTarget = alignment === 'right' ? required : alignment === 'center' ? Math.ceil(required / 2) : 0;
  const rightTarget = alignment === 'left' ? required : alignment === 'center' ? Math.floor(required / 2) : 0;
  const left = consumeWidth(occupancy.left, Math.min(leftTarget, leftAvailable));
  const right = consumeWidth(occupancy.right, Math.min(rightTarget, rightAvailable));
  let remaining = Math.max(0, required - left.widthPx - right.widthPx);
  if (remaining > 0 && alignment !== 'right') {
    const extra = consumeWidth(occupancy.right.slice(right.columns), Math.min(remaining, Math.max(0, rightAvailable - right.widthPx)));
    right.widthPx += extra.widthPx;
    right.columns += extra.columns;
    remaining -= extra.widthPx;
  }
  if (remaining > 0 && alignment !== 'left') {
    const extra = consumeWidth(occupancy.left.slice(left.columns), Math.min(remaining, Math.max(0, leftAvailable - left.widthPx)));
    left.widthPx += extra.widthPx;
    left.columns += extra.columns;
  }
  if (remaining > 0 && alignment === 'right') {
    const extra = consumeWidth(occupancy.right.slice(right.columns), Math.min(remaining, Math.max(0, rightAvailable - right.widthPx)));
    right.widthPx += extra.widthPx;
    right.columns += extra.columns;
    remaining -= extra.widthPx;
  }
  if (remaining > 0 && alignment === 'left') {
    const extra = consumeWidth(occupancy.left.slice(left.columns), Math.min(remaining, Math.max(0, leftAvailable - left.widthPx)));
    left.widthPx += extra.widthPx;
    left.columns += extra.columns;
  }
  const desired = { x: baseRect.x - left.widthPx, y: baseRect.y, width: baseRect.width + left.widthPx + right.widthPx, height: Math.max(baseRect.height, naturalHeightPx) };
  const rect = clampRectToViewport(desired, input.viewportRect);
  return { rect, requiresInternalScroll: desired.width > rect.width || desired.height > rect.height };
}

function clampRectToViewport(rect: Rect, viewport: Rect | undefined): Rect {
  if (!viewport) return rect;
  const x = Math.max(viewport.x, Math.min(rect.x, viewport.x + viewport.width - Math.min(rect.width, viewport.width)));
  const y = Math.max(viewport.y, Math.min(rect.y, viewport.y + viewport.height - Math.min(rect.height, viewport.height)));
  return { x, y, width: Math.min(rect.width, viewport.width), height: Math.min(rect.height, viewport.height) };
}

function sameRect(left: Rect, right: Rect): boolean {
  return left.x === right.x && left.y === right.y && left.width === right.width && left.height === right.height;
}

function resolveCaretGeometry(input: CellContentLayoutInput, layout: { lines: readonly string[]; lineHeightPx: number; font: string; contentRect: Rect; horizontalAlignment: CellHorizontalAlignment }): CellCaretGeometry {
  const caretOffset = Math.max(0, Math.min(input.text.length, input.caret?.end ?? input.text.length));
  const lineOffsets = lineStartOffsets(input.text, layout.lines);
  let lineIndex = 0;
  for (let index = 0; index < lineOffsets.length; index += 1) {
    if (caretOffset >= lineOffsets[index]!.start) lineIndex = index;
    else break;
  }
  const line = layout.lines[Math.min(lineIndex, layout.lines.length - 1)] ?? '';
  const lineStart = lineOffsets[Math.min(lineIndex, lineOffsets.length - 1)]?.start ?? 0;
  const within = Math.max(0, Math.min(line.length, caretOffset - lineStart));
  input.context.save();
  input.context.font = layout.font;
  const prefixWidth = input.context.measureText(line.slice(0, within)).width;
  const lineWidth = input.context.measureText(line).width;
  input.context.restore();
  const x = layout.horizontalAlignment === 'right'
    ? layout.contentRect.x + layout.contentRect.width - lineWidth + prefixWidth
    : layout.horizontalAlignment === 'center' || layout.horizontalAlignment === 'centerContinuous'
      ? layout.contentRect.x + (layout.contentRect.width - lineWidth) / 2 + prefixWidth
      : layout.contentRect.x + prefixWidth;
  const y = layout.contentRect.y + lineIndex * layout.lineHeightPx;
  const caretRect = { x, y, width: 1, height: layout.lineHeightPx };
  const viewport = input.viewportRect;
  const visible = !viewport || caretRect.x >= viewport.x && caretRect.x <= viewport.x + viewport.width && caretRect.y + caretRect.height >= viewport.y && caretRect.y <= viewport.y + viewport.height;
  return { ...caretRect, lineIndex, visible };
}

function lineStartOffsets(text: string, lines: readonly string[]): Array<{ start: number }> {
  const offsets: Array<{ start: number }> = [];
  let cursor = 0;
  for (const line of lines) {
    const start = text.indexOf(line, cursor);
    const resolved = start < 0 ? cursor : start;
    offsets.push({ start: resolved });
    cursor = resolved + line.length;
    if (text[cursor] === '\r') cursor += 1;
    if (text[cursor] === '\n') cursor += 1;
  }
  return offsets.length > 0 ? offsets : [{ start: 0 }];
}
