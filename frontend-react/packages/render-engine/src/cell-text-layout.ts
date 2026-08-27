import type { CellRenderData, RenderTheme } from './types';

const MIN_SHRINK_FONT_SIZE_PX = 8;
const TEXT_LAYOUT_CACHE_LIMIT = 4_000;
const layoutCache = new Map<string, CellTextLayout>();

export interface CellTextLayout {
  font: string;
  fontSizePx: number;
  lineHeightPx: number;
  lines: readonly string[];
  rawTextWidthPx: number;
  widthPx: number;
  heightPx: number;
}

export function cellRenderFont(style: CellRenderData['style'], theme: RenderTheme, fontSizePx = style?.fontSizePx ?? 13): string {
  const family = style?.fontFamily ? `"${style.fontFamily}", sans-serif` : '"Microsoft YaHei", "Segoe UI", sans-serif';
  const weight = style?.bold ? '700' : '400';
  const slant = style?.italic ? ' italic' : '';
  return `${slant} ${weight} ${fontSizePx}px ${family}`;
}

/**
 * Single text geometry owner for both Canvas rendering and AutoFit.  Its
 * cache is keyed by display content, effective font, wrapping width, and all
 * layout-affecting cell style values; normal text never shrinks implicitly.
 */
export function resolveCellTextLayout(
  context: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  cell: CellRenderData,
  theme: RenderTheme,
  text: string,
  availableWidthPx?: number,
  reserveFilterButton = false,
): CellTextLayout {
  const style = cell.style;
  const padding = style?.padding ?? theme.cellPadding;
  const indent = Math.max(0, Math.trunc(style?.indent ?? 0)) * 12;
  const borderWidth = (style?.borders?.left ? 1 : 0) + (style?.borders?.right ? 1 : 0);
  const borderHeight = (style?.borders?.top ? 1 : 0) + (style?.borders?.bottom ? 1 : 0);
  const contentWidth = availableWidthPx === undefined
    ? undefined
    : Math.max(0, availableWidthPx - padding * 2 - indent - (reserveFilterButton ? 18 : 0));
  const baseFontSizePx = style?.fontSizePx ?? 13;
  const baseFont = cellRenderFont(style, theme, baseFontSizePx);
  const cacheKey = JSON.stringify({
    text,
    font: baseFont,
    contentWidth,
    padding,
    indent,
    borderWidth,
    borderHeight,
    wrapText: Boolean(style?.wrapText),
    shrinkToFit: Boolean(style?.shrinkToFit),
    textOrientation: style?.textOrientation,
    textRotate: style?.textRotate ?? 0,
    reserveFilterButton,
  });
  const cached = layoutCache.get(cacheKey);
  if (cached) return cached;

  context.save();
  context.font = baseFont;
  let fontSizePx = baseFontSizePx;
  let font = baseFont;
  let rawTextWidthPx = widestLine(context, text);

  // Excel's shrink-to-fit is an explicit style.  It has a usability floor and
  // never alters cells without that style.
  if (style?.shrinkToFit && !style?.wrapText && contentWidth !== undefined && rawTextWidthPx > contentWidth && contentWidth > 0) {
    fontSizePx = Math.max(MIN_SHRINK_FONT_SIZE_PX, baseFontSizePx * contentWidth / rawTextWidthPx);
    font = cellRenderFont(style, theme, fontSizePx);
    context.font = font;
    rawTextWidthPx = widestLine(context, text);
  }

  const lineHeightPx = Math.max(fontSizePx * 1.25, 16);
  const lines = style?.textOrientation === 'stacked'
    ? Array.from(text.replace(/\r?\n/g, ''))
    : style?.wrapText && contentWidth !== undefined
      ? wrapText(context, text, contentWidth)
      : text.split(/\r?\n/);
  let width = style?.textOrientation === 'stacked'
    ? Math.max(fontSizePx, ...lines.map((line) => context.measureText(line).width))
    : Math.max(0, ...lines.map((line) => context.measureText(line).width));
  width += padding * 2 + indent + (reserveFilterButton ? 18 : 0) + borderWidth;
  if ((style?.wrapText || style?.shrinkToFit) && availableWidthPx !== undefined) width = Math.min(width, availableWidthPx);
  let height = Math.max(1, lines.length) * lineHeightPx + padding * 2 + borderHeight;

  const rotationDegrees = style?.textOrientation === 'rotateUp'
    ? 90
    : style?.textOrientation === 'rotateDown'
      ? 180
      : style?.textRotate ?? 0;
  const rotation = Math.abs(rotationDegrees * Math.PI / 180);
  if (rotation > 0) {
    const rotatedWidth = Math.abs(Math.cos(rotation)) * width + Math.abs(Math.sin(rotation)) * height;
    const rotatedHeight = Math.abs(Math.sin(rotation)) * width + Math.abs(Math.cos(rotation)) * height;
    width = rotatedWidth;
    height = rotatedHeight;
  }
  context.restore();

  const result: CellTextLayout = {
    font,
    fontSizePx,
    lineHeightPx,
    lines,
    rawTextWidthPx,
    widthPx: Math.ceil(width),
    heightPx: Math.ceil(height),
  };
  if (layoutCache.size >= TEXT_LAYOUT_CACHE_LIMIT) layoutCache.clear();
  layoutCache.set(cacheKey, result);
  return result;
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
    } else {
      current = candidate;
    }
  }
  if (current || lines.length === 0) lines.push(current.trimEnd());
  return lines;
}

function splitToken(context: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D, token: string, widthPx: number): { head: string; tail: string } {
  const graphemes = Array.from(token);
  let head = '';
  let index = 0;
  while (index < graphemes.length && context.measureText(head + graphemes[index]).width <= widthPx) {
    head += graphemes[index++];
  }
  if (!head) head = graphemes[index++] ?? '';
  return { head, tail: graphemes.slice(index).join('') };
}
