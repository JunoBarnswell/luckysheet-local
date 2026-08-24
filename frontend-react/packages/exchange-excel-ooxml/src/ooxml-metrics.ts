export const CSS_PIXELS_PER_INCH = 96;
export const POINTS_PER_INCH = 72;
export const DEFAULT_EXCEL_FONT_FAMILY = 'Calibri';
export const DEFAULT_EXCEL_FONT_SIZE_PT = 11;
export const DEFAULT_EXCEL_MAX_DIGIT_WIDTH_PX = 7;
export const MAX_EXCEL_COLUMN_WIDTH = 255;

export interface OoxmlNormalFont {
  family: string;
  sizePt: number;
}

/** Host-injectable font metric. Workers use the deterministic Calibri-compatible fallback. */
export interface OoxmlFontMeasurer {
  maximumDigitWidthPx(font: OoxmlNormalFont): number;
}

export const DEFAULT_OOXML_FONT_MEASURER: OoxmlFontMeasurer = {
  maximumDigitWidthPx(font) {
    const sizeScale = font.sizePt > 0 ? font.sizePt / DEFAULT_EXCEL_FONT_SIZE_PT : 1;
    return Math.max(1, DEFAULT_EXCEL_MAX_DIGIT_WIDTH_PX * sizeScale);
  },
};

export function pointsToPixels(points: number): number {
  if (!Number.isFinite(points)) throw new Error('Point size must be finite');
  return points * CSS_PIXELS_PER_INCH / POINTS_PER_INCH;
}

export function pixelsToPoints(pixels: number): number {
  if (!Number.isFinite(pixels)) throw new Error('Pixel size must be finite');
  return pixels * POINTS_PER_INCH / CSS_PIXELS_PER_INCH;
}

/** ECMA-376 column width formula. The result is model-space 96-DPI CSS pixels. */
export function excelColumnWidthToPixels(width: number, maximumDigitWidthPx = DEFAULT_EXCEL_MAX_DIGIT_WIDTH_PX): number {
  if (!Number.isFinite(width) || width < 0) throw new Error('Excel column width must be a non-negative finite number');
  if (!Number.isFinite(maximumDigitWidthPx) || maximumDigitWidthPx <= 0) throw new Error('Maximum digit width must be positive');
  if (width === 0) return 0;
  return Math.floor(((256 * Math.min(width, MAX_EXCEL_COLUMN_WIDTH) + Math.floor(128 / maximumDigitWidthPx)) / 256) * maximumDigitWidthPx);
}

/**
 * Inverse of `excelColumnWidthToPixels`, quantized to OOXML's 1/256 character unit.
 * Binary search avoids the several incompatible approximations commonly copied
 * from UI code and guarantees a <= 1px geometry round-trip.
 */
export function pixelsToExcelColumnWidth(pixels: number, maximumDigitWidthPx = DEFAULT_EXCEL_MAX_DIGIT_WIDTH_PX): number {
  if (!Number.isFinite(pixels) || pixels < 0) throw new Error('Column pixels must be a non-negative finite number');
  if (!Number.isFinite(maximumDigitWidthPx) || maximumDigitWidthPx <= 0) throw new Error('Maximum digit width must be positive');
  if (pixels === 0) return 0;
  const maxUnits = MAX_EXCEL_COLUMN_WIDTH * 256;
  let low = 1;
  let high = maxUnits;
  while (low < high) {
    const middle = Math.floor((low + high) / 2);
    const rendered = excelColumnWidthToPixels(middle / 256, maximumDigitWidthPx);
    if (rendered < pixels) low = middle + 1;
    else high = middle;
  }
  const upper = low / 256;
  const lower = Math.max(0, low - 1) / 256;
  return Math.abs(excelColumnWidthToPixels(lower, maximumDigitWidthPx) - pixels)
    <= Math.abs(excelColumnWidthToPixels(upper, maximumDigitWidthPx) - pixels)
    ? lower
    : upper;
}

export function sanitizeImportedWorkbookName(fileName: string): string {
  const leaf = fileName.replace(/\\/g, '/').split('/').pop() ?? fileName;
  const stem = leaf.replace(/\.(xlsx|xlsm|xltx|xltm)$/i, '').trim();
  const cleaned = stem.replace(/[\u0000-\u001f<>:"/\\|?*]/g, ' ').replace(/\s+/g, ' ').trim();
  return (cleaned || 'Imported Workbook').slice(0, 255);
}
