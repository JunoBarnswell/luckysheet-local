/** Spreadsheet alignment values are deliberately not Canvas/CSS values. */
export const HORIZONTAL_ALIGNMENTS = [
  'general',
  'left',
  'center',
  'right',
  'centerContinuous',
  'justify',
  'distributed',
  'fill',
] as const;

export type HorizontalAlignment = typeof HORIZONTAL_ALIGNMENTS[number];

export const VERTICAL_ALIGNMENTS = ['top', 'middle', 'bottom', 'justify', 'distributed'] as const;
export type VerticalAlignment = typeof VERTICAL_ALIGNMENTS[number];

export const READING_ORDERS = ['context', 'ltr', 'rtl'] as const;
export type ReadingOrder = typeof READING_ORDERS[number];

export const TEXT_ORIENTATIONS = ['horizontal', 'stacked', 'rotateUp', 'rotateDown'] as const;
export type TextOrientation = typeof TEXT_ORIENTATIONS[number];

/**
 * Native alignment attributes that the editor cannot execute yet. They are
 * explicit snapshot data, not a value that may leak into Canvas textAlign.
 */
export interface UnsupportedCellAlignment {
  horizontal?: string;
  vertical?: string;
  attributes?: Record<string, string>;
}

export function isHorizontalAlignment(value: unknown): value is HorizontalAlignment {
  return (HORIZONTAL_ALIGNMENTS as readonly unknown[]).includes(value);
}

export function isVerticalAlignment(value: unknown): value is VerticalAlignment {
  return (VERTICAL_ALIGNMENTS as readonly unknown[]).includes(value);
}

export function isReadingOrder(value: unknown): value is ReadingOrder {
  return (READING_ORDERS as readonly unknown[]).includes(value);
}
