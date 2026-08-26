/**
 * The workbook model stores font family names as trimmed strings.  Names from
 * the canonical suggestion list are case-insensitive at input time but are
 * persisted using the listed spelling.  Other names are valid too: workbooks
 * commonly contain fonts installed only on the authoring machine and those
 * names must remain editable and round-trip through OOXML.
 */
export const CANONICAL_FONT_FAMILIES = [
  'Microsoft YaHei',
  'Arial',
  'Calibri',
  'Segoe UI',
  'Times New Roman',
  'Courier New',
  'Aptos',
  'Cambria',
  'Georgia',
  'Tahoma',
  'Verdana',
  'Inter',
] as const;

export type CanonicalFontFamily = typeof CANONICAL_FONT_FAMILIES[number];

export function normalizeFontFamily(value: unknown, listed: readonly string[] = CANONICAL_FONT_FAMILIES): string {
  if (typeof value !== 'string') throw new Error('Font family must be a string');
  const trimmed = value.trim();
  if (trimmed.length === 0) throw new Error('Font family must not be empty');
  if ([...trimmed].some((character) => character.charCodeAt(0) < 0x20 || character === '\u007f')) {
    throw new Error('Font family contains control characters');
  }
  const folded = trimmed.toLowerCase();
  return listed.find((candidate) => candidate.toLowerCase() === folded) ?? trimmed;
}

export function isCanonicalFontFamily(value: unknown): value is string {
  if (typeof value !== 'string') return false;
  try {
    return normalizeFontFamily(value) === value;
  } catch {
    return false;
  }
}
