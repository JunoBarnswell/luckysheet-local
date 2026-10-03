import { normalizeFontFamily } from './font-family';
import { isCellEditorConfig } from './cell-editor';
/** Resource limits enforced at the canonical workbook boundary before projection. */
export function assertWorkbookResourceBudget(value: unknown): void {
  const pending: Array<{ value: unknown; depth: number; key: string }> = [{ value, depth: 0, key: '' }];
  let nodes = 0;
  let text = 0;
  while (pending.length) {
    const item = pending.pop()!;
    if (++nodes > 2_000_000 || item.depth > 64) throw new Error('UNSUPPORTED_FEATURE: Workbook structure exceeds its resource budget');
    if (item.key === "fontFamily") normalizeFontFamily(item.value);
    if (item.key === "editor" && item.value !== null && !isCellEditorConfig(item.value)) throw new Error("Invalid cell editor");
    if (item.key === "numberFormat" && typeof item.value !== "string") throw new Error("Number format must be a string");
    if (typeof item.value === 'string') {
      text += item.value.length;
      const limit = item.key === 'numberFormat' ? 255 : item.key === 'svgPath' ? 16_384 : 32_767;
      if (item.value.length > limit || text > 16 * 1024 * 1024) throw new Error('UNSUPPORTED_FEATURE: Workbook text exceeds its resource budget');
    } else if (typeof item.value === 'number' && ['width', 'height', 'widthPx', 'heightPx', 'fontSize', 'fontSizePx', 'defaultRowHeightPx', 'defaultColumnWidthPx'].includes(item.key)) {
      if (!Number.isFinite(item.value) || item.value < 0 || item.value > 8192) throw new Error('UNSUPPORTED_FEATURE: Workbook pixel geometry exceeds its resource budget');
    } else if (item.value && typeof item.value === 'object') {
      for (const [key, child] of Object.entries(item.value)) pending.push({ value: child, depth: item.depth + 1, key: item.key === 'rowHeightsPx' ? 'heightPx' : item.key === 'columnWidthsPx' ? 'widthPx' : key });
      if (pending.length > 2_000_000) throw new Error('UNSUPPORTED_FEATURE: Workbook structure exceeds its resource budget');
    }
  }
}
