import type { CellValue } from '@react-sheets/core-model';
export { MAX_OBJECT_RANGE_CELLS } from '@react-sheets/sheet-features';
export function isCellValue(value: unknown): value is CellValue {
  return value === null || typeof value === 'string' || typeof value === 'boolean' || typeof value === 'number' && Number.isFinite(value);
}
export function immutableSnapshot<T>(value: T): T {
  const copy = structuredClone(value);
  const freeze = (node: unknown): void => {
    if (node && typeof node === 'object') { for (const item of Object.values(node)) freeze(item); Object.freeze(node); }
  };
  freeze(copy);
  return copy;
}
