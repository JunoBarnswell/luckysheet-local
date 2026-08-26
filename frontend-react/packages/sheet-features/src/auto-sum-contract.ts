import type { RangeRef } from '@react-sheets/core-model';

export const AUTO_SUM_FUNCTIONS = ['SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN'] as const;
export type AutoSumFunctionName = typeof AUTO_SUM_FUNCTIONS[number];
export type AutoSumInferenceMode = 'adjacent';

export interface AutoSumTarget {
  row: number;
  column: number;
}

/** Canonical semantic description carried with the single range.set write. */
export interface AutoSumDescriptor {
  functionName: AutoSumFunctionName;
  sourceRange: RangeRef;
  targets: AutoSumTarget[];
  inferenceMode: AutoSumInferenceMode;
}

export function isAutoSumDescriptor(value: unknown): value is AutoSumDescriptor {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false;
  const candidate = value as Record<string, unknown>;
  const sourceRange = candidate.sourceRange;
  if (!sourceRange || typeof sourceRange !== 'object' || Array.isArray(sourceRange)) return false;
  const range = sourceRange as Record<string, unknown>;
  if (typeof range.sheetId !== 'string'
    || !Number.isSafeInteger(range.startRow) || Number(range.startRow) < 0
    || !Number.isSafeInteger(range.endRow) || Number(range.endRow) < Number(range.startRow)
    || !Number.isSafeInteger(range.startColumn) || Number(range.startColumn) < 0
    || !Number.isSafeInteger(range.endColumn) || Number(range.endColumn) < Number(range.startColumn)) return false;
  if (!AUTO_SUM_FUNCTIONS.includes(candidate.functionName as AutoSumFunctionName)
    || candidate.inferenceMode !== 'adjacent'
    || !Array.isArray(candidate.targets)
    || candidate.targets.length === 0) return false;
  const seen = new Set<string>();
  return candidate.targets.every((target) => {
    if (!target || typeof target !== 'object' || Array.isArray(target)) return false;
    const cell = target as Record<string, unknown>;
    if (!Number.isSafeInteger(cell.row) || Number(cell.row) < 0
      || !Number.isSafeInteger(cell.column) || Number(cell.column) < 0) return false;
    const key = `${cell.row}:${cell.column}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}
