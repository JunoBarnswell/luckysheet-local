export type WorkbookEnterDirection = 'down' | 'up' | 'right' | 'left';

export interface WorkbookEditingOptions {
  allowEditDirectly: boolean;
  moveAfterEnter: boolean;
  enterDirection: WorkbookEnterDirection;
  formulaAutoComplete: boolean;
  valueAutoComplete: boolean;
  /** Null disables fixed-decimal input; otherwise 0..15 decimal places. */
  fixedDecimalPlaces: number | null;
}

export const DEFAULT_WORKBOOK_EDITING_OPTIONS: WorkbookEditingOptions = Object.freeze({
  allowEditDirectly: true,
  moveAfterEnter: true,
  enterDirection: 'down',
  formulaAutoComplete: true,
  valueAutoComplete: true,
  fixedDecimalPlaces: null,
});

export function normalizeWorkbookEditingOptions(value: WorkbookEditingOptions): WorkbookEditingOptions {
  if (!isWorkbookEditingOptions(value)) throw new Error('Workbook editing options are invalid');
  return structuredClone(value);
}

export function isWorkbookEditingOptions(value: unknown): value is WorkbookEditingOptions {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false;
  const options = value as Record<string, unknown>;
  return typeof options.allowEditDirectly === 'boolean'
    && typeof options.moveAfterEnter === 'boolean'
    && ['down', 'up', 'right', 'left'].includes(String(options.enterDirection))
    && typeof options.formulaAutoComplete === 'boolean'
    && typeof options.valueAutoComplete === 'boolean'
    && (options.fixedDecimalPlaces === null || (Number.isInteger(options.fixedDecimalPlaces) && Number(options.fixedDecimalPlaces) >= 0 && Number(options.fixedDecimalPlaces) <= 15));
}
