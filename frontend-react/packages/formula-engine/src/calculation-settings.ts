export type WorkbookCalculationMode = 'automatic' | 'automatic-except-data-tables' | 'manual';

export interface WorkbookCalculationSettings {
  readonly mode: WorkbookCalculationMode;
  readonly iterativeCalculation: boolean;
  readonly maximumIterations: number;
  readonly maximumChange: number;
  readonly precisionAsDisplayed: boolean;
  readonly calculateBeforeSave: boolean;
  readonly fullCalculationOnLoad: boolean;
}

export const DEFAULT_WORKBOOK_CALCULATION_SETTINGS: WorkbookCalculationSettings = {
  mode: 'automatic',
  iterativeCalculation: false,
  maximumIterations: 100,
  maximumChange: 0.001,
  precisionAsDisplayed: false,
  calculateBeforeSave: true,
  fullCalculationOnLoad: false,
};

export function normalizeWorkbookCalculationSettings(
  settings: Partial<WorkbookCalculationSettings> = {},
): WorkbookCalculationSettings {
  const result = {
    ...DEFAULT_WORKBOOK_CALCULATION_SETTINGS,
    ...settings,
  };
  if (!Number.isSafeInteger(result.maximumIterations) || result.maximumIterations < 1) {
    throw new Error('Workbook calculation maximumIterations must be a positive integer');
  }
  if (!Number.isFinite(result.maximumChange) || result.maximumChange < 0) {
    throw new Error('Workbook calculation maximumChange must be non-negative');
  }
  return Object.freeze(result);
}

export function isWorkbookCalculationSettings(value: unknown): value is WorkbookCalculationSettings {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) return false;
  const settings = value as Partial<WorkbookCalculationSettings>;
  return (settings.mode === 'automatic' || settings.mode === 'automatic-except-data-tables' || settings.mode === 'manual')
    && typeof settings.iterativeCalculation === 'boolean'
    && Number.isSafeInteger(settings.maximumIterations)
    && settings.maximumIterations >= 1
    && typeof settings.maximumChange === 'number'
    && Number.isFinite(settings.maximumChange)
    && settings.maximumChange >= 0
    && typeof settings.precisionAsDisplayed === 'boolean'
    && typeof settings.calculateBeforeSave === 'boolean'
    && typeof settings.fullCalculationOnLoad === 'boolean';
}
