import { createFormulaError, isFormulaError, type FormulaError, type FormulaErrorCode } from './values';

export type FormulaInputFault = FormulaError & { readonly inputFault: NonNullable<FormulaError['inputFault']> };

export function isFormulaInputFault(value: unknown): value is FormulaInputFault {
  if (!isFormulaError(value) || !value.inputFault || typeof value.inputFault !== 'object') return false;
  return ['access-denied', 'source-unavailable', 'source-missing', 'runtime-unavailable'].includes(value.inputFault.reason)
    && typeof value.inputFault.source === 'string' && value.inputFault.source.length > 0;
}

export function createFormulaInputFault(code: FormulaErrorCode, message: string,
  reason: FormulaInputFault['inputFault']['reason'], source: string): FormulaInputFault {
  if (!source) throw new Error('FORMULA_INPUT_FAULT_SOURCE_REQUIRED');
  return { ...createFormulaError(code, message), inputFault: Object.freeze({ reason, source }) };
}
