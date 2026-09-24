import { FormulaReferenceError } from './errors';

export interface FormulaSheetIdentity {
  readonly id: string;
  readonly name: string;
}

export function formulaSheetReferenceIndex(
  reference: string,
  sheetOrder: readonly FormulaSheetIdentity[] | undefined,
): number {
  if (!sheetOrder) return -1;
  const normalized = reference.toLowerCase();
  const byName = sheetOrder.findIndex((sheet) => sheet.name.toLowerCase() === normalized);
  return byName >= 0 ? byName : sheetOrder.findIndex((sheet) => sheet.id === reference);
}

export function sameFormulaSheetName(left: string | undefined, right: string): boolean {
  return left?.toLowerCase() === right.toLowerCase();
}

export function resolveFormulaSheetId(
  reference: string | undefined,
  ownerSheetId: string,
  sheetOrder: readonly FormulaSheetIdentity[],
): string {
  if (reference === undefined) return ownerSheetId;
  if (sheetOrder.length === 0) {
    throw new FormulaReferenceError(`Worksheet identity order is required to resolve: ${reference}`);
  }
  const index = formulaSheetReferenceIndex(reference, sheetOrder);
  const sheet = index >= 0 ? sheetOrder[index] : undefined;
  if (!sheet) throw new FormulaReferenceError(`Reference worksheet cannot be resolved: ${reference}`);
  return sheet.id;
}
