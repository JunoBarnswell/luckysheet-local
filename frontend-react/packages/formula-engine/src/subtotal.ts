/** Excel's function-number contract, shared by SUBTOTAL and Data commands. */
export const SUBTOTAL_FUNCTIONS = Object.freeze([
  'AVERAGE', 'COUNT', 'COUNTA', 'MAX', 'MIN', 'PRODUCT',
  'STDEV', 'STDEVP', 'SUM', 'VAR', 'VARP',
] as const);

export type SubtotalFunctionName = typeof SUBTOTAL_FUNCTIONS[number];

export function subtotalFunctionNumber(name: unknown): number | undefined {
  const index = SUBTOTAL_FUNCTIONS.findIndex(candidate => candidate === name);
  return index < 0 ? undefined : index + 1;
}
