export const AUTO_SUM_FUNCTIONS = ['SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN'] as const;
export type AutoSumFunctionName = typeof AUTO_SUM_FUNCTIONS[number];
