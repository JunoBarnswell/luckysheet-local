import { createFormulaError, type FormulaError } from '../values';

export interface FunctionArgumentContract { readonly minimum: number; readonly maximum: number }
export const FUNCTION_ARGUMENT_CONTRACTS = new Map<string, FunctionArgumentContract>();
for (const [names, minimum, maximum] of [
  ['ABS CHAR CLEAN CODE EXP DATEVALUE DAY HOUR INT ISBLANK ISERR ISERROR ISLOGICAL ISNA ISNONTEXT ISNUMBER ISTEXT LEN LN LOG10 LOWER MINUTE MONTH N NOT PROPER SECOND SIGN SQRT T TRANSPOSE TRIM UPPER VALUE YEAR', 1, 1],
  ['DATE MID', 3, 3],
  ['REPLACE', 4, 4],
  ['DAYS EDATE EOMONTH EXACT IFERROR IFNA LARGE MOD POWER RANDBETWEEN REPT SMALL TEXT', 2, 2],
  ['CEILING FLOOR LEFT LOG RIGHT ROUND ROUNDDOWN ROUNDUP TRUNC WEEKDAY', 1, 2],
  ['IF SUMIF AVERAGEIF', 2, 3],
  ['COUNTIF', 2, 2],
  ['COUNTBLANK ROWS COLUMNS', 1, 1],
  ['FIND SEARCH TAKE DROP MATCH', 2, 3],
  ['SUBSTITUTE VLOOKUP HLOOKUP XMATCH INDEX', 3, 4],
  ['INDEX', 2, 4],
  ['XMATCH', 2, 4],
  ['FILTER', 2, 3],
  ['SORT SEQUENCE', 1, 4],
  ['UNIQUE', 1, 3],
  ['RANDARRAY', 0, 5],
  ['XLOOKUP', 3, 6],
  ['SUM AVERAGE COUNT COUNTA MIN MAX PRODUCT VAR VAR.S VARP VAR.P STDEV STDEV.S STDEVP STDEV.P MEDIAN', 0, 255],
  ['AND OR XOR CONCAT CONCATENATE HSTACK VSTACK SUMPRODUCT', 1, 255],
  ['SUMIFS AVERAGEIFS MINIFS MAXIFS', 3, 255],
  ['COUNTIFS IFS', 2, 254],
  ['SORTBY', 2, 255],
  ['CHOOSE SUBTOTAL', 2, 255],
  ['SWITCH SJS.TABLE LET', 3, 255],
  ['AGGREGATE', 3, 255],
  ['TEXTJOIN', 3, 252],
  ['GROUPBY', 3, 3],
  ['PIVOTBY', 4, 4],
  ['LAMBDA', 1, 254],
  ['FALSE TRUE PI RAND NOW TODAY', 0, 0],
  ['ROW COLUMN', 0, 1],
  ['ADDRESS', 2, 5],
  ['OFFSET', 3, 5],
  ['INDIRECT', 1, 2],
] as const) for (const name of names.split(' ')) FUNCTION_ARGUMENT_CONTRACTS.set(name, { minimum, maximum });

export function validateFunctionArguments(name: string, count: number): FormulaError | undefined {
  const contract = FUNCTION_ARGUMENT_CONTRACTS.get(name);
  return contract && (count < contract.minimum || count > contract.maximum) ? createFormulaError('#VALUE!', `${name} requires ${contract.minimum}-${contract.maximum} arguments; received ${count}`) : undefined;
}

export const SCALAR_ARRAY_FUNCTIONS = new Set(('ABS EXP SQRT POWER MOD ROUND ROUNDUP ROUNDDOWN INT TRUNC CEILING FLOOR LN LOG LOG10 SIGN LEFT RIGHT MID LEN LOWER UPPER PROPER TRIM CLEAN EXACT FIND SEARCH REPLACE SUBSTITUTE REPT TEXT VALUE CHAR CODE DATE DATEVALUE DAY MONTH YEAR HOUR MINUTE SECOND WEEKDAY EDATE EOMONTH DAYS ISBLANK ISNUMBER ISTEXT ISNONTEXT ISLOGICAL ISERROR ISERR ISNA N T NOT').split(' '));
