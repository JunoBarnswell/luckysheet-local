import type { FormulaValue } from '../values';
import { FINANCIAL_FUNCTION_CORPUS } from './financial-corpus';

export interface FunctionCase { readonly formula: string; readonly expected: FormulaValue }
/** Fixed examples with independently specified results, shared with browser acceptance. */
const vectors: readonly (readonly [string, string, FormulaValue])[] = [
  ['ABS', 'ABS(-2)', 2], ['EXP', 'EXP(0)', 1], ['SQRT', 'SQRT(4)', 2], ['POWER', 'POWER(2,3)', 8],
  ['MOD', 'MOD(-3,2)', 1], ['ROUND', 'ROUND(2.25,1)', 2.3], ['ROUNDUP', 'ROUNDUP(2.21,1)', 2.3], ['ROUNDDOWN', 'ROUNDDOWN(2.29,1)', 2.2],
  ['INT', 'INT(-2.5)', -3], ['TRUNC', 'TRUNC(-2.5)', -2], ['CEILING', 'CEILING(2.2,1)', 3], ['FLOOR', 'FLOOR(2.8,1)', 2],
  ['PI', 'PI()', Math.PI], ['RAND', 'RAND()', null], ['RANDBETWEEN', 'RANDBETWEEN(2,2)', 2], ['LN', 'LN(1)', 0],
  ['LOG', 'LOG(100,10)', 2], ['LOG10', 'LOG10(100)', 2], ['SIGN', 'SIGN(-3)', -1], ['SUM', 'SUM(A1:A3)', 12], ['PRODUCT', 'PRODUCT(A1:A3)', 48],
  ['AVERAGE', 'AVERAGE(A1:A3)', 4], ['COUNT', 'COUNT(A1:A3)', 3], ['COUNTA', 'COUNTA(B1:B3)', 3], ['MIN', 'MIN(A1:A3)', 2], ['MAX', 'MAX(A1:A3)', 6],
  ['VAR', 'VAR(A1:A2)', 2], ['VAR.S', 'VAR.S(A1:A2)', 2], ['VARP', 'VARP(A1:A2)', 1], ['VAR.P', 'VAR.P(A1:A2)', 1],
  ['STDEV', 'STDEV(A1:A2)', Math.SQRT2], ['STDEV.S', 'STDEV.S(A1:A2)', Math.SQRT2], ['STDEVP', 'STDEVP(A1:A2)', 1], ['STDEV.P', 'STDEV.P(A1:A2)', 1],
  ['COUNTBLANK', 'COUNTBLANK(A1:A5)', 2], ['MEDIAN', 'MEDIAN(A1:A3)', 4], ['LARGE', 'LARGE(A1:A3,2)', 4], ['SMALL', 'SMALL(A1:A3,2)', 4],
  ['COUNTIF', 'COUNTIF(B1:B3,"x")', 2], ['SUMIF', 'SUMIF(B1:B3,"x",A1:A3)', 8], ['AVERAGEIF', 'AVERAGEIF(B1:B3,"x",A1:A3)', 4],
  ['SUMIFS', 'SUMIFS(A1:A3,B1:B3,"x")', 8], ['COUNTIFS', 'COUNTIFS(B1:B3,"x",A1:A3,">3")', 1],
  ['AVERAGEIFS', 'AVERAGEIFS(A1:A3,B1:B3,"x")', 4], ['MAXIFS', 'MAXIFS(A1:A3,B1:B3,"x")', 6], ['MINIFS', 'MINIFS(A1:A3,B1:B3,"x")', 2],
  ['SUMPRODUCT', 'SUMPRODUCT(A1:A3,A1:A3)', 56], ['SUBTOTAL', 'SUBTOTAL(9,A1:A3)', 12], ['AGGREGATE', 'AGGREGATE(9,5,A1:A3)', 12],
  ['IF', 'IF(FALSE,1/0,7)', 7], ['IFS', 'IFS(FALSE,1,TRUE,2)', 2], ['IFERROR', 'IFERROR(1/0,7)', 7], ['IFNA', 'IFNA(#N/A,7)', 7],
  ['AND', 'AND(C1:C3)', false], ['OR', 'OR(C1:C3)', true], ['NOT', 'NOT(TRUE)', false], ['XOR', 'XOR(C1:C3)', false],
  ['SWITCH', 'SWITCH(2,1,"a",2,"b","c")', 'b'], ['TRUE', 'TRUE()', true], ['FALSE', 'FALSE()', false],
  ['CONCAT', 'CONCAT(B1:B3)', 'xyx'], ['CONCATENATE', 'CONCATENATE("a","b")', 'ab'], ['TEXTJOIN', 'TEXTJOIN(",",TRUE,B1:B3)', 'x,y,x'],
  ['LEFT', 'LEFT("abc",2)', 'ab'], ['RIGHT', 'RIGHT("abc",2)', 'bc'], ['MID', 'MID("abcd",2,2)', 'bc'], ['LEN', 'LEN("abc")', 3],
  ['LOWER', 'LOWER("AbC")', 'abc'], ['UPPER', 'UPPER("AbC")', 'ABC'], ['PROPER', 'PROPER("hELLO wORLD")', 'Hello World'],
  ['TRIM', 'TRIM("  a  b  ")', 'a b'], ['CLEAN', 'CLEAN("a"&CHAR(10)&"b")', 'ab'], ['EXACT', 'EXACT("a","A")', false],
  ['FIND', 'FIND("b","abc")', 2], ['SEARCH', 'SEARCH("B","abc")', 2], ['REPLACE', 'REPLACE("abcd",2,2,"x")', 'axd'],
  ['SUBSTITUTE', 'SUBSTITUTE("ababa","a","x",2)', 'abxba'], ['REPT', 'REPT("ab",2)', 'abab'], ['TEXT', 'TEXT(2.5,"0.00")', '2.50'],
  ['VALUE', 'VALUE("25%")', 0.25], ['CHAR', 'CHAR(65)', 'A'], ['CODE', 'CODE("A")', 65],
  ['VLOOKUP', 'VLOOKUP(4,A1:B3,2,FALSE)', 'y'], ['HLOOKUP', 'HLOOKUP(2,A1:C3,2,FALSE)', 4], ['INDEX', 'INDEX(A1:A3,2)', 4],
  ['MATCH', 'MATCH(4,A1:A3,0)', 2], ['XLOOKUP', 'XLOOKUP(4,A1:A3,B1:B3)', 'y'], ['CHOOSE', 'CHOOSE(2,"a","b")', 'b'],
  ['ROWS', 'ROWS(A1:B3)', 3], ['COLUMNS', 'COLUMNS(A1:B3)', 2], ['TRANSPOSE', 'TRANSPOSE(A1:A3)', [[2,4,6]]],
  ['ROW', 'ROW(A2)', 2], ['COLUMN', 'COLUMN(B2)', 2], ['ADDRESS', 'ADDRESS(2,3,2)', 'C$2'], ['OFFSET', 'SUM(OFFSET(A1:A2,1,0))', 10], ['INDIRECT', 'INDIRECT("A2")', 4],
  ['DATE', 'DATE(2024,1,1)', 45292], ['DATEVALUE', 'DATEVALUE("2024-01-01")', 45292], ['DAY', 'DAY(45292)', 1], ['MONTH', 'MONTH(45292)', 1], ['YEAR', 'YEAR(45292)', 2024],
  ['HOUR', 'HOUR(0.5)', 12], ['MINUTE', 'MINUTE(0.5)', 0], ['SECOND', 'SECOND(0.5)', 0], ['TODAY', 'TODAY()', 45293], ['NOW', 'NOW()', 45293.5],
  ['WEEKDAY', 'WEEKDAY(45292,2)', 1], ['EDATE', 'EDATE(45292,1)', 45323], ['EOMONTH', 'EOMONTH(45292,0)', 45322], ['DAYS', 'DAYS(45293,45292)', 1],
  ['ISBLANK', 'ISBLANK(A4)', true], ['ISNUMBER', 'ISNUMBER(A1)', true], ['ISTEXT', 'ISTEXT(B1)', true], ['ISNONTEXT', 'ISNONTEXT(A1)', true],
  ['ISLOGICAL', 'ISLOGICAL(C1)', true], ['ISERROR', 'ISERROR(1/0)', true], ['ISERR', 'ISERR(#N/A)', false], ['ISNA', 'ISNA(#N/A)', true], ['N', 'N(TRUE)', 1], ['T', 'T(2)', ''],
  ['FILTER', 'FILTER(A1:A3,C1:C3)', [[2],[6]]], ['UNIQUE', 'UNIQUE(B1:B3)', [['x'],['y']]], ['SORT', 'SORT(A1:A3,1,-1)', [[6],[4],[2]]],
  ['SEQUENCE', 'SEQUENCE(2,2)', [[1,2],[3,4]]], ['XMATCH', 'XMATCH(4,A1:A3)', 2], ['HSTACK', 'HSTACK(A1:A2,B1:B2)', [[2,'x'],[4,'y']]],
  ['VSTACK', 'VSTACK(A1:A2,A3)', [[2],[4],[6]]], ['TAKE', 'TAKE(A1:A3,-2)', [[4],[6]]], ['DROP', 'DROP(A1:A3,1)', [[4],[6]]],
  ['SORTBY', 'SORTBY(B1:B3,A1:A3,-1)', [['x'],['y'],['x']]], ['RANDARRAY', 'RANDARRAY(2,2,2,2,TRUE)', [[2,2],[2,2]]],
  ['GROUPBY', 'GROUPBY(B1:B3,A1:A3,SUM)', [['x',8],['y',4]]], ['PIVOTBY', 'PIVOTBY(B1:B3,C1:C3,A1:A3,SUM)', [[null,true,false],['x',8,0],['y',0,4]]],
  ['LET', 'LET(amount,2,amount*3)', 6], ['LAMBDA', 'LAMBDA(amount,amount*3)(2)', 6], ['SJS.TABLE', 'SJS.TABLE(A1*D1,A1:A3,D1)', [[4],[8],[12]]],
];
export const FUNCTION_CORPUS: Readonly<Record<string, FunctionCase>> = {
  ...Object.fromEntries(vectors.map(([id, formula, expected]) => [id, { formula: `=${formula}`, expected }])),
  ...FINANCIAL_FUNCTION_CORPUS,
};
