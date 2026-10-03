import { isFormulaError, type FormulaError, type FormulaValue } from './values';

export type WorkbookCollationType = 'number' | 'text' | 'boolean' | 'error' | 'blank';

export interface WorkbookCollationContext {
  readonly cultureId: string;
  readonly caseSensitive: boolean;
  readonly accentSensitive: boolean;
  readonly numericTextMode: 'lexical' | 'numeric';
  readonly blankOrder: 'first' | 'last';
  readonly typeOrder: readonly WorkbookCollationType[];
  readonly customLists: readonly (readonly string[])[];
}

export const DEFAULT_WORKBOOK_COLLATION: WorkbookCollationContext = Object.freeze({
  cultureId: 'invariant',
  caseSensitive: true,
  accentSensitive: true,
  numericTextMode: 'lexical',
  blankOrder: 'last',
  typeOrder: ['number', 'text', 'boolean', 'error', 'blank'] as const,
  customLists: [],
});

interface TextComparerEntry {
  signature: string;
  compare: (left: string, right: string) => number;
}

const textComparerCache = new WeakMap<WorkbookCollationContext, TextComparerEntry>();

export function normalizeWorkbookCollation(context?: Partial<WorkbookCollationContext>): WorkbookCollationContext {
  const next = {
    ...DEFAULT_WORKBOOK_COLLATION,
    ...context,
    typeOrder: context?.typeOrder ? [...context.typeOrder] : [...DEFAULT_WORKBOOK_COLLATION.typeOrder],
    customLists: context?.customLists?.map((list) => [...list]) ?? [],
  } satisfies WorkbookCollationContext;
  if (typeof next.cultureId !== 'string' || !next.cultureId.trim() || next.cultureId.length > 255 || typeof next.caseSensitive !== 'boolean' || typeof next.accentSensitive !== 'boolean' || !['lexical', 'numeric'].includes(next.numericTextMode) || !['first', 'last'].includes(next.blankOrder)) throw new Error('Workbook collation options are invalid');
  if (next.customLists.length > 256 || next.customLists.reduce((sum, list) => sum + list.length, 0) > 10000 || next.customLists.some((list) => list.some((value) => typeof value !== 'string' || value.length > 32767))) throw new Error('Workbook custom collation lists exceed their budget');
  if (next.typeOrder.length !== 5 || new Set(next.typeOrder).size !== 5) throw new Error('Workbook collation typeOrder must contain each value type exactly once');
  if (next.typeOrder.some((kind) => !['number', 'text', 'boolean', 'error', 'blank'].includes(kind))) throw new Error('Workbook collation has an invalid typeOrder');
  return next;
}

export function compareWorkbookValues(
  left: FormulaValue | unknown,
  right: FormulaValue | unknown,
  context: WorkbookCollationContext = DEFAULT_WORKBOOK_COLLATION,
): number {
  const leftKind = collationType(left);
  const rightKind = collationType(right);
  if (leftKind !== rightKind) return typeRank(leftKind, context) - typeRank(rightKind, context);
  if (leftKind === 'blank') return 0;
  if (leftKind === 'number' && rightKind === 'number') return compareNumbers(left as number, right as number);
  if (leftKind === 'boolean' && rightKind === 'boolean') return Number(left) - Number(right);
  if (leftKind === 'error' && rightKind === 'error') return compareInvariant(String((left as FormulaError).code), String((right as FormulaError).code));
  return compareWorkbookText(String(left ?? ''), String(right ?? ''), context);
}

export function compareWorkbookText(left: string, right: string, context: WorkbookCollationContext = DEFAULT_WORKBOOK_COLLATION): number {
  const compareText = workbookTextComparer(context);
  const leftCustom = customListRank(left, context, compareText);
  const rightCustom = customListRank(right, context, compareText);
  if (leftCustom !== rightCustom) {
    if (leftCustom === undefined) return 1;
    if (rightCustom === undefined) return -1;
    return leftCustom - rightCustom;
  }
  if (context.numericTextMode === 'numeric') {
    const numeric = compareNumericText(left, right, compareText);
    if (numeric !== undefined) return numeric;
  }
  return compareText(left, right);
}

function collationType(value: unknown): WorkbookCollationType {
  if (value === null || value === undefined || value === '') return 'blank';
  if (isFormulaError(value)) return 'error';
  if (typeof value === 'number') return 'number';
  if (typeof value === 'boolean') return 'boolean';
  return 'text';
}

function typeRank(kind: WorkbookCollationType, context: WorkbookCollationContext): number {
  if (kind === 'blank') return context.blankOrder === 'last' ? context.typeOrder.length : -1;
  return context.typeOrder.indexOf(kind);
}

function compareNumbers(left: number, right: number): number {
  if (Number.isNaN(left)) return Number.isNaN(right) ? 0 : 1;
  if (Number.isNaN(right)) return -1;
  return left < right ? -1 : left > right ? 1 : 0;
}

function normalizeText(value: string, context: WorkbookCollationContext): string {
  const accentNormalized = context.accentSensitive ? value : value.normalize('NFD').replace(/\p{M}/gu, '');
  return context.caseSensitive ? accentNormalized : accentNormalized.toLocaleLowerCase('en-US');
}

function compareInvariant(left: string, right: string): number {
  return left < right ? -1 : left > right ? 1 : 0;
}

function workbookTextComparer(context: WorkbookCollationContext): (left: string, right: string) => number {
  const cultureId = context.cultureId.trim();
  const sensitivity = context.caseSensitive
    ? context.accentSensitive ? 'variant' : 'case'
    : context.accentSensitive ? 'accent' : 'base';
  const signature = `${cultureId}\u0000${sensitivity}`;
  const cached = textComparerCache.get(context);
  if (cached?.signature === signature) return cached.compare;

  let compare: (left: string, right: string) => number;
  if (cultureId.toLowerCase() === 'invariant') {
    compare = (left, right) => compareInvariant(normalizeText(left, context), normalizeText(right, context));
  } else {
    try {
      const collator = new Intl.Collator(cultureId, { sensitivity, usage: 'sort', numeric: false });
      compare = collator.compare;
    } catch {
      throw new Error(`Workbook collation cultureId is not supported: ${cultureId}`);
    }
  }
  textComparerCache.set(context, { signature, compare });
  return compare;
}

function compareNumericText(
  left: string,
  right: string,
  compareText: (left: string, right: string) => number,
): number | undefined {
  const leftMatch = /^(.*?)(\d+)(.*?)$/.exec(left);
  const rightMatch = /^(.*?)(\d+)(.*?)$/.exec(right);
  if (!leftMatch || !rightMatch) return undefined;
  const prefix = compareText(leftMatch[1]!, rightMatch[1]!);
  if (prefix !== 0) return prefix;
  const numeric = compareIntegerText(leftMatch[2]!, rightMatch[2]!);
  return numeric !== 0 ? numeric : compareText(leftMatch[3]!, rightMatch[3]!);
}

function compareIntegerText(left: string, right: string): number {
  const leftSignificant = left.replace(/^0+(?=\d)/, '');
  const rightSignificant = right.replace(/^0+(?=\d)/, '');
  if (leftSignificant.length !== rightSignificant.length) return leftSignificant.length - rightSignificant.length;
  return compareInvariant(leftSignificant, rightSignificant);
}

const customRankCache = new WeakMap<WorkbookCollationContext, Array<{ value: string; rank: number }>>();
function customListRank(value: string, context: WorkbookCollationContext, compareText: (left: string, right: string) => number): number | undefined {
  let entries = customRankCache.get(context);
  if (!entries) {
    entries = [];
    for (const list of context.customLists) for (const [rank, item] of list.entries()) entries.push({ value: item, rank });
    if (entries.length > 10000) throw new Error('Workbook custom collation lists exceed their budget');
    entries.sort((left, right) => compareText(left.value, right.value));
    customRankCache.set(context, entries);
  }
  let low = 0; let high = entries.length;
  while (low < high) { const middle = (low + high) >>> 1; if (compareText(entries[middle]!.value, value) < 0) low = middle + 1; else high = middle; }
  return low < entries.length && compareText(entries[low]!.value, value) === 0 ? entries[low]!.rank : undefined;
}
