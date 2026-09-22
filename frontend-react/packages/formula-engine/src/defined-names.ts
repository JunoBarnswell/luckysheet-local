import type { CellAddress } from './ast';
import type { RangeDependency } from './range-index';
import { evaluateFormula } from './evaluator';
import { parseFormula } from './parser';
import { offsetAst } from './ast-rewrite';
import type { FormulaAst } from './ast';
import { createFormulaError, isArrayValue, type ArrayValue, type FormulaValue } from './values';
import type { ExcelNumericContext } from './numeric';

/** Formula-engine representation of the workbook's canonical scoped names. */
export interface FormulaDefinedName {
  readonly name: string;
  readonly formula: string;
  readonly scope: 'workbook' | 'sheet';
  readonly sheetId?: string;
  readonly anchor?: CellAddress;
}

export interface DefinedNameContext {
  currentCell: CellAddress;
  readCell: (address: CellAddress) => FormulaValue;
  readRangeMatrix: (range: RangeDependency) => ArrayValue;
  resolveName?: (name: string) => FormulaValue | undefined;
  anchor?: CellAddress;
  numericContext?: ExcelNumericContext;
}

export function normalizeDefinedNames(names: Record<string, string>): Record<string, string> {
  const normalized: Record<string, string> = {};
  for (const [name, value] of Object.entries(names)) {
    normalized[name.trim().toUpperCase()] = value;
  }
  return normalized;
}

/**
 * Normalize canonical scoped names into a deterministic, duplicate-free
 * collection. FormulaEngine stores this collection directly; a workbook-level
 * Record is only accepted as a legacy input projection.
 */
export function normalizeDefinedNameModels(names: readonly FormulaDefinedName[]): FormulaDefinedName[] {
  const normalized = new Map<string, FormulaDefinedName>();
  for (const input of names) {
    const name = input.name.trim();
    const formula = input.formula.trim();
    if (!name || !formula) throw new Error('Defined name requires a name and formula');
    if (input.scope !== 'workbook' && input.scope !== 'sheet') throw new Error(`Invalid defined name scope: ${String(input.scope)}`);
    if (input.scope === 'sheet' && !input.sheetId?.trim()) throw new Error(`Sheet-scoped defined name ${name} requires a sheetId`);
    if (input.scope === 'workbook' && input.sheetId !== undefined) throw new Error(`Workbook-scoped defined name ${name} cannot specify sheetId`);
    if (input.anchor && (!input.anchor.sheetId || !Number.isSafeInteger(input.anchor.row) || input.anchor.row < 0 || !Number.isSafeInteger(input.anchor.column) || input.anchor.column < 0)) throw new Error(`Defined name ${name} has an invalid anchor`);
    const sheetId = input.sheetId?.trim();
    const key = input.scope === 'sheet'
      ? `sheet:${stableNameKey(sheetId!)}:${stableNameKey(name)}`
      : `workbook:${stableNameKey(name)}`;
    normalized.set(key, {
      name,
      formula,
      scope: input.scope,
      ...(sheetId ? { sheetId } : {}),
      ...(input.anchor ? { anchor: structuredClone(input.anchor) } : {}),
    });
  }
  return [...normalized.values()].sort((left, right) => {
    const leftKey = `${left.scope}:${stableNameKey(left.sheetId ?? '')}:${stableNameKey(left.name)}`;
    const rightKey = `${right.scope}:${stableNameKey(right.sheetId ?? '')}:${stableNameKey(right.name)}`;
    return leftKey < rightKey ? -1 : leftKey > rightKey ? 1 : 0;
  });
}

export function resolveDefinedNameSource(source: string, context: DefinedNameContext): FormulaValue {
  const trimmed = source.trim();
  if (!trimmed) return createFormulaError('#NAME?', 'Empty defined name');

  const scalar = tryParseDefinedScalar(trimmed);
  if (scalar !== undefined) return scalar;

  try {
    const formulaText = trimmed.startsWith('=') ? trimmed : `=${trimmed}`;
    const parsed = parseFormula(formulaText);
    const ast = context.anchor
      ? offsetAst(parsed, context.currentCell.row - context.anchor.row, context.currentCell.column - context.anchor.column)
      : parsed;
    return evaluateFormula(ast, {
      currentCell: context.currentCell,
      readCell: context.readCell,
      readRange: (rangeRef) => {
        const matrix = context.readRangeMatrix(rangeRef);
        const values: FormulaValue[] = [];
        for (const row of matrix) values.push(...row);
        return values;
      },
      readRangeMatrix: context.readRangeMatrix,
      resolveName: context.resolveName,
      numericContext: context.numericContext,
    });
  } catch {
    return createFormulaError('#NAME?', `Cannot resolve defined name: ${trimmed}`);
  }
}

export function definedNameToDisplay(value: FormulaValue): string {
  if (isArrayValue(value)) {
    const rows = value.map((row) => row.map((cell) => String(cell ?? '')).join('\t')).join('\n');
    return rows;
  }
  return value == null ? '' : String(value);
}

function tryParseDefinedScalar(text: string): FormulaValue | undefined {
  if (text.startsWith('=')) return undefined;
  if (/^[A-Za-z]+!/.test(text) || text.includes(':')) return undefined;
  if (text === 'TRUE') return true;
  if (text === 'FALSE') return false;
  const numeric = Number(text);
  if (Number.isFinite(numeric) && text.trim() !== '') return numeric;
  if ((text.startsWith('"') && text.endsWith('"')) || (text.startsWith("'") && text.endsWith("'"))) {
    return text.slice(1, -1);
  }
  return undefined;
}

export function parseDefinedNameFormula(source: string): FormulaAst | undefined {
  const trimmed = source.trim();
  if (!trimmed || tryParseDefinedScalar(trimmed) !== undefined) return undefined;
  try {
    return parseFormula(trimmed.startsWith('=') ? trimmed : `=${trimmed}`);
  } catch {
    return undefined;
  }
}

function stableNameKey(value: string): string {
  return value.trim().toUpperCase();
}
