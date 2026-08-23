import type { CellAddress } from './ast';
import { parseCellAddress } from './address';
import { normalizeRange, type RangeDependency } from './range-index';
import { evaluateFormula } from './evaluator';
import { parseFormula } from './parser';
import { createFormulaError, isArrayValue, type ArrayValue, type FormulaValue } from './values';

export interface DefinedNameContext {
  currentCell: CellAddress;
  readCell: (address: CellAddress) => FormulaValue;
  readRangeMatrix: (range: RangeDependency) => ArrayValue;
}

export function normalizeDefinedNames(names: Record<string, string>): Record<string, string> {
  const normalized: Record<string, string> = {};
  for (const [name, value] of Object.entries(names)) {
    normalized[name.trim().toUpperCase()] = value;
  }
  return normalized;
}

export function resolveDefinedNameSource(source: string, context: DefinedNameContext): FormulaValue {
  const trimmed = source.trim();
  if (!trimmed) return createFormulaError('#NAME?', 'Empty defined name');

  const range = tryParseDefinedRange(trimmed, context.currentCell.sheetId);
  if (range) return context.readRangeMatrix(range);

  const scalar = tryParseDefinedScalar(trimmed);
  if (scalar !== undefined) return scalar;

  try {
    const formulaText = trimmed.startsWith('=') ? trimmed : `=${trimmed}`;
    const ast = parseFormula(formulaText);
    return evaluateFormula(ast, {
      currentCell: context.currentCell,
      readCell: context.readCell,
      readRange: (rangeRef) => {
        const matrix = context.readRangeMatrix(rangeRef);
        const values: FormulaValue[] = [];
        for (const row of matrix) values.push(...row);
        return values;
      },
      resolveName: undefined,
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

function tryParseDefinedRange(text: string, defaultSheetId: string): RangeDependency | undefined {
  if (text.startsWith('=')) return undefined;
  const rangeText = text.includes(':') ? text : undefined;
  if (!rangeText) return undefined;
  const [startText, endText] = rangeText.split(':');
  if (!startText || !endText) return undefined;
  try {
    const start = parseCellAddress(startText, defaultSheetId);
    const end = parseCellAddress(endText, defaultSheetId);
    return normalizeRange(start, end);
  } catch {
    return undefined;
  }
}
