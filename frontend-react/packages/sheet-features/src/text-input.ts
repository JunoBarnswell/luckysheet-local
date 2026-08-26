import { clearFormulaProvenance, type CellData, type CellStyle } from '@react-sheets/core-model';
import { formatFormula, parseFormula } from '@react-sheets/formula-engine';

/** Parsed result of one user/host text input. */
export interface ParsedCellText {
  readonly value: CellData['value'];
  readonly formula?: string;
}

/**
 * Parse one text value using spreadsheet input semantics.
 *
 * Leading/trailing whitespace is significant for text. We only coerce a
 * number/boolean when the complete input is the literal token, so `" 42 "`
 * remains text rather than silently losing user-entered spaces. A leading
 * `=` is the formula delimiter; its original text is retained after AST
 * validation so formula-bar spacing is not rewritten by the command layer.
 */
export function parseCellText(text: string): ParsedCellText {
  if (text.startsWith('=')) {
    // Validate before opening a mutation. The AST is the formula contract;
    // no regex is used to inspect or rewrite formula references.
    parseFormula(text);
    return { value: null, formula: text };
  }
  if (text === '') return { value: null };
  if (isNumberLiteral(text)) {
    const number = Number(text);
    if (Number.isFinite(number)) return { value: number };
  }
  if (text === 'TRUE' || text === 'true') return { value: true };
  if (text === 'FALSE' || text === 'false') return { value: false };
  return { value: text };
}

/**
 * Construct the next cell while retaining its presentation and auxiliary
 * metadata. Content-only fields which would otherwise be stale are removed.
 */
export function buildCellFromText(
  text: string,
  previous: CellData | undefined,
  style?: Partial<CellStyle>,
): CellData {
  const parsed = parseCellText(text);
  const next: CellData = clearFormulaProvenance(previous ? structuredClone(previous) : { value: null });
  next.value = parsed.value;
  delete next.formula;
  delete next.formulaValue;
  delete next.displayValue;
  if (parsed.formula !== undefined) next.formula = parsed.formula;
  if (style !== undefined) next.style = { ...(next.style ?? {}), ...structuredClone(style) };
  return next;
}

/**
 * Return the canonical formula text for callers that need to inspect a parsed
 * formula without mutating the user's original input.
 */
export function canonicalizeFormulaText(formula: string): string {
  return formatFormula(parseFormula(formula));
}

function isNumberLiteral(text: string): boolean {
  if (text.length === 0) return false;
  let index = 0;
  const first = text[index];
  if (first === '+' || first === '-') index += 1;

  let integerDigits = 0;
  while (isDigit(text[index])) {
    integerDigits += 1;
    index += 1;
  }
  let fractionDigits = 0;
  if (text[index] === '.') {
    index += 1;
    while (isDigit(text[index])) {
      fractionDigits += 1;
      index += 1;
    }
  }
  if (integerDigits + fractionDigits === 0) return false;

  if (text[index] === 'e' || text[index] === 'E') {
    index += 1;
    if (text[index] === '+' || text[index] === '-') index += 1;
    let exponentDigits = 0;
    while (isDigit(text[index])) {
      exponentDigits += 1;
      index += 1;
    }
    if (exponentDigits === 0) return false;
  }
  return index === text.length;
}

function isDigit(value: string | undefined): boolean {
  return value !== undefined && value >= '0' && value <= '9';
}
