import type { CellData, CellValue, RangeRef, WorkbookModel } from '@react-sheets/core-model';

/**
 * The automation language is deliberately a small, data-only DSL.  It is
 * parsed into this AST before anything is sent to CommandRuntime.  There is
 * no JavaScript evaluation step (and therefore no string based "sandbox").
 */
export type FacadeStatement =
  | { kind: 'set-values'; range: A1Range; values: unknown[][] }
  | { kind: 'set-font-weight'; range: A1Range; weight: 'normal' | 'bold' }
  | { kind: 'clear'; range: A1Range };

export interface FacadeProgram {
  readonly statements: readonly FacadeStatement[];
  readonly sourceLength: number;
  readonly cellCount: number;
}

export interface A1Range {
  readonly startRow: number;
  readonly startColumn: number;
  readonly endRow: number;
  readonly endColumn: number;
  readonly text: string;
}

export interface FacadeDslLimits {
  readonly maxSourceLength: number;
  readonly maxStatements: number;
  readonly maxCells: number;
}

export const DEFAULT_FACADE_DSL_LIMITS: FacadeDslLimits = {
  maxSourceLength: 256 * 1024,
  maxStatements: 1000,
  maxCells: 100_000,
};

export interface FacadeCellOperation {
  readonly kind: 'set-cell' | 'set-style' | 'clear-cell';
  readonly sheetId: string;
  readonly row: number;
  readonly column: number;
  readonly value?: CellData;
  readonly style?: { bold: boolean };
}

export interface FacadePlan {
  readonly statements: readonly FacadeStatement[];
  readonly operations: readonly FacadeCellOperation[];
  readonly affectedRanges: readonly RangeRef[];
}

export interface FacadeExecutionControl {
  readonly signal?: AbortSignal;
  readonly deadlineAt?: number;
}

export function checkFacadeExecution(control?: FacadeExecutionControl): void {
  if (control?.signal?.aborted) throw new Error('Automation execution cancelled');
  if (control?.deadlineAt !== undefined && Date.now() > control.deadlineAt) {
    throw new Error('Automation execution timed out');
  }
}

class DslParser {
  private index = 0;
  private readonly statements: FacadeStatement[] = [];
  private cellCount = 0;

  constructor(
    private readonly source: string,
    private readonly limits: FacadeDslLimits,
  ) {}

  parse(): FacadeProgram {
    if (this.source.length > this.limits.maxSourceLength) {
      throw new Error(`Automation source exceeds ${this.limits.maxSourceLength} characters`);
    }

    this.skipWhitespace();
    while (!this.atEnd()) {
      if (this.statements.length >= this.limits.maxStatements) {
        throw new Error(`Automation script exceeds ${this.limits.maxStatements} statements`);
      }
      this.statements.push(this.parseStatement());
      this.skipWhitespace();
      if (this.peek() === ';') {
        this.index += 1;
        this.skipWhitespace();
      } else if (!this.atEnd()) {
        this.fail("Expected ';' between automation statements");
      }
    }

    if (this.statements.length === 0) throw new Error('Automation script cannot be empty');

    return {
      statements: this.statements,
      sourceLength: this.source.length,
      cellCount: this.cellCount,
    };
  }

  private parseStatement(): FacadeStatement {
    this.expectWord('sheet');
    this.expectLiteral('.getRange');
    this.expect('(');
    const rangeText = this.readString();
    this.expect(')');
    this.expectLiteral('.');
    const method = this.readIdentifier();
    this.expect('(');

    const range = parseA1Range(rangeText);
    const cells = rangeCellCount(range);
    this.cellCount += cells;
    if (this.cellCount > this.limits.maxCells) {
      throw new Error(`Automation script addresses more than ${this.limits.maxCells} cells`);
    }

    if (method === 'setValues') {
      const values = this.readLiteralArray();
      this.expect(')');
      validateValues(values);
      return { kind: 'set-values', range, values };
    }
    if (method === 'setFontWeight') {
      const weight = this.readString();
      this.expect(')');
      if (weight !== 'normal' && weight !== 'bold') {
        this.fail("setFontWeight only accepts 'normal' or 'bold'");
      }
      return { kind: 'set-font-weight', range, weight };
    }
    if (method === 'clear') {
      this.skipWhitespace();
      if (this.peek() !== ')') this.fail('clear does not accept arguments');
      this.expect(')');
      return { kind: 'clear', range };
    }

    this.fail(`Unsupported Facade method: ${method}`);
  }

  private readLiteralArray(): unknown[][] {
    const value = this.readLiteral();
    if (!Array.isArray(value) || value.some((row) => !Array.isArray(row))) {
      this.fail('setValues requires a two-dimensional literal array');
    }
    return value as unknown[][];
  }

  private readLiteral(): unknown {
    this.skipWhitespace();
    const char = this.peek();
    if (char === '[') {
      this.index += 1;
      const values: unknown[] = [];
      this.skipWhitespace();
      if (this.peek() === ']') {
        this.index += 1;
        return values;
      }
      while (true) {
        values.push(this.readLiteral());
        this.skipWhitespace();
        if (this.peek() === ',') {
          this.index += 1;
          this.skipWhitespace();
          if (this.peek() === ']') this.fail('Trailing commas are not part of the automation DSL');
          continue;
        }
        this.expect(']');
        return values;
      }
    }
    if (char === '{') return this.readObjectLiteral();
    if (char === '"' || char === "'") return this.readString();
    if (this.source.startsWith('true', this.index)) {
      this.index += 4;
      return true;
    }
    if (this.source.startsWith('false', this.index)) {
      this.index += 5;
      return false;
    }
    if (this.source.startsWith('null', this.index)) {
      this.index += 4;
      return null;
    }
    return this.readNumber();
  }

  private readObjectLiteral(): Record<string, unknown> {
    this.expect('{');
    const result: Record<string, unknown> = {};
    this.skipWhitespace();
    if (this.peek() === '}') {
      this.index += 1;
      return result;
    }
    while (true) {
      this.skipWhitespace();
      const key = this.peek() === '"' || this.peek() === "'" ? this.readString() : this.readIdentifier();
      this.skipWhitespace();
      this.expect(':');
      if (key in result) this.fail(`Duplicate object key: ${key}`);
      result[key] = this.readLiteral();
      this.skipWhitespace();
      if (this.peek() === ',') {
        this.index += 1;
        continue;
      }
      this.expect('}');
      return result;
    }
  }

  private readNumber(): number {
    const start = this.index;
    if (this.peek() === '-') this.index += 1;
    let digits = 0;
    while (isDigit(this.peek())) {
      this.index += 1;
      digits += 1;
    }
    if (this.peek() === '.') {
      this.index += 1;
      while (isDigit(this.peek())) {
        this.index += 1;
        digits += 1;
      }
    }
    if (digits === 0) this.fail('Expected a literal value');
    if (this.peek() === 'e' || this.peek() === 'E') {
      this.index += 1;
      if (this.peek() === '+' || this.peek() === '-') this.index += 1;
      let exponentDigits = 0;
      while (isDigit(this.peek())) {
        this.index += 1;
        exponentDigits += 1;
      }
      if (exponentDigits === 0) this.fail('Invalid numeric literal');
    }
    const parsed = Number(this.source.slice(start, this.index));
    if (!Number.isFinite(parsed)) this.fail('Numeric literal must be finite');
    return parsed;
  }

  private readString(): string {
    this.skipWhitespace();
    const quote = this.peek();
    if (quote !== '"' && quote !== "'") this.fail('Expected a quoted string');
    this.index += 1;
    let result = '';
    while (!this.atEnd()) {
      const char = this.source[this.index]!;
      this.index += 1;
      if (char === quote) return result;
      if (char !== '\\') {
        result += char;
        continue;
      }
      if (this.atEnd()) this.fail('Unterminated string literal');
      const escaped = this.source[this.index]!;
      this.index += 1;
      switch (escaped) {
        case 'n': result += '\n'; break;
        case 'r': result += '\r'; break;
        case 't': result += '\t'; break;
        case 'b': result += '\b'; break;
        case 'f': result += '\f'; break;
        case '\\': result += '\\'; break;
        case '"': result += '"'; break;
        case "'": result += "'"; break;
        default: this.fail(`Unsupported escape sequence: \\${escaped}`);
      }
    }
    this.fail('Unterminated string literal');
  }

  private readIdentifier(): string {
    this.skipWhitespace();
    const start = this.index;
    while (isIdentifierPart(this.peek())) this.index += 1;
    if (start === this.index) this.fail('Expected an identifier');
    return this.source.slice(start, this.index);
  }

  private expectWord(word: string): void {
    const actual = this.readIdentifier();
    if (actual !== word) this.fail(`Expected ${word}`);
  }

  private expectLiteral(literal: string): void {
    this.skipWhitespace();
    if (!this.source.startsWith(literal, this.index)) this.fail(`Expected ${literal}`);
    this.index += literal.length;
  }

  private expect(char: string): void {
    this.skipWhitespace();
    if (this.peek() !== char) this.fail(`Expected '${char}'`);
    this.index += 1;
  }

  private skipWhitespace(): void {
    while (isWhitespace(this.peek())) this.index += 1;
  }

  private peek(): string | undefined {
    return this.source[this.index];
  }

  private atEnd(): boolean {
    return this.index >= this.source.length;
  }

  private fail(message: string): never {
    throw new Error(`Invalid automation DSL at character ${this.index + 1}: ${message}`);
  }
}

export function parseFacadeScript(
  source: string,
  limits: FacadeDslLimits = DEFAULT_FACADE_DSL_LIMITS,
): FacadeProgram {
  if (typeof source !== 'string') throw new Error('Automation source must be a string');
  return new DslParser(source, limits).parse();
}

export function buildFacadePlan(
  workbook: WorkbookModel,
  program: FacadeProgram,
  control?: FacadeExecutionControl,
): FacadePlan {
  const sheet = workbook.getSheet(workbook.primarySheetId);
  const operations: FacadeCellOperation[] = [];
  const affectedRanges: RangeRef[] = [];

  for (const statement of program.statements) {
    checkFacadeExecution(control);
    assertRangeWithinSheet(sheet.rowCount, sheet.columnCount, statement.range);
    affectedRanges.push({
      sheetId: sheet.id,
      startRow: statement.range.startRow,
      endRow: statement.range.endRow,
      startColumn: statement.range.startColumn,
      endColumn: statement.range.endColumn,
    });
    if (statement.kind === 'set-values') {
      for (let rowOffset = 0; rowOffset < statement.values.length; rowOffset += 1) {
        checkFacadeExecution(control);
        const row = statement.values[rowOffset] ?? [];
        for (let columnOffset = 0; columnOffset < row.length; columnOffset += 1) {
          checkFacadeExecution(control);
          const targetRow = statement.range.startRow + rowOffset;
          const targetColumn = statement.range.startColumn + columnOffset;
          if (targetRow > statement.range.endRow || targetColumn > statement.range.endColumn) {
            throw new Error('setValues data exceeds the target range');
          }
          operations.push({
            kind: 'set-cell',
            sheetId: sheet.id,
            row: targetRow,
            column: targetColumn,
            value: normalizeCellData(row[columnOffset]),
          });
        }
      }
      continue;
    }
    for (let row = statement.range.startRow; row <= statement.range.endRow; row += 1) {
      checkFacadeExecution(control);
      for (let column = statement.range.startColumn; column <= statement.range.endColumn; column += 1) {
        checkFacadeExecution(control);
        if (statement.kind === 'set-font-weight') {
          operations.push({
            kind: 'set-style',
            sheetId: sheet.id,
            row,
            column,
            style: { bold: statement.weight === 'bold' },
          });
        } else {
          operations.push({ kind: 'clear-cell', sheetId: sheet.id, row, column });
        }
      }
    }
  }

  return { statements: program.statements, operations, affectedRanges };
}

export function parseAndBuildFacadePlan(
  workbook: WorkbookModel,
  source: string,
  limits: FacadeDslLimits = DEFAULT_FACADE_DSL_LIMITS,
): FacadePlan {
  return buildFacadePlan(workbook, parseFacadeScript(source, limits));
}

export function rangeCellCount(range: A1Range): number {
  return (range.endRow - range.startRow + 1) * (range.endColumn - range.startColumn + 1);
}

export function parseA1Range(text: string): A1Range {
  const trimmed = text.trim();
  if (!trimmed) throw new Error('A1 reference cannot be empty');
  const separator = trimmed.indexOf(':');
  const first = separator < 0 ? trimmed : trimmed.slice(0, separator);
  const second = separator < 0 ? first : trimmed.slice(separator + 1);
  if (separator >= 0 && trimmed.indexOf(':', separator + 1) >= 0) throw new Error(`Invalid A1 reference: ${text}`);
  const start = parseA1Cell(first);
  const end = parseA1Cell(second);
  const startRow = Math.min(start.row, end.row);
  const endRow = Math.max(start.row, end.row);
  const startColumn = Math.min(start.column, end.column);
  const endColumn = Math.max(start.column, end.column);
  return { startRow, startColumn, endRow, endColumn, text };
}

function parseA1Cell(text: string): { row: number; column: number } {
  let index = 0;
  while (index < text.length && isAsciiLetter(text[index])) index += 1;
  if (index === 0 || index === text.length) throw new Error(`Invalid A1 reference: ${text}`);
  let column = 0;
  for (let cursor = 0; cursor < index; cursor += 1) {
    const code = text[cursor]!.toUpperCase().charCodeAt(0);
    column = column * 26 + code - 64;
  }
  if (column <= 0) throw new Error(`Invalid A1 reference: ${text}`);
  let row = 0;
  while (index < text.length && isDigit(text[index])) {
    row = row * 10 + Number(text[index]);
    index += 1;
  }
  if (index !== text.length || row <= 0) throw new Error(`Invalid A1 reference: ${text}`);
  return { row: row - 1, column: column - 1 };
}

function assertRangeWithinSheet(rowCount: number, columnCount: number, range: A1Range): void {
  if (range.startRow < 0 || range.endRow >= rowCount || range.startColumn < 0 || range.endColumn >= columnCount) {
    throw new Error(`A1 range ${range.text} is outside the worksheet bounds`);
  }
}

function validateValues(values: unknown[][]): void {
  for (const row of values) {
    for (const value of row) {
      if (value === null || typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') continue;
      if (!isCellDataLiteral(value)) throw new Error('setValues only accepts scalar values or CellData literals');
    }
  }
}

function isCellDataLiteral(value: unknown): value is Record<string, unknown> {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) return false;
  const keys = Object.keys(value);
  if (!keys.includes('value')) return false;
  return keys.every((key) => key === 'value' || key === 'formula' || key === 'style' || key === 'numberFormat');
}

function normalizeCellData(value: unknown): CellData {
  if (isCellDataLiteral(value)) {
    const cellValue = value.value;
    if (cellValue !== null && typeof cellValue !== 'string' && typeof cellValue !== 'number' && typeof cellValue !== 'boolean') {
      throw new Error('CellData.value must be a scalar');
    }
    return structuredClone(value) as unknown as CellData;
  }
  return { value: value as CellValue };
}

function isDigit(value: string | undefined): boolean {
  return value !== undefined && value >= '0' && value <= '9';
}

function isAsciiLetter(value: string | undefined): boolean {
  return value !== undefined && ((value >= 'A' && value <= 'Z') || (value >= 'a' && value <= 'z'));
}

function isIdentifierPart(value: string | undefined): boolean {
  return isAsciiLetter(value) || isDigit(value) || value === '_';
}

function isWhitespace(value: string | undefined): boolean {
  return value === ' ' || value === '\n' || value === '\r' || value === '\t';
}
