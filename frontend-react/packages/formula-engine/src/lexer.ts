import { FormulaLexError } from './errors';
import type { SourceSpan } from './ast';

export type TokenKind =
  | 'number'
  | 'string'
  | 'identifier'
  | 'error-reference'
  | 'plus'
  | 'minus'
  | 'star'
  | 'slash'
  | 'caret'
  | 'ampersand'
  | 'percent'
  | 'equal'
  | 'not-equal'
  | 'less-than'
  | 'less-than-equal'
  | 'greater-than'
  | 'greater-than-equal'
  | 'left-paren'
  | 'right-paren'
  | 'comma'
  | 'colon'
  | 'bang'
  | 'left-bracket'
  | 'right-bracket'
  | 'at-sign'
  | 'spill-operator'
  | 'table-specifier'
  | 'eof';

export interface Token {
  readonly kind: TokenKind;
  readonly lexeme: string;
  readonly span: SourceSpan;
  readonly value?: string;
}

export function lexFormula(source: string): readonly Token[] {
  const tokens: Token[] = [];
  let index = skipWhitespace(source, 0);
  if (source[index] === '=') index += 1;

  while (index < source.length) {
    const character = source[index] ?? '';
    if (/\s/.test(character)) {
      index = skipWhitespace(source, index);
      continue;
    }

    if (isNumberStart(source, index)) {
      const result = scanNumber(source, index);
      tokens.push(result.token);
      index = result.nextIndex;
      continue;
    }

    if (character === '"' || character === "'") {
      const result = scanString(source, index, character);
      tokens.push(result.token);
      index = result.nextIndex;
      continue;
    }

    if (character === '<') {
      const nextChar = source[index + 1];
      if (nextChar === '>') {
        tokens.push({ kind: 'not-equal', lexeme: '<>', span: { start: index, end: index + 2 } });
        index += 2;
        continue;
      }
      if (nextChar === '=') {
        tokens.push({ kind: 'less-than-equal', lexeme: '<=', span: { start: index, end: index + 2 } });
        index += 2;
        continue;
      }
      tokens.push({ kind: 'less-than', lexeme: '<', span: { start: index, end: index + 1 } });
      index += 1;
      continue;
    }

    if (character === '>') {
      if (source[index + 1] === '=') {
        tokens.push({ kind: 'greater-than-equal', lexeme: '>=', span: { start: index, end: index + 2 } });
        index += 2;
        continue;
      }
      tokens.push({ kind: 'greater-than', lexeme: '>', span: { start: index, end: index + 1 } });
      index += 1;
      continue;
    }

    if (character === '=') {
      tokens.push({ kind: 'equal', lexeme: '=', span: { start: index, end: index + 1 } });
      index += 1;
      continue;
    }

    if (character === '#') {
      const rest = source.slice(index);
      if (/^#REF!(?![A-Za-z0-9_])/i.test(rest)) {
        tokens.push({ kind: 'error-reference', lexeme: '#REF!', span: { start: index, end: index + 5 } });
        index += 5;
        continue;
      }
      const match = /^#(ALL|HEADERS|DATA|TOTALS)(?![A-Za-z0-9_])/i.exec(rest);
      if (match) {
        const lexeme = source.slice(index, index + match[0].length);
        tokens.push({
          kind: 'table-specifier',
          lexeme,
          value: match[1]!.toLowerCase(),
          span: { start: index, end: index + match[0].length },
        });
        index += match[0].length;
        continue;
      }
      tokens.push({ kind: 'spill-operator', lexeme: '#', span: { start: index, end: index + 1 } });
      index += 1;
      continue;
    }

    if (isWordStart(character)) {
      const start = index;
      index += 1;
      while (index < source.length && isWordCharacter(source[index] ?? '')) index += 1;
      tokens.push({ kind: 'identifier', lexeme: source.slice(start, index), span: { start, end: index } });
      continue;
    }

    const kind = punctuationKind(character);
    if (kind) {
      const start = index;
      index += 1;
      tokens.push({ kind, lexeme: character, span: { start, end: index } });
      continue;
    }

    throw new FormulaLexError(`Unexpected character: ${character}`, index);
  }

  tokens.push({ kind: 'eof', lexeme: '', span: { start: source.length, end: source.length } });
  return tokens;
}

function scanNumber(source: string, start: number): { token: Token; nextIndex: number } {
  let index = start;
  if (source[index] === '.') {
    index += 1;
    while (isDigit(source[index] ?? '')) index += 1;
  } else {
    while (isDigit(source[index] ?? '')) index += 1;
    if (source[index] === '.') {
      index += 1;
      while (isDigit(source[index] ?? '')) index += 1;
    }
  }

  const exponentMarker = source[index];
  if (exponentMarker === 'e' || exponentMarker === 'E') {
    index += 1;
    if (source[index] === '+' || source[index] === '-') index += 1;
    const exponentStart = index;
    while (isDigit(source[index] ?? '')) index += 1;
    if (index === exponentStart) throw new FormulaLexError('Exponent requires digits', exponentStart);
  }

  const lexeme = source.slice(start, index);
  const value = Number(lexeme);
  if (!Number.isFinite(value)) throw new FormulaLexError(`Invalid number: ${lexeme}`, start);
  return { token: { kind: 'number', lexeme, span: { start, end: index } }, nextIndex: index };
}

function scanString(source: string, start: number, quote: string): { token: Token; nextIndex: number } {
  let index = start + 1;
  let value = '';
  while (index < source.length) {
    const character = source[index] ?? '';
    if (character === quote) {
      if (source[index + 1] === quote) {
        value += quote;
        index += 2;
        continue;
      }
      index += 1;
      return {
        token: { kind: 'string', lexeme: source.slice(start, index), value, span: { start, end: index } },
        nextIndex: index,
      };
    }
    if (character === '\\') {
      const escaped = source[index + 1];
      if (escaped === undefined) throw new FormulaLexError('Unterminated string literal', start);
      value += decodeEscape(escaped);
      index += 2;
      continue;
    }
    value += character;
    index += 1;
  }
  throw new FormulaLexError('Unterminated string literal', start);
}

function decodeEscape(character: string): string {
  switch (character) {
    case 'n':
      return '\n';
    case 'r':
      return '\r';
    case 't':
      return '\t';
    case '\\':
      return '\\';
    case '"':
      return '"';
    case "'":
      return "'";
    default:
      return character;
  }
}

function punctuationKind(character: string): TokenKind | undefined {
  switch (character) {
    case '+':
      return 'plus';
    case '-':
      return 'minus';
    case '*':
      return 'star';
    case '/':
      return 'slash';
    case '^':
      return 'caret';
    case '&':
      return 'ampersand';
    case '%':
      return 'percent';
    case '(':
      return 'left-paren';
    case ')':
      return 'right-paren';
    case ',':
      return 'comma';
    case ':':
      return 'colon';
    case '!':
      return 'bang';
    case '[':
      return 'left-bracket';
    case ']':
      return 'right-bracket';
    case '@':
      return 'at-sign';
    default:
      return undefined;
  }
}

function isNumberStart(source: string, index: number): boolean {
  const character = source[index] ?? '';
  return isDigit(character) || (character === '.' && isDigit(source[index + 1] ?? ''));
}

function isDigit(character: string): boolean {
  return character >= '0' && character <= '9';
}

function isWordStart(character: string): boolean {
  return /^[\p{L}_$]$/u.test(character);
}

function isWordCharacter(character: string): boolean {
  return /^[\p{L}\p{N}_$.]$/u.test(character);
}

function skipWhitespace(source: string, start: number): number {
  let index = start;
  while (index < source.length && /\s/.test(source[index] ?? '')) index += 1;
  return index;
}
