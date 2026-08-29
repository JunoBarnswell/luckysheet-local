/**
 * Excel 数字格式代码解析与应用引擎(单文件包,无外部依赖)。
 *
 * 支持:正/负/零/文本四段格式、数字占位符 0 # ?、千分位逗号、小数点、
 * 百分比、字面量(引号 / 反斜杠转义)、日期时间令牌
 * yyyy yy mmm mm m ddd dd d hhh hh h sss ss s AM/PM A/P、
 * 下划线宽度符 _x 与星号填充 *x(按单字符宽度展开)。
 */

export interface FormatSection {
  readonly tokens: FormatToken[];
  readonly hasTextPlaceholder: boolean;
}

export type NumberFormatPrecisionFailure =
  | 'missing-format'
  | 'invalid-delta'
  | 'unsupported-format'
  | 'precision-limit'
  | 'no-change';

export type NumberFormatPrecisionResult =
  | { readonly ok: true; readonly format: string; readonly decimalPlaces: number }
  | { readonly ok: false; readonly reason: NumberFormatPrecisionFailure };

export type FormatToken =
  | { kind: 'literal'; text: string }
  | { kind: 'space'; widthChar: string }
  | { kind: 'fill'; char: string }
  | { kind: 'digit'; ch: '0' | '#' | '?' }
  | { kind: 'decimal-point' }
  | { kind: 'percent' }
  | { kind: 'date'; ch: string }
  | { kind: 'ampm'; style: 'AM/PM' | 'A/P' }
  | { kind: 'text-placeholder' };

const DATE_TOKEN_CHARS = new Set(['y', 'm', 'd', 'h', 's']);

export function parseSections(format: string): FormatSection[] {
  const sections: FormatSection[] = [];
  let tokens: FormatToken[] = [];
  let index = 0;

  while (index <= format.length) {
    if (index === format.length || format[index] === ';') {
      sections.push({
        tokens,
        hasTextPlaceholder: tokens.some((token) => token.kind === 'text-placeholder'),
      });
      tokens = [];
      index += 1;
      continue;
    }
    const scanned = scanToken(format, index);
    if (scanned) {
      tokens.push(scanned.token);
      index = scanned.next;
    } else {
      tokens.push({ kind: 'literal', text: format[index] ?? '' });
      index += 1;
    }
  }
  return sections.length > 0 ? sections : [{ tokens: [], hasTextPlaceholder: false }];
}

interface PrecisionToken {
  readonly kind: 'digit' | 'decimal' | 'date' | 'text' | 'fraction' | 'exponent' | 'unsupported';
  readonly start: number;
  readonly end: number;
}

interface PrecisionSectionAnalysis {
  readonly source: string;
  readonly tokens: readonly PrecisionToken[];
  readonly digitTokens: readonly PrecisionToken[];
  readonly decimalToken?: PrecisionToken;
  readonly decimalStart?: number;
  readonly decimalEnd?: number;
  readonly decimalPlaces: number;
  readonly exponentToken?: PrecisionToken;
  readonly unsupported: boolean;
}

/**
 * Splits an Excel format code without treating semicolons inside quoted,
 * escaped, or bracketed metadata as section separators. The raw section text
 * is retained so precision changes can leave every unrelated byte untouched.
 */
function splitPrecisionSections(format: string): string[] | undefined {
  const sections: string[] = [];
  let sectionStart = 0;
  let quoted = false;
  let bracketed = false;

  for (let index = 0; index < format.length; index += 1) {
    const character = format[index] ?? '';
    if (quoted) {
      if (character === '"') quoted = false;
      continue;
    }
    if (bracketed) {
      if (character === ']') bracketed = false;
      continue;
    }
    if (character === '"') {
      quoted = true;
      continue;
    }
    if (character === '[') {
      bracketed = true;
      continue;
    }
    if (character === '\\') {
      if (index + 1 >= format.length) return undefined;
      index += 1;
      continue;
    }
    if (character === ';') {
      sections.push(format.slice(sectionStart, index));
      sectionStart = index + 1;
    }
  }

  if (quoted || bracketed) return undefined;
  sections.push(format.slice(sectionStart));
  return sections;
}

function precisionTokenize(section: string): { tokens: PrecisionToken[]; invalid: boolean } {
  const tokens: PrecisionToken[] = [];
  let invalid = false;

  for (let index = 0; index < section.length;) {
    const character = section[index] ?? '';
    if (character === '"') {
      const end = section.indexOf('"', index + 1);
      if (end < 0) return { tokens, invalid: true };
      index = end + 1;
      continue;
    }
    if (character === '\\' || character === '_' || character === '*') {
      if (index + 1 >= section.length) return { tokens, invalid: true };
      index += 2;
      continue;
    }
    if (character === '[') {
      const end = section.indexOf(']', index + 1);
      if (end < 0) return { tokens, invalid: true };
      index = end + 1;
      continue;
    }
    if (character === '@') {
      tokens.push({ kind: 'text', start: index, end: index + 1 });
      index += 1;
      continue;
    }
    if (character === '0' || character === '#' || character === '?') {
      tokens.push({ kind: 'digit', start: index, end: index + 1 });
      index += 1;
      continue;
    }
    if (character === '.') {
      tokens.push({ kind: 'decimal', start: index, end: index + 1 });
      index += 1;
      continue;
    }
    if (character === '/') {
      tokens.push({ kind: 'fraction', start: index, end: index + 1 });
      index += 1;
      continue;
    }

    const ampm = section.slice(index).match(/^AM\/PM/i) ?? section.slice(index).match(/^A\/P/i);
    if (ampm) {
      tokens.push({ kind: 'date', start: index, end: index + ampm[0].length });
      index += ampm[0].length;
      continue;
    }
    if ('ymdhs'.includes(character.toLowerCase())) {
      let end = index + 1;
      while (end < section.length && section[end]?.toLowerCase() === character.toLowerCase()) end += 1;
      tokens.push({ kind: 'date', start: index, end });
      index = end;
      continue;
    }
    if (character === 'E' || character === 'e') {
      tokens.push({ kind: 'exponent', start: index, end: index + 1 });
      index += 1;
      continue;
    }
    if (/[A-Za-z]/.test(character)) {
      tokens.push({ kind: 'unsupported', start: index, end: index + 1 });
      invalid = true;
      index += 1;
      continue;
    }

    // Currency signs, grouping commas, signs, parentheses, percent signs,
    // spaces, and other punctuation are preserved as raw literals.
    index += 1;
  }

  return { tokens, invalid };
}

function analyzePrecisionSection(source: string): PrecisionSectionAnalysis {
  const tokenized = precisionTokenize(source);
  const tokens = tokenized.tokens;
  const digitTokens = tokens.filter((token) => token.kind === 'digit');
  const decimalTokens = tokens.filter((token) => token.kind === 'decimal');
  const exponentTokens = tokens.filter((token) => token.kind === 'exponent');
  const dateToken = tokens.some((token) => token.kind === 'date');
  const textToken = tokens.some((token) => token.kind === 'text');
  const unsupportedToken = tokens.some((token) => token.kind === 'unsupported');
  const fraction = tokens.some((token) => token.kind === 'fraction');
  const decimalToken = decimalTokens.length === 1 ? decimalTokens[0] : undefined;
  const exponentToken = exponentTokens.length === 1 ? exponentTokens[0] : undefined;
  let unsupported = tokenized.invalid || decimalTokens.length > 1 || exponentTokens.length > 1 || dateToken || unsupportedToken || fraction;
  if (textToken && digitTokens.length > 0) unsupported = true;

  if (exponentToken) {
    const mantissaDigits = digitTokens.filter((token) => token.end <= exponentToken.start);
    const exponentDigits = digitTokens.filter((token) => token.start > exponentToken.end);
    const exponentText = source.slice(exponentToken.end);
    if (mantissaDigits.length === 0 || exponentDigits.length === 0 || !/^[+-]?[0#?]+$/.test(exponentText)) unsupported = true;
  }

  if (decimalToken && exponentToken && decimalToken.start > exponentToken.start) unsupported = true;
  const decimalStart = decimalToken?.start;
  const decimalEnd = decimalToken?.end;
  let decimalPlaces = 0;
  if (decimalEnd !== undefined) {
    let cursor = decimalEnd;
    while (cursor < source.length && '0#?'.includes(source[cursor] ?? '')) {
      decimalPlaces += 1;
      cursor += 1;
    }
    if (decimalPlaces === 0) unsupported = true;
  }

  return { source, tokens, digitTokens, decimalToken, decimalStart, decimalEnd, decimalPlaces, exponentToken, unsupported };
}

function transformPrecisionSection(section: PrecisionSectionAnalysis, delta: number): { source: string; changed: boolean; decimalPlaces: number } | undefined {
  if (section.unsupported || section.digitTokens.length === 0) return { source: section.source, changed: false, decimalPlaces: 0 };
  if (delta < 0 && section.decimalPlaces === 0) return { source: section.source, changed: false, decimalPlaces: 0 };
  const nextPlaces = section.decimalPlaces + delta;
  if (nextPlaces < 0 || nextPlaces > 30) return undefined;

  const insertionPoint = section.decimalEnd !== undefined
    ? section.decimalEnd + section.decimalPlaces
    : (section.exponentToken?.start ?? section.digitTokens[section.digitTokens.length - 1]!.end);
  if (delta > 0) {
    const source = section.decimalEnd === undefined
      ? `${section.source.slice(0, insertionPoint)}.${'0'.repeat(delta)}${section.source.slice(insertionPoint)}`
      : `${section.source.slice(0, insertionPoint)}${'0'.repeat(delta)}${section.source.slice(insertionPoint)}`;
    return { source, changed: true, decimalPlaces: nextPlaces };
  }
  if (section.decimalStart === undefined || section.decimalEnd === undefined) return undefined;
  const removeStart = section.decimalEnd + section.decimalPlaces - 1;
  const removeEnd = section.decimalEnd + section.decimalPlaces;
  if (nextPlaces === 0) {
    return { source: `${section.source.slice(0, section.decimalStart)}${section.source.slice(removeEnd)}`, changed: true, decimalPlaces: 0 };
  }
  return { source: `${section.source.slice(0, removeStart)}${section.source.slice(removeEnd)}`, changed: true, decimalPlaces: nextPlaces };
}

/**
 * Applies a decimal precision delta while preserving every unrelated part of
 * an Excel number-format code. Unsupported families fail closed.
 */
export function transformNumberFormatPrecision(format: string | undefined, delta: number): NumberFormatPrecisionResult {
  if (!format || format.trim().length === 0) return { ok: false, reason: 'missing-format' };
  if (!Number.isInteger(delta) || delta === 0) return { ok: false, reason: 'invalid-delta' };
  if (format.trim().toLowerCase() === 'general') return { ok: false, reason: 'unsupported-format' };

  const sections = splitPrecisionSections(format);
  if (!sections) return { ok: false, reason: 'unsupported-format' };
  const analyses = sections.map(analyzePrecisionSection);
  if (analyses.some((section) => section.unsupported)) return { ok: false, reason: 'unsupported-format' };
  if (!analyses.some((section) => section.digitTokens.length > 0)) return { ok: false, reason: 'unsupported-format' };
  if (analyses.some((section) => section.digitTokens.length > 0 && section.decimalPlaces + delta > 30)) return { ok: false, reason: 'precision-limit' };
  if (analyses.some((section) => section.digitTokens.length > 0 && delta < 0 && section.decimalPlaces > 0 && section.decimalPlaces + delta < 0)) return { ok: false, reason: 'precision-limit' };

  let changed = false;
  let decimalPlaces = 0;
  const transformed: string[] = [];
  for (const section of analyses) {
    const result = transformPrecisionSection(section, delta);
    if (!result) return { ok: false, reason: delta > 0 ? 'precision-limit' : 'unsupported-format' };
    transformed.push(result.source);
    changed ||= result.changed;
    if (section.digitTokens.length > 0) decimalPlaces = Math.max(decimalPlaces, result.decimalPlaces);
  }
  if (!changed) return { ok: false, reason: delta < 0 ? 'no-change' : 'precision-limit' };
  return { ok: true, format: transformed.join(';'), decimalPlaces };
}

function scanToken(format: string, start: number): { token: FormatToken; next: number } | undefined {
  const character = format[start] ?? '';
  const next = start + 1;

  if (character === '"') {
    const end = format.indexOf('"', next);
    if (end < 0) return { token: { kind: 'literal', text: format.slice(next) }, next: format.length };
    return { token: { kind: 'literal', text: format.slice(next, end) }, next: end + 1 };
  }
  if (character === '\\') {
    return { token: { kind: 'literal', text: format[next] ?? '' }, next: next + 1 };
  }
  if (character === '_') return { token: { kind: 'space', widthChar: format[next] ?? ' ' }, next: next + 1 };
  if (character === '*') return { token: { kind: 'fill', char: format[next] ?? ' ' }, next: next + 1 };
  if (character === '@') return { token: { kind: 'text-placeholder' }, next };
  if (character === '.') return { token: { kind: 'decimal-point' }, next };
  if (character === '%') return { token: { kind: 'percent' }, next };
  if (character === '0' || character === '#' || character === '?') {
    return { token: { kind: 'digit', ch: character }, next };
  }
  if (/am\/pm/i.test(format.slice(start, start + 5))) {
    return { token: { kind: 'ampm', style: 'AM/PM' }, next: start + 5 };
  }
  if (/a\/p/i.test(format.slice(start, start + 3))) {
    return { token: { kind: 'ampm', style: 'A/P' }, next: start + 3 };
  }
  if (DATE_TOKEN_CHARS.has(character.toLowerCase())) {
    // Keep one token per date-format character.  A run such as `yyyy` is
    // interpreted by countRun(), and consuming the complete run here would
    // make every year/month/day token look like a one-character format.
    return { token: { kind: 'date', ch: character.toLowerCase() }, next };
  }
  if (character === '[') {
    const end = format.indexOf(']', start);
    if (end > start) return { token: { kind: 'literal', text: '' }, next: end + 1 };
  }
  return undefined;
}

interface NumericPattern {
  integerDigits: Array<'0' | '#' | '?'>;
  decimalDigits: Array<'0' | '#' | '?'>;
  prefix: string;
  suffix: string;
  percentScale: number;
  isDate: boolean;
  hasText: boolean;
}

function analyzeTokens(tokens: readonly FormatToken[]): NumericPattern & { thousands: boolean } {
  const pattern = {
    integerDigits: [] as Array<'0' | '#' | '?'>,
    decimalDigits: [] as Array<'0' | '#' | '?'>,
    prefix: '',
    suffix: '',
    percentScale: 1,
    isDate: false,
    hasText: false,
    thousands: false,
  };

  let seenDecimal = false;
  let seenDigit = false;

  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i]!;
    switch (token.kind) {
      case 'digit':
        seenDigit = true;
        if (seenDecimal) pattern.decimalDigits.push(token.ch);
        else pattern.integerDigits.push(token.ch);
        break;
      case 'decimal-point':
        seenDecimal = true;
        break;
      case 'percent': {
        pattern.percentScale *= 100;
        pattern.suffix += '%';
        break;
      }
      case 'date':
        pattern.isDate = true;
        break;
      case 'ampm':
        pattern.isDate = true;
        break;
      case 'text-placeholder':
        pattern.hasText = true;
        break;
      case 'literal': {
        // ',' 一律视为千分位标记(thousands 检测在下方),不进入前后缀
        if (token.text === ',') break;
        if (!seenDigit && !seenDecimal && pattern.decimalDigits.length === 0) pattern.prefix += token.text;
        else pattern.suffix += token.text;
        break;
      }
      case 'space':
        if (seenDigit || seenDecimal) pattern.suffix += ' ';
        else pattern.prefix += ' ';
        break;
      case 'fill':
        break;
      default:
        break;
    }
    void i;
  }

  // 千分位:',' 字面量与其右侧(同在整数部分内)的数字占位相邻即生效
  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i]!;
    if (token.kind === 'literal' && token.text === ',') {
      for (let j = i + 1; j < tokens.length; j++) {
        const nextToken = tokens[j]!;
        if (nextToken.kind === 'decimal-point') break;
        if (nextToken.kind === 'digit') {
          pattern.thousands = true;
          break;
        }
      }
      if (pattern.thousands) break;
    }
  }

  return pattern;
}

function groupThousands(digits: string): string {
  let output = '';
  for (let i = 0; i < digits.length; i++) {
    const remaining = digits.length - i;
    output += digits[i];
    if (remaining > 1 && (remaining - 1) % 3 === 0) output += ',';
  }
  return output;
}

function pad(text: number, length: number): string {
  return String(text).padStart(length, '0');
}

function countRun(tokens: readonly FormatToken[], start: number, ch: string): number {
  let run = 0;
  for (let i = start; i < tokens.length; i++) {
    const token = tokens[i]!;
    if (token.kind === 'date' && token.ch === ch) run += 1;
    else break;
  }
  return run;
}


const MONTH_NAMES = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const WEEKDAY_NAMES = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

/** Numeric/date display context shared by cell, formula, filter, chart and print projections. */
export interface NumberFormatContext {
  /** Workbook calendar. Omitting it uses the canonical Excel 1900 system. */
  readonly dateSystem?: '1900' | '1904';
  /** BCP-47 locale used for separators and month/day names. */
  readonly locale?: string;
  /** A serial date is always rendered in UTC unless a host explicitly selects local time. */
  readonly timeZone?: 'UTC' | 'local';
}

export type NumberFormatConditionOperator = '<' | '<=' | '=' | '<>' | '>=' | '>';

export interface NumberFormatCondition {
  readonly operator: NumberFormatConditionOperator;
  readonly value: number;
}

export interface NumberFormatSectionAst {
  readonly source: string;
  readonly body: string;
  readonly tokens: readonly FormatToken[];
  readonly condition?: NumberFormatCondition;
  readonly color?: string;
  readonly locale?: string;
  readonly isText: boolean;
  readonly isDate: boolean;
  readonly elapsedUnit?: 'h' | 'm' | 's';
  readonly scale: number;
  readonly percentScale: number;
}

export interface NumberFormatAst {
  readonly source: string;
  readonly sections: readonly NumberFormatSectionAst[];
}

export interface FormattedNumber {
  readonly text: string;
  readonly color?: string;
  readonly sectionIndex: number;
}

const NUMBER_FORMAT_CACHE = new Map<string, NumberFormatAst>();

/** Parse a complete Excel format into an immutable, cache-bounded AST. */
export function parseNumberFormat(format: string): NumberFormatAst {
  const source = format.trim();
  if (!source) return { source: '', sections: [{ source: '', body: '', tokens: [], isText: false, isDate: false, scale: 1, percentScale: 1 }] };
  const cached = NUMBER_FORMAT_CACHE.get(source);
  if (cached) return cached;
  const sections = splitFormatSections(source).map(parseNumberFormatSection);
  const ast = { source, sections } satisfies NumberFormatAst;
  if (NUMBER_FORMAT_CACHE.size >= 1024) NUMBER_FORMAT_CACHE.clear();
  NUMBER_FORMAT_CACHE.set(source, ast);
  return ast;
}

function splitFormatSections(source: string): string[] {
  const sections: string[] = [];
  let start = 0;
  let quoted = false;
  let escaped = false;
  let bracketed = false;
  for (let index = 0; index < source.length; index += 1) {
    const character = source[index]!;
    if (escaped) { escaped = false; continue; }
    if (character === '\\') { escaped = true; continue; }
    if (quoted) { if (character === '"') quoted = false; continue; }
    if (character === '"') { quoted = true; continue; }
    if (bracketed) { if (character === ']') bracketed = false; continue; }
    if (character === '[') { bracketed = true; continue; }
    if (character === ';') { sections.push(source.slice(start, index)); start = index + 1; }
  }
  if (quoted || escaped || bracketed) throw new Error('NUMBER_FORMAT_INVALID: unterminated literal, escape, or directive');
  sections.push(source.slice(start));
  return sections;
}

function parseNumberFormatSection(source: string): NumberFormatSectionAst {
  let cursor = 0;
  let condition: NumberFormatCondition | undefined;
  let color: string | undefined;
  let locale: string | undefined;
  let elapsedUnit: 'h' | 'm' | 's' | undefined;
  while (source[cursor] === '[') {
    const end = source.indexOf(']', cursor + 1);
    if (end < 0) throw new Error('NUMBER_FORMAT_INVALID: unterminated bracket directive');
    const directive = source.slice(cursor + 1, end).trim();
    const conditionMatch = directive.match(/^(<=|>=|<>|=|<|>)([-+]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][-+]?\d+)?)$/);
    if (conditionMatch) {
      if (condition) throw new Error('NUMBER_FORMAT_INVALID: multiple conditions in one section');
      condition = { operator: conditionMatch[1] as NumberFormatConditionOperator, value: Number(conditionMatch[2]) };
    } else if (/^\$-[0-9a-z]+$/i.test(directive)) {
      locale = directive.slice(2);
    } else if (/^[hms]$/i.test(directive)) {
      elapsedUnit = directive.toLowerCase() as 'h' | 'm' | 's';
    } else if (/^[A-Za-z]+$/.test(directive)) {
      color = directive;
    } else {
      throw new Error(`NUMBER_FORMAT_INVALID: unsupported directive [${directive}]`);
    }
    cursor = end + 1;
  }
  const body = source.slice(cursor);
  const parsed = parseSections(body)[0] ?? { tokens: [], hasTextPlaceholder: false };
  const dateTokens = parsed.tokens.filter((token) => token.kind === 'date' || token.kind === 'ampm');
  const percentScale = parsed.tokens.filter((token) => token.kind === 'percent').reduce((scale) => scale * 100, 1);
  const scale = trailingScale(body, parsed.tokens);
  return {
    source,
    body,
    tokens: parsed.tokens,
    ...(condition ? { condition } : {}),
    ...(color ? { color } : {}),
    ...(locale ? { locale } : {}),
    isText: parsed.hasTextPlaceholder,
    isDate: dateTokens.length > 0 || elapsedUnit !== undefined,
    ...(elapsedUnit ? { elapsedUnit } : {}),
    scale,
    percentScale,
  };
}

function trailingScale(body: string, tokens: readonly FormatToken[]): number {
  let scale = 1;
  for (let index = tokens.length - 1; index >= 0; index -= 1) {
    const token = tokens[index]!;
    if (token.kind === 'literal' && token.text === ',') scale *= 1000;
    else if (token.kind === 'digit' || token.kind === 'decimal-point') break;
    else if (token.kind !== 'literal' || token.text.trim() !== '') break;
  }
  // A comma between digit placeholders is grouping, not scaling.
  return scale;
}

function conditionMatches(condition: NumberFormatCondition, value: number): boolean {
  switch (condition.operator) {
    case '<': return value < condition.value;
    case '<=': return value <= condition.value;
    case '=': return value === condition.value;
    case '<>': return value !== condition.value;
    case '>=': return value >= condition.value;
    case '>': return value > condition.value;
  }
}

function selectNumberFormatSection(ast: NumberFormatAst, value: number | string): { section: NumberFormatSectionAst; index: number } {
  const numeric = typeof value === 'number' && Number.isFinite(value) ? value : undefined;
  if (numeric !== undefined) {
    const conditional = ast.sections.findIndex((section) => section.condition && conditionMatches(section.condition, numeric));
    if (conditional >= 0) return { section: ast.sections[conditional]!, index: conditional };
    const unconditional = ast.sections.findIndex((section) => !section.condition);
    const sections = ast.sections.filter((section) => !section.condition);
    const position = sections.length === 1 ? 0 : sections.length === 2 ? (numeric < 0 ? 1 : 0) : numeric > 0 ? 0 : numeric < 0 ? 1 : 2;
    const section = sections[position] ?? sections[sections.length - 1] ?? ast.sections[unconditional >= 0 ? unconditional : 0]!;
    return { section, index: ast.sections.indexOf(section) };
  }
  const index = ast.sections.findIndex((section) => section.isText);
  const fallback = index >= 0 ? index : Math.min(3, ast.sections.length - 1);
  return { section: ast.sections[fallback]!, index: fallback };
}

function localeSeparators(locale?: string): { decimal: string; group: string } {
  if (!locale) return { decimal: '.', group: ',' };
  try {
    const parts = new Intl.NumberFormat(normalizeLocale(locale)).formatToParts(1000.5);
    return {
      decimal: parts.find((part) => part.type === 'decimal')?.value ?? '.',
      group: parts.find((part) => part.type === 'group')?.value ?? ',',
    };
  } catch {
    // Invalid host locales are rejected by Intl; retain the canonical Excel
    // separators rather than emitting a partially localized value.
    return { decimal: '.', group: ',' };
  }
}

function normalizeLocale(locale: string): string {
  const aliases: Record<string, string> = {
    '409': 'en-US', '804': 'zh-CN', '404': 'zh-TW', '407': 'de-DE', '40c': 'fr-FR', '410': 'it-IT', '411': 'ja-JP', '412': 'ko-KR', '816': 'pt-PT', '416': 'pt-BR',
  };
  return aliases[locale.toLowerCase()] ?? locale;
}

function formatGeneralNumber(value: number): string {
  if (!Number.isFinite(value)) return String(value);
  if (Number.isInteger(value)) return String(value);
  // toPrecision avoids the previous fixed-six-decimal truncation while still
  // respecting Excel's 15 significant-digit numeric contract.
  return Number(value.toPrecision(15)).toString();
}

function serialToDate(serial: number, dateSystem: '1900' | '1904'): Date | '1900-02-29' {
  if (dateSystem === '1900' && serial === 60) return '1900-02-29';
  const epoch = dateSystem === '1904' ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 31);
  const offset = dateSystem === '1900' && serial > 60 ? serial - 1 : serial;
  return new Date(epoch + Math.round(offset * 86400000));
}

function localizedMonthName(date: Date, locale: string | undefined, short: boolean): string {
  try { return new Intl.DateTimeFormat(normalizeLocale(locale || 'en-US'), { month: short ? 'short' : 'long', timeZone: 'UTC' }).format(date); }
  catch { return (short ? MONTH_NAMES[date.getUTCMonth()]?.slice(0, 3) : MONTH_NAMES[date.getUTCMonth()]) ?? ''; }
}

function localizedWeekdayName(date: Date, locale: string | undefined, short: boolean): string {
  try { return new Intl.DateTimeFormat(normalizeLocale(locale || 'en-US'), { weekday: short ? 'short' : 'long', timeZone: 'UTC' }).format(date); }
  catch { return (short ? WEEKDAY_NAMES[date.getUTCDay()]?.slice(0, 3) : WEEKDAY_NAMES[date.getUTCDay()]) ?? ''; }
}

function formatDateWithContext(serial: number, section: NumberFormatSectionAst, context: NumberFormatContext): string {
  if (section.elapsedUnit) {
    const total = Math.max(0, Math.round(Math.abs(serial) * 86400));
    const elapsedValue = section.elapsedUnit === 'h' ? Math.floor(total / 3600) : section.elapsedUnit === 'm' ? Math.floor(total / 60) : total;
    let output = String(elapsedValue);
    for (let index = 0; index < section.tokens.length; index += 1) {
      const token = section.tokens[index]!;
      if (token.kind === 'literal') output += token.text;
      else if (token.kind === 'space') output += ' ';
      else if (token.kind === 'date') {
        const run = countRun(section.tokens, index, token.ch);
        const part = token.ch === 'h' ? Math.floor(total / 3600) % 24 : token.ch === 'm' ? Math.floor(total / 60) % 60 : token.ch === 's' ? total % 60 : 0;
        output += pad(part, Math.min(run, 2));
        index += run - 1;
      }
    }
    return output;
  }
  const date = serialToDate(Math.abs(serial), context.dateSystem ?? '1900');
  let output = '';
  if (date === '1900-02-29') {
    for (let index = 0; index < section.tokens.length; index += 1) {
      const token = section.tokens[index]!;
      if (token.kind === 'date') {
        const run = countRun(section.tokens, index, token.ch);
        output += token.ch === 'd' ? pad(29, Math.min(run, 2)) : token.ch === 'm' ? pad(2, Math.min(run, 2)) : token.ch === 'y' ? '1900' : '';
        index += run - 1;
      } else if (token.kind === 'literal') output += token.text;
    }
    return output;
  }
  const hasAmPm = section.tokens.some((token) => token.kind === 'ampm');
  for (let index = 0; index < section.tokens.length; index += 1) {
    const token = section.tokens[index]!;
    if (token.kind === 'date') {
      const run = countRun(section.tokens, index, token.ch);
      switch (token.ch) {
        case 'y': output += run >= 3 ? String(date.getUTCFullYear()) : pad(date.getUTCFullYear() % 100, 2); break;
        case 'm': {
          let previous = index - 1;
          while (previous >= 0 && (section.tokens[previous]?.kind === 'literal' || section.tokens[previous]?.kind === 'space')) previous -= 1;
          const previousToken = previous >= 0 ? section.tokens[previous] : undefined;
          const minutes = previousToken?.kind === 'date' && previousToken.ch === 'h';
          output += minutes ? pad(date.getUTCMinutes(), Math.min(run, 2)) : run >= 4 ? localizedMonthName(date, context.locale || section.locale, false) : run === 3 ? localizedMonthName(date, context.locale || section.locale, true) : pad(date.getUTCMonth() + 1, Math.min(run, 2));
          break;
        }
        case 'd': output += run >= 4 ? localizedWeekdayName(date, context.locale || section.locale, false) : run === 3 ? localizedWeekdayName(date, context.locale || section.locale, true) : pad(date.getUTCDate(), Math.min(run, 2)); break;
        case 'h': { let hours = date.getUTCHours(); if (hasAmPm) hours = hours % 12 || 12; output += pad(hours, Math.min(run, 2)); break; }
        case 's': output += pad(date.getUTCSeconds(), Math.min(run, 2)); break;
      }
      index += run - 1;
    } else if (token.kind === 'ampm') output += token.style === 'A/P' ? (date.getUTCHours() < 12 ? 'A' : 'P') : (date.getUTCHours() < 12 ? 'AM' : 'PM');
    else if (token.kind === 'literal') output += token.text;
    else if (token.kind === 'space') output += ' ';
  }
  return output;
}

function formatNumberSection(value: number, section: NumberFormatSectionAst, context: NumberFormatContext): string {
  if (section.isDate) return formatDateWithContext(value, section, context);
  const separators = localeSeparators(context.locale || section.locale);
  const scaled = Math.abs(value) * section.percentScale / section.scale;
  const pattern = analyzeTokens(section.tokens);
  if (pattern.integerDigits.length === 0 && pattern.decimalDigits.length === 0) {
    return section.tokens.map((token) => token.kind === 'literal' ? token.text : token.kind === 'space' ? ' ' : '').join('');
  }
  const decimals = Math.min(pattern.decimalDigits.length, 20);
  const rounded = scaled.toFixed(decimals);
  const [intRaw, decRaw = ''] = rounded.split('.');
  const minIntegers = pattern.integerDigits.filter((digit) => digit === '0').length;
  let integer = intRaw || '0';
  if (scaled === 0 && minIntegers === 0) integer = '';
  while (integer.length < minIntegers) integer = `0${integer}`;
  if (pattern.thousands) integer = groupThousands(integer);
  let decimal = decRaw.slice(0, pattern.decimalDigits.length);
  while (decimal.length > 0) {
    const position = decimal.length - 1;
    const placeholder = pattern.decimalDigits[position];
    if (placeholder === '#' && decimal[position] === '0') decimal = decimal.slice(0, -1);
    else break;
  }
  const decimalOutput = decimal.length > 0 ? `${separators.decimal}${decimal}` : '';
  const sign = value < 0 && !pattern.prefix.includes('-') && !pattern.prefix.includes('(') ? '-' : '';
  const prefix = pattern.prefix.replace(/,/g, separators.group);
  const suffix = pattern.suffix.replace(/,/g, separators.group);
  return `${prefix}${sign}${integer.replace(/,/g, separators.group)}${decimalOutput}${suffix}`;
}

/** Format a value and expose presentation metadata such as a conditional color. */
export function formatNumberValue(value: number | string | boolean | null | undefined, format?: string | NumberFormatAst, context: NumberFormatContext = {}): FormattedNumber {
  if (value == null) return { text: '', sectionIndex: 0 };
  if (typeof value === 'boolean') return { text: value ? 'TRUE' : 'FALSE', sectionIndex: 0 };
  const ast = typeof format === 'string' ? parseNumberFormat(format) : format;
  if (!ast || ast.source.toLowerCase() === 'general' || ast.source === '') {
    return { text: typeof value === 'number' ? formatGeneralNumber(value) : String(value), sectionIndex: 0 };
  }
  const selected = selectNumberFormatSection(ast, value);
  if (typeof value === 'string') {
    // A numeric-only section has no defined text rendering contract.  Excel
    // leaves text values unchanged instead of emitting the section's numeric
    // literals (for example, `#,##0` must not turn "abc" into ",").
    if (!selected.section.isText) return { text: value, ...(selected.section.color ? { color: selected.section.color } : {}), sectionIndex: selected.index };
    const text = selected.section.tokens.map((token) => token.kind === 'text-placeholder' ? value : token.kind === 'literal' ? token.text : token.kind === 'space' ? ' ' : '').join('');
    return { text: text || value, ...(selected.section.color ? { color: selected.section.color } : {}), sectionIndex: selected.index };
  }
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return { text: String(value), sectionIndex: selected.index };
  return { text: formatNumberSection(numeric, selected.section, context), ...(selected.section.color ? { color: selected.section.color } : {}), sectionIndex: selected.index };
}

/** Canonical display string used by all non-rendering consumers. */
export function formatValue(
  value: number | string | boolean | null | undefined,
  format?: string,
  context?: NumberFormatContext,
): string {
  return formatNumberValue(value, format, context).text;
}
