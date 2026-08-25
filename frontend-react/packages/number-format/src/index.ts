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

/** Excel 序列日期起点:1899-12-30(TC) */
const EPOCH_MS = Date.UTC(1899, 11, 30);
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
    let end = start;
    while (end < format.length && format[end]?.toLowerCase() === character.toLowerCase()) end += 1;
    return { token: { kind: 'date', ch: character.toLowerCase() }, next: end };
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

function formatDateValue(serial: number, tokens: readonly FormatToken[]): string {
  const ms = EPOCH_MS + Math.round(serial * 86400000);
  const date = new Date(ms);
  let output = '';
  const hasAmPm = tokens.some((token) => token.kind === 'ampm');

  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i]!;
    switch (token.kind) {
      case 'date': {
        const run = countRun(tokens, i, token.ch);
        switch (token.ch) {
          case 'y':
            output += run >= 3 ? String(date.getUTCFullYear()) : pad(date.getUTCFullYear() % 100, 2);
            break;
          case 'm': {
            const asMinutes = previousTokenIs(tokens, i, 'h');
            if (asMinutes) output += pad(date.getUTCMinutes(), Math.min(run, 2));
            else if (run >= 4) output += MONTH_NAMES[date.getUTCMonth()] ?? '';
            else if (run === 3) output += (MONTH_NAMES[date.getUTCMonth()] ?? '').slice(0, 3);
            else output += pad(date.getUTCMonth() + 1, Math.min(run, 2));
            break;
          }
          case 'd':
            if (run >= 4) output += WEEKDAY_NAMES[date.getUTCDay()] ?? '';
            else if (run === 3) output += (WEEKDAY_NAMES[date.getUTCDay()] ?? '').slice(0, 3);
            else output += pad(date.getUTCDate(), Math.min(run, 2));
            break;
          case 'h': {
            let hours = date.getUTCHours();
            if (hasAmPm) hours = hours % 12 === 0 ? 12 : hours % 12;
            output += pad(hours, Math.min(run, 2));
            break;
          }
          case 's':
            output += pad(date.getUTCSeconds(), Math.min(run, 2));
            break;
          default:
            break;
        }
        i += run - 1;
        break;
      }
      case 'ampm':
        output += token.style === 'A/P'
          ? (date.getUTCHours() < 12 ? 'A' : 'P')
          : (date.getUTCHours() < 12 ? 'AM' : 'PM');
        break;
      case 'literal':
        output += token.text;
        break;
      case 'space':
        output += ' ';
        break;
      default:
        break;
    }
  }
  return output;
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

function previousTokenIs(tokens: readonly FormatToken[], index: number, ch: string): boolean {
  for (let i = index - 1; i >= 0; i--) {
    const token = tokens[i]!;
    if (token.kind === 'literal') continue;
    if (token.kind === 'space') continue;
    return token.kind === 'date' ? token.ch === ch : false;
  }
  return false;
}

const MONTH_NAMES = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const WEEKDAY_NAMES = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

const sectionCache = new Map<string, FormatSection[]>();

function getSections(format: string): FormatSection[] {
  const cached = sectionCache.get(format);
  if (cached) return cached;
  const parsed = parseSections(format);
  if (sectionCache.size > 512) sectionCache.clear();
  sectionCache.set(format, parsed);
  return parsed;
}

function formatGeneralNumber(value: number): string {
  if (Number.isInteger(value)) return String(value);
  const text = value.toFixed(6).replace(/0+$/, '').replace(/\.$/, '');
  return text;
}

/** 将数值/文本按 Excel 格式代码渲染为显示字符串 */
export function formatValue(
  value: number | string | boolean | null | undefined,
  format?: string,
): string {
  if (value == null) return '';
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';

  const trimmedFormat = format?.trim();

  if (typeof value === 'string') {
    if (!trimmedFormat || trimmedFormat.toLowerCase() === 'general') return value;
    const sections = getSections(trimmedFormat);
    const textSection = sections.find((section) => section.hasTextPlaceholder)
      ?? sections[sections.length - 1];
    if (!textSection) return value;
    let output = '';
    for (const token of textSection.tokens) {
      if (token.kind === 'text-placeholder') output += value;
      else if (token.kind === 'literal' && token.text !== ',') output += token.text;
    }
    return output.length > 0 ? output : value;
  }

  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return String(value);
  if (!trimmedFormat || trimmedFormat.toLowerCase() === 'general') {
    return formatGeneralNumber(numeric);
  }

  const sections = getSections(trimmedFormat);
  let sectionIndex = 0;
  if (sections.length >= 3 && numeric === 0) sectionIndex = 2;
  else if (sections.length >= 2 && numeric < 0) sectionIndex = 1;
  const section = sections[sectionIndex] ?? sections[0]!;
  const pattern = analyzeTokens(section.tokens);

  if (pattern.isDate) {
    return formatDateValue(Math.abs(numeric), section.tokens);
  }

  const scaled = Math.abs(numeric) * pattern.percentScale;
  const decimals = Math.min(pattern.decimalDigits.length, 20);
  const rounded = scaled.toFixed(decimals);
  const [intRaw, decRaw = ''] = rounded.split('.');

  const minIntegers = pattern.integerDigits.filter((ch) => ch === '0').length;
  let intText = intRaw ?? '0';
  while (intText.length < minIntegers) intText = `0${intText}`;
  if (pattern.thousands) intText = groupThousands(intText);

  // 小数位裁剪到模式长度;'?' 占位允许尾部空缺(此处直接截断即可满足常见格式)
  const maxDecimals = pattern.decimalDigits.length;
  const decText = decRaw.slice(0, maxDecimals);

  const sign = numeric < 0 ? '-' : '';
  const decimalOutput = decText.length > 0 ? `.${decText}` : '';
  return `${pattern.prefix}${sign}${intText}${decimalOutput}${pattern.suffix}`;
}
