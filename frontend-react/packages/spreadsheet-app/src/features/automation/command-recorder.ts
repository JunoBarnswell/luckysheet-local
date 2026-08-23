import type { CommandListener } from '@react-sheets/command-runtime';

export interface RecordedStatement {
  line: number;
  code: string;
}

/** Command 事件录制器 — 输出 Facade 风格脚本，禁止录坐标 */
export class CommandRecorder {
  private recording = false;
  private statements: RecordedStatement[] = [];
  private line = 1;

  start(): void {
    this.recording = true;
    this.statements = [];
    this.line = 1;
  }

  stop(): RecordedStatement[] {
    this.recording = false;
    return [...this.statements];
  }

  isRecording(): boolean {
    return this.recording;
  }

  createListener(): CommandListener {
    return (commandId, params, _result) => {
      if (!this.recording) return;
      const code = this.toFacadeCode(commandId, params);
      if (code) {
        this.statements.push({ line: this.line++, code });
      }
    };
  }

  toScript(): string {
    return this.statements.map((s) => s.code).join('\n');
  }

  private toFacadeCode(commandId: string, params: unknown): string | undefined {
    const p = params as Record<string, unknown> | null;
    switch (commandId) {
      case 'sheet.cell.set': {
        const row = p?.row as number | undefined;
        const col = p?.column as number | undefined;
        const value = p?.value as { value?: unknown } | undefined;
        if (row == null || col == null) throw new Error('Cannot serialize sheet.cell.set: missing coordinates');
        const a1 = `${columnToLetter(col)}${row + 1}`;
        return `sheet.getRange(${quoteA1(a1)}).setValues([[${serializeLiteral(value)}]]);`;
      }
      case 'sheet.range.set': {
        const values = p?.values;
        if (!Array.isArray(values) || values.some((row) => !Array.isArray(row))) throw new Error('Cannot serialize sheet.range.set: invalid values');
        const range = p;
        const startRow = range?.startRow as number | undefined;
        const startColumn = range?.startColumn as number | undefined;
        if (startRow == null || startColumn == null) throw new Error('Cannot serialize sheet.range.set: missing start');
        const rowCount = values.length;
        const columnCount = Math.max(1, ...values.map((row) => (row as unknown[]).length));
        const a1 = `${columnToLetter(startColumn)}${startRow + 1}:${columnToLetter(startColumn + columnCount - 1)}${startRow + rowCount}`;
        return `sheet.getRange(${quoteA1(a1)}).setValues(${serializeLiteral(values)});`;
      }
      case 'sheet.style.set': {
        const style = p?.style as { bold?: boolean } | undefined;
        const range = p?.range as { startRow?: number; startColumn?: number; endRow?: number; endColumn?: number } | undefined;
        if (!range || range.startRow == null || range.startColumn == null) throw new Error('Cannot serialize sheet.style.set: missing range');
        const endRow = range.endRow ?? range.startRow;
        const endCol = range.endColumn ?? range.startColumn;
        const a1 = `${columnToLetter(range.startColumn)}${range.startRow + 1}:${columnToLetter(endCol)}${endRow + 1}`;
        if (style && Object.keys(style).every((key) => key === 'bold') && typeof style.bold === 'boolean') {
          return `sheet.getRange(${quoteA1(a1)}).setFontWeight(${quoteA1(style.bold ? 'bold' : 'normal')});`;
        }
        throw new Error(`Cannot serialize sheet.style.set: unsupported style fields`);
      }
      case 'sheet.range.clear': {
        const range = p?.range as { startRow?: number; startColumn?: number; endRow?: number; endColumn?: number } | undefined;
        if (p?.mode !== undefined) throw new Error('Cannot serialize sheet.range.clear with a mode');
        const startRow = range?.startRow as number | undefined;
        const startCol = range?.startColumn as number | undefined;
        const endRow = (range?.endRow as number | undefined) ?? startRow;
        const endCol = (range?.endColumn as number | undefined) ?? startCol;
        if (startRow == null || startCol == null || endRow == null || endCol == null) throw new Error('Cannot serialize sheet.range.clear: missing range');
        const a1 = `${columnToLetter(startCol)}${startRow + 1}:${columnToLetter(endCol)}${endRow + 1}`;
        return `sheet.getRange(${quoteA1(a1)}).clear();`;
      }
      default:
        throw new Error(`Cannot serialize recorded command: ${commandId}`);
    }
  }
}

function serializeLiteral(value: unknown): string {
  const seen = new Set<object>();
  const check = (entry: unknown): void => {
    if (entry === null || typeof entry === 'string' || typeof entry === 'boolean') return;
    if (typeof entry === 'number') {
      if (!Number.isFinite(entry)) throw new Error('Cannot serialize non-finite automation value');
      return;
    }
    if (typeof entry !== 'object' || seen.has(entry)) throw new Error('Cannot serialize recorded command value');
    seen.add(entry);
    if (Array.isArray(entry)) entry.forEach(check);
    else Object.values(entry as Record<string, unknown>).forEach(check);
  };
  check(value);
  const json = JSON.stringify(value);
  if (json === undefined) throw new Error('Cannot serialize recorded command value');
  return json;
}

function quoteA1(value: string): string {
  return `'${value.replaceAll('\\', '\\\\').replaceAll("'", "\\'")}'`;
}

function columnToLetter(col: number): string {
  let n = col + 1;
  let result = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    result = String.fromCharCode(65 + rem) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
}
