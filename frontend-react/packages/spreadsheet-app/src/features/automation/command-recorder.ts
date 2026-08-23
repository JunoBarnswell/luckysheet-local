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
        if (row == null || col == null) return undefined;
        const a1 = `${columnToLetter(col)}${row + 1}`;
        return `sheet.getRange('${a1}').setValues([[${JSON.stringify(value?.value ?? value)}]]);`;
      }
      case 'sheet.style.set': {
        const style = p?.style as { bold?: boolean } | undefined;
        const range = p?.range as { startRow?: number; startColumn?: number; endRow?: number; endColumn?: number } | undefined;
        if (!range || range.startRow == null || range.startColumn == null) return undefined;
        const endRow = range.endRow ?? range.startRow;
        const endCol = range.endColumn ?? range.startColumn;
        const a1 = `${columnToLetter(range.startColumn)}${range.startRow + 1}:${columnToLetter(endCol)}${endRow + 1}`;
        if (style?.bold) return `sheet.getRange('${a1}').setFontWeight('bold');`;
        return undefined;
      }
      case 'sheet.range.clear': {
        const startRow = p?.startRow as number | undefined;
        const startCol = p?.startColumn as number | undefined;
        const endRow = (p?.endRow as number | undefined) ?? startRow;
        const endCol = (p?.endColumn as number | undefined) ?? startCol;
        if (startRow == null || startCol == null || endRow == null || endCol == null) return undefined;
        const a1 = `${columnToLetter(startCol)}${startRow + 1}:${columnToLetter(endCol)}${endRow + 1}`;
        return `sheet.getRange('${a1}').clear();`;
      }
      default:
        return undefined;
    }
  }
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
