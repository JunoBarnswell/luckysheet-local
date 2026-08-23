import type { WorkbookModel } from '@react-sheets/core-model';
import type { CommandRuntime } from '@react-sheets/command-runtime';
import { ScriptSandbox } from './sandbox';

export interface FacadeRange {
  setValues(values: unknown[][]): void;
  setFontWeight(weight: 'normal' | 'bold'): void;
  clear(): void;
}

export interface SpreadsheetFacade {
  getActiveSheet(): { getName(): string; getRange(a1: string): FacadeRange };
  getWorkbook(): { getName(): string };
}

/** Facade 脚本运行时 — 脚本只允许调 Facade */
export class FacadeScriptRuntime {
  constructor(
    private readonly workbook: WorkbookModel,
    private readonly runtime: CommandRuntime,
  ) {}

  createFacade(): SpreadsheetFacade {
    const workbook = this.workbook;
    const runtime = this.runtime;
    const activeSheetId = () => workbook.activeSheetId;

    const parseA1 = (a1: string): { row: number; column: number; endRow?: number; endColumn?: number } => {
      const match = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(a1.toUpperCase());
      if (!match) throw new Error(`Invalid A1 reference: ${a1}`);
      const colToNum = (col: string) => col.split('').reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1;
      const startCol = colToNum(match[1]!);
      const startRow = parseInt(match[2]!, 10) - 1;
      if (match[3] && match[4]) {
        return { row: startRow, column: startCol, endRow: parseInt(match[4], 10) - 1, endColumn: colToNum(match[3]) };
      }
      return { row: startRow, column: startCol };
    };

    const createRange = (a1: string): FacadeRange => ({
      setValues(values: unknown[][]) {
        const ref = parseA1(a1);
        for (let r = 0; r < values.length; r++) {
          const row = values[r] ?? [];
          for (let c = 0; c < row.length; c++) {
            const value = row[c];
            runtime.execute('sheet.cell.set', {
              row: ref.row + r,
              column: ref.column + c,
              value,
            });
          }
        }
      },
      setFontWeight(weight: 'normal' | 'bold') {
        const ref = parseA1(a1);
        runtime.execute('sheet.style.set', {
          row: ref.row,
          column: ref.column,
          style: { bold: weight === 'bold' },
        });
      },
      clear() {
        const ref = parseA1(a1);
        runtime.execute('sheet.range.clear', {
          sheetId: activeSheetId(),
          startRow: ref.row,
          endRow: ref.endRow ?? ref.row,
          startColumn: ref.column,
          endColumn: ref.endColumn ?? ref.column,
        });
      },
    });

    return {
      getActiveSheet() {
        const sheet = workbook.getSheet(activeSheetId());
        return {
          getName: () => sheet.name,
          getRange: createRange,
        };
      },
      getWorkbook() {
        return { getName: () => workbook.name };
      },
    };
  }

  /** 执行 Facade 脚本字符串 — 禁止 eval，使用 Function 沙箱 */
  runScript(source: string, sandbox: ScriptSandbox): ScriptRunResult {
    const facade = this.createFacade();
    const started = Date.now();
    try {
      sandbox.assertAllowed(source);
      const fn = new Function('sheet', 'workbook', `"use strict";\n${source}`);
      fn(facade.getActiveSheet(), facade.getWorkbook());
      return { ok: true, durationMs: Date.now() - started };
    } catch (error) {
      return { ok: false, durationMs: Date.now() - started, error: error instanceof Error ? error.message : String(error) };
    }
  }
}

export interface ScriptRunResult {
  ok: boolean;
  durationMs: number;
  error?: string;
}
