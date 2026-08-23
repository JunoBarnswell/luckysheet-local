import type { RangeRef, SelectionSnapshot } from '@react-sheets/core-model';
import { createEmptySelection } from '@react-sheets/core-model';
import type { CommandRuntime } from '@react-sheets/command-runtime';

export interface SetSelectionParams {
  unitId: string;
  sheetId: string;
  ranges: RangeRef[];
  primaryRangeIndex?: number;
  primaryCell?: { row: number; column: number };
  anchorCell?: { row: number; column: number };
}

/** M1 编辑特性 — 多选区 selection.set(选区由 Application 订阅,命令返回目标范围) */
export function registerEditingFeatures(runtime: CommandRuntime): void {
  runtime.registry.registerCommand<SetSelectionParams>({
    id: 'selection.set',
    execute: (params, context) => {
      const ranges = params.ranges.map((range) => ({
        sheetId: range.sheetId,
        startRow: Math.min(range.startRow, range.endRow),
        endRow: Math.max(range.startRow, range.endRow),
        startColumn: Math.min(range.startColumn, range.endColumn),
        endColumn: Math.max(range.startColumn, range.endColumn),
      }));
      const primary = params.primaryCell ?? (ranges[0]
        ? { row: ranges[0].startRow, column: ranges[0].startColumn }
        : { row: 0, column: 0 });
      return {
        operationId: context.operationId,
        mutationCount: 0,
        affectedRanges: ranges.length > 0 ? ranges : [{
          sheetId: params.sheetId,
          startRow: primary.row,
          endRow: primary.row,
          startColumn: primary.column,
          endColumn: primary.column,
        }],
      };
    },
  });
}

export function buildSelectionSnapshot(params: SetSelectionParams): SelectionSnapshot {
  const ranges = params.ranges;
  const primaryCell = params.primaryCell ?? (ranges[0]
    ? { row: ranges[0].startRow, column: ranges[0].startColumn }
    : { row: 0, column: 0 });
  return {
    ...createEmptySelection(params.unitId, params.sheetId),
    ranges,
    primaryRangeIndex: params.primaryRangeIndex ?? 0,
    primaryCell,
    anchorCell: params.anchorCell ?? primaryCell,
    phase: 'selected',
  };
}
