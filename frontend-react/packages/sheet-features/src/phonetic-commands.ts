import type { CommandRuntime } from '@react-sheets/command-runtime';
import { isCellPhoneticMetadata, type CellData, type CellPhoneticMetadata, type RangeRef } from '@react-sheets/core-model';
import { createCellSetMutationParams } from './cell-write-authority';

export interface PhoneticApplyParams {
  sheetId: string;
  range: RangeRef;
  metadata: CellPhoneticMetadata;
}

function validateMetadata(metadata: CellPhoneticMetadata, textLength: number): void {
  if (!isCellPhoneticMetadata(metadata)) throw new Error('INVALID_PHONETIC_METADATA: Phonetic runs or properties are invalid');
  if (metadata.runs.some((run) => run.end > textLength)) throw new Error('INVALID_PHONETIC_METADATA: Phonetic run exceeds the source text length');
}

export function registerPhoneticCommands(runtime: CommandRuntime): void {
  runtime.registry.registerCommand<PhoneticApplyParams>({
    id: 'sheet.phonetic.set',
    execute: (params, context) => {
      const sheet = context.workbook.getSheet(params.sheetId);
      const range = { ...structuredClone(params.range), sheetId: params.sheetId };
      if (range.startRow < 0 || range.startColumn < 0 || range.endRow < range.startRow || range.endColumn < range.startColumn || range.endRow >= sheet.rowCount || range.endColumn >= sheet.columnCount) throw new Error('INVALID_PHONETIC_RANGE: Phonetic target is outside worksheet bounds');
      const writes: Array<{ row: number; column: number; before: CellData; after: CellData }> = [];
      for (let row = range.startRow; row <= range.endRow; row += 1) for (let column = range.startColumn; column <= range.endColumn; column += 1) {
        const current = sheet.cells.get(row, column);
        if (!current || current.value === null || current.value === '') continue;
        if (typeof current.value !== 'string') throw new Error(`PHONETIC_TEXT_REQUIRED: ${params.sheetId}!${row}:${column} is not text`);
        validateMetadata(params.metadata, current.value.length);
        writes.push({ row, column, before: structuredClone(current), after: { ...structuredClone(current), phonetic: structuredClone(params.metadata) } });
      }
      if (!writes.length) throw new Error('PHONETIC_TEXT_REQUIRED: Select at least one non-empty text cell');
      const affectedRanges = [range];
      for (const write of writes) {
        const setParams = createCellSetMutationParams(sheet, { sheetId: params.sheetId, row: write.row, column: write.column, value: write.after }, 'script');
        context.applyMutation({
          id: 'cell.set', unitId: context.workbook.unitId, sheetId: params.sheetId, params: setParams, affectedRanges,
          inverse: [{ id: 'cell.restore', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, row: write.row, column: write.column, previous: write.before }, affectedRanges }],
          apply: () => sheet.cells.set(write.row, write.column, structuredClone(write.after)),
        });
      }
      return { operationId: context.operationId, mutationCount: writes.length, affectedRanges };
    },
  });
}
