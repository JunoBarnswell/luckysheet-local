import test from 'node:test';
import assert from 'node:assert/strict';
import { buildGanttProjection } from './projection';
import type { CanvasSheetSnapshot } from '../../ui-snapshot';

const table = {
  id: 'tasks', name: 'Tasks', sourceSheetId: 'source', rowCount: 3, blockSize: 4096, blocks: [], revision: 0,
  fields: [
    { id: 'id', name: 'ID', ordinal: 0, type: 'text' as const },
    { id: 'title', name: 'Title', ordinal: 1, type: 'text' as const },
    { id: 'start', name: 'Start', ordinal: 2, type: 'date' as const },
    { id: 'end', name: 'End', ordinal: 3, type: 'date' as const },
    { id: 'progress', name: 'Progress', ordinal: 4, type: 'number' as const },
    { id: 'parent', name: 'Parent', ordinal: 5, type: 'text' as const },
    { id: 'deps', name: 'Dependencies', ordinal: 6, type: 'text' as const },
  ],
};
const definition = { viewId: 'tasks', fieldMap: { id: 'id', title: 'title', start: 'start', end: 'end', progress: 'progress', parentId: 'parent', dependencies: 'deps' }, calendar: { workingDays: [1, 2, 3, 4, 5], dayStartHour: 9, dayEndHour: 18 }, timeline: { unit: 'day' as const }, dependencyStyle: { color: '#64748b', width: 1 } };

function sheetWithRows(rows: string[][], nextDefinition = definition): CanvasSheetSnapshot {
  return {
    id: 'gantt', name: 'Gantt', rowCount: 20, columnCount: 20, columns: [], occupiedCellCount: rows.length * 7,
    kind: 'gantt-sheet', ganttSheet: structuredClone(nextDefinition),
    getCell: (row: number, column: number) => row > 0 && rows[row - 1]?.[column] !== undefined ? { value: rows[row - 1]![column]!, address: `A${row}`, displayValue: rows[row - 1]![column]! } : undefined,
  } as unknown as CanvasSheetSnapshot;
}

test('Gantt projection preserves hierarchy, dependencies, dates and progress', () => {
  const projection = buildGanttProjection(sheetWithRows([
    ['A', 'Planning', '2026-08-25', '2026-08-26', '20', '', ''],
    ['B', 'Build', '2026-08-26', '2026-08-28', '50', 'A', 'A'],
    ['C', 'Review', '2026-08-28', '2026-08-29', '100', 'B', 'B'],
  ]), [table]);
  assert.equal(projection.status, 'ready');
  assert.equal(projection.tasks[1]?.level, 1);
  assert.deepEqual(projection.tasks[2]?.dependencies, ['B']);
  assert.equal(projection.tasks[0]?.progress, 20);
});

test('Gantt projection rejects missing references and dependency cycles', () => {
  const missing = buildGanttProjection(sheetWithRows([['A', 'Task', '2026-08-25', '2026-08-26', '0', '', 'Z']]), [table]);
  assert.equal(missing.status, 'error');
  assert.match(missing.error ?? '', /missing dependency/);
  const cycle = buildGanttProjection(sheetWithRows([
    ['A', 'A', '2026-08-25', '2026-08-26', '0', '', 'B'],
    ['B', 'B', '2026-08-25', '2026-08-26', '0', '', 'A'],
  ]), [table]);
  assert.equal(cycle.status, 'error');
  assert.match(cycle.error ?? '', /cycle/);
});
