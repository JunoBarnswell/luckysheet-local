import test from 'node:test';
import assert from 'node:assert/strict';
import { WorkbookModel } from '@react-sheets/core-model';
import { planFill } from './fill-series';

function setup() {
  const workbook = new WorkbookModel('fill-series-unit', 'Fill Series');
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.rowCount = 16;
  sheet.columnCount = 16;
  return sheet;
}

function range(sheetId: string, startRow: number, endRow: number, startColumn: number, endColumn: number) {
  return { sheetId, startRow, endRow, startColumn, endColumn };
}

function writesByCoordinate(plan: ReturnType<typeof planFill>): Map<string, number | string | boolean | null | undefined> {
  return new Map(plan.writes.map((entry) => [`${entry.row}:${entry.column}`, entry.cell?.value]));
}

test('series planner derives a multi-seed vertical progression without rewriting seeds', () => {
  const sheet = setup();
  sheet.cells.set(0, 0, { value: 1 });
  sheet.cells.set(1, 0, { value: 3 });
  const source = range(sheet.id, 0, 1, 0, 0);
  const before = structuredClone(sheet.cells.toJSON());

  const plan = planFill(sheet, {
    sheetId: sheet.id,
    sourceRange: source,
    targetRange: range(sheet.id, 0, 4, 0, 0),
    direction: 'down',
    mode: 'series',
  });

  assert.deepEqual([...writesByCoordinate(plan).entries()], [['2:0', 5], ['3:0', 7], ['4:0', 9]]);
  assert.deepEqual(sheet.cells.toJSON(), before);
  assert.equal(plan.writes.some((entry) => entry.row < 2), false);
});

test('series planner supports descending, decimal and negative values in every axis direction', () => {
  const sheet = setup();
  sheet.cells.set(3, 0, { value: 10 });
  sheet.cells.set(4, 0, { value: 8 });
  const up = planFill(sheet, {
    sheetId: sheet.id,
    sourceRange: range(sheet.id, 3, 4, 0, 0),
    targetRange: range(sheet.id, 0, 4, 0, 0),
    direction: 'up',
    mode: 'series',
  });
  assert.deepEqual([...writesByCoordinate(up).entries()], [['0:0', 16], ['1:0', 14], ['2:0', 12]]);

  sheet.cells.set(0, 2, { value: -1.5 });
  sheet.cells.set(0, 3, { value: -0.5 });
  const right = planFill(sheet, {
    sheetId: sheet.id,
    sourceRange: range(sheet.id, 0, 0, 2, 3),
    targetRange: range(sheet.id, 0, 0, 2, 5),
    direction: 'right',
    mode: 'series',
  });
  assert.deepEqual([...writesByCoordinate(right).entries()], [['0:4', 0.5], ['0:5', 1.5]]);

  sheet.cells.set(2, 2, { value: 10 });
  sheet.cells.set(2, 3, { value: 8 });
  const left = planFill(sheet, {
    sheetId: sheet.id,
    sourceRange: range(sheet.id, 2, 2, 2, 3),
    targetRange: range(sheet.id, 2, 2, 0, 3),
    direction: 'left',
    mode: 'series',
  });
  assert.deepEqual([...writesByCoordinate(left).entries()], [['2:0', 14], ['2:1', 12]]);
});

test('copy planner shifts formulas and preserves the source cells', () => {
  const sheet = setup();
  sheet.cells.set(1, 0, { value: null, formula: '=A1+$B$1' });
  const down = planFill(sheet, {
    sheetId: sheet.id,
    sourceRange: range(sheet.id, 1, 1, 0, 0),
    targetRange: range(sheet.id, 1, 3, 0, 0),
    direction: 'down',
    mode: 'copy',
  });
  assert.deepEqual([...writesByCoordinate(down).entries()], [['2:0', null], ['3:0', null]]);
  assert.equal(down.writes[0]?.cell?.formula, '=A2+$B$1');
  assert.equal(down.writes[1]?.cell?.formula, '=A3+$B$1');
  assert.equal(sheet.cells.get(1, 0)?.formula, '=A1+$B$1');

  sheet.cells.set(0, 1, { value: null, formula: '=A1+B1' });
  const right = planFill(sheet, {
    sheetId: sheet.id,
    sourceRange: range(sheet.id, 0, 0, 1, 1),
    targetRange: range(sheet.id, 0, 0, 1, 3),
    direction: 'right',
    mode: 'copy',
  });
  assert.equal(right.writes[0]?.cell?.formula, '=B1+C1');
  assert.equal(right.writes[1]?.cell?.formula, '=C1+D1');
});

test('series planner rejects invalid seeds and non-canonical geometry before producing writes', () => {
  const sheet = setup();
  sheet.cells.set(0, 0, { value: '1' });
  assert.throws(() => planFill(sheet, {
    sheetId: sheet.id,
    sourceRange: range(sheet.id, 0, 0, 0, 0),
    targetRange: range(sheet.id, 0, 2, 0, 0),
    direction: 'down',
    mode: 'series',
  }), /finite numeric seeds/);

  sheet.cells.delete(0, 0);
  sheet.cells.set(0, 0, { value: null, formula: '=A1+1' });
  assert.throws(() => planFill(sheet, {
    sheetId: sheet.id,
    sourceRange: range(sheet.id, 0, 0, 0, 0),
    targetRange: range(sheet.id, 0, 2, 0, 0),
    direction: 'down',
    mode: 'series',
  }), /finite numeric seeds/);

  sheet.cells.set(0, 0, { value: 1 });
  assert.throws(() => planFill(sheet, {
    sheetId: sheet.id,
    sourceRange: range(sheet.id, 0, 0, 0, 0),
    targetRange: range(sheet.id, 0, 2, 0, 1),
    direction: 'down',
    mode: 'series',
  }), /one-axis target extension/);

  assert.throws(() => planFill(sheet, {
    sheetId: sheet.id,
    sourceRange: range(sheet.id, 0, 0, 0, 0),
    targetRange: range(sheet.id, 0, 16, 0, 0),
    direction: 'down',
    mode: 'copy',
  }), /outside worksheet bounds/);
});

test('planner normalizes reversed range endpoints before applying axis semantics', () => {
  const sheet = setup();
  sheet.cells.set(0, 0, { value: 1 });
  sheet.cells.set(1, 0, { value: 3 });
  const plan = planFill(sheet, {
    sheetId: sheet.id,
    sourceRange: range(sheet.id, 1, 0, 0, 0),
    targetRange: range(sheet.id, 4, 0, 0, 0),
    direction: 'down',
    mode: 'series',
  });
  assert.deepEqual(plan.sourceRange, range(sheet.id, 0, 1, 0, 0));
  assert.deepEqual(plan.targetRange, range(sheet.id, 0, 4, 0, 0));
  assert.deepEqual([...writesByCoordinate(plan).entries()], [['2:0', 5], ['3:0', 7], ['4:0', 9]]);
});

test('series planner applies growth, date-unit, stop-value and autofill semantics', () => {
  const sheet = setup();
  sheet.cells.set(0, 0, { value: 2 });
  sheet.cells.set(1, 0, { value: 4 });
  const growth = planFill(sheet, {
    sheetId: sheet.id,
    sourceRange: range(sheet.id, 0, 1, 0, 0),
    targetRange: range(sheet.id, 0, 4, 0, 0),
    direction: 'down',
    mode: 'series',
    series: { type: 'growth', seriesIn: 'columns', stopValue: 16 },
  });
  assert.deepEqual([...writesByCoordinate(growth).entries()], [['2:0', 8], ['3:0', 16]]);

  sheet.cells.set(0, 2, { value: '2026-01-01T00:00:00.000Z', numberFormat: 'yyyy-mm-dd' });
  const date = planFill(sheet, {
    sheetId: sheet.id,
    sourceRange: range(sheet.id, 0, 0, 2, 2),
    targetRange: range(sheet.id, 0, 2, 2, 2),
    direction: 'down',
    mode: 'series',
    series: { type: 'date', seriesIn: 'columns', dateUnit: 'month', stepValue: 1 },
  });
  assert.deepEqual([...writesByCoordinate(date).entries()], [['1:2', '2026-02-01T00:00:00.000Z'], ['2:2', '2026-03-01T00:00:00.000Z']]);

  sheet.cells.set(0, 4, { value: 'A' });
  const auto = planFill(sheet, {
    sheetId: sheet.id,
    sourceRange: range(sheet.id, 0, 0, 4, 4),
    targetRange: range(sheet.id, 0, 2, 4, 4),
    direction: 'down',
    mode: 'series',
    series: { type: 'autofill' },
  });
  assert.deepEqual([...writesByCoordinate(auto).entries()], [['1:4', 'A'], ['2:4', 'A']]);
});

test('series planner uses least-squares trend semantics when explicitly requested', () => {
  const sheet = setup();
  sheet.cells.set(0, 0, { value: 1 });
  sheet.cells.set(1, 0, { value: 3 });
  sheet.cells.set(2, 0, { value: 8 });
  const plan = planFill(sheet, {
    sheetId: sheet.id,
    sourceRange: range(sheet.id, 0, 2, 0, 0),
    targetRange: range(sheet.id, 0, 4, 0, 0),
    direction: 'down',
    mode: 'series',
    series: { trend: true },
  });
  assert.deepEqual([...writesByCoordinate(plan).entries()], [['3:0', 11], ['4:0', 14.5]]);
});
