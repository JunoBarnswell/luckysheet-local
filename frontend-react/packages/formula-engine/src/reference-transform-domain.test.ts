import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import test from 'node:test';
import { MAX_COLUMN_INDEX, MAX_ROW_INDEX, ReferenceTransformDomain } from './reference-transform-domain';
import { mapAstMovedReferences, mapAstStructuralReferences } from './ast-rewrite';
import { formatFormula } from './ast-format';
import { parseFormula } from './parser';
import { FormulaSyntaxError } from './errors';

interface SharedReferenceTransformVectors {
  readonly schema: string;
  readonly version: number;
  readonly points: readonly {
    readonly id: string;
    readonly axis: 'row' | 'column';
    readonly position: number;
    readonly at: number;
    readonly count: number;
    readonly operation: 'insert' | 'delete';
    readonly expected: { readonly kind: string; readonly position?: number };
  }[];
  readonly intervals: readonly {
    readonly id: string;
    readonly axis: 'row' | 'column';
    readonly start: number;
    readonly end: number;
    readonly at: number;
    readonly count: number;
    readonly operation: 'insert' | 'delete';
    readonly expected: { readonly kind: string; readonly start?: number; readonly end?: number };
  }[];
  readonly cellShiftIndices: readonly {
    readonly id: string;
    readonly axis: 'row' | 'column';
    readonly position: number;
    readonly start: number;
    readonly end: number;
    readonly operation: 'insert' | 'delete';
    readonly expected: number;
  }[];
  readonly cellShiftPoints: readonly {
    readonly id: string;
    readonly axis: 'row' | 'column';
    readonly row: number;
    readonly column: number;
    readonly selection: {
      readonly startRow: number;
      readonly endRow: number;
      readonly startColumn: number;
      readonly endColumn: number;
    };
    readonly operation: 'insert' | 'delete';
    readonly expected: { readonly kind: string; readonly row?: number; readonly column?: number };
  }[];
  readonly rowPermutations: readonly {
    readonly id: string;
    readonly startRow: number;
    readonly sourceRowsByTarget: readonly number[];
    readonly expectedTargetRowsBySource: readonly number[];
  }[];
  readonly permutationPoints: readonly {
    readonly id: string;
    readonly position: number;
    readonly startRow: number;
    readonly targetRowsBySource: readonly number[];
    readonly expected: number;
  }[];
  readonly formulaSheetOrder: readonly { readonly id: string; readonly name: string }[];
  readonly formulaAxes: readonly {
    readonly id: string;
    readonly formula: string;
    readonly ownerSheetId: string;
    readonly targetSheetId: string;
    readonly axis: 'row' | 'column';
    readonly at: number;
    readonly count: number;
    readonly operation: 'insert' | 'delete';
    readonly expected?: string;
    readonly expectedError?: string;
  }[];
  readonly formulaCellShifts: readonly {
    readonly id: string;
    readonly formula: string;
    readonly ownerSheetId: string;
    readonly targetSheetId: string;
    readonly axis: 'row' | 'column';
    readonly selection: {
      readonly startRow: number;
      readonly endRow: number;
      readonly startColumn: number;
      readonly endColumn: number;
    };
    readonly operation: 'insert' | 'delete';
    readonly expected?: string;
    readonly expectedError?: string;
  }[];
  readonly formulaMoves: readonly {
    readonly id: string;
    readonly formula: string;
    readonly ownerSheetId: string;
    readonly targetSheetId: string;
    readonly selection: { readonly startRow: number; readonly endRow: number; readonly startColumn: number; readonly endColumn: number };
    readonly rowDelta: number;
    readonly columnDelta: number;
    readonly expected?: string;
    readonly expectedError?: string;
  }[];
}

const sharedVectors = JSON.parse(readFileSync(
  new URL('../../../../contracts/reference-transform-vectors.json', import.meta.url),
  'utf8',
)) as SharedReferenceTransformVectors;

test('maps point references through insert and delete with distinct outcomes', () => {
  assert.deepEqual(ReferenceTransformDomain.mapPoint(2, 3, 2, 1, MAX_ROW_INDEX), { kind: 'mapped', position: 2 });
  assert.deepEqual(ReferenceTransformDomain.mapPoint(3, 3, 2, 1, MAX_ROW_INDEX), { kind: 'mapped', position: 5 });
  assert.deepEqual(ReferenceTransformDomain.mapPoint(2, 3, 2, -1, MAX_ROW_INDEX), { kind: 'mapped', position: 2 });
  assert.deepEqual(ReferenceTransformDomain.mapPoint(3, 3, 2, -1, MAX_ROW_INDEX), { kind: 'deleted' });
  assert.deepEqual(ReferenceTransformDomain.mapPoint(5, 3, 2, -1, MAX_ROW_INDEX), { kind: 'mapped', position: 3 });
  assert.deepEqual(ReferenceTransformDomain.mapPoint(MAX_ROW_INDEX, MAX_ROW_INDEX, 1, 1, MAX_ROW_INDEX), {
    kind: 'out-of-bounds',
    position: MAX_ROW_INDEX + 1,
  });
});

test('maps inclusive reference intervals without mapping endpoints independently', () => {
  assert.deepEqual(ReferenceTransformDomain.mapInterval(2, 4, { axis: 'row', at: 3, count: 2, op: 'insert' }), { kind: 'mapped', start: 2, end: 6 });
  assert.deepEqual(ReferenceTransformDomain.mapInterval(3, 4, { axis: 'column', at: 3, count: 2, op: 'insert' }), { kind: 'mapped', start: 5, end: 6 });
  assert.deepEqual(ReferenceTransformDomain.mapInterval(2, 5, { axis: 'row', at: 3, count: 2, op: 'delete' }), { kind: 'mapped', start: 2, end: 3 });
  assert.deepEqual(ReferenceTransformDomain.mapInterval(3, 4, { axis: 'row', at: 3, count: 2, op: 'delete' }), { kind: 'deleted' });
});

test('rejects mapped intervals outside Excel address limits and invalid domain inputs', () => {
  assert.deepEqual(ReferenceTransformDomain.mapInterval(MAX_ROW_INDEX, MAX_ROW_INDEX, { axis: 'row', at: MAX_ROW_INDEX, count: 1, op: 'insert' }), {
    kind: 'out-of-bounds', start: MAX_ROW_INDEX + 1, end: MAX_ROW_INDEX + 1,
  });
  assert.deepEqual(ReferenceTransformDomain.mapInterval(MAX_COLUMN_INDEX, MAX_COLUMN_INDEX, { axis: 'column', at: MAX_COLUMN_INDEX, count: 1, op: 'insert' }), {
    kind: 'out-of-bounds', start: MAX_COLUMN_INDEX + 1, end: MAX_COLUMN_INDEX + 1,
  });
  assert.throws(() => ReferenceTransformDomain.mapPoint(0, -1, 1, 1, MAX_ROW_INDEX), /invalid/);
  assert.throws(() => ReferenceTransformDomain.mapInterval(0, 1, { axis: 'row', at: 0, count: 0, op: 'delete' }), /invalid/);
  assert.throws(() => ReferenceTransformDomain.mapInterval(0, 0, { axis: 'row', at: MAX_ROW_INDEX + 1, count: 1, op: 'insert' }), /bounds/);
  assert.throws(() => ReferenceTransformDomain.mapInterval(0, 0, { axis: 'row', at: MAX_ROW_INDEX, count: 2, op: 'delete' }), /bounds/);
});

test('matches the shared TypeScript and Java structural-mapping vectors', () => {
  assert.equal(sharedVectors.schema, 'ReferenceTransformVectors');
  assert.equal(sharedVectors.version, 5);
  for (const vector of sharedVectors.points) {
    const maximum = vector.axis === 'row' ? MAX_ROW_INDEX : MAX_COLUMN_INDEX;
    assert.deepEqual(
      ReferenceTransformDomain.mapPoint(vector.position, vector.at, vector.count, vector.operation === 'insert' ? 1 : -1, maximum),
      vector.expected,
      vector.id,
    );
  }
  for (const vector of sharedVectors.intervals) {
    assert.deepEqual(
      ReferenceTransformDomain.mapInterval(vector.start, vector.end, {
        axis: vector.axis,
        at: vector.at,
        count: vector.count,
        op: vector.operation,
      }),
      vector.expected,
      vector.id,
    );
  }
  for (const vector of sharedVectors.cellShiftIndices) {
    const maximum = vector.axis === 'row' ? MAX_ROW_INDEX : MAX_COLUMN_INDEX;
    assert.equal(
      ReferenceTransformDomain.mapCellShiftIndex(vector.position, vector.start, vector.end, vector.operation, maximum),
      vector.expected,
      vector.id,
    );
  }
  for (const vector of sharedVectors.cellShiftPoints) {
    assert.deepEqual(
      ReferenceTransformDomain.mapCellShiftPoint(
        vector.row,
        vector.column,
        vector.selection,
        vector.axis,
        vector.operation,
      ),
      vector.expected,
      vector.id,
    );
  }
  for (const vector of sharedVectors.rowPermutations) {
    assert.deepEqual(
      ReferenceTransformDomain.createRowPermutationMap(vector.startRow, vector.sourceRowsByTarget),
      vector.expectedTargetRowsBySource,
      vector.id,
    );
  }
  for (const vector of sharedVectors.permutationPoints) {
    assert.equal(
      ReferenceTransformDomain.mapPermutationIndex(vector.position, vector.startRow, vector.targetRowsBySource),
      vector.expected,
      vector.id,
    );
  }
});

test('rejects invalid cell-shift domain inputs', () => {
  assert.throws(() => ReferenceTransformDomain.mapCellShiftIndex(-1, 0, 0, 'insert', MAX_ROW_INDEX), /invalid/);
  assert.throws(() => ReferenceTransformDomain.mapCellShiftIndex(0, 1, 0, 'delete', MAX_ROW_INDEX), /invalid/);
  assert.throws(() => ReferenceTransformDomain.mapCellShiftIndex(0, 0, MAX_ROW_INDEX + 1, 'insert', MAX_ROW_INDEX), /invalid/);
  assert.throws(() => ReferenceTransformDomain.mapCellShiftPoint(0, 0, {
    startRow: 0, endRow: 0, startColumn: 0, endColumn: MAX_COLUMN_INDEX + 1,
  }, 'column', 'insert'), /invalid/);
});

test('rejects invalid row-permutation maps and positions', () => {
  assert.throws(() => ReferenceTransformDomain.createRowPermutationMap(4, [4, 4]), /duplicate/);
  assert.throws(() => ReferenceTransformDomain.createRowPermutationMap(MAX_ROW_INDEX, [MAX_ROW_INDEX, MAX_ROW_INDEX + 1]), /invalid|outside/);
  assert.throws(() => ReferenceTransformDomain.mapPermutationIndex(MAX_ROW_INDEX + 1, 0, [0]), /invalid/);
  assert.throws(() => ReferenceTransformDomain.mapPermutationIndex(0, 0, [1]), /mapping is invalid/);
});

test('matches the shared TypeScript and Java formula-axis vectors', () => {
  assert.ok(sharedVectors.formulaAxes.length > 0);
  for (const vector of sharedVectors.formulaAxes) {
    const owner = sharedVectors.formulaSheetOrder.find((sheet) => sheet.id === vector.ownerSheetId);
    const target = sharedVectors.formulaSheetOrder.find((sheet) => sheet.id === vector.targetSheetId);
    assert.ok(owner, `${vector.id}: missing formula owner`);
    assert.ok(target, `${vector.id}: missing target worksheet`);
    const rewrite = () => formatFormula(mapAstStructuralReferences(parseFormula(vector.formula), {
      ownerSheetId: owner.id,
      targetSheetId: target.id,
      targetSheetName: target.name,
      sheetOrder: sharedVectors.formulaSheetOrder,
      shift: { axis: vector.axis, at: vector.at, count: vector.count, op: vector.operation },
    }));
    if (vector.expectedError !== undefined) {
      assert.throws(rewrite, (error) => error instanceof Error && error.message.includes(vector.expectedError!), vector.id);
    } else {
      assert.equal(rewrite(), vector.expected, vector.id);
    }
  }
});

test('matches the shared TypeScript and Java formula cell-shift vectors', () => {
  assert.ok(sharedVectors.formulaCellShifts.length > 0);
  for (const vector of sharedVectors.formulaCellShifts) {
    const owner = sharedVectors.formulaSheetOrder.find((sheet) => sheet.id === vector.ownerSheetId);
    const target = sharedVectors.formulaSheetOrder.find((sheet) => sheet.id === vector.targetSheetId);
    assert.ok(owner, `${vector.id}: missing formula owner`);
    assert.ok(target, `${vector.id}: missing target worksheet`);
    const rewrite = () => formatFormula(mapAstStructuralReferences(parseFormula(vector.formula), {
      ownerSheetId: owner.id,
      targetSheetId: target.id,
      targetSheetName: target.name,
      sheetOrder: sharedVectors.formulaSheetOrder,
      shift: { axis: vector.axis, at: 0, count: 1, op: vector.operation },
      cellShift: {
        axis: vector.axis,
        selection: vector.selection,
        direction: vector.operation === 'insert' ? 1 : -1,
      },
    }));
    if (vector.expectedError !== undefined) {
      assert.throws(rewrite, (error) => error instanceof Error && error.message.includes(vector.expectedError!), vector.id);
    } else {
      assert.equal(rewrite(), vector.expected, vector.id);
    }
  }
});

test('matches the shared TypeScript and Java moved-formula vectors', () => {
  assert.ok(sharedVectors.formulaMoves.length > 0);
  for (const vector of sharedVectors.formulaMoves) {
    const owner = sharedVectors.formulaSheetOrder.find((sheet) => sheet.id === vector.ownerSheetId);
    const target = sharedVectors.formulaSheetOrder.find((sheet) => sheet.id === vector.targetSheetId);
    assert.ok(owner, `${vector.id}: missing formula owner`);
    assert.ok(target, `${vector.id}: missing target worksheet`);
    const rewrite = () => formatFormula(mapAstMovedReferences(parseFormula(vector.formula), {
      ownerSheetId: owner.id,
      targetSheetId: target.id,
      targetSheetName: target.name,
      sheetOrder: sharedVectors.formulaSheetOrder,
      selection: { sheetId: target.id, ...vector.selection },
      rowDelta: vector.rowDelta,
      columnDelta: vector.columnDelta,
    }));
    if (vector.expectedError !== undefined) {
      assert.throws(rewrite, (error) => error instanceof Error && error.message.includes(vector.expectedError!), vector.id);
    } else {
      assert.equal(rewrite(), vector.expected, vector.id);
    }
  }
});

test('reference parser rejects incomplete reference syntax without swallowing it as a name', () => {
  assert.equal(formatFormula(parseFormula('=Revenue+1')), '=Revenue+1');
  for (const formula of ['=A:', '=A1:', '=Sheet1!', '=A1:bad_name']) {
    assert.throws(() => parseFormula(formula), /reference endpoint|Expected/);
  }
});

test('whole-axis parser accepts Excel limits and rejects endpoints outside the reference domain', () => {
  assert.equal(formatFormula(parseFormula('=SUM($XFD:XFD,$1048576:1048576)')), '=SUM($XFD:XFD,$1048576:1048576)');
  for (const formula of ['=SUM(XFE:XFE)', '=SUM(A:XFE)', '=SUM(1048577:1048577)', '=SUM(1:$1048577)', '=SUM(0:1)', '=SUM(Revenue:Other)']) {
    assert.throws(() => parseFormula(formula), FormulaSyntaxError);
  }
});

test('quoted 3D whole-axis references reject an edit inside their worksheet span', () => {
  assert.throws(() => mapAstStructuralReferences(parseFormula("=SUM('Budget A1':Other!A:A)"), {
    ownerSheetId: 'local-id',
    targetSheetId: 'budget-id',
    targetSheetName: 'Budget A1',
    sheetOrder: sharedVectors.formulaSheetOrder,
    shift: { axis: 'row', at: 0, count: 1, op: 'insert' },
  }), /cannot rewrite one sheet inside a 3D reference/);
});
