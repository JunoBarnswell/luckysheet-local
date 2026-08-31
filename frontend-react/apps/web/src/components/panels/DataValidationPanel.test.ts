import assert from 'node:assert/strict';
import test from 'node:test';
import { buildDataValidationRule } from './DataValidationPanel';

const range = { sheetId: 'sheet-1', startRow: 4, endRow: 8, startColumn: 2, endColumn: 3 };

test('buildDataValidationRule owns the actual selection and typed list source', () => {
  const rule = buildDataValidationRule({
    id: 'dv-list',
    sheetId: 'sheet-1',
    range,
    type: 'list',
    operator: 'between',
    formula1: 'Open, Closed, Review',
    formula2: '',
    listSourceKind: 'values',
    errorMessage: 'Choose a state',
  });
  assert.deepEqual(rule.ranges, [range]);
  assert.deepEqual(rule.formulaAnchor, { sheetId: 'sheet-1', row: 4, column: 2 });
  assert.deepEqual(rule.listSource, { kind: 'values', values: ['Open', 'Closed', 'Review'] });
  assert.equal(rule.formula1, undefined);
  const rangeRule = buildDataValidationRule({ id: 'dv-range', sheetId: 'sheet-1', range, type: 'list', operator: 'between', formula1: '$D$2:$D$8', formula2: '', listSourceKind: 'range', errorMessage: 'Choose' });
  assert.deepEqual(rangeRule.listSource, { kind: 'range', range: { sheetId: 'sheet-1', startRow: 1, endRow: 7, startColumn: 3, endColumn: 3 } });
});

test('buildDataValidationRule creates complete bounds and canonical custom formulas', () => {
  const bounded = buildDataValidationRule({ id: 'dv-number', sheetId: 'sheet-1', range, type: 'whole', operator: 'between', formula1: '1', formula2: '10', errorMessage: '1-10' });
  assert.equal(bounded.operator, 'between');
  assert.equal(bounded.formula1, '1');
  assert.equal(bounded.formula2, '10');
  const custom = buildDataValidationRule({ id: 'dv-custom', sheetId: 'sheet-1', range, type: 'custom', operator: 'equal', formula1: 'A5<>""', formula2: '', errorMessage: 'Required' });
  assert.equal(custom.formula1, '=A5<>""');
});

test('buildDataValidationRule rejects missing, cross-sheet, and incomplete ranges', () => {
  const base = { id: 'dv-invalid', sheetId: 'sheet-1', type: 'whole' as const, operator: 'between' as const, formula1: '1', formula2: '', errorMessage: 'Invalid' };
  assert.throws(() => buildDataValidationRule({ ...base, range: undefined }), /Select a worksheet range/);
  assert.throws(() => buildDataValidationRule({ ...base, range: { ...range, sheetId: 'sheet-2' } }), /another worksheet/);
  assert.throws(() => buildDataValidationRule({ ...base, range }), /second validation bound/);
});
