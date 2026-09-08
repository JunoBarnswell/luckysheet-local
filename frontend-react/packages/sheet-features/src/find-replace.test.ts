import assert from 'node:assert/strict';
import test from 'node:test';
import { matchesFindText, replaceFindText, parseReplacementValue } from './find-replace';

const context = {
  sourceKind: 'find-replace' as const, cultureId: 'en-US', decimalSeparator: '.', groupSeparator: ',', dateSystem: '1900' as const,
  referenceDate: { year: 2026, month: 8, day: 27, hour: 0, minute: 0, second: 0, millisecond: 0 },
};

test('Find matcher treats wildcard escapes and replacements as literal text', () => {
  assert.equal(matchesFindText('A*B', { query: 'A~*B', wildcard: true }), true);
  assert.equal(replaceFindText('ab ab', { query: 'a?', matchCase: false, wildcard: true }, '$1'), '$1 $1');
  assert.equal(replaceFindText('prefix', { query: 'prefix', entireCell: true }, '0'), '0');
  assert.equal(replaceFindText('prefix', { query: 'fix', entireCell: true }, '0'), undefined);
});

test('replacement parser keeps typed values and rejects empty replacements at the command boundary', () => {
  assert.deepEqual(parseReplacementValue('0', context), { kind: 'number', value: 0 });
  assert.deepEqual(parseReplacementValue('FALSE', context), { kind: 'boolean', value: false });
  assert.deepEqual(parseReplacementValue('=A1+1', context), { kind: 'formula', value: null, formula: '=A1+1' });
  assert.equal(parseReplacementValue('', context).kind, 'empty');
});
