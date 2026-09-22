import assert from 'node:assert/strict';
import test from 'node:test';
import { normalizeExternalFormula } from './external-formula-normalizer';

test('normalizes Excel and WPS function namespaces without rewriting string literals', () => {
  assert.equal(normalizeExternalFormula('=_xlfn.FILTER(_xlws.A1:A3,B1:B3=1)'), '=FILTER(A1:A3,B1:B3=1)');
  assert.equal(normalizeExternalFormula('="_xlfn.FILTER"&_xlfn.UNIQUE(A1:A3)'), '="_xlfn.FILTER"&UNIQUE(A1:A3)');
});

test('normalizes the legacy SINGLE spelling into the canonical implicit-intersection operator', () => {
  assert.equal(normalizeExternalFormula('=_xlfn.SINGLE(A1)'), '=@(A1)');
  assert.equal(normalizeExternalFormula('=@A1'), '=@A1');
});
