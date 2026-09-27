import assert from 'node:assert/strict';
import test from 'node:test';
import {
  compareWorkbookText,
  DEFAULT_WORKBOOK_COLLATION,
  normalizeWorkbookCollation,
} from './collation';

test('workbook culture controls locale-specific text order', () => {
  const swedish = normalizeWorkbookCollation({ ...DEFAULT_WORKBOOK_COLLATION, cultureId: 'sv-SE' });
  const english = normalizeWorkbookCollation({ ...DEFAULT_WORKBOOK_COLLATION, cultureId: 'en-US' });

  assert.ok(compareWorkbookText('z', 'ä', swedish) < 0);
  assert.ok(compareWorkbookText('ä', 'z', english) < 0);
});

test('workbook case and accent sensitivity map to locale collation', () => {
  const caseInsensitive = normalizeWorkbookCollation({
    ...DEFAULT_WORKBOOK_COLLATION,
    cultureId: 'en-US',
    caseSensitive: false,
    accentSensitive: true,
  });
  const accentInsensitive = normalizeWorkbookCollation({
    ...DEFAULT_WORKBOOK_COLLATION,
    cultureId: 'en-US',
    caseSensitive: true,
    accentSensitive: false,
  });

  assert.equal(compareWorkbookText('A', 'a', caseInsensitive), 0);
  assert.notEqual(compareWorkbookText('a', 'á', caseInsensitive), 0);
  assert.notEqual(compareWorkbookText('A', 'a', accentInsensitive), 0);
  assert.equal(compareWorkbookText('a', 'á', accentInsensitive), 0);
});

test('custom lists and numeric-text ordering remain effective with a culture', () => {
  const context = normalizeWorkbookCollation({
    ...DEFAULT_WORKBOOK_COLLATION,
    cultureId: 'en-US',
    caseSensitive: false,
    numericTextMode: 'numeric',
    customLists: [['April', 'August']],
  });

  assert.ok(compareWorkbookText('april', 'ordinary', context) < 0);
  assert.ok(compareWorkbookText('item2', 'item10', context) < 0);
  assert.ok(compareWorkbookText('item9007199254740993', 'item9007199254740992', context) > 0);
});

test('invariant collation retains its ordinal comparison behavior', () => {
  assert.ok(compareWorkbookText('A', 'a', DEFAULT_WORKBOOK_COLLATION) < 0);
});

test('invariant accent-insensitive comparison removes all Unicode combining marks', () => {
  const context = normalizeWorkbookCollation({
    ...DEFAULT_WORKBOOK_COLLATION,
    accentSensitive: false,
  });

  assert.equal(compareWorkbookText('a\u1ab0', 'a', context), 0);
});
