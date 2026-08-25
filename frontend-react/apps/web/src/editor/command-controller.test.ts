import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { buildPivotSourceOptions } from './command-controller';

const currentDataRange = { sheetId: 'sheet-1', startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 };

describe('Pivot named-range source picker', () => {
  it('keeps workbook and worksheet scopes distinct in both ids and source payloads', () => {
    const options = buildPivotSourceOptions({
      currentDataRange,
      currentSheetName: 'Sheet1',
      sheetTables: [],
      sheetNames: new Map([
        ['sheet-1', 'Sheet1'],
        ['sheet-2', 'Sheet2'],
      ]),
      locale: 'en-US',
      definedNameModels: [
        { name: 'SharedName', formula: "='Sheet1'!A1:B2", scope: 'workbook' },
        { name: 'SharedName', formula: "='Sheet1'!D1:E2", scope: 'sheet', sheetId: 'sheet-1' },
        { name: 'SharedName', formula: "='Sheet2'!A1:B2", scope: 'sheet', sheetId: 'sheet-2' },
      ],
    });

    const named = options.filter((option) => option.source.kind === 'named-range');
    assert.equal(named.length, 3);
    assert.equal(new Set(named.map((option) => option.id)).size, 3);
    assert.deepEqual(named.map((option) => option.source), [
      { kind: 'named-range', name: 'SharedName' },
      { kind: 'named-range', name: 'SharedName', sheetId: 'sheet-1' },
      { kind: 'named-range', name: 'SharedName', sheetId: 'sheet-2' },
    ]);
    assert.match(named[0]!.label, /SharedName · Workbook$/);
    assert.match(named[1]!.label, /SharedName · Worksheet: Sheet1$/);
    assert.match(named[2]!.label, /SharedName · Worksheet: Sheet2$/);
  });

  it('fails closed when a defined-name model has an invalid scope identity', () => {
    const base = {
      currentDataRange,
      currentSheetName: 'Sheet1',
      sheetTables: [],
      sheetNames: new Map([['sheet-1', 'Sheet1']]),
      locale: 'en-US' as const,
    };

    assert.throws(() => buildPivotSourceOptions({
      ...base,
      definedNameModels: [{ name: 'LocalName', formula: '=A1', scope: 'sheet' }],
    }), /requires a sheetId/);
    assert.throws(() => buildPivotSourceOptions({
      ...base,
      definedNameModels: [{ name: 'LocalName', formula: '=A1', scope: 'sheet', sheetId: 'missing-sheet' }],
    }), /unknown sheet/);
    assert.throws(() => buildPivotSourceOptions({
      ...base,
      definedNameModels: [{ name: 'GlobalName', formula: '=A1', scope: 'workbook', sheetId: 'sheet-1' }],
    }), /cannot specify sheetId/);
    assert.throws(() => buildPivotSourceOptions({
      ...base,
      definedNameModels: [{ name: 'InvalidScope', formula: '=A1', scope: 'session', sheetId: 'sheet-1' } as never],
    }), /unsupported scope/);
  });
});
