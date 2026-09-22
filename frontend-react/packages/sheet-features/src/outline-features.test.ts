import test from 'node:test';
import assert from 'node:assert/strict';
import { WorksheetModel } from '@react-sheets/core-model';
import type { OutlineGroup } from '@react-sheets/core-model';
import { computeOutlineHiddenColumns, computeOutlineHiddenRows } from './data-features';
import {
  buildColumnOutlineGroup,
  buildRowOutlineGroup,
  groupsWithinRange,
  nextOutlineLevel,
  resolveOutlineControls,
} from './outline-features';

const rowGroup: OutlineGroup = {
  id: 'g1',
  axis: 'row',
  start: 1,
  end: 4,
  level: 1,
  collapsed: true,
};

test('computeOutlineHiddenRows keeps the first grouped row visible', () => {
  const sheet = new WorksheetModel('s1', 'Sheet1');
  sheet.outline = { groups: [rowGroup] };
  const hidden = computeOutlineHiddenRows(sheet);
  assert.deepEqual([...hidden].sort((a, b) => a - b), [2, 3, 4]);
});

test('nextOutlineLevel increases for overlapping row groups', () => {
  const sheet = new WorksheetModel('s1', 'Sheet1');
  sheet.outline = { groups: [rowGroup] };
  assert.equal(nextOutlineLevel(sheet, 'row', 2, 5), 2);
});

test('buildRowOutlineGroup and resolveOutlineControls expose chrome metadata', () => {
  const sheet = new WorksheetModel('s1', 'Sheet1');
  const group = buildRowOutlineGroup('s1', { sheetId: 's1', startRow: 5, endRow: 8, startColumn: 0, endColumn: 3 }, sheet, 'g2');
  sheet.outline = { groups: [group] };
  assert.equal(group.level, 1);
  assert.deepEqual(resolveOutlineControls(sheet), [{ axis: 'row', index: 5, level: 1, collapsed: false, groupId: 'g2' }]);
});

test('groupsWithinRange finds groups fully inside the selection', () => {
  const sheet = new WorksheetModel('s1', 'Sheet1');
  sheet.outline = { groups: [rowGroup, { ...rowGroup, id: 'g2', start: 10, end: 12 }] };
  const matches = groupsWithinRange(sheet.outline, 'row', {
    sheetId: 's1',
    startRow: 0,
    endRow: 12,
    startColumn: 0,
    endColumn: 2,
  });
  assert.equal(matches.length, 2);
});

const columnGroup: OutlineGroup = {
  id: 'cg1',
  axis: 'column',
  start: 1,
  end: 4,
  level: 1,
  collapsed: true,
};

test('computeOutlineHiddenColumns keeps the first grouped column visible', () => {
  const sheet = new WorksheetModel('s1', 'Sheet1');
  sheet.outline = { groups: [columnGroup] };
  const hidden = computeOutlineHiddenColumns(sheet);
  assert.deepEqual([...hidden].sort((a, b) => a - b), [2, 3, 4]);
});

test('buildColumnOutlineGroup and resolveOutlineControls expose column chrome metadata', () => {
  const sheet = new WorksheetModel('s1', 'Sheet1');
  const group = buildColumnOutlineGroup('s1', { sheetId: 's1', startRow: 0, endRow: 3, startColumn: 5, endColumn: 8 }, sheet, 'cg2');
  sheet.outline = { groups: [group] };
  assert.equal(group.level, 1);
  assert.deepEqual(resolveOutlineControls(sheet), [{ axis: 'column', index: 5, level: 1, collapsed: false, groupId: 'cg2' }]);
});
