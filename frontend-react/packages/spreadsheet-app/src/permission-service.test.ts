import test from 'node:test';
import assert from 'node:assert/strict';
import { PermissionService } from './permission-service';

test('PermissionService blocks viewer from editing cells', () => {
  const perm = new PermissionService();
  perm.applyServerAccess('viewer');
  perm.setOnline(true);
  const result = perm.canCheck({
    commandId: 'sheet.cell.set',
    affectedRanges: [{ sheetId: 's1', startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }],
    actor: { actorId: 'user-1' },
  });
  assert.equal(result.allowed, false);
  assert.equal(result.blockedBy, 'share-role');
});

test('PermissionService allows navigation commands for viewers', () => {
  const perm = new PermissionService();
  perm.applyServerAccess('viewer');
  perm.setOnline(true);
  const result = perm.canCheck({
    commandId: 'ui.panel.open',
    affectedRanges: [{ sheetId: 's1', startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }],
    actor: { actorId: 'user-1' },
  });
  assert.equal(result.allowed, true);
});

test('PermissionService fails closed while an online access projection is unavailable', () => {
  const perm = new PermissionService();
  perm.setOnline(true);
  const result = perm.canCheck({
    commandId: 'sheet.cell.set',
    affectedRanges: [{ sheetId: 's1', startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }],
    actor: { actorId: 'untrusted-client-actor' },
  });
  assert.equal(result.allowed, false);
  assert.equal(result.blockedBy, 'share-role');
});

test('PermissionService blocks locked range', () => {
  const perm = new PermissionService();
  perm.applyServerAccess('editor');
  perm.setOnline(true);
  perm.setRangeRules([{
    id: 'r1',
    scope: 'range',
    sheetId: 's1',
    range: { sheetId: 's1', startRow: 0, endRow: 10, startColumn: 0, endColumn: 5 },
    locked: true,
    allow: {},
    allowedActions: ['format'],
  }]);
  const result = perm.canCheck({
    commandId: 'sheet.cell.set',
    affectedRanges: [{ sheetId: 's1', startRow: 2, endRow: 2, startColumn: 1, endColumn: 1 }],
    actor: { actorId: 'user-1' },
  });
  assert.equal(result.allowed, false);
  assert.equal(result.blockedBy !== 'share-role' && result.blockedBy?.id, 'r1');
});
