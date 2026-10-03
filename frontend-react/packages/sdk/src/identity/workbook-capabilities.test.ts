import assert from 'node:assert/strict';
import test from 'node:test';
import { isWorkbookRole, WORKBOOK_ROLES, type WorkbookRole } from '@react-sheets/protocol';
import { workbookCapabilities } from './workbook-capabilities';

test('canonical role generation and lifecycle capabilities cover every role', () => {
  assert.deepEqual(WORKBOOK_ROLES, ['owner', 'editor', 'commenter', 'viewer']);
  for (const role of WORKBOOK_ROLES) {
    const active = workbookCapabilities(role, 'active');
    const trash = workbookCapabilities(role, 'trashed');
    assert.equal(active.canEdit, role === 'owner' || role === 'editor');
    assert.equal(active.canComment, role !== 'viewer');
    assert.equal(active.canShare, role === 'owner');
    assert.equal(active.canTrash, role === 'owner');
    assert.equal(active.canRestore, false);
    assert.equal(trash.canOpen, false);
    assert.equal(trash.canEdit, false);
    assert.equal(trash.canExport, false);
    assert.equal(trash.canRestore, role === 'owner');
    assert.equal(trash.canPurge, role === 'owner');
    assert.ok(Object.isFrozen(active));
  }
});
test('unknown roles and lifecycle states fail closed', () => {
  for (const value of [null, undefined, 'admin', 'OWNER', '', 3]) {
    assert.equal(isWorkbookRole(value), false);
    assert.throws(() => workbookCapabilities(value as WorkbookRole, 'active'), /canonical workbook role/);
  }
  assert.throws(() => workbookCapabilities('owner', 'deleted' as 'active'), /lifecycle/);
});
