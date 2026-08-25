import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookModel } from '@react-sheets/core-model';
import { PermissionService } from '../../permission-service';
import { buildPermissionCapabilities, canExecuteCommand, resolveCommandAction } from './policy';

describe('permission policy', () => {
  it('maps review and drawing commands to permission actions', () => {
    assert.equal(resolveCommandAction('comment.add'), 'comment');
    assert.equal(resolveCommandAction('hyperlink.set'), 'edit-cell');
    assert.equal(resolveCommandAction('checkbox.toggle'), 'edit-cell');
    assert.equal(resolveCommandAction('dataChart.create'), 'drawing');
    assert.equal(resolveCommandAction('ui.panel.open'), 'navigate');
  });

  it('builds capability snapshots per share role', () => {
    const viewer = buildPermissionCapabilities('viewer');
    const commenter = buildPermissionCapabilities('commenter');
    assert.equal(viewer.editCell, false);
    assert.equal(viewer.navigate, true);
    assert.equal(commenter.comment, true);
    assert.equal(commenter.editCell, false);
  });

  it('blocks viewer edits and protected range writes for editors', () => {
    const workbook = new WorkbookModel('wb', 'Permission');
    const viewerPermission = new PermissionService();
    viewerPermission.applyServerAccess('viewer');
    viewerPermission.setOnline(true);
    assert.equal(canExecuteCommand(viewerPermission, workbook, 'sheet.cell.set', { row: 0, column: 0, value: { value: 1 } }, 'viewer-1', 'sheet-1').allowed, false);

    const editorPermission = new PermissionService();
    editorPermission.applyServerAccess('editor');
    editorPermission.setOnline(true);
    workbook.getSheet('sheet-1').protectionRules.push({
      id: 'rule-1',
      scope: 'range',
      sheetId: 'sheet-1',
      range: { sheetId: 'sheet-1', startRow: 0, endRow: 2, startColumn: 0, endColumn: 2 },
      locked: true,
      allow: {},
      allowedActions: ['format'],
    });
    const blocked = canExecuteCommand(editorPermission, workbook, 'sheet.cell.set', { row: 1, column: 1, value: { value: 2 } }, 'editor-1', 'sheet-1');
    assert.equal(blocked.allowed, false);
    const formatAllowed = canExecuteCommand(editorPermission, workbook, 'sheet.style.set', { range: { sheetId: 'sheet-1', startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 }, style: { bold: true } }, 'editor-1', 'sheet-1');
    assert.equal(formatAllowed.allowed, true);
  });
});
