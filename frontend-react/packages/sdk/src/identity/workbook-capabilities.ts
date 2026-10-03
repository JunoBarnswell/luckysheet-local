import { isWorkbookRole, type WorkbookRole } from '@react-sheets/protocol';

/** Lifecycle capabilities are distinct from system administration and range protection. */
export interface WorkbookCapabilities {
  readonly canOpen: boolean;
  readonly canEdit: boolean;
  readonly canComment: boolean;
  readonly canRename: boolean;
  readonly canMove: boolean;
  readonly canShare: boolean;
  readonly canProtect: boolean;
  readonly canCopy: boolean;
  readonly canExport: boolean;
  readonly canTrash: boolean;
  readonly canRestore: boolean;
  readonly canPurge: boolean;
  readonly readOnly: boolean;
}

export function workbookCapabilities(role: WorkbookRole, lifecycle: 'active' | 'trashed'): WorkbookCapabilities {
  if (!isWorkbookRole(role)) throw new TypeError('Invalid canonical workbook role');
  if (lifecycle !== 'active' && lifecycle !== 'trashed') throw new TypeError('Invalid workbook lifecycle');
  const active = lifecycle === 'active';
  const owner = role === 'owner';
  const editor = owner || role === 'editor';
  return Object.freeze({
    canOpen: active, canEdit: active && editor, canComment: active && role !== 'viewer',
    canRename: active && editor, canMove: active && editor, canShare: active && owner,
    canProtect: active && owner, canCopy: active, canExport: active,
    canTrash: active && owner, canRestore: !active && owner, canPurge: !active && owner,
    readOnly: !active || !editor,
  });
}
