import type { ProtectionRule, RangeRef, WorkbookModel } from '@react-sheets/core-model';
import type { PermissionService, ShareRole } from './permission-service';

export type PermissionAction =
  | 'navigate'
  | 'edit-cell'
  | 'format'
  | 'structure'
  | 'drawing'
  | 'protect'
  | 'share'
  | 'comment'
  | 'restore'
  | 'query'
  | 'script';

export interface PermissionCapabilities {
  navigate: boolean;
  editCell: boolean;
  format: boolean;
  structure: boolean;
  drawing: boolean;
  protect: boolean;
  share: boolean;
  comment: boolean;
  restore: boolean;
  query: boolean;
  script: boolean;
}

export const PERMISSION_EXEMPT_COMMAND_PREFIXES = [
  'ui.panel.open',
  'ui.dialog.open',
  'ui.zoom.',
  'ui.notice',
  'ui.freeze.',
  'ui.file.export',
  'selection.set',
  'navigation.',
  'permission.role.set',
] as const;

export function inferAffectedRanges(commandId: string, params: unknown, sheetId: string): RangeRef[] {
  void commandId;
  const p = params as Record<string, unknown> | null;
  if (!p) return [{ sheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }];
  if (typeof p.row === 'number' && typeof p.column === 'number') {
    return [{ sheetId, startRow: p.row, endRow: p.row, startColumn: p.column, endColumn: p.column }];
  }
  if (p.range && typeof p.range === 'object') return [p.range as RangeRef];
  if (Array.isArray(p.ranges)) return p.ranges as RangeRef[];
  if (typeof p.startRow === 'number') {
    return [{
      sheetId: (p.sheetId as string) ?? sheetId,
      startRow: p.startRow as number,
      endRow: (p.endRow as number) ?? (p.startRow as number),
      startColumn: (p.startColumn as number) ?? 0,
      endColumn: (p.endColumn as number) ?? 0,
    }];
  }
  return [{ sheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }];
}

const COMMAND_ACTION_MAP: Readonly<Record<string, PermissionAction>> = {
  'sheet.cell.set': 'edit-cell',
  'sheet.range.set': 'edit-cell',
  'sheet.range.clear': 'edit-cell',
  'sheet.range.paste': 'edit-cell',
  'sheet.style.set': 'format',
  'sheet.style.toggle': 'format',
  'sheet.merge.set': 'format',
  'sheet.merge.remove': 'format',
  'sheet.cf.add': 'format',
  'sheet.cf.remove': 'format',
  'sheet.rows.insert': 'structure',
  'sheet.rows.delete': 'structure',
  'sheet.columns.insert': 'structure',
  'sheet.columns.delete': 'structure',
  'sheet.row.hide': 'structure',
  'sheet.column.hide': 'structure',
  'sheet.protect.set': 'protect',
  'sheet.protect.remove': 'protect',
  'workbook.share.set': 'share',
  'comment.add': 'comment',
  'comment.reply': 'comment',
  'comment.resolve': 'comment',
  'comment.remove': 'comment',
  'note.set': 'comment',
  'note.remove': 'comment',
  'note.visibility': 'comment',
  'hyperlink.set': 'edit-cell',
  'hyperlink.remove': 'edit-cell',
  'history.restore': 'restore',
  'persistence.save': 'edit-cell',
  'persistence.draft.clear': 'edit-cell',
  'xlsx.import': 'edit-cell',
  'xlsx.export': 'navigate',
  'print.preview': 'navigate',
  'print.export': 'navigate',
  'print.area.set': 'navigate',
  'print.pageSetup': 'navigate',
  'query.load': 'query',
  'query.refresh': 'query',
  'automation.run': 'script',
  'automation.record.start': 'script',
  'automation.record.stop': 'script',
  'extended.whatIf.goalSeek': 'script',
  'extended.whatIf.scenario': 'script',
  'extended.whatIf.dataTable': 'script',
  'extended.capability.evaluate': 'navigate',
  'drawing.insert.rectangle': 'drawing',
  'drawing.zorder.forward': 'drawing',
  'drawing.zorder.backward': 'drawing',
  'drawing.remove.selected': 'drawing',
  'chart.insert.column': 'drawing',
  'pivot.insert.quick': 'structure',
  'sparkline.insert.quick': 'drawing',
};

const ROLE_CAPABILITIES: Readonly<Record<ShareRole, ReadonlySet<PermissionAction>>> = {
  owner: new Set(['navigate', 'edit-cell', 'format', 'structure', 'drawing', 'protect', 'share', 'comment', 'restore', 'query', 'script']),
  editor: new Set(['navigate', 'edit-cell', 'format', 'structure', 'drawing', 'comment', 'query', 'script']),
  commenter: new Set(['navigate', 'comment']),
  viewer: new Set(['navigate']),
};

export function isPermissionExempt(commandId: string): boolean {
  return PERMISSION_EXEMPT_COMMAND_PREFIXES.some((prefix) => commandId.startsWith(prefix));
}

export function resolveCommandAction(commandId: string): PermissionAction {
  if (isPermissionExempt(commandId)) return 'navigate';
  if (COMMAND_ACTION_MAP[commandId]) return COMMAND_ACTION_MAP[commandId]!;
  if (commandId.startsWith('comment.') || commandId.startsWith('note.')) return 'comment';
  if (commandId.startsWith('hyperlink.')) return 'edit-cell';
  if (commandId.startsWith('drawing.') || commandId.startsWith('chart.') || commandId.startsWith('shape.')) return 'drawing';
  if (commandId.startsWith('sparkline.')) return 'drawing';
  if (commandId.startsWith('pivot.')) return 'structure';
  if (commandId.startsWith('sheet.style') || commandId.startsWith('sheet.merge') || commandId.startsWith('sheet.cf')) return 'format';
  if (commandId.startsWith('sheet.row') || commandId.startsWith('sheet.column') || commandId.startsWith('sheet.rows') || commandId.startsWith('sheet.columns') || commandId.startsWith('outline.') || commandId.startsWith('data.')) {
    return 'structure';
  }
  if (commandId.startsWith('history.')) return 'restore';
  if (commandId.startsWith('persistence.')) return 'edit-cell';
  if (commandId.startsWith('xlsx.')) return commandId === 'xlsx.export' ? 'navigate' : 'edit-cell';
  if (commandId.startsWith('print.')) return 'navigate';
  if (commandId.startsWith('query.')) return 'query';
  if (commandId.startsWith('automation.')) return 'script';
  if (commandId.startsWith('extended.whatIf')) return 'script';
  if (commandId.startsWith('extended.capability')) return 'navigate';
  if (commandId.startsWith('sheet.protect')) return 'protect';
  if (commandId.startsWith('ui.clipboard.copy')) return 'navigate';
  if (commandId.startsWith('ui.history.')) return 'edit-cell';
  return 'edit-cell';
}

export function buildPermissionCapabilities(role: ShareRole): PermissionCapabilities {
  const capabilities = ROLE_CAPABILITIES[role];
  return {
    navigate: capabilities.has('navigate'),
    editCell: capabilities.has('edit-cell'),
    format: capabilities.has('format'),
    structure: capabilities.has('structure'),
    drawing: capabilities.has('drawing'),
    protect: capabilities.has('protect'),
    share: capabilities.has('share'),
    comment: capabilities.has('comment'),
    restore: capabilities.has('restore'),
    query: capabilities.has('query'),
    script: capabilities.has('script'),
  };
}

export function syncProtectionRulesFromWorkbook(permission: PermissionService, workbook: WorkbookModel): void {
  const rangeRules: ProtectionRule[] = [];
  const sheetRules = new Map<string, ProtectionRule[]>();
  for (const sheet of workbook.getSheets()) {
    for (const rule of sheet.protectionRules) {
      if (rule.scope === 'range') {
        rangeRules.push(rule);
        continue;
      }
      if (rule.scope === 'sheet') {
        const existing = sheetRules.get(sheet.id) ?? [];
        existing.push(rule);
        sheetRules.set(sheet.id, existing);
      }
    }
  }
  permission.setRangeRules(rangeRules);
  for (const sheet of workbook.getSheets()) {
    permission.setSheetRules(sheet.id, sheetRules.get(sheet.id) ?? []);
  }
}

export function canExecuteCommand(
  permission: PermissionService,
  workbook: WorkbookModel,
  commandId: string,
  params: unknown,
  actorId: string,
  activeSheetId: string,
): { allowed: boolean; reason?: string } {
  syncProtectionRulesFromWorkbook(permission, workbook);
  const role = permission.getShareRole(actorId);
  const action = resolveCommandAction(commandId);
  const capabilities = ROLE_CAPABILITIES[role];
  if (!capabilities.has(action)) {
    return { allowed: false, reason: `Role "${role}" cannot perform "${action}"` };
  }
  const affectedRanges = inferAffectedRanges(commandId, params, activeSheetId);
  const result = permission.canCheck({
    commandId,
    affectedRanges,
    actor: { actorId, role },
    params,
  });
  return { allowed: result.allowed, reason: result.reason };
}

export function findProtectionRuleCoveringRange(
  workbook: WorkbookModel,
  sheetId: string,
  range: RangeRef,
): ProtectionRule | undefined {
  const sheet = workbook.getSheet(sheetId);
  return sheet.protectionRules.find((rule) => {
    if (!rule.locked || rule.scope !== 'range' || !rule.range) return false;
    return rule.range.sheetId === range.sheetId
      && rule.range.startRow <= range.startRow
      && rule.range.endRow >= range.endRow
      && rule.range.startColumn <= range.startColumn
      && rule.range.endColumn >= range.endColumn;
  });
}
