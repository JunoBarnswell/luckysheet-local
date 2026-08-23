import type { ProtectionRule, RangeRef } from '@react-sheets/core-model';

/** 共享角色 — 与 Excel Share 语义对齐 */
export type ShareRole = 'owner' | 'editor' | 'commenter' | 'viewer';

export interface ActorContext {
  actorId: string;
  role: ShareRole;
}

export interface PermissionCheckInput {
  commandId: string;
  affectedRanges: RangeRef[];
  actor: ActorContext;
  params?: unknown;
}

export interface PermissionResult {
  allowed: boolean;
  reason?: string;
  blockedBy?: ProtectionRule | 'share-role';
}

const COMMAND_ACTION_MAP: Readonly<Record<string, string>> = {
  'sheet.cell.set': 'edit-cell',
  'sheet.range.set': 'edit-cell',
  'sheet.range.clear': 'edit-cell',
  'sheet.range.paste': 'edit-cell',
  'sheet.style.set': 'format',
  'sheet.merge.set': 'format',
  'sheet.row.insert': 'structure',
  'sheet.row.delete': 'structure',
  'sheet.column.insert': 'structure',
  'sheet.column.delete': 'structure',
  'sheet.protect.set': 'protect',
  'workbook.share.set': 'share',
  'comment.add': 'comment',
  'comment.reply': 'comment',
  'history.restore': 'restore',
  'query.load': 'query',
  'query.refresh': 'query',
  'automation.run': 'script',
};

const ROLE_CAPABILITIES: Readonly<Record<ShareRole, ReadonlySet<string>>> = {
  owner: new Set(['edit-cell', 'format', 'structure', 'protect', 'share', 'comment', 'restore', 'query', 'script']),
  editor: new Set(['edit-cell', 'format', 'structure', 'comment', 'query', 'script']),
  commenter: new Set(['comment']),
  viewer: new Set([]),
};

function rangesOverlap(a: RangeRef, b: RangeRef): boolean {
  return a.sheetId === b.sheetId
    && a.startRow <= b.endRow
    && b.startRow <= a.endRow
    && a.startColumn <= b.endColumn
    && b.startColumn <= a.endColumn;
}

function rangeContains(outer: RangeRef, inner: RangeRef): boolean {
  return outer.sheetId === inner.sheetId
    && outer.startRow <= inner.startRow
    && outer.endRow >= inner.endRow
    && outer.startColumn <= inner.startColumn
    && outer.endColumn >= inner.endColumn;
}

export function inferAffectedRanges(commandId: string, params: unknown, sheetId: string): RangeRef[] {
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

/** Workbook/Sheet/Range 权限 — 命令 dispatch 前拦截 */
export class PermissionService {
  private workbookRules: ProtectionRule[] = [];
  private sheetRules = new Map<string, ProtectionRule[]>();
  private rangeRules: ProtectionRule[] = [];
  private shareRoles = new Map<string, ShareRole>();

  setWorkbookRules(rules: ProtectionRule[]): void {
    this.workbookRules = [...rules];
  }

  setSheetRules(sheetId: string, rules: ProtectionRule[]): void {
    this.sheetRules.set(sheetId, [...rules]);
  }

  setRangeRules(rules: ProtectionRule[]): void {
    this.rangeRules = [...rules];
  }

  setShareRole(actorId: string, role: ShareRole): void {
    this.shareRoles.set(actorId, role);
  }

  getShareRole(actorId: string): ShareRole {
    return this.shareRoles.get(actorId) ?? 'owner';
  }

  canCheck(input: PermissionCheckInput): PermissionResult {
    const action = COMMAND_ACTION_MAP[input.commandId] ?? 'edit-cell';
    const role = input.actor.role ?? this.getShareRole(input.actor.actorId);
    const capabilities = ROLE_CAPABILITIES[role];

    if (!capabilities.has(action)) {
      return { allowed: false, reason: `Role "${role}" cannot perform "${action}"`, blockedBy: 'share-role' };
    }

    for (const rule of this.allRules()) {
      if (!rule.locked) continue;
      const allowedActions = new Set(rule.allowedActions ?? []);
      if (allowedActions.has(action)) continue;
      const blocked = input.affectedRanges.some((range) => this.ruleCoversRange(rule, range));
      if (blocked) {
        return { allowed: false, reason: `Protected area blocks "${action}"`, blockedBy: rule };
      }
    }
    return { allowed: true };
  }

  assertAllowed(input: PermissionCheckInput): void {
    const result = this.canCheck(input);
    if (!result.allowed) throw new Error(result.reason ?? 'Permission denied');
  }

  /** @deprecated 兼容旧 API */
  can(commandId: string, params?: unknown, actorId?: string): boolean {
    const sheetId = 'sheet-1';
    const role = actorId ? this.getShareRole(actorId) : 'owner';
    return this.canCheck({
      commandId,
      affectedRanges: inferAffectedRanges(commandId, params, sheetId),
      actor: { actorId: actorId ?? 'local', role },
      params,
    }).allowed;
  }

  /** @deprecated 兼容旧 API */
  assert(commandId: string, params?: unknown, actorId?: string): void {
    if (!this.can(commandId, params, actorId)) {
      throw new Error(`Permission denied for command: ${commandId}`);
    }
  }

  private allRules(): ProtectionRule[] {
    const sheetLevel = [...this.sheetRules.values()].flat();
    return [...this.workbookRules, ...sheetLevel, ...this.rangeRules];
  }

  private ruleCoversRange(rule: ProtectionRule, range: RangeRef): boolean {
    if (rule.scope === 'workbook') return true;
    if (rule.scope === 'sheet' && rule.sheetId === range.sheetId) return true;
    if (rule.scope === 'range' && rule.range) {
      return rangeContains(rule.range, range) || rangesOverlap(rule.range, range);
    }
    return false;
  }
}
