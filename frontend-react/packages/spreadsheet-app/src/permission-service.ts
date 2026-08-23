import type { ProtectionRule, RangeRef, WorkbookModel } from '@react-sheets/core-model';
import {
  buildPermissionCapabilities,
  inferAffectedRanges,
  isPermissionExempt,
  resolveCommandAction,
  syncProtectionRulesFromWorkbook,
  type PermissionCapabilities,
} from './permission-bridge';

export { inferAffectedRanges } from './permission-bridge';

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

  getCapabilities(actorId: string): PermissionCapabilities {
    return buildPermissionCapabilities(this.getShareRole(actorId));
  }

  canCheck(input: PermissionCheckInput): PermissionResult {
    if (isPermissionExempt(input.commandId)) return { allowed: true };
    const action = resolveCommandAction(input.commandId);
    const role = input.actor.role ?? this.getShareRole(input.actor.actorId);
    const capabilities = buildPermissionCapabilities(role);
    const capabilityAllowed = action === 'navigate' ? capabilities.navigate
      : action === 'edit-cell' ? capabilities.editCell
        : action === 'format' ? capabilities.format
          : action === 'structure' ? capabilities.structure
            : action === 'drawing' ? capabilities.drawing
              : action === 'protect' ? capabilities.protect
                : action === 'share' ? capabilities.share
                  : action === 'comment' ? capabilities.comment
                    : action === 'restore' ? capabilities.restore
                      : action === 'query' ? capabilities.query
                        : capabilities.script;

    if (!capabilityAllowed) {
      return { allowed: false, reason: `Role "${role}" cannot perform "${action}"`, blockedBy: 'share-role' };
    }

    // A protection rule controls workbook content, not its owner-managed
    // lifecycle. Otherwise a rule that does not explicitly allow "protect"
    // could never be removed by the owner who created it.
    if (action === 'protect' || action === 'share' || action === 'restore') {
      return { allowed: true };
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

  checkCommand(commandId: string, params: unknown, actorId: string, activeSheetId: string): PermissionResult {
    const role = this.getShareRole(actorId);
    return this.canCheck({
      commandId,
      affectedRanges: inferAffectedRanges(commandId, params, activeSheetId),
      actor: { actorId, role },
      params,
    });
  }

  syncFromWorkbook(workbook: WorkbookModel): void {
    syncProtectionRulesFromWorkbook(this, workbook);
  }

  /** @deprecated 兼容旧 API */
  can(commandId: string, params?: unknown, actorId?: string, activeSheetId = 'sheet-1'): boolean {
    const role = actorId ? this.getShareRole(actorId) : 'owner';
    return this.canCheck({
      commandId,
      affectedRanges: inferAffectedRanges(commandId, params, activeSheetId),
      actor: { actorId: actorId ?? 'local', role },
      params,
    }).allowed;
  }

  /** @deprecated 兼容旧 API */
  assert(commandId: string, params?: unknown, actorId?: string, activeSheetId = 'sheet-1'): void {
    const result = this.checkCommand(commandId, params, actorId ?? 'local', activeSheetId);
    if (!result.allowed) throw new Error(result.reason ?? `Permission denied for command: ${commandId}`);
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

export type { PermissionCapabilities };
