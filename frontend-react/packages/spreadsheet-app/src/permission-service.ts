import type { ProtectionRule, RangeRef, WorkbookModel } from '@react-sheets/core-model';
import {
  buildPermissionCapabilities,
  inferAffectedRanges,
  isPermissionExempt,
  resolveCommandAction,
  syncProtectionRulesFromWorkbook,
  type PermissionCapabilities,
} from './features/permission';

export { inferAffectedRanges } from './features/permission';

/** 共享角色 — 与 Excel Share 语义对齐 */
export type ShareRole = 'owner' | 'editor' | 'commenter' | 'viewer';

export interface ActorContext {
  actorId: string;
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

const LOCAL_CAPABILITIES: PermissionCapabilities = Object.freeze({
  navigate: true,
  editCell: true,
  format: true,
  structure: true,
  drawing: true,
  protect: true,
  share: false,
  comment: true,
  restore: true,
  query: true,
  script: true,
});

const UNKNOWN_REMOTE_CAPABILITIES: PermissionCapabilities = Object.freeze({
  navigate: true,
  editCell: false,
  format: false,
  structure: false,
  drawing: false,
  protect: false,
  share: false,
  comment: false,
  restore: false,
  query: false,
  script: false,
});

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
  private serverRole: ShareRole | null = null;
  private online = false;

  setWorkbookRules(rules: ProtectionRule[]): void {
    this.workbookRules = [...rules];
  }

  setSheetRules(sheetId: string, rules: ProtectionRule[]): void {
    this.sheetRules.set(sheetId, [...rules]);
  }

  setRangeRules(rules: ProtectionRule[]): void {
    this.rangeRules = [...rules];
  }

  /** Consume the server-calculated projection; no UI or command can set it. */
  applyServerAccess(role: ShareRole): void {
    this.serverRole = role;
  }

  clearServerAccess(): void {
    this.serverRole = null;
  }

  setOnline(online: boolean): void {
    this.online = online;
  }

  getShareRole(): ShareRole | null {
    return this.serverRole;
  }

  getCapabilities(): PermissionCapabilities {
    if (!this.online) return LOCAL_CAPABILITIES;
    return this.serverRole ? buildPermissionCapabilities(this.serverRole) : UNKNOWN_REMOTE_CAPABILITIES;
  }

  canCheck(input: PermissionCheckInput): PermissionResult {
    if (isPermissionExempt(input.commandId)) return { allowed: true };
    const action = resolveCommandAction(input.commandId);
    const role = this.serverRole;
    const capabilities = this.getCapabilities();
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
      return { allowed: false, reason: `Server role "${role ?? 'unknown'}" cannot perform "${action}"`, blockedBy: 'share-role' };
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
    return this.canCheck({
      commandId,
      affectedRanges: inferAffectedRanges(commandId, params, activeSheetId),
      actor: { actorId },
      params,
    });
  }

  syncFromWorkbook(workbook: WorkbookModel): void {
    syncProtectionRulesFromWorkbook(this, workbook);
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
