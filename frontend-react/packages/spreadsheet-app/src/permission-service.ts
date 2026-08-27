import { protectionResolver, type ProtectionAction, type ProtectionRule, type RangeRef, type WorkbookModel } from '@react-sheets/core-model';
import { mutationPermission, type PermissionPolicy } from '@react-sheets/protocol';
import {
  buildPermissionCapabilities,
  inferAffectedRanges,
  isPermissionExempt,
  resolveCommandPermission,
  type PermissionAction,
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

function capabilityAllowed(capabilities: PermissionCapabilities, action: PermissionAction): boolean {
  return action === 'navigate' ? capabilities.navigate
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
}

const PERMISSION_ACTIONS: ReadonlySet<string> = new Set([
  'navigate', 'edit-cell', 'format', 'structure', 'drawing', 'protect', 'share', 'comment', 'restore', 'query', 'script',
]);

function mutationPolicyOverride(value: { capability: string; protectionAction: ProtectionAction | 'none'; checksProtection: boolean; affectedRangeMode: 'none' | 'declared' | 'exact'; objectScope: 'cell' | 'range' | 'row' | 'column' | 'drawing' | 'worksheet' | 'workbook' }): PermissionPolicy | undefined {
  if (!PERMISSION_ACTIONS.has(value.capability)) return undefined;
  return {
    capability: value.capability as PermissionAction,
    protectionAction: value.protectionAction,
    checksProtection: value.checksProtection,
    affectedRangeMode: value.affectedRangeMode,
    objectScope: value.objectScope,
  };
}

/** Workbook/Sheet/Range 权限 — 命令 dispatch 前拦截 */
export class PermissionService {
  private workbook: WorkbookModel | null = null;
  private serverRole: ShareRole | null = null;
  private online = false;

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
    const policy = resolveCommandPermission(input.commandId);
    if (!policy) return { allowed: false, reason: `Unknown command permission contract: ${input.commandId}`, blockedBy: 'share-role' };
    const action = policy.capability;
    const role = this.serverRole;
    const capabilities = this.getCapabilities();

    if (!capabilityAllowed(capabilities, action)) {
      return { allowed: false, reason: `Server role "${role ?? 'unknown'}" cannot perform "${action}"`, blockedBy: 'share-role' };
    }

    if (!policy.checksProtection) {
      return { allowed: true };
    }

    const allowsPendingSheet = input.commandId === 'pivot.create' || input.commandId === 'pivot.drillDown';
    if (policy.protectionAction === 'none') {
      return { allowed: false, reason: `Command permission contract requires a protection action: ${input.commandId}`, blockedBy: 'share-role' };
    }
    return this.checkProtection(policy.protectionAction, input.affectedRanges, allowsPendingSheet);
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
    this.workbook = workbook;
  }

  checkMutation(mutation: { id: string; affectedRanges: readonly RangeRef[]; params?: unknown; permission?: { capability: string; protectionAction: ProtectionAction | 'none'; checksProtection: boolean; affectedRangeMode: 'none' | 'declared' | 'exact'; objectScope: 'cell' | 'range' | 'row' | 'column' | 'drawing' | 'worksheet' | 'workbook' } }): PermissionResult {
    const policy = mutation.permission ? mutationPolicyOverride(mutation.permission) : mutationPermission(mutation.id);
    if (!policy) return { allowed: false, reason: `Unknown mutation permission contract: ${mutation.id}`, blockedBy: 'share-role' };
    if (!capabilityAllowed(this.getCapabilities(), policy.capability)) {
      return { allowed: false, reason: `Server role "${this.serverRole ?? 'unknown'}" cannot perform "${policy.capability}"`, blockedBy: 'share-role' };
    }
    if (!policy.checksProtection) return { allowed: true };
    if (policy.protectionAction === 'none') {
      return { allowed: false, reason: `Mutation permission contract requires a protection action: ${mutation.id}`, blockedBy: 'share-role' };
    }
    const allowsPendingSheet = mutation.id === 'pivot.add' || mutation.id === 'pivot.drilldown.add';
    return this.checkProtection(policy.protectionAction, mutation.affectedRanges, allowsPendingSheet);
  }

  canSelectCell(sheetId: string, row: number, column: number): PermissionResult {
    if (!this.workbook) return { allowed: true };
    const sheet = this.workbook.getSheet(sheetId);
    const rules = this.workbook.getSheets().flatMap((candidate) => candidate.protectionRules);
    const range: RangeRef = { sheetId, startRow: row, endRow: row, startColumn: column, endColumn: column };
    const style = sheet.cells.get(row, column)?.style;
    const resolution = protectionResolver.resolveCell(rules, sheetId, row, column, style);
    if (!resolution.active && resolution.rules.length === 0) return { allowed: true };
    const action: ProtectionAction = resolution.locked ? 'select-locked' : 'select-unlocked';
    return this.checkProtection(action, [range]);
  }

  private checkProtection(action: Exclude<ProtectionAction, 'none'>, affectedRanges: readonly RangeRef[], allowPendingSheet = false): PermissionResult {
    if (!this.workbook) return { allowed: true };
    const rangesBySheet = new Map<string, RangeRef[]>();
    for (const range of affectedRanges) {
      const ranges = rangesBySheet.get(range.sheetId) ?? [];
      ranges.push(range);
      rangesBySheet.set(range.sheetId, ranges);
    }
    const rules = this.workbook.getSheets().flatMap((candidate) => candidate.protectionRules);
    for (const [sheetId, ranges] of rangesBySheet) {
      const sheet = this.workbook.sheets.get(sheetId);
      if (!sheet) {
        if (allowPendingSheet) continue;
        return { allowed: false, reason: `Unknown protected worksheet: ${sheetId}` };
      }
      const decision = protectionResolver.resolve({
        sheetId,
        rules,
        ranges,
        action,
        rowCount: sheet.rowCount,
        columnCount: sheet.columnCount,
        readCellStyle: (row, column) => sheet.cells.get(row, column)?.style,
        countUnlockedCells: (range) => {
          let count = 0;
          sheet.cells.forEachInRange(range.startRow, range.endRow, range.startColumn, range.endColumn, (cell) => {
            if (cell.style?.locked === false) count += 1;
          });
          return count;
        },
      });
      if (!decision.allowed) return { allowed: false, reason: decision.reason, blockedBy: decision.blockedBy };
    }
    return { allowed: true };
  }
}

export type { PermissionCapabilities };
