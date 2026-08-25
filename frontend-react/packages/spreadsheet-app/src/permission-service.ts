import { protectionResolver, type ProtectionAction, type ProtectionRule, type RangeRef, type WorkbookModel } from '@react-sheets/core-model';
import {
  buildPermissionCapabilities,
  inferAffectedRanges,
  isPermissionExempt,
  resolveCommandAction,
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

function protectionActionForCommand(commandId: string, action: PermissionAction): ProtectionAction {
  if (commandId === 'sheet.rows.insert') return 'insert-rows';
  if (commandId === 'sheet.rows.delete') return 'delete-rows';
  if (commandId === 'sheet.columns.insert') return 'insert-columns';
  if (commandId === 'sheet.columns.delete') return 'delete-columns';
  if (action === 'format') return 'format';
  if (commandId.startsWith('sheetTable.style') || commandId.startsWith('pivot.')) return 'format';
  if (action === 'drawing') return 'edit-objects';
  if (commandId.startsWith('sheet.autoFilter') || commandId.startsWith('sheetTable.autoFilter')) return 'auto-filter';
  if (commandId.startsWith('data.sort') || commandId.startsWith('sheet.sort')) return 'sort';
  return 'edit-cell';
}

function protectionActionForMutation(mutationId: string): ProtectionAction | undefined {
  if (mutationId === 'style.set') return 'format';
  if (mutationId === 'rows.inserted') return 'insert-rows';
  if (mutationId === 'rows.deleted') return 'delete-rows';
  if (mutationId === 'columns.inserted') return 'insert-columns';
  if (mutationId === 'columns.deleted') return 'delete-columns';
  if (mutationId === 'rows.permuted') return 'sort';
  if (mutationId === 'autoFilter.set' || mutationId === 'autoFilter.remove' || mutationId === 'sheetTable.autoFilter.set') return 'auto-filter';
  if (mutationId.startsWith('drawing.') || mutationId.startsWith('chart.') || mutationId.startsWith('picture.')) return 'edit-objects';
  if (mutationId.startsWith('merge.')) return 'format';
  if (mutationId.startsWith('sheetTable.style.') || mutationId.startsWith('pivot.')) return 'format';
  if (mutationId.startsWith('sheet.protect.')) return undefined;
  return 'edit-cell';
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

    const allowsPendingSheet = input.commandId === 'pivot.create' || input.commandId === 'pivot.add' || input.commandId === 'pivot.drillDown';
    return this.checkProtection(protectionActionForCommand(input.commandId, action), input.affectedRanges, allowsPendingSheet);
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

  checkMutation(mutation: { id: string; affectedRanges: readonly RangeRef[]; params?: unknown }): PermissionResult {
    const action = mutation.id === 'cell.restore' && this.isFormatOnlyRestore(mutation.params)
      ? 'format'
      : protectionActionForMutation(mutation.id);
    const allowsPendingSheet = mutation.id === 'pivot.add' || mutation.id === 'pivot.drilldown.add';
    return action ? this.checkProtection(action, mutation.affectedRanges, allowsPendingSheet) : { allowed: true };
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

  /** A style command's inverse is a cell.restore payload, but it must retain
   * the original format permission rather than becoming a value edit during
   * undo/redo. Content changes are deliberately not classified this way. */
  isFormatOnlyRestore(params: unknown): boolean {
    if (!this.workbook || !params || typeof params !== 'object' || Array.isArray(params)) return false;
    const candidate = params as Record<string, unknown>;
    if (typeof candidate.sheetId !== 'string' || !Number.isInteger(candidate.row) || !Number.isInteger(candidate.column)) return false;
    const sheet = this.workbook.getSheet(candidate.sheetId);
    const current = sheet.cells.get(Number(candidate.row), Number(candidate.column));
    const previous = candidate.previous;
    if (previous !== undefined && (!previous || typeof previous !== 'object' || Array.isArray(previous))) return false;
    const currentContent = current ? { value: current.value, formula: current.formula, formulaValue: current.formulaValue } : undefined;
    const previousRecord = previous as Record<string, unknown> | undefined;
    const previousContent = previousRecord ? { value: previousRecord.value, formula: previousRecord.formula, formulaValue: previousRecord.formulaValue } : undefined;
    if (previous === undefined) {
      return Boolean(current && current.value === null && current.formula === undefined && current.formulaValue === undefined && current.style);
    }
    return JSON.stringify(currentContent) === JSON.stringify(previousContent);
  }

  private checkProtection(action: ProtectionAction, affectedRanges: readonly RangeRef[], allowPendingSheet = false): PermissionResult {
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
      });
      if (!decision.allowed) return { allowed: false, reason: decision.reason, blockedBy: decision.blockedBy };
    }
    return { allowed: true };
  }
}

export type { PermissionCapabilities };
