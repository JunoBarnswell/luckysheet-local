import { guardRecordWorksheetWrites } from '@react-sheets/core-model';
import {
  MAX_SHEET_COLUMN_COUNT,
  MAX_SHEET_ROW_COUNT,
  protectionResolver,
  type ProtectionAction,
  type ProtectionRule,
  type RangeRef,
  type WorkbookModel,
} from '@react-sheets/core-model';
import {
  mutationPermission,
  type EffectiveAccessRegion,
  type PermissionPolicy,
  type RangeAccessLevel,
  type WorkbookAccessResponse,
} from '@react-sheets/protocol';
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
  blockedBy?: ProtectionRule | 'share-role' | 'range-access';
}

const RANGE_ACCESS_RANK: Readonly<Record<RangeAccessLevel, number>> = Object.freeze({ hidden: 0, read: 1, edit: 2 });

interface RegionNode {
  minRow: number;
  maxRow: number;
  minColumn: number;
  maxColumn: number;
  regions: readonly EffectiveAccessRegion[];
  left?: RegionNode;
  right?: RegionNode;
}

class AccessProjectionIndex {
  private readonly roots = new Map<string, RegionNode>();

  constructor(regions: readonly EffectiveAccessRegion[]) {
    const bySheet = new Map<string, EffectiveAccessRegion[]>();
    for (const region of regions) {
      const list = bySheet.get(region.range.sheetId) ?? [];
      list.push(region);
      bySheet.set(region.range.sheetId, list);
    }
    for (const [sheetId, source] of bySheet) this.roots.set(sheetId, this.build(source));
  }

  intersecting(range: RangeRef): readonly EffectiveAccessRegion[] {
    const root = this.roots.get(range.sheetId);
    if (!root) return [];
    const result: EffectiveAccessRegion[] = [];
    const visit = (node: RegionNode): void => {
      if (!this.intersectsBounds(node, range)) return;
      if (node.regions.length) {
        for (const region of node.regions) if (this.intersects(region.range, range)) result.push(region);
      } else {
        if (node.left) visit(node.left);
        if (node.right) visit(node.right);
      }
    };
    visit(root);
    return result;
  }

  private build(source: readonly EffectiveAccessRegion[]): RegionNode {
    const box = { minRow: Number.POSITIVE_INFINITY, maxRow: -1, minColumn: Number.POSITIVE_INFINITY, maxColumn: -1 };
    for (const region of source) {
      box.minRow = Math.min(box.minRow, region.range.startRow);
      box.maxRow = Math.max(box.maxRow, region.range.endRow);
      box.minColumn = Math.min(box.minColumn, region.range.startColumn);
      box.maxColumn = Math.max(box.maxColumn, region.range.endColumn);
    }
    if (source.length <= 8) return { ...box, regions: source };
    const rows = box.maxRow - box.minRow >= box.maxColumn - box.minColumn;
    const sorted = [...source].sort((a, b) => rows
      ? (a.range.startRow + a.range.endRow) - (b.range.startRow + b.range.endRow)
      : (a.range.startColumn + a.range.endColumn) - (b.range.startColumn + b.range.endColumn));
    const middle = Math.floor(sorted.length / 2);
    return { ...box, regions: [], left: this.build(sorted.slice(0, middle)), right: this.build(sorted.slice(middle)) };
  }

  private intersectsBounds(node: RegionNode, range: RangeRef): boolean {
    return node.minRow <= range.endRow && range.startRow <= node.maxRow
      && node.minColumn <= range.endColumn && range.startColumn <= node.maxColumn;
  }

  private intersects(left: RangeRef, right: RangeRef): boolean {
    return left.startRow <= right.endRow && right.startRow <= left.endRow
      && left.startColumn <= right.endColumn && right.startColumn <= left.endColumn;
  }
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
  private accessRevision = 0;
  private accessRegions: readonly EffectiveAccessRegion[] = [];
  private accessIndex = new AccessProjectionIndex([]);
  private online = false;

  /** Consume the server-calculated projection; no UI or command can set it. */
  applyServerAccess(access: WorkbookAccessResponse): void {
    this.serverRole = access.role;
    this.accessRevision = access.accessRevision;
    this.accessRegions = structuredClone(access.regions);
    this.accessIndex = new AccessProjectionIndex(this.accessRegions);
  }

  clearServerAccess(): void {
    this.serverRole = null;
    this.accessRevision = 0;
    this.accessRegions = [];
    this.accessIndex = new AccessProjectionIndex([]);
  }

  setOnline(online: boolean): void {
    this.online = online;
  }

  getShareRole(): ShareRole | null {
    return this.serverRole;
  }

  getAccessRevision(): number { return this.accessRevision; }

  getAccessRegions(): readonly EffectiveAccessRegion[] { return this.accessRegions; }

  isWorkbookAccessManager(): boolean { return this.serverRole === 'owner'; }

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

    const rangeCheck = this.checkRangeAccess(action, input.affectedRanges);
    if (!rangeCheck.allowed) return rangeCheck;

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
    if (this.workbook) { try { guardRecordWorksheetWrites(this.workbook, mutation.id, mutation.affectedRanges); } catch (error) { return { allowed: false, reason: error instanceof Error ? error.message : 'Record field is read-only', blockedBy: 'range-access' }; } }
    const policy = mutation.permission ? mutationPolicyOverride(mutation.permission) : mutationPermission(mutation.id);
    if (!policy) return { allowed: false, reason: `Unknown mutation permission contract: ${mutation.id}`, blockedBy: 'share-role' };
    if (!capabilityAllowed(this.getCapabilities(), policy.capability)) {
      return { allowed: false, reason: `Server role "${this.serverRole ?? 'unknown'}" cannot perform "${policy.capability}"`, blockedBy: 'share-role' };
    }
    const rangeCheck = this.checkRangeAccess(policy.capability, mutation.affectedRanges);
    if (!rangeCheck.allowed) return rangeCheck;
    if (!policy.checksProtection) return { allowed: true };
    if (policy.protectionAction === 'none') {
      return { allowed: false, reason: `Mutation permission contract requires a protection action: ${mutation.id}`, blockedBy: 'share-role' };
    }
    const allowsPendingSheet = mutation.id === 'pivot.add' || mutation.id === 'pivot.drilldown.add';
    return this.checkProtection(policy.protectionAction, mutation.affectedRanges, allowsPendingSheet);
  }

  canSelectCell(sheetId: string, row: number, column: number): PermissionResult {
    if (this.serverRole && this.effectiveAccess({ sheetId, startRow: row, endRow: row, startColumn: column, endColumn: column }) === 'hidden') {
      return { allowed: false, reason: 'This range is hidden from the current subject', blockedBy: 'range-access' };
    }
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

  canSelectRange(range: RangeRef): PermissionResult {
    if (!this.serverRole) return { allowed: true };
    const hidden = this.accessIndex.intersecting(range).some((region) => region.access === 'hidden');
    return hidden
      ? { allowed: false, reason: 'This range contains cells hidden from the current subject', blockedBy: 'range-access' }
      : { allowed: true };
  }

  private checkRangeAccess(action: PermissionAction, ranges: readonly RangeRef[]): PermissionResult {
    if (!this.serverRole || ranges.length === 0 || action === 'navigate' || action === 'share') return { allowed: true };
    const required: RangeAccessLevel = action === 'comment' ? 'read' : 'edit';
    for (const range of ranges) {
      if (RANGE_ACCESS_RANK[this.effectiveAccess(range)] < RANGE_ACCESS_RANK[required]) {
        return { allowed: false, reason: required === 'read'
          ? 'This range is hidden from the current subject'
          : 'This range is read-only or hidden for the current subject', blockedBy: 'range-access' };
      }
    }
    return { allowed: true };
  }

  private effectiveAccess(range: RangeRef): RangeAccessLevel {
    const ceiling: RangeAccessLevel = this.serverRole === 'owner' || this.serverRole === 'editor' ? 'edit' : 'read';
    let effective: RangeAccessLevel = ceiling;
    for (const region of this.accessIndex.intersecting(range)) {
      const nextAccess: RangeAccessLevel = RANGE_ACCESS_RANK[region.access] < RANGE_ACCESS_RANK[effective] ? region.access : effective;
      effective = nextAccess;
    }
    return effective;
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
        // Worksheet extent is sparse and grows on demand. Protection must
        // validate a pending write against the canonical Excel bounds so a
        // legitimate paste/insert can grow the sheet before the mutation is
        // applied; using the current allocated extent rejects valid writes at
        // the edge of the visible grid.
        rowCount: MAX_SHEET_ROW_COUNT,
        columnCount: MAX_SHEET_COLUMN_COUNT,
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
