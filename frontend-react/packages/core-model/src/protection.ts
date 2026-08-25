import type { CellStyle, RangeRef } from './index';
import type { ProtectionRule } from './domain';

/**
 * Canonical worksheet-protection actions.  These names are deliberately
 * independent of command ids so every frontend command and server reducer
 * can map to the same native worksheet allow flag.
 */
export type ProtectionAction =
  | 'edit-cell'
  | 'format'
  | 'insert-rows'
  | 'insert-columns'
  | 'delete-rows'
  | 'delete-columns'
  | 'sort'
  | 'auto-filter'
  | 'edit-objects'
  | 'select-locked'
  | 'select-unlocked';

export interface ProtectionCellResolution {
  locked: boolean;
  formulaHidden: boolean;
  active: boolean;
  rules: readonly ProtectionRule[];
}

export interface ProtectionDecision {
  allowed: boolean;
  reason?: string;
  blockedBy?: ProtectionRule;
  lockedCells: number;
  unlockedCells: number;
}

export interface ProtectionResolveRequest {
  sheetId: string;
  rules: readonly ProtectionRule[];
  ranges: readonly RangeRef[];
  action: ProtectionAction;
  rowCount: number;
  columnCount: number;
  readCellStyle?: (row: number, column: number) => Pick<CellStyle, 'locked' | 'formulaHidden'> | undefined;
}

const ACTION_ALLOW_FIELD: Readonly<Partial<Record<ProtectionAction, keyof ProtectionRule['allow']>>> = {
  format: 'formatCells',
  'insert-rows': 'insertRows',
  'insert-columns': 'insertColumns',
  'delete-rows': 'deleteRows',
  'delete-columns': 'deleteColumns',
  sort: 'sort',
  'auto-filter': 'autoFilter',
  'edit-objects': 'editObjects',
  'select-locked': 'selectLocked',
  'select-unlocked': 'selectUnlocked',
};

function rangesIntersect(left: RangeRef, right: RangeRef): boolean {
  return left.sheetId === right.sheetId
    && left.startRow <= right.endRow
    && left.endRow >= right.startRow
    && left.startColumn <= right.endColumn
    && left.endColumn >= right.startColumn;
}

function contains(range: RangeRef, sheetId: string, row: number, column: number): boolean {
  return range.sheetId === sheetId
    && row >= range.startRow && row <= range.endRow
    && column >= range.startColumn && column <= range.endColumn;
}

function ruleCoversCell(rule: ProtectionRule, sheetId: string, row: number, column: number): boolean {
  if (!rule.locked) return false;
  if (rule.scope === 'workbook') return true;
  if (rule.scope === 'sheet') return rule.sheetId === sheetId;
  return rule.range ? contains(rule.range, sheetId, row, column) : false;
}

function ruleCoversRange(rule: ProtectionRule, range: RangeRef): boolean {
  if (!rule.locked) return false;
  if (rule.scope === 'workbook') return true;
  if (rule.scope === 'sheet') return rule.sheetId === range.sheetId;
  return rule.range ? rangesIntersect(rule.range, range) : false;
}

function activeSheetRule(rules: readonly ProtectionRule[], sheetId: string): ProtectionRule | undefined {
  return rules.find((rule) => rule.locked && (rule.scope === 'workbook' || (rule.scope === 'sheet' && rule.sheetId === sheetId)));
}

function boundedRange(range: RangeRef, rowCount: number, columnCount: number): RangeRef | undefined {
  const startRow = Math.max(0, range.startRow);
  const startColumn = Math.max(0, range.startColumn);
  const endRow = Math.min(rowCount - 1, range.endRow);
  const endColumn = Math.min(columnCount - 1, range.endColumn);
  if (endRow < startRow || endColumn < startColumn) return undefined;
  return { ...range, startRow, endRow, startColumn, endColumn };
}

/**
 * Sole resolver for worksheet protection semantics.  It applies the native
 * worksheet allow flags first and only then evaluates per-cell Locked/Hidden
 * state.  Missing style.locked means Excel's default locked=true whenever a
 * sheet/workbook protection rule is active.
 */
export class ProtectionResolver {
  resolveCell(
    rules: readonly ProtectionRule[],
    sheetId: string,
    row: number,
    column: number,
    style?: Pick<CellStyle, 'locked' | 'formulaHidden'>,
  ): ProtectionCellResolution {
    const matching = rules.filter((rule) => ruleCoversCell(rule, sheetId, row, column));
    const active = activeSheetRule(rules, sheetId) !== undefined;
    const rangeLock = matching.some((rule) => rule.scope === 'range' && rule.locked);
    const locked = rangeLock || (active && style?.locked !== false);
    const formulaHidden = active && style?.formulaHidden === true;
    return { locked, formulaHidden, active, rules: matching };
  }

  isFormulaHidden(
    rules: readonly ProtectionRule[],
    sheetId: string,
    row: number,
    column: number,
    style?: Pick<CellStyle, 'locked' | 'formulaHidden'>,
  ): boolean {
    return this.resolveCell(rules, sheetId, row, column, style).formulaHidden;
  }

  resolve(request: ProtectionResolveRequest): ProtectionDecision {
    if (!Number.isSafeInteger(request.rowCount) || request.rowCount <= 0
      || !Number.isSafeInteger(request.columnCount) || request.columnCount <= 0) {
      return { allowed: false, reason: 'Protection resolution requires canonical worksheet bounds', lockedCells: 0, unlockedCells: 0 };
    }
    const ranges = request.ranges.map((range) => boundedRange(range, request.rowCount, request.columnCount)).filter((range): range is RangeRef => Boolean(range));
    if (ranges.length === 0) return { allowed: true, lockedCells: 0, unlockedCells: 0 };

    const allowField = ACTION_ALLOW_FIELD[request.action];
    if (allowField) {
      // Native worksheet flags are operation-level permissions.  A false or
      // missing flag is a denial; an explicit true flag allows the operation
      // even when the target cells themselves are locked (Excel semantics).
      for (const rule of request.rules) {
        if (ranges.some((range) => ruleCoversRange(rule, range)) && rule.allow[allowField] !== true) {
          return {
            allowed: false,
            reason: `Protected worksheet blocks ${request.action}`,
            blockedBy: rule,
            lockedCells: 0,
            unlockedCells: 0,
          };
        }
      }
      return { allowed: true, lockedCells: 0, unlockedCells: 0 };
    }

    if (request.action === 'edit-cell') {
      let lockedCells = 0;
      let unlockedCells = 0;
      for (const range of ranges) {
        for (let row = range.startRow; row <= range.endRow; row += 1) {
          for (let column = range.startColumn; column <= range.endColumn; column += 1) {
            const style = request.readCellStyle?.(row, column);
            const resolution = this.resolveCell(request.rules, request.sheetId, row, column, style);
            if (resolution.locked) {
              lockedCells += 1;
              return {
                allowed: false,
                reason: `Protected cell ${request.sheetId}!${row}:${column} is locked`,
                blockedBy: resolution.rules.find((rule) => rule.scope === 'range') ?? resolution.rules[0],
                lockedCells,
                unlockedCells,
              };
            }
            unlockedCells += 1;
          }
        }
      }
      return { allowed: true, lockedCells, unlockedCells };
    }

    return { allowed: false, reason: `Unsupported protection action ${request.action}`, lockedCells: 0, unlockedCells: 0 };
  }
}

export const protectionResolver = new ProtectionResolver();
