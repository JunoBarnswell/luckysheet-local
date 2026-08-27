import type { CellStyle, RangeRef } from './index';
import type { ProtectionRule } from './domain';
import { PROTECTION_ACTION_ALLOW_FIELD, type ProtectionAction as GeneratedProtectionAction } from './generated-protection';

/**
 * Canonical worksheet-protection actions.  These names are deliberately
 * independent of command ids so every frontend command and server reducer
 * can map to the same native worksheet allow flag.
 */
export type ProtectionAction = GeneratedProtectionAction;

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
  /** Count explicit unlocked cells from the worksheet's sparse cell index. */
  countUnlockedCells?: (range: RangeRef) => number;
}

const ACTION_ALLOW_FIELD: Readonly<Partial<Record<ProtectionAction, keyof ProtectionRule['allow']>>> = PROTECTION_ACTION_ALLOW_FIELD;

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

function validateRange(range: RangeRef, sheetId: string, rowCount: number, columnCount: number): boolean {
  return range.sheetId === sheetId
    && Number.isSafeInteger(range.startRow) && Number.isSafeInteger(range.endRow)
    && Number.isSafeInteger(range.startColumn) && Number.isSafeInteger(range.endColumn)
    && range.startRow >= 0 && range.startColumn >= 0
    && range.endRow >= range.startRow && range.endColumn >= range.startColumn
    && range.endRow < rowCount && range.endColumn < columnCount;
}

function rangeCellCount(range: RangeRef): number | undefined {
  const rows = range.endRow - range.startRow + 1;
  const columns = range.endColumn - range.startColumn + 1;
  const count = rows * columns;
  return Number.isSafeInteger(count) ? count : undefined;
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
    if (request.ranges.some((range) => !validateRange(range, request.sheetId, request.rowCount, request.columnCount))) {
      return { allowed: false, reason: 'Protection resolution received a range outside canonical worksheet bounds', lockedCells: 0, unlockedCells: 0 };
    }
    const ranges = [...request.ranges];
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
        const count = rangeCellCount(range);
        if (count === undefined) {
          return { allowed: false, reason: 'Protection resolution cannot safely represent the requested cell extent', lockedCells, unlockedCells };
        }
        const rangeRule = request.rules.find((rule) => rule.scope === 'range' && ruleCoversRange(rule, range));
        if (rangeRule) {
          lockedCells += 1;
          return {
            allowed: false,
            reason: `Protected worksheet: protected range ${rangeRule.id} blocks cell editing`,
            blockedBy: rangeRule,
            lockedCells,
            unlockedCells,
          };
        }
        if (activeSheetRule(request.rules, request.sheetId)) {
          if (!request.countUnlockedCells) {
            return { allowed: false, reason: 'Protection resolution requires the canonical sparse cell index', lockedCells, unlockedCells };
          }
          const explicitUnlocked = request.countUnlockedCells(range);
          if (!Number.isSafeInteger(explicitUnlocked) || explicitUnlocked < 0 || explicitUnlocked > count) {
            return { allowed: false, reason: 'Protection resolution received an invalid sparse cell exception count', lockedCells, unlockedCells };
          }
          unlockedCells += explicitUnlocked;
          lockedCells += count - explicitUnlocked;
          if (explicitUnlocked !== count) {
            return {
              allowed: false,
              reason: `Protected worksheet: ${count - explicitUnlocked} cells in ${request.sheetId}!${range.startRow}:${range.startColumn}-${range.endRow}:${range.endColumn} remain locked`,
              blockedBy: activeSheetRule(request.rules, request.sheetId),
              lockedCells,
              unlockedCells,
            };
          }
        }
      }
      return { allowed: true, lockedCells, unlockedCells };
    }

    return { allowed: false, reason: `Unsupported protection action ${request.action}`, lockedCells: 0, unlockedCells: 0 };
  }
}

export const protectionResolver = new ProtectionResolver();
