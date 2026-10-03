import type { ProtectionRule, WorksheetModel } from './index';
import { PROTECTION_ACTION_ALLOW_FIELD } from './generated-protection';

const object = (value: unknown): value is Record<string, unknown> => value !== null && typeof value === 'object' && !Array.isArray(value);
const identity = (value: unknown): value is string => typeof value === 'string' && value.length > 0 && value === value.trim();

/** Canonical rule shape. Model-dependent ownership is checked separately. */
export function protectionRuleValidationError(value: unknown): string | undefined {
  if (!object(value)) return 'object';
  if (Object.keys(value).some(key => !['id', 'scope', 'sheetId', 'range', 'passwordHash', 'locked', 'allow'].includes(key))) return 'unknown-field';
  if (!identity(value.id)) return 'id';
  if (typeof value.scope !== 'string' || !['workbook', 'sheet', 'range'].includes(value.scope)) return 'scope';
  if (typeof value.locked !== 'boolean') return 'locked';
  if (value.sheetId !== undefined && !identity(value.sheetId)) return 'sheetId';
  if (value.passwordHash !== undefined && (typeof value.passwordHash !== 'string' || !value.passwordHash || value.passwordHash.length > 512)) return 'passwordHash';
  if (!object(value.allow)) return 'allow';
  const allowed = new Set<string>(Object.values(PROTECTION_ACTION_ALLOW_FIELD));
  if (Object.entries(value.allow).some(([key, flag]) => !allowed.has(key) || typeof flag !== 'boolean')) return 'allow';
  if (value.scope !== 'range') return value.range === undefined ? undefined : 'range';
  const range = value.range;
  if (!object(range) || Object.keys(range).some(key => !['sheetId', 'startRow', 'endRow', 'startColumn', 'endColumn'].includes(key))
    || !identity(range.sheetId) || value.sheetId !== undefined && range.sheetId !== value.sheetId) return 'range';
  for (const field of ['startRow', 'endRow', 'startColumn', 'endColumn']) if (!Number.isSafeInteger(range[field]) || Number(range[field]) < 0) return 'range';
  if (Number(range.endRow) < Number(range.startRow) || Number(range.endColumn) < Number(range.startColumn)
    || Number(range.endRow) > 1_048_575 || Number(range.endColumn) > 16_383) return 'range';
  return undefined;
}

export function assertWorksheetProtectionRule(sheet: WorksheetModel, rule: ProtectionRule): void {
  const error = protectionRuleValidationError(rule);
  if (error !== undefined) throw new Error(`PROTECTION_RULE_INVALID: ${error}`);
  if (rule.scope === 'workbook') throw new Error('UNSUPPORTED_FEATURE: workbook protection requires its workbook owner');
  if (rule.sheetId !== undefined && rule.sheetId !== sheet.id
    || rule.range && (rule.range.sheetId !== sheet.id || rule.range.endRow >= sheet.rowCount || rule.range.endColumn >= sheet.columnCount)) {
    throw new Error('PROTECTION_RULE_INVALID: rule does not belong to this worksheet extent');
  }
}
