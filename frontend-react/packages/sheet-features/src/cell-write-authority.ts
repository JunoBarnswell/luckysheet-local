import type { CellData, WorksheetModel } from '@react-sheets/core-model';
import { validateDataInput } from './data-features';

export type CellWriteKind = 'direct-entry' | 'paste' | 'fill' | 'formula-result' | 'query-load' | 'script' | 'external-sync';

export interface CellWriteAuthority {
  kind: CellWriteKind;
  target: { sheetId: string; row: number; column: number };
  candidate: CellData;
  validationDecision: {
    status: 'accepted' | 'confirmed';
    ruleId?: string;
    alertStyle?: 'stop' | 'warning' | 'information';
  };
}

export interface CellSetMutationParams {
  sheetId: string;
  row: number;
  column: number;
  value: CellData;
  writeAuthority: CellWriteAuthority;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null;
}

function isCellData(value: unknown): value is CellData {
  return isRecord(value) && ('value' in value || typeof value.formula === 'string');
}

function sameCanonicalValue(left: unknown, right: unknown): boolean {
  if (Object.is(left, right)) return true;
  if (Array.isArray(left) || Array.isArray(right)) {
    return Array.isArray(left) && Array.isArray(right) && left.length === right.length
      && left.every((item, index) => sameCanonicalValue(item, right[index]));
  }
  if (!isRecord(left) || !isRecord(right)) return false;
  const leftKeys = Object.keys(left).sort();
  const rightKeys = Object.keys(right).sort();
  return leftKeys.length === rightKeys.length
    && leftKeys.every((key, index) => key === rightKeys[index] && sameCanonicalValue(left[key], right[key]));
}

export function createCellWriteAuthority(
  sheet: WorksheetModel,
  params: { sheetId: string; row: number; column: number; value: CellData },
  kind: CellWriteKind,
  confirmed = false,
): CellWriteAuthority {
  const validation = params.value.formula
    ? { valid: true, blocking: false, ruleId: undefined, alertStyle: undefined }
    : validateDataInput(sheet, params.row, params.column, params.value.value);
  if (validation.blocking) throw new Error(validation.message ?? 'Cell value failed data validation');
  if (!validation.valid && !confirmed) {
    throw new Error('CELL_ENTRY_CONFIRMATION_REQUIRED: warning/information validation requires explicit confirmation');
  }
  return {
    kind,
    target: { sheetId: params.sheetId, row: params.row, column: params.column },
    candidate: structuredClone(params.value),
    validationDecision: {
      status: validation.valid ? 'accepted' : 'confirmed',
      ...(validation.ruleId ? { ruleId: validation.ruleId } : {}),
      ...(validation.alertStyle ? { alertStyle: validation.alertStyle } : {}),
    },
  };
}

export function createCellSetMutationParams(
  sheet: WorksheetModel,
  params: { sheetId: string; row: number; column: number; value: CellData },
  kind: CellWriteKind,
  confirmed = false,
): CellSetMutationParams {
  return { ...params, writeAuthority: createCellWriteAuthority(sheet, params, kind, confirmed) };
}

export function isCellSetMutationParams(value: unknown): value is CellSetMutationParams {
  if (!isRecord(value) || typeof value.sheetId !== 'string'
    || !Number.isSafeInteger(value.row) || Number(value.row) < 0
    || !Number.isSafeInteger(value.column) || Number(value.column) < 0
    || !isCellData(value.value) || !isRecord(value.writeAuthority)) return false;
  const authority = value.writeAuthority;
  return ['direct-entry', 'paste', 'fill', 'formula-result', 'query-load', 'script', 'external-sync'].includes(String(authority.kind))
    && isRecord(authority.target)
    && authority.target.sheetId === value.sheetId
    && authority.target.row === value.row
    && authority.target.column === value.column
    && isCellData(authority.candidate)
    && sameCanonicalValue(authority.candidate, value.value)
    && isRecord(authority.validationDecision)
    && ['accepted', 'confirmed'].includes(String(authority.validationDecision.status))
    && (authority.validationDecision.ruleId === undefined || typeof authority.validationDecision.ruleId === 'string')
    && (authority.validationDecision.alertStyle === undefined || ['stop', 'warning', 'information'].includes(String(authority.validationDecision.alertStyle)));
}

export function assertCellWriteAuthority(params: CellSetMutationParams, sheet: WorksheetModel): void {
  const expected = createCellWriteAuthority(
    sheet,
    { sheetId: params.sheetId, row: params.row, column: params.column, value: params.value },
    params.writeAuthority.kind,
    params.writeAuthority.validationDecision.status === 'confirmed',
  );
  if (!sameCanonicalValue(expected, params.writeAuthority)) {
    throw new Error('CELL_ENTRY_VALIDATION_STALE: cell write authority no longer matches the canonical worksheet state');
  }
}
