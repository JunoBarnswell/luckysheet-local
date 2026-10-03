import type { CommandContext, CommandRuntime } from '@react-sheets/command-runtime';
import { assertWorksheetProtectionRule, protectionRuleValidationError, type ProtectionRule, type RangeRef } from '@react-sheets/core-model';
import type { SpreadsheetFeatureManifest } from '../../feature-registry';

export interface ProtectSetParams { sheetId: string; rule: ProtectionRule; }
export interface ProtectRemoveParams { sheetId: string; ruleId: string; }
const object = (value: unknown): value is Record<string, unknown> => value !== null && typeof value === 'object' && !Array.isArray(value);
const identity = (value: unknown): value is string => typeof value === 'string' && value.length > 0 && value === value.trim();
function isProtectSetParams(value: unknown): value is ProtectSetParams {
  return object(value) && Object.keys(value).every(key => ['sheetId', 'rule'].includes(key))
    && identity(value.sheetId) && protectionRuleValidationError(value.rule) === undefined;
}
function isProtectRemoveParams(value: unknown): value is ProtectRemoveParams {
  return object(value) && Object.keys(value).every(key => ['sheetId', 'ruleId'].includes(key)) && identity(value.sheetId) && identity(value.ruleId);
}
function rangesForRule(rule: ProtectionRule): RangeRef[] { return rule.scope === 'range' ? [structuredClone(rule.range!)] : []; }
function rules(context: CommandContext, sheetId: string): ProtectionRule[] {
  const sheet = context.workbook.getSheet(sheetId), seen = new Set<string>();
  for (const rule of sheet.protectionRules) {
    if (protectionRuleValidationError(rule) !== undefined || seen.has(rule.id)) throw new Error('PROTECTION_RULE_INVALID: persisted rule identities or shape are invalid');
    assertWorksheetProtectionRule(sheet, rule);
    seen.add(rule.id);
  }
  return sheet.protectionRules;
}
function setRule(context: CommandContext, params: ProtectSetParams): void {
  assertWorksheetProtectionRule(context.workbook.getSheet(params.sheetId), params.rule);
  const state = rules(context, params.sheetId), index = state.findIndex(rule => rule.id === params.rule.id);
  if (index >= 0) state[index] = structuredClone(params.rule); else state.push(structuredClone(params.rule));
}
function removeRule(context: CommandContext, params: ProtectRemoveParams): void {
  const state = rules(context, params.sheetId), index = state.findIndex(rule => rule.id === params.ruleId);
  if (index < 0) throw new Error(`PROTECTION_RULE_NOT_FOUND: ${params.ruleId}`);
  state.splice(index, 1);
}
export function registerPermissionCommands(runtime: CommandRuntime): string[] {
  runtime.registry.registerMutation<ProtectSetParams>({
    id: 'sheet.protect.set', handler: (item, context) => { if (!isProtectSetParams(item.params)) throw new Error('PROTECTION_RULE_INVALID: parameters'); setRule(context, item.params); },
    metadata: {
      calculation: { inputs: 'none', visibility: false, spillBlockers: 'none', mode: false },
      schema: { name: 'ProtectSetParams', validate: isProtectSetParams }, permission: { capability: 'workbook.protect', roles: ['owner'] },
      affectedRanges: { resolve: params => rangesForRule(params.rule), mode: 'exact' },
      inversePolicy: { allowedMutationIds: ['sheet.protect.set', 'sheet.protect.remove'], minCount: 1, maxCount: 1 },
    },
  });
  runtime.registry.registerMutation<ProtectRemoveParams>({
    id: 'sheet.protect.remove', handler: (item, context) => { if (!isProtectRemoveParams(item.params)) throw new Error('PROTECTION_RULE_INVALID: parameters'); removeRule(context, item.params); },
    metadata: {
      calculation: { inputs: 'none', visibility: false, spillBlockers: 'none', mode: false },
      schema: { name: 'ProtectRemoveParams', validate: isProtectRemoveParams }, permission: { capability: 'workbook.protect', roles: ['owner'] },
      affectedRanges: { resolve: () => [], mode: 'exact' },
      inversePolicy: { allowedMutationIds: ['sheet.protect.set'], minCount: 1, maxCount: 1 },
    },
  });
  runtime.registry.registerCommand<ProtectSetParams>({
    id: 'sheet.protect.set', execute: (params, context) => {
      if (!isProtectSetParams(params)) throw new Error('PROTECTION_RULE_INVALID: parameters');
      assertWorksheetProtectionRule(context.workbook.getSheet(params.sheetId), params.rule);
      const previous = rules(context, params.sheetId).find(rule => rule.id === params.rule.id);
      const affectedRanges = rangesForRule(params.rule);
      context.applyMutation({ id: 'sheet.protect.set', unitId: context.workbook.unitId, sheetId: params.sheetId, params: structuredClone(params), affectedRanges,
        inverse: [previous ? { id: 'sheet.protect.set', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, rule: structuredClone(previous) }, affectedRanges: rangesForRule(previous) }
          : { id: 'sheet.protect.remove', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, ruleId: params.rule.id }, affectedRanges: [] }],
        apply: () => setRule(context, params),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  runtime.registry.registerCommand<ProtectRemoveParams>({
    id: 'sheet.protect.remove', execute: (params, context) => {
      if (!isProtectRemoveParams(params)) throw new Error('PROTECTION_RULE_INVALID: parameters');
      const previous = rules(context, params.sheetId).find(rule => rule.id === params.ruleId);
      if (!previous) throw new Error(`PROTECTION_RULE_NOT_FOUND: ${params.ruleId}`);
      context.applyMutation({ id: 'sheet.protect.remove', unitId: context.workbook.unitId, sheetId: params.sheetId, params: structuredClone(params), affectedRanges: [],
        inverse: [{ id: 'sheet.protect.set', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, rule: structuredClone(previous) }, affectedRanges: rangesForRule(previous) }],
        apply: () => removeRule(context, params),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: [] };
    },
  });
  return ['sheet.protect.set', 'sheet.protect.remove'];
}

export function registerPermissionFeature(runtime: CommandRuntime): SpreadsheetFeatureManifest {
  const commandIds = registerPermissionCommands(runtime);
  return {
    id: 'permission',
    version: '1.0.0',
    commandIds,
    mutationIds: ['sheet.protect.set', 'sheet.protect.remove'],
    ribbon: [],
    permissions: ['workbook.protect', 'workbook.share'],
  };
}
