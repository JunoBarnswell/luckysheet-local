import type { CommandContext, CommandRuntime } from '@react-sheets/command-runtime';
import type { ProtectionRule, RangeRef } from '@react-sheets/core-model';
import type { SpreadsheetFeatureManifest } from '../../feature-registry';

export interface ProtectSetParams {
  sheetId: string;
  rule: ProtectionRule;
}

export interface ProtectRemoveParams {
  sheetId: string;
  ruleId: string;
}

function sheetWideRange(sheetId: string): RangeRef[] {
  return [{ sheetId, startRow: 0, endRow: Number.MAX_SAFE_INTEGER, startColumn: 0, endColumn: Number.MAX_SAFE_INTEGER }];
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function isProtectSetParams(value: unknown): value is ProtectSetParams {
  return isRecord(value)
    && typeof value.sheetId === 'string'
    && isRecord(value.rule)
    && typeof value.rule.id === 'string'
    && typeof value.rule.scope === 'string';
}

function isProtectRemoveParams(value: unknown): value is ProtectRemoveParams {
  return isRecord(value) && typeof value.sheetId === 'string' && typeof value.ruleId === 'string';
}

function applyProtectSet(context: CommandContext, params: ProtectSetParams, inverse: ProtectSetParams | ProtectRemoveParams): void {
  const affectedRanges = params.rule.scope === 'range' && params.rule.range ? [params.rule.range] : sheetWideRange(params.sheetId);
  const apply = () => {
    const rules = context.workbook.getSheet(params.sheetId).protectionRules;
    const index = rules.findIndex((entry) => entry.id === params.rule.id);
    if (index >= 0) rules[index] = structuredClone(params.rule);
    else rules.push(structuredClone(params.rule));
  };
  if ('rule' in inverse) {
    const inverseRanges = inverse.rule.scope === 'range' && inverse.rule.range ? [inverse.rule.range] : sheetWideRange(params.sheetId);
    context.applyMutation({
      id: 'sheet.protect.set', unitId: context.workbook.unitId, sheetId: params.sheetId, params, affectedRanges,
      inverse: [{ id: 'sheet.protect.set', unitId: context.workbook.unitId, sheetId: params.sheetId, params: inverse, affectedRanges: inverseRanges }], apply,
    });
    return;
  }
  context.applyMutation({
    id: 'sheet.protect.set', unitId: context.workbook.unitId, sheetId: params.sheetId, params, affectedRanges,
    inverse: [{ id: 'sheet.protect.remove', unitId: context.workbook.unitId, sheetId: params.sheetId, params: inverse, affectedRanges: sheetWideRange(params.sheetId) }], apply,
  });
}

function applyProtectRemove(context: CommandContext, params: ProtectRemoveParams, inverse: ProtectSetParams | ProtectRemoveParams): void {
  const affectedRanges = sheetWideRange(params.sheetId);
  const apply = () => {
    const rules = context.workbook.getSheet(params.sheetId).protectionRules;
    const index = rules.findIndex((entry) => entry.id === params.ruleId);
    if (index >= 0) rules.splice(index, 1);
  };
  if ('rule' in inverse) {
    const inverseRanges = inverse.rule.scope === 'range' && inverse.rule.range ? [inverse.rule.range] : affectedRanges;
    context.applyMutation({
      id: 'sheet.protect.remove', unitId: context.workbook.unitId, sheetId: params.sheetId, params, affectedRanges,
      inverse: [{ id: 'sheet.protect.set', unitId: context.workbook.unitId, sheetId: params.sheetId, params: inverse, affectedRanges: inverseRanges }], apply,
    });
    return;
  }
  context.applyMutation({
    id: 'sheet.protect.remove', unitId: context.workbook.unitId, sheetId: params.sheetId, params, affectedRanges,
    inverse: [{ id: 'sheet.protect.remove', unitId: context.workbook.unitId, sheetId: params.sheetId, params: inverse, affectedRanges }], apply,
  });
}

export function registerPermissionCommands(runtime: CommandRuntime): string[] {
  runtime.registry.registerMutation<ProtectSetParams>({
    id: 'sheet.protect.set',
    handler: (item, context) => {
      const params = item.params;
      const rules = context.workbook.getSheet(params.sheetId).protectionRules;
      const index = rules.findIndex((entry) => entry.id === params.rule.id);
      if (index >= 0) rules[index] = structuredClone(params.rule);
      else rules.push(structuredClone(params.rule));
    },
    metadata: {
      schema: { name: 'ProtectSetParams', validate: isProtectSetParams },
      permission: { capability: 'workbook.protect', roles: ['owner'] },
      affectedRanges: { resolve: (params) => params.rule.scope === 'range' && params.rule.range ? [params.rule.range] : sheetWideRange(params.sheetId), mode: 'exact' },
      inversePolicy: { allowedMutationIds: ['sheet.protect.set', 'sheet.protect.remove'], minCount: 1, maxCount: 1 },
    },
  });
  runtime.registry.registerMutation<ProtectRemoveParams>({
    id: 'sheet.protect.remove',
    handler: (item, context) => {
      const rules = context.workbook.getSheet(item.params.sheetId).protectionRules;
      const index = rules.findIndex((entry) => entry.id === item.params.ruleId);
      if (index >= 0) rules.splice(index, 1);
    },
    metadata: {
      schema: { name: 'ProtectRemoveParams', validate: isProtectRemoveParams },
      permission: { capability: 'workbook.protect', roles: ['owner'] },
      affectedRanges: { resolve: (params) => sheetWideRange(params.sheetId), mode: 'exact' },
      inversePolicy: { allowedMutationIds: ['sheet.protect.set', 'sheet.protect.remove'], minCount: 1, maxCount: 1 },
    },
  });

  runtime.registry.registerCommand<ProtectSetParams>({
    id: 'sheet.protect.set',
    execute: (params, context) => {
      const previous = context.workbook.getSheet(params.sheetId).protectionRules.find((entry) => entry.id === params.rule.id);
      applyProtectSet(context, params, previous ? { sheetId: params.sheetId, rule: structuredClone(previous) } : { sheetId: params.sheetId, ruleId: params.rule.id });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: params.rule.scope === 'range' && params.rule.range ? [params.rule.range] : sheetWideRange(params.sheetId) };
    },
  });
  runtime.registry.registerCommand<ProtectRemoveParams>({
    id: 'sheet.protect.remove',
    execute: (params, context) => {
      const previous = context.workbook.getSheet(params.sheetId).protectionRules.find((entry) => entry.id === params.ruleId);
      applyProtectRemove(context, params, previous ? { sheetId: params.sheetId, rule: structuredClone(previous) } : params);
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: sheetWideRange(params.sheetId) };
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
