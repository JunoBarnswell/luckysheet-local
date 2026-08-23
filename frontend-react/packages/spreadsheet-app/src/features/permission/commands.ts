import type { CommandRuntime } from '@react-sheets/command-runtime';
import type { ProtectionRule } from '@react-sheets/core-model';
import type { SpreadsheetFeatureManifest } from '../../feature-registry';
import { applyTrackedMutation, registerMutationHandler, removeById } from '../../command-helpers';

export interface ProtectSetParams {
  sheetId: string;
  rule: ProtectionRule;
}

export interface ProtectRemoveParams {
  sheetId: string;
  ruleId: string;
}

function sheetWideRange(sheetId: string) {
  return [{ sheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }];
}

export function registerPermissionCommands(runtime: CommandRuntime): string[] {
  const commandIds: string[] = [];

  registerMutationHandler<ProtectSetParams>(runtime, 'sheet.protect.set', (params, context) => {
    context.workbook.getSheet(params.sheetId).protectionRules.push(structuredClone(params.rule));
  });
  registerMutationHandler<ProtectRemoveParams>(runtime, 'sheet.protect.remove', (params, context) => {
    removeById(context.workbook.getSheet(params.sheetId).protectionRules, params.ruleId);
  });

  const registerSimple = <P extends { sheetId: string }>(
    commandId: string,
    mutationId: string,
    readPrevious: (params: P, context: import('@react-sheets/command-runtime').CommandContext) => unknown,
  ): void => {
    runtime.registry.registerCommand<P>({
      id: commandId,
      execute: (params, context) => {
        const previous = readPrevious(params, context);
        applyTrackedMutation(context, {
          id: mutationId,
          sheetId: params.sheetId,
          params,
          inverseParams: previous ?? params,
          affectedRanges: sheetWideRange(params.sheetId),
          apply: () => runtime.registry.getMutation(mutationId)({ id: mutationId, unitId: context.workbook.unitId, sheetId: params.sheetId, params, affectedRanges: sheetWideRange(params.sheetId) }, context),
        });
        return { operationId: context.operationId, mutationCount: 1, affectedRanges: sheetWideRange(params.sheetId) };
      },
    });
    commandIds.push(commandId);
  };

  registerSimple<ProtectSetParams>('sheet.protect.set', 'sheet.protect.set', () => undefined);
  registerSimple<ProtectRemoveParams>('sheet.protect.remove', 'sheet.protect.remove', (params, context) => {
    const rule = context.workbook.getSheet(params.sheetId).protectionRules.find((entry) => entry.id === params.ruleId);
    return rule ? { sheetId: params.sheetId, rule: structuredClone(rule) } : params;
  });

  return commandIds;
}

export const PERMISSION_RIBBON_ENTRIES = [
  { id: 'permission-protect', tab: 'Review', group: 'Protection', label: 'Protect Selection', commandId: 'permission.protect.selection', icon: 'lock' },
  { id: 'permission-unprotect', tab: 'Review', group: 'Protection', label: 'Unprotect Selection', commandId: 'permission.unprotect.selection', icon: 'lock' },
] as const;

export function registerPermissionFeature(runtime: CommandRuntime): SpreadsheetFeatureManifest {
  const commandIds = registerPermissionCommands(runtime);
  return {
    id: 'permission',
    version: '1.0.0',
    commandIds,
    mutationIds: ['sheet.protect.set', 'sheet.protect.remove'],
    ribbon: [...PERMISSION_RIBBON_ENTRIES],
    permissions: ['workbook.protect', 'workbook.share'],
  };
}
