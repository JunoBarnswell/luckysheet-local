import type { CommandRegistry, CommandResult } from '@react-sheets/command-runtime';
import type { AutomationRunParams } from '../../automation-bridge';

export function registerAutomationCommands(registry: CommandRegistry): void {
  registry.registerCommand({
    id: 'automation.run',
    execute(params: AutomationRunParams, context): CommandResult {
      void params;
      return {
        operationId: context.operationId,
        mutationCount: 0,
        affectedRanges: [],
      };
    },
  });

  registry.registerCommand({
    id: 'automation.record.start',
    execute(_params, context): CommandResult {
      return {
        operationId: context.operationId,
        mutationCount: 0,
        affectedRanges: [],
      };
    },
  });

  registry.registerCommand({
    id: 'automation.record.stop',
    execute(_params, context): CommandResult {
      return {
        operationId: context.operationId,
        mutationCount: 0,
        affectedRanges: [],
      };
    },
  });
}
