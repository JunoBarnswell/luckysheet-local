import type { CommandRegistry, CommandResult } from '@react-sheets/command-runtime';
import type { PrintAreaSetCommandParams, PrintPreviewCommandParams } from './layout';

export function registerPrintCommands(registry: CommandRegistry): void {
  registry.registerCommand({
    id: 'print.preview',
    execute(params: PrintPreviewCommandParams, context): CommandResult {
      void params;
      return {
        operationId: context.operationId,
        mutationCount: 0,
        affectedRanges: params.range ? [params.range] : [],
      };
    },
  });

  registry.registerCommand({
    id: 'print.export',
    execute(params: PrintPreviewCommandParams, context): CommandResult {
      void params;
      return {
        operationId: context.operationId,
        mutationCount: 0,
        affectedRanges: params.range ? [params.range] : [],
      };
    },
  });

  registry.registerCommand({
    id: 'print.area.set',
    execute(params: PrintAreaSetCommandParams, context): CommandResult {
      return {
        operationId: context.operationId,
        mutationCount: 0,
        affectedRanges: [params.range],
      };
    },
  });

  registry.registerCommand({
    id: 'print.pageSetup',
    execute(params: PrintPreviewCommandParams, context): CommandResult {
      void params;
      return {
        operationId: context.operationId,
        mutationCount: 0,
        affectedRanges: [],
      };
    },
  });
}
