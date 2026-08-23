import type { CommandRegistry, CommandResult } from '@react-sheets/command-runtime';

export interface XlsxImportCommandParams {
  fileName: string;
  base64: string;
}

export interface XlsxExportCommandParams {
  fileName?: string;
}

export function registerXlsxCommands(registry: CommandRegistry): void {
  registry.registerCommand({
    id: 'xlsx.import',
    execute(params: XlsxImportCommandParams, context): CommandResult {
      void params;
      return {
        operationId: context.operationId,
        mutationCount: 0,
        affectedRanges: [],
      };
    },
  });

  registry.registerCommand({
    id: 'xlsx.export',
    execute(params: XlsxExportCommandParams, context): CommandResult {
      void params;
      return {
        operationId: context.operationId,
        mutationCount: 0,
        affectedRanges: [],
      };
    },
  });
}
