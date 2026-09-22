import type { CommandRegistry, CommandResult } from '@react-sheets/command-runtime';

export interface NativeDocumentImportCommandParams {
  fileName: string;
  buffer: ArrayBuffer;
}

export interface NativeDocumentExportCommandParams {
  fileName?: string;
}

export function registerNativeDocumentCommands(registry: CommandRegistry): void {
  registry.registerCommand({
    id: 'document.import',
    execute(params: NativeDocumentImportCommandParams, context): CommandResult {
      void params;
      return {
        operationId: context.operationId,
        mutationCount: 0,
        affectedRanges: [],
      };
    },
  });

  registry.registerCommand({
    id: 'document.export',
    execute(params: NativeDocumentExportCommandParams, context): CommandResult {
      void params;
      return {
        operationId: context.operationId,
        mutationCount: 0,
        affectedRanges: [],
      };
    },
  });
}

export * from './exchange';
