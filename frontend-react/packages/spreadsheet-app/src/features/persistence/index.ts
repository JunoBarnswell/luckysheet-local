import type { CommandRegistry, CommandResult } from '@react-sheets/command-runtime';

export interface PersistenceSaveParams {
  reason?: string;
  baseRevision?: number;
}

/** 显式保存为审计命令；实际 flush 由 Application + persistence-bridge 执行 */
export function registerPersistenceCommands(registry: CommandRegistry): void {
  registry.registerCommand({
    id: 'persistence.save',
    execute(params: PersistenceSaveParams, context): CommandResult {
      void params;
      return {
        operationId: context.operationId,
        mutationCount: 0,
        affectedRanges: [],
      };
    },
  });

  registry.registerCommand({
    id: 'persistence.draft.clear',
    execute(_params, context): CommandResult {
      return {
        operationId: context.operationId,
        mutationCount: 0,
        affectedRanges: [],
      };
    },
  });
}
