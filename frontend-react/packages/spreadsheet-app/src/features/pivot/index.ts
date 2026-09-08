import type { CommandRuntime } from '@react-sheets/command-runtime';
import type { SpreadsheetFeatureManifest } from '../../feature-registry';
import { PIVOT_MUTATION_IDS, registerPivotCommands } from './commands';

export * from './commands';
export * from './helpers';
export * from './recommendation';
export * from './engine';
export * from './panel-state';
export * from './writeback';
export * from './rust-analytics';
export * from './server-task-port';

export function registerPivotFeature(runtime: CommandRuntime): SpreadsheetFeatureManifest {
  return {
    id: 'pivot',
    version: '1.0.0',
    dependencies: ['sheet-features'],
    commandIds: registerPivotCommands(runtime),
    mutationIds: [...PIVOT_MUTATION_IDS],
    permissions: ['pivot.edit', 'pivot.refresh', 'pivot.delete'],
  };
}
