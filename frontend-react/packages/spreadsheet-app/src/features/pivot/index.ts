import type { CommandRuntime } from '@react-sheets/command-runtime';
import type { SpreadsheetFeatureManifest } from '../../feature-registry';
import { PIVOT_MUTATION_IDS, registerPivotCommands } from './commands';

export * from './commands';
export * from './helpers';
export * from './engine';
export * from './block-source';
export * from './panel-state';
export * from './writeback';

export function registerPivotFeature(runtime: CommandRuntime): SpreadsheetFeatureManifest {
  return {
    id: 'pivot',
    version: '1.0.0',
    dependencies: ['sheet-features'],
    commandIds: registerPivotCommands(runtime),
    mutationIds: [...PIVOT_MUTATION_IDS],
    contextualTabs: [
      { id: 'pivot-analyze', tab: 'Analyze', group: 'Pivot', label: 'Refresh', commandId: 'pivot.refresh', icon: 'pivot' },
      { id: 'pivot-remove', tab: 'Analyze', group: 'Pivot', label: 'Remove', commandId: 'pivot.remove', icon: 'pivot' },
    ],
    permissions: ['pivot.edit', 'pivot.refresh', 'pivot.delete'],
  };
}
