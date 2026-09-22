import type { CommandRuntime } from '@react-sheets/command-runtime';
import type { SpreadsheetFeatureManifest } from '../../feature-registry';
import { registerSparklineCommands, SPARKLINE_MUTATION_IDS } from './commands';

export * from './commands';
export * from './helpers';

export function registerSparklineFeature(runtime: CommandRuntime): SpreadsheetFeatureManifest {
  return {
    id: 'sparkline',
    version: '1.0.0',
    dependencies: ['sheet-features'],
    commandIds: registerSparklineCommands(runtime),
    mutationIds: [...SPARKLINE_MUTATION_IDS],
    contextualTabs: [
      { id: 'sparkline-design', tab: 'sparklineDesign', group: 'Sparkline', label: 'Design', commandId: 'sparkline.group.create', icon: 'sparkline' },
      { id: 'sparkline-remove', tab: 'sparklineDesign', group: 'Sparkline', label: 'Remove', commandId: 'sparkline.remove', icon: 'sparkline' },
    ],
    permissions: ['sparkline.edit', 'sparkline.delete'],
  };
}
