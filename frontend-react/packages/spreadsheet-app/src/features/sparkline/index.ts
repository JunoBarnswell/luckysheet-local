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
    permissions: ['sparkline.edit', 'sparkline.delete'],
  };
}
