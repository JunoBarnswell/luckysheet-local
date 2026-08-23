import type { CommandRuntime } from '@react-sheets/command-runtime';
import type { SpreadsheetFeatureManifest } from '../../feature-registry';
import { CHART_MUTATION_IDS, registerChartCommands } from './commands';

export * from './commands';
export * from './data';

export function registerChartFeature(runtime: CommandRuntime): SpreadsheetFeatureManifest {
  return {
    id: 'chart',
    version: '1.0.0',
    dependencies: ['sheet-features', 'drawing'],
    commandIds: registerChartCommands(runtime),
    mutationIds: [...CHART_MUTATION_IDS],
    contextualTabs: [
      { id: 'chart-design', tab: 'Design', group: 'Chart', label: 'Chart Type', commandId: 'chart.setType', icon: 'chart' },
      { id: 'chart-remove', tab: 'Format', group: 'Arrange', label: 'Remove', commandId: 'chart.remove', icon: 'chart' },
    ],
    permissions: ['chart.edit', 'chart.delete'],
  };
}
