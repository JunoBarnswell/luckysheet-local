import type { CommandRuntime } from '@react-sheets/command-runtime';
import type { SpreadsheetFeatureManifest } from '../../feature-registry';
import { CHART_MUTATION_IDS, registerChartCommands } from './commands';

export * from './commands';
export * from './data';
export * from './recommendation';

export function registerChartFeature(runtime: CommandRuntime): SpreadsheetFeatureManifest {
  return {
    id: 'chart',
    version: '1.0.0',
    dependencies: ['sheet-features', 'drawing'],
    commandIds: registerChartCommands(runtime),
    mutationIds: [...CHART_MUTATION_IDS],
    contextualTabs: [
      { id: 'chart-design', tab: 'Design', group: 'Chart', label: 'Chart Elements', commandId: 'chart.setElements', icon: 'chart' },
      { id: 'chart-data', tab: 'Design', group: 'Data', label: 'Select Data', commandId: 'chart.setSeries', icon: 'table' },
      { id: 'chart-format', tab: 'Format', group: 'Chart Styles', label: 'Format Chart', commandId: 'chart.setElements', icon: 'sparkles' },
    ],
    permissions: ['chart.edit', 'chart.delete'],
  };
}
