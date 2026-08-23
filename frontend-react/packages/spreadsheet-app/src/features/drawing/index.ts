import type { CommandRuntime } from '@react-sheets/command-runtime';
import type { SpreadsheetFeatureManifest } from '../../feature-registry';
import { DrawingRuntime } from './runtime';
import { DRAWING_MUTATION_IDS, registerDrawingCommands } from './commands';

export * from './runtime';
export * from './commands';

export function registerDrawingFeature(runtime: CommandRuntime, runtimeState = new DrawingRuntime()): SpreadsheetFeatureManifest {
  const commandIds = registerDrawingCommands(runtime, runtimeState);
  return {
    id: 'drawing',
    version: '1.0.0',
    commandIds,
    mutationIds: [...DRAWING_MUTATION_IDS],
    contextualTabs: [
      { id: 'drawing-format', tab: 'Format', group: 'Arrange', label: 'Bring Forward', commandId: 'drawing.zorder', icon: 'shape' },
      { id: 'drawing-remove', tab: 'Format', group: 'Arrange', label: 'Remove', commandId: 'drawing.remove', icon: 'shape' },
    ],
    permissions: ['drawing.edit', 'drawing.arrange'],
  };
}
