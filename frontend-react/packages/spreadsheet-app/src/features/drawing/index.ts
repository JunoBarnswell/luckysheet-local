import type { CommandRuntime } from '@react-sheets/command-runtime';
import type { SpreadsheetFeatureManifest } from '../../feature-registry';
import { DrawingRuntime } from './runtime';
import { DRAWING_MUTATION_IDS, registerDrawingCommands } from './commands';

export * from './runtime';
export * from './commands';
export * from './geometry';
export * from './chart-payload';

export function registerDrawingFeature(runtime: CommandRuntime, runtimeState = new DrawingRuntime()): SpreadsheetFeatureManifest {
  const commandIds = registerDrawingCommands(runtime, runtimeState);
  return {
    id: 'drawing',
    version: '1.0.0',
    commandIds,
    mutationIds: [...DRAWING_MUTATION_IDS],
    permissions: ['drawing.edit', 'drawing.arrange'],
  };
}
