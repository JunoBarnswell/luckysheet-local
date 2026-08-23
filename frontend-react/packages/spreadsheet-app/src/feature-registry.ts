import type { CommandRuntime } from '@react-sheets/command-runtime';
import { registerSheetCommands } from '@react-sheets/sheet-features';
import { registerPlatformFeatures } from './platform-features';
import { registerChartFeature } from './features/chart';
import { registerDrawingFeature, type DrawingRuntime } from './features/drawing';
import { registerEditingFeatures } from './features/editing';
import { registerPivotFeature } from './features/pivot';
import { registerReviewFeature } from './features/review/commands';
import { registerSparklineFeature } from './features/sparkline';

export interface SpreadsheetFeatureManifest {
  id: string;
  version: string;
  dependencies?: string[];
  commandIds: string[];
  mutationIds?: string[];
  ribbon?: ReadonlyArray<{ id: string; tab: string; group: string; label: string; commandId: string; icon: string }>;
  contextualTabs?: ReadonlyArray<{ id: string; tab: string; group: string; label: string; commandId: string; icon: string }>;
  permissions?: string[];
}

const CORE_MANIFEST: SpreadsheetFeatureManifest = {
  id: 'sheet-features',
  version: '1.0.0',
  commandIds: [],
};

let registeredManifests: SpreadsheetFeatureManifest[] = [CORE_MANIFEST];

/**
 * Register every spreadsheet feature against one CommandRuntime.
 *
 * Chart, Pivot and Sparkline own their command and mutation registrations in
 * spreadsheet-app. There is deliberately no compatibility namespace or
 * command forwarding layer: callers must use the canonical feature command.
 */
export function registerSpreadsheetFeatures(runtime: CommandRuntime, drawingRuntime: DrawingRuntime): SpreadsheetFeatureManifest[] {
  registerSheetCommands(runtime);
  registerEditingFeatures(runtime);

  const drawingManifest = registerDrawingFeature(runtime, drawingRuntime);
  const chartManifest = registerChartFeature(runtime);
  const pivotManifest = registerPivotFeature(runtime);
  const sparklineManifest = registerSparklineFeature(runtime);
  const reviewManifest = registerReviewFeature(runtime);
  const platformCommandIds = registerPlatformFeatures(runtime);

  CORE_MANIFEST.commandIds = [...runtime.registry.listCommandIds()];
  registeredManifests = [
    CORE_MANIFEST,
    drawingManifest,
    chartManifest,
    pivotManifest,
    sparklineManifest,
    reviewManifest,
    {
      id: 'platform',
      version: '1.0.0',
      commandIds: platformCommandIds,
      permissions: ['history.restore', 'persistence.write', 'print.export', 'xlsx.exchange', 'query.execute', 'automation.run'],
    },
  ];
  return registeredManifests;
}

export function getFeatureRegistry(): SpreadsheetFeatureManifest[] {
  return registeredManifests;
}
