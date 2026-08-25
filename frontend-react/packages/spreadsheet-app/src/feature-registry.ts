import type { CommandRuntime } from '@react-sheets/command-runtime';
import { registerSheetCommands } from '@react-sheets/sheet-features';
import { registerPlatformFeatures } from './platform-features';
import { registerChartFeature } from './features/chart';
import { registerDrawingFeature, type DrawingRuntime } from './features/drawing';
import { registerDataSourceFeature } from './features/data-source';
import { registerEditingFeatures } from './features/editing';
import { registerPivotFeature } from './features/pivot';
import { registerPivotControlFeature } from './features/pivot-controls';
import { registerReviewFeature } from './features/review/commands';
import { registerSparklineFeature } from './features/sparkline';
import { registerInsertCommands } from './features/insert';
import { registerFindReplaceFeature } from './features/find-replace/commands';

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
  const pivotControlManifest = registerPivotControlFeature(runtime);
  const dataSourceManifest = registerDataSourceFeature(runtime);
  const chartManifest = registerChartFeature(runtime);
  const pivotManifest = registerPivotFeature(runtime);
  const sparklineManifest = registerSparklineFeature(runtime);
  const reviewManifest = registerReviewFeature(runtime);
  const findReplaceManifest = registerFindReplaceFeature(runtime);
  const insertCommandIds = registerInsertCommands(runtime);
  const platformCommandIds = registerPlatformFeatures(runtime);

  CORE_MANIFEST.commandIds = [...runtime.registry.listCommandIds()];
  registeredManifests = [
    CORE_MANIFEST,
    drawingManifest,
    pivotControlManifest,
    dataSourceManifest,
    chartManifest,
    pivotManifest,
    sparklineManifest,
    reviewManifest,
    findReplaceManifest,
    { id: 'insert', version: '1.0.0', commandIds: insertCommandIds, permissions: ['sheet.structure.write', 'sheet.format.write'] },
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
