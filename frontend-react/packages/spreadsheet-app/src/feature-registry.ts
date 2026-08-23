import type { CommandRuntime } from '@react-sheets/command-runtime';
import { registerSheetCommands } from '@react-sheets/sheet-features';
import { registerProSheetCommands } from '@react-sheets/pro-features';
import { registerPlatformFeatures } from './platform-features';
import { registerDrawingFeature, type DrawingRuntime } from './features/drawing';
import { registerReviewFeature } from './features/review/commands';
import { registerEditingFeatures } from './features/editing';

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

const PRO_MANIFEST: SpreadsheetFeatureManifest = {
  id: 'pro-features',
  version: '1.0.0',
  dependencies: ['sheet-features'],
  commandIds: [],
};

/** Pro command IDs registered without the `pro.` prefix */
const PRO_COMMAND_ALIASES: Array<[alias: string, source: string]> = [
  ['chart.add', 'pro.chart.add'],
  ['chart.move', 'pro.chart.move'],
  ['pivot.add', 'pro.pivot.add'],
  ['pivot.update', 'pro.pivot.update'],
  ['pivot.refresh', 'pro.pivot.refresh'],
  ['shape.add', 'pro.shape.add'],
  ['shape.move', 'pro.shape.move'],
  ['sparkline.add', 'pro.sparkline.add'],
  ['sparkline.insert', 'sparkline.insert'],
  ['chart.insert', 'chart.insert'],
  ['pivot.setAggregate', 'pivot.setAggregate'],
  ['pivot.setShowAs', 'pivot.setShowAs'],
  ['pivot.setGroup', 'pivot.setGroup'],
  ['pivot.drillDown', 'pivot.drillDown'],
  ['pivot.slicer.set', 'pivot.slicer.set'],
  ['pivot.timeline.set', 'pivot.timeline.set'],
];

function registerCommandAlias(runtime: CommandRuntime, aliasId: string, sourceId: string): void {
  if (runtime.registry.hasCommand(aliasId)) return;
  const source = runtime.registry.getCommand(sourceId);
  runtime.registry.registerCommand({
    id: aliasId,
    execute: (params, context) => source.execute(params, context),
  });
}

function registerRemoveCommand(
  runtime: CommandRuntime,
  commandId: string,
  mutationId: string,
  collection: (sheet: ReturnType<CommandRuntime['workbook']['getSheet']>) => Array<{ id: string }>,
): void {
  if (runtime.registry.hasCommand(commandId)) return;
  runtime.registry.registerCommand<string>({
    id: commandId,
    execute: (id, context) => {
      const sheetId = context.workbook.activeSheetId;
      const sheet = context.workbook.getSheet(sheetId);
      const items = collection(sheet);
      const index = items.findIndex((item) => item.id === id);
      if (index < 0) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const removed = structuredClone(items[index]!);
      const affectedRanges = [{ sheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }];
      context.applyMutation({
        id: mutationId,
        unitId: context.workbook.unitId,
        sheetId,
        params: id,
        affectedRanges,
        inverse: [{ id: `${mutationId.replace('.remove', '.add')}`, unitId: context.workbook.unitId, sheetId, params: removed, affectedRanges }],
        apply: () => {
          items.splice(index, 1);
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
}

export function registerSpreadsheetFeatures(runtime: CommandRuntime, drawingRuntime: DrawingRuntime): SpreadsheetFeatureManifest[] {
  registerSheetCommands(runtime);
  registerEditingFeatures(runtime);
  registerProSheetCommands(runtime);
  registerDrawingFeature(runtime, drawingRuntime);
  registerReviewFeature(runtime);

  for (const [alias, source] of PRO_COMMAND_ALIASES) {
    if (runtime.registry.hasCommand(source)) {
      registerCommandAlias(runtime, alias, source);
    }
  }

  const platformCommandIds = registerPlatformFeatures(runtime);
  void platformCommandIds;

  registerRemoveCommand(runtime, 'chart.remove', 'chart.remove', (sheet) => sheet.charts);
  registerRemoveCommand(runtime, 'pivot.remove', 'pivot.remove', (sheet) => sheet.pivots);
  registerRemoveCommand(runtime, 'shape.remove', 'shape.remove', (sheet) => sheet.shapes);
  registerRemoveCommand(runtime, 'sparkline.remove', 'sparkline.remove', (sheet) => sheet.sparklines);

  CORE_MANIFEST.commandIds = runtime.registry.listCommandIds().filter((id) => !id.startsWith('pro.'));
  PRO_MANIFEST.commandIds = [...PRO_COMMAND_ALIASES.map(([alias]) => alias), 'chart.remove', 'pivot.remove', 'shape.remove', 'sparkline.remove'];

  return [CORE_MANIFEST, PRO_MANIFEST];
}

export function getFeatureRegistry(): SpreadsheetFeatureManifest[] {
  return [CORE_MANIFEST, PRO_MANIFEST];
}
