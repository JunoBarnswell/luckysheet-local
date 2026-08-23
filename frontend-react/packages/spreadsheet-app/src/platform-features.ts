import type { CommandRuntime } from '@react-sheets/command-runtime';
import { registerHistoryCommands } from './features/history';
import { registerPersistenceCommands } from './features/persistence';
import { registerPrintCommands } from './features/print';
import { registerXlsxCommands } from './features/xlsx';
import { registerPermissionFeature } from './features/permission/commands';
import { registerQueryCommands } from './features/query';
import { registerAutomationCommands } from './features/automation';
import { registerExtendedCommands } from './features/extended';

/** 注册 M10–M18 平台特性命令 */
export function registerPlatformFeatures(runtime: CommandRuntime): string[] {
  registerHistoryCommands(runtime.registry);
  registerPersistenceCommands(runtime.registry);
  registerPrintCommands(runtime.registry);
  registerXlsxCommands(runtime.registry);
  registerQueryCommands(runtime.registry);
  registerAutomationCommands(runtime.registry);
  registerExtendedCommands(runtime.registry);
  registerPermissionFeature(runtime);
  return [
    'history.restore',
    'persistence.save',
    'persistence.draft.clear',
    'print.preview',
    'print.export',
    'print.area.set',
    'print.pageSetup',
    'xlsx.import',
    'xlsx.export',
    'query.load',
    'query.refresh',
    'automation.run',
    'automation.record.start',
    'automation.record.stop',
    'extended.whatIf.goalSeek',
    'extended.whatIf.scenario',
    'extended.whatIf.dataTable',
    'extended.capability.evaluate',
    'sheet.protect.set',
    'sheet.protect.remove',
  ];
}
