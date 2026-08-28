import type { CommandRuntime } from '@react-sheets/command-runtime';
import { registerHistoryCommands } from './features/history';
import { registerPrintCommands } from './features/print';
import { registerNativeDocumentCommands } from './features/native-document';
import { registerPermissionFeature } from './features/permission/commands';
import { registerQueryCommands } from './features/query';
import { registerExtendedCommands } from './features/extended';

/** 注册 M10–M18 平台特性命令 */
export function registerPlatformFeatures(runtime: CommandRuntime): string[] {
  registerHistoryCommands(runtime.registry);
  registerPrintCommands(runtime.registry);
  registerNativeDocumentCommands(runtime.registry);
  registerQueryCommands(runtime.registry);
  registerExtendedCommands(runtime.registry);
  registerPermissionFeature(runtime);
  return [
    'history.restore',
    'print.preview',
    'print.export',
    'pageLayout.printArea.set',
    'pageLayout.pageSetup.set',
    'document.import',
    'document.export',
    'query.load',
    'query.refresh',
    'extended.whatIf.goalSeek',
    'extended.whatIf.scenario',
    'sheet.protect.set',
    'sheet.protect.remove',
  ];
}
