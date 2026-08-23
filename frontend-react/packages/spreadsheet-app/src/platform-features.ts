import type { CommandRuntime } from '@react-sheets/command-runtime';
import { registerHistoryCommands } from './features/history';
import { createDefaultConnectorRegistry, registerQueryCommands } from './features/query';

/** 注册 M10–M18 平台特性命令 */
export function registerPlatformFeatures(runtime: CommandRuntime): string[] {
  registerHistoryCommands(runtime.registry);
  const connectors = createDefaultConnectorRegistry();
  registerQueryCommands(runtime.registry, connectors);
  return ['history.restore', 'query.load', 'query.refresh'];
}
