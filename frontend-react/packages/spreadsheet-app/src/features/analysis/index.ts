import type { CommandRuntime } from '@react-sheets/command-runtime';
import { registerAnalysisCommands } from './commands';

export { registerAnalysisCommands } from './commands';
export type {
  AnalysisViewRemoveParams,
  AnalysisViewReplaceParams,
  AnalysisViewSetParams,
} from './commands';
export {
  buildAnalysisViewProjection,
  type AnalysisCellValue,
  type AnalysisChartPoint,
  type AnalysisChartProjection,
  type AnalysisErrorValue,
  type AnalysisProjectedRow,
  type AnalysisViewProjection,
} from './projection';

export function registerAnalysisFeature(runtime: CommandRuntime) {
  return {
    id: 'analysis',
    version: '1.0.0',
    commandIds: registerAnalysisCommands(runtime),
    mutationIds: ['analysis.view.replace'],
    permissions: ['analysis.view.write'],
  };
}
