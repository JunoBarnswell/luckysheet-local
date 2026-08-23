import type { CommandContext, CommandRuntime, MutationInfo } from '@react-sheets/command-runtime';
import type { RangeRef } from '@react-sheets/core-model';

export function sheetRange(sheetId: string): RangeRef[] {
  return [{ sheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }];
}

export function removeById<T extends { id: string }>(items: T[], id: string): T | undefined {
  const index = items.findIndex((item) => item.id === id);
  if (index < 0) return undefined;
  return items.splice(index, 1)[0];
}

export interface ApplyMutationOptions<P, InverseP = P> {
  id: string;
  sheetId: string;
  params: P;
  affectedRanges: RangeRef[];
  inverseId?: string;
  inverseParams: InverseP;
  apply: (context: CommandContext) => void;
}

export function applyTrackedMutation<P, InverseP = P>(context: CommandContext, options: ApplyMutationOptions<P, InverseP>): void {
  context.applyMutation({
    id: options.id,
    unitId: context.workbook.unitId,
    sheetId: options.sheetId,
    params: options.params,
    affectedRanges: options.affectedRanges,
    inverse: [
      {
        id: options.inverseId ?? options.id,
        unitId: context.workbook.unitId,
        sheetId: options.sheetId,
        params: options.inverseParams,
        affectedRanges: options.affectedRanges,
      },
    ],
    apply: () => options.apply(context),
  });
}

export function registerMutationHandler<P>(
  runtime: CommandRuntime,
  mutationId: string,
  handler: (params: P, context: CommandContext) => void,
): void {
  runtime.registry.registerMutation(mutationId, (item, context) => {
    handler(item.params as P, context);
  });
}

export function replayMutation<P>(
  runtime: CommandRuntime,
  mutationId: string,
  item: MutationInfo<P>,
  context: CommandContext,
): void {
  runtime.registry.getMutation<P>(mutationId)(item, context);
}
