import type { FunctionDescriptor } from './functions';

export type ComputeLocation = 'client' | 'server';

export interface ComputePolicyInput {
  dependencyCellCount: number;
  functions: readonly FunctionDescriptor[];
  externalData?: boolean;
}

export function chooseComputeLocation(input: ComputePolicyInput): ComputeLocation {
  if (input.externalData) return 'server';
  if (input.functions.some((descriptor) => descriptor.volatile || descriptor.cost === 'external' || descriptor.cost === 'sort')) return 'server';
  if (input.functions.some((descriptor) => descriptor.cost === 'range') && input.dependencyCellCount > 10_000) return 'server';
  return input.dependencyCellCount <= 10_000 ? 'client' : 'server';
}
