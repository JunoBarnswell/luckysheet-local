import { KernelInvocationError, kernelInvoke } from '../../../kernel-client/src/index';
export interface FunctionDescriptor { id: string; cost: 'scalar' | 'range' | 'sort' | 'volatile' | 'external'; streaming: boolean; volatile: boolean; }
export type FunctionLibraryCategory = 'logical' | 'text' | 'date-time' | 'lookup-reference' | 'math-trig' | 'more-functions';
export interface FunctionLibraryEntry extends FunctionDescriptor { category: FunctionLibraryCategory; }
export type FormulaFunctionCapabilityStatus = 'native' | 'unsupported';
export interface FormulaFunctionCapability extends FunctionLibraryEntry { status: FormulaFunctionCapabilityStatus; }
/** The executable Rust registry is the only function capability source. */
export function listFunctionLibrary(category?: FunctionLibraryCategory): readonly FormulaFunctionCapability[] {
  const { functions } = kernelInvoke<{ functions: FormulaFunctionCapability[] }>('formula.functions', {});
  if (!Array.isArray(functions) || functions.some(entry => !entry || typeof entry.id !== 'string' || typeof entry.category !== 'string' || entry.status !== 'native')) throw new KernelInvocationError({ code: 'KERNEL_PROTOCOL_ERROR', message: 'Kernel function capability registry is invalid.', recovery: 'reload-kernel' });
  return category ? functions.filter(entry => entry.category === category) : functions;
}
export function getFunctionDescriptor(name: string): FunctionDescriptor | undefined { return listFunctionLibrary().find(entry => entry.id === name.trim().toUpperCase()); }
export function getFunctionCapability(name: string): FormulaFunctionCapability {
  const normalized = name.trim().toUpperCase();
  return listFunctionLibrary().find(entry => entry.id === normalized) ?? { id: normalized, cost: 'external', streaming: false, volatile: false, status: 'unsupported', category: 'more-functions' };
}
