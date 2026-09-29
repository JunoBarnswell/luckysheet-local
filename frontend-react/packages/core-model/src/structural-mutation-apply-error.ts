export class StructuralMutationApplyError extends Error {
  readonly code = 'STRUCTURAL_MUTATION_APPLY_FAILED';
  readonly originalCause: unknown;

  constructor(cause: unknown) {
    const detail = cause instanceof Error ? cause.message : String(cause);
    super(`STRUCTURAL_MUTATION_APPLY_FAILED: structural commit may be partially applied. Cause: ${detail}`);
    this.name = 'StructuralMutationApplyError';
    this.originalCause = cause;
  }
}

export function commitStructuralMutation<T>(commit: () => T): T {
  try {
    return commit();
  } catch (error) {
    if (error instanceof StructuralMutationApplyError) throw error;
    throw new StructuralMutationApplyError(error);
  }
}
