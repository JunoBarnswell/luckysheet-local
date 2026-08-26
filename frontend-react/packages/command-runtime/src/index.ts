import { WorkbookModel, type ProtectionAction, type RangeRef, type WorksheetModel } from '@react-sheets/core-model';

export interface MutationInfo<P = unknown> {
  id: string;
  unitId: string;
  sheetId: string;
  params: P;
  affectedRanges: RangeRef[];
  /** Explicit semantic override used by inverses whose storage mutation id is shared. */
  permission?: {
    capability: string;
    protectionAction: ProtectionAction | 'none';
    checksProtection: boolean;
    affectedRangeMode: 'none' | 'declared' | 'exact';
    objectScope: 'cell' | 'range' | 'row' | 'column' | 'drawing' | 'worksheet' | 'workbook';
  };
}

/**
 * Runtime validation contract for mutation parameters.
 *
 * The command runtime deliberately does not depend on a schema library. A
 * feature can provide a small, deterministic validator (typically generated
 * from its command contract) and the runtime will execute it for both local
 * and remote/replayed mutations.
 */
export interface MutationParamsSchema<P = unknown> {
  readonly name?: string;
  readonly validate: (params: unknown) => boolean;
}

export interface MutationPermissionMetadata {
  /** Stable capability name checked by the host/server authorization layer. */
  readonly capability: string;
  /** Optional role hint for UI projection; never used as an authorization source. */
  readonly roles?: readonly string[];
}

export interface MutationAffectedRangesMetadata<P = unknown> {
  /** Resolve the ranges from the mutation payload before it is applied. */
  readonly resolve: (params: P) => readonly RangeRef[];
  /** `exact` rejects a payload whose ranges differ from the declaration. */
  readonly mode?: 'exact' | 'declared';
}

export interface MutationRegistrationMetadata<P = unknown> {
  /** Canonical parameter contract; every production mutation must provide it. */
  readonly schema: MutationParamsSchema<P>;
  /** Canonical authorization capability; every production mutation must provide it. */
  readonly permission: MutationPermissionMetadata;
  /** Canonical affected-range resolver; every production mutation must provide it. */
  readonly affectedRanges: MutationAffectedRangesMetadata<P>;
  /**
   * Inverse contract; all applied inverse ids must be declared here. Exactly
   * one of `inversePolicy` and `inverseIds` is required; `inverseIds` is
   * normalized into this policy at registration time.
   */
  readonly inversePolicy?: MutationInversePolicy;
  /** Explicit inverse allow-list accepted by the registry and normalized to `inversePolicy`. */
  readonly inverseIds?: readonly string[];
}

/** Short public name for feature packages that expose a mutation contract. */
export type MutationMetadata<P = unknown> = MutationRegistrationMetadata<P>;

export interface MutationInversePolicy {
  readonly allowedMutationIds: readonly string[];
  readonly minCount: number;
  readonly maxCount?: number;
}

export interface CommandResult {
  operationId: string;
  mutationCount: number;
  affectedRanges: RangeRef[];
  /** Typed domain event emitted by commands that intentionally have no model mutation. */
  event?: {
    type: string;
    payload: Record<string, unknown>;
  };
}

/** The sole UI, script and host intent contract for a domain command. */
export interface CommandDescriptor<Params = unknown> {
  readonly commandId: string;
  readonly params?: Params;
}

export interface OperationResult {
  operationId: string;
}

export interface Command<P = unknown> {
  id: string;
  execute(params: P, context: CommandContext): CommandResult;
}

export interface Operation<P = unknown> {
  id: string;
  execute(params: P, context: CommandContext): OperationResult;
}

export interface Mutation<P = unknown> extends MutationInfo<P> {
  apply(context: CommandContext): void;
  inverse: MutationInfo[];
}

export interface CommandContext {
  readonly workbook: WorkbookModel;
  readonly operationId: string;
  /** Optional canonical worksheet-value authority supplied by the host runtime. */
  readonly resolveCellValue?: (sheet: WorksheetModel, row: number, column: number) => unknown;
  applyMutation<P>(mutation: Mutation<P>): void;
  recordOperation<P>(operation: Operation<P>, params: P): OperationResult;
}

export type MutationHandler<P = unknown> = (item: MutationInfo<P>, context: CommandContext) => void;

export interface MutationRegistration<P = unknown> {
  readonly id: string;
  readonly handler: MutationHandler<P>;
  readonly metadata: MutationRegistrationMetadata<P>;
}

/** Constructor options are intentionally non-permissive; metadata is always required. */
export interface CommandRegistryOptions {
  readonly requireMutationMetadata?: true;
}

export type CanonicalMutationMetadata<P = unknown> = MutationRegistrationMetadata<P> & {
  readonly inversePolicy: MutationInversePolicy;
};

type RegisteredMutation<P = unknown> = Omit<MutationRegistration<P>, 'metadata'> & {
  readonly metadata: CanonicalMutationMetadata<P>;
};

export type MutationRegistryIssueCode =
  | 'invalid-registration'
  | 'missing-schema'
  | 'missing-permission'
  | 'missing-affected-ranges'
  | 'unknown-mutation'
  | 'unknown-inverse'
  | 'invalid-inverse'
  | 'invalid-params'
  | 'invalid-affected-ranges'
  | 'inverse-not-allowed'
  | 'invalid-inverse-policy';

export interface MutationRegistryIssue {
  readonly code: MutationRegistryIssueCode;
  readonly mutationId: string;
  readonly inverseId?: string;
  readonly message: string;
}

export interface MutationRegistryCompletenessResult {
  readonly ok: boolean;
  readonly issues: readonly MutationRegistryIssue[];
}

function issue(
  code: MutationRegistryIssueCode,
  mutationId: string,
  message: string,
  inverseId?: string,
): MutationRegistryIssue {
  return inverseId === undefined ? { code, mutationId, message } : { code, mutationId, inverseId, message };
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null;
}

function isValidRangeRef(value: unknown): value is RangeRef {
  if (!isRecord(value)) return false;
  const { sheetId, startRow, endRow, startColumn, endColumn } = value;
  return (
    typeof sheetId === 'string' &&
    sheetId.length > 0 &&
    Number.isInteger(startRow) &&
    Number.isInteger(endRow) &&
    Number.isInteger(startColumn) &&
    Number.isInteger(endColumn) &&
    (startRow as number) >= 0 &&
    (endRow as number) >= (startRow as number) &&
    (startColumn as number) >= 0 &&
    (endColumn as number) >= (startColumn as number)
  );
}

function rangesEqual(left: readonly RangeRef[], right: readonly RangeRef[]): boolean {
  if (left.length !== right.length) return false;
  return left.every((range, index) => {
    const other = right[index];
    return (
      other !== undefined &&
      range.sheetId === other.sheetId &&
      range.startRow === other.startRow &&
      range.endRow === other.endRow &&
      range.startColumn === other.startColumn &&
      range.endColumn === other.endColumn
    );
  });
}

function formatIssues(issues: readonly MutationRegistryIssue[]): string {
  return issues.map((entry) => entry.message).join('; ');
}

function isValidInversePolicy(value: unknown): value is MutationInversePolicy {
  if (!isRecord(value) || !Array.isArray(value.allowedMutationIds) || value.allowedMutationIds.length === 0
    || !value.allowedMutationIds.every((entry) => typeof entry === 'string' && entry.length > 0)
    || !Number.isSafeInteger(value.minCount) || Number(value.minCount) < 1) return false;
  const minCount = Number(value.minCount);
  return value.maxCount === undefined
    || (Number.isSafeInteger(value.maxCount) && Number(value.maxCount) >= minCount);
}

function validateRegistrationMetadata(
  id: string,
  metadata: unknown,
): { metadata?: CanonicalMutationMetadata; issues: MutationRegistryIssue[] } {
  const issues: MutationRegistryIssue[] = [];
  if (!isRecord(metadata)) {
    issues.push(issue('invalid-registration', id, `Mutation ${id} requires canonical metadata`));
    return { issues };
  }
  const schema = metadata.schema;
  if (!isRecord(schema) || typeof schema.validate !== 'function') {
    issues.push(issue('missing-schema', id, `Mutation ${id} must declare a parameter schema`));
  }
  const permission = metadata.permission;
  if (!isRecord(permission) || typeof permission.capability !== 'string' || permission.capability.length === 0) {
    issues.push(issue('missing-permission', id, `Mutation ${id} must declare a permission capability`));
  } else if (permission.roles !== undefined
    && (!Array.isArray(permission.roles) || !permission.roles.every((entry) => typeof entry === 'string' && entry.length > 0))) {
    issues.push(issue('invalid-registration', id, `Mutation ${id} declares invalid permission roles`));
  }
  const affectedRanges = metadata.affectedRanges;
  if (!isRecord(affectedRanges) || typeof affectedRanges.resolve !== 'function') {
    issues.push(issue('missing-affected-ranges', id, `Mutation ${id} must declare an affected-range resolver`));
  } else if (affectedRanges.mode !== undefined && affectedRanges.mode !== 'exact' && affectedRanges.mode !== 'declared') {
    issues.push(issue('invalid-registration', id, `Mutation ${id} declares an invalid affected-range mode`));
  }
  const inversePolicy = metadata.inversePolicy;
  const inverseIds = metadata.inverseIds;
  if (inversePolicy !== undefined && inverseIds !== undefined) {
    issues.push(issue('invalid-inverse-policy', id, `Mutation ${id} must declare one inverse policy form`));
  }
  const normalizedPolicy = inversePolicy ?? (Array.isArray(inverseIds)
    ? { allowedMutationIds: inverseIds, minCount: 1 }
    : undefined);
  if (!isValidInversePolicy(normalizedPolicy)) {
    issues.push(issue('invalid-inverse-policy', id, `Mutation ${id} must declare a valid inverse policy`));
  }
  if (issues.length > 0) return { issues };
  const { inversePolicy: _inversePolicy, inverseIds: _inverseIds, ...baseMetadata } = metadata;
  return {
    metadata: {
      ...baseMetadata,
      inversePolicy: normalizedPolicy as MutationInversePolicy,
    } as CanonicalMutationMetadata,
    issues,
  };
}

function createOperationId(): string {
  if (typeof crypto !== 'undefined' && crypto.randomUUID) {
    return crypto.randomUUID();
  }
  return 'op-' + Math.random().toString(36).substring(2, 9) + '-' + Date.now().toString(36);
}

export class CommandRegistry {
  private readonly commands = new Map<string, Command<unknown>>();
  private readonly operations = new Map<string, Operation<unknown>>();
  private readonly mutations = new Map<string, RegisteredMutation<unknown>>();

  constructor(_options: CommandRegistryOptions = {}) {}

  registerCommand<P>(command: Command<P>): void {
    if (!command.id || typeof command.id !== 'string' || typeof command.execute !== 'function') {
      throw new Error('Command registration requires a non-empty id and execute function');
    }
    if (this.commands.has(command.id)) throw new Error(`Duplicate command: ${command.id}`);
    this.commands.set(command.id, command as Command<unknown>);
  }

  registerOperation<P>(operation: Operation<P>): void {
    if (!operation.id || typeof operation.id !== 'string' || typeof operation.execute !== 'function') {
      throw new Error('Operation registration requires a non-empty id and execute function');
    }
    if (this.operations.has(operation.id)) throw new Error(`Duplicate operation: ${operation.id}`);
    this.operations.set(operation.id, operation as Operation<unknown>);
  }

  registerMutation<P>(registration: MutationRegistration<P>): void {
    const candidate = registration as unknown as MutationRegistration<P>;
    const normalized = validateRegistrationMetadata(candidate.id, candidate.metadata);
    if (normalized.issues.length > 0 || !normalized.metadata) {
      throw new Error(`Invalid mutation registration ${candidate.id}: ${formatIssues(normalized.issues)}`);
    }
    if (!candidate.id || typeof candidate.id !== 'string') {
      throw new Error('Mutation registration requires a non-empty id');
    }
    if (typeof candidate.handler !== 'function') {
      throw new Error(`Mutation registration requires a handler: ${candidate.id}`);
    }
    if (this.mutations.has(candidate.id)) throw new Error(`Duplicate mutation: ${candidate.id}`);
    this.mutations.set(candidate.id, {
      id: candidate.id,
      handler: candidate.handler as MutationHandler<unknown>,
      metadata: normalized.metadata as CanonicalMutationMetadata<unknown>,
    });
  }

  getCommand<P>(id: string): Command<P> {
    const command = this.commands.get(id);
    if (!command) throw new Error(`Unknown command: ${id}`);
    return command as Command<P>;
  }

  getOperation<P>(id: string): Operation<P> {
    const operation = this.operations.get(id);
    if (!operation) throw new Error(`Unknown operation: ${id}`);
    return operation as Operation<P>;
  }

  getMutation<P>(id: string): MutationHandler<P> {
    const registration = this.mutations.get(id);
    if (!registration) throw new Error(`Unknown mutation: ${id}`);
    return registration.handler as MutationHandler<P>;
  }

  getMutationRegistration<P>(id: string): MutationRegistration<P> {
    const registration = this.mutations.get(id);
    if (!registration) throw new Error(`Unknown mutation: ${id}`);
    return registration as MutationRegistration<P>;
  }

  getMutationMetadata<P>(id: string): MutationRegistrationMetadata<P> {
    return this.getMutationRegistration<P>(id).metadata;
  }

  hasCommand(id: string): boolean {
    return this.commands.has(id);
  }

  hasMutation(id: string): boolean {
    return this.mutations.has(id);
  }

  listCommandIds(): string[] {
    return [...this.commands.keys()].sort();
  }

  listMutationIds(): string[] {
    return [...this.mutations.keys()].sort();
  }

  listMutationRegistrations(): readonly MutationRegistration[] {
    return [...this.mutations.values()].map((registration) => ({
      id: registration.id,
      handler: registration.handler,
      metadata: registration.metadata,
    }));
  }

  /**
   * Validate all registered mutation contracts defensively. Registration
   * rejects incomplete metadata immediately; this gate also detects registry
   * state mutated through untyped JavaScript or an invalid runtime boundary.
   */
  validateCompleteness(): MutationRegistryCompletenessResult {
    const issues: MutationRegistryIssue[] = [];

    for (const registration of this.mutations.values()) {
      const { id, handler, metadata } = registration;
      if (!id || typeof id !== 'string' || typeof handler !== 'function') {
        issues.push(issue('invalid-registration', id || '<empty>', `Invalid mutation registration: ${id || '<empty>'}`));
        continue;
      }
      const normalized = validateRegistrationMetadata(id, metadata);
      issues.push(...normalized.issues);
      if (!normalized.metadata) continue;
      for (const inverseId of normalized.metadata.inversePolicy.allowedMutationIds) {
        if (!this.mutations.has(inverseId)) {
          issues.push(issue('unknown-inverse', id, `Mutation ${id} declares unknown inverse mutation ${inverseId}`, inverseId));
        }
      }
    }

    return { ok: issues.length === 0, issues };
  }

  assertComplete(): void {
    const result = this.validateCompleteness();
    if (!result.ok) throw new Error(`Mutation registry is incomplete: ${formatIssues(result.issues)}`);
  }

  /** Validate a mutation before its apply callback can touch the workbook. */
  validateMutation<P>(mutation: Mutation<P>): readonly MutationRegistryIssue[] {
    const issues: MutationRegistryIssue[] = [];
    const registration = this.mutations.get(mutation.id);
    if (!registration) {
      issues.push(issue('unknown-mutation', mutation.id, `Unknown mutation: ${mutation.id}`));
      return issues;
    }

    this.validateMutationInfo(mutation, issues);
    const inversePolicy = registration.metadata.inversePolicy;
    if (!Array.isArray(mutation.inverse)
      || mutation.inverse.length < inversePolicy.minCount
      || (inversePolicy.maxCount !== undefined && mutation.inverse.length > inversePolicy.maxCount)) {
      issues.push(issue(
        'invalid-inverse',
        mutation.id,
        `Mutation ${mutation.id} inverse count violates its policy (${inversePolicy.minCount}${inversePolicy.maxCount === undefined ? '+' : `-${inversePolicy.maxCount}`})`,
      ));
    } else {
      for (const inverse of mutation.inverse) {
        if (!isRecord(inverse) || typeof inverse.id !== 'string' || !inverse.id) {
          issues.push(issue('invalid-inverse', mutation.id, `Mutation ${mutation.id} contains an invalid inverse mutation`));
          continue;
        }
        if (!this.mutations.has(inverse.id)) {
          issues.push(issue('unknown-inverse', mutation.id, `Mutation ${mutation.id} references unknown inverse ${inverse.id}`, inverse.id));
          continue;
        }
        if (inverse.unitId !== mutation.unitId) {
          issues.push(issue('invalid-inverse', mutation.id, `Mutation ${mutation.id} inverse ${inverse.id} targets a different workbook unit`, inverse.id));
        }
        const inverseIssues: MutationRegistryIssue[] = [];
        this.validateMutationInfo(inverse as MutationInfo, inverseIssues, mutation.id, inverse.id);
        issues.push(...inverseIssues);
        if (!inversePolicy.allowedMutationIds.includes(inverse.id)) {
          issues.push(issue('inverse-not-allowed', mutation.id, `Mutation ${mutation.id} does not allow inverse ${inverse.id}`, inverse.id));
        }
      }
    }
    return issues;
  }

  assertMutation<P>(mutation: Mutation<P>): void {
    const issues = this.validateMutation(mutation);
    if (issues.length > 0) throw new Error(`Invalid mutation ${mutation.id}: ${formatIssues(issues)}`);
  }

  validateMutationInfo<P>(
    item: MutationInfo<P>,
    issues: MutationRegistryIssue[] = [],
    ownerMutationId = item.id,
    inverseId?: string,
  ): MutationRegistryIssue[] {
    const registration = this.mutations.get(item.id);
    if (!registration) {
      issues.push(issue('unknown-mutation', ownerMutationId, `Unknown mutation: ${item.id}`, inverseId));
      return issues;
    }

    if (!isRecord(item) || typeof item.unitId !== 'string' || !item.unitId || typeof item.sheetId !== 'string' || !item.sheetId) {
      issues.push(issue('invalid-registration', ownerMutationId, `Mutation ${item.id} has invalid unitId or sheetId`, inverseId));
    }
    if (!Array.isArray(item.affectedRanges) || !item.affectedRanges.every(isValidRangeRef)) {
      issues.push(issue('invalid-affected-ranges', ownerMutationId, `Mutation ${item.id} has invalid affected ranges`, inverseId));
    }

    const schema = registration.metadata.schema;
    let valid = false;
    try {
      valid = schema.validate(item.params);
    } catch {
      valid = false;
    }
    if (!valid) {
      issues.push(issue('invalid-params', ownerMutationId, `Mutation ${item.id} parameters do not match ${schema.name ?? 'its schema'}`, inverseId));
    }

    const affectedRanges = registration.metadata.affectedRanges;
    if (Array.isArray(item.affectedRanges)) {
      try {
        const declared = affectedRanges.resolve(item.params as never);
        if (!Array.isArray(declared) || !declared.every(isValidRangeRef)) {
          issues.push(issue('invalid-affected-ranges', ownerMutationId, `Mutation ${item.id} declared an invalid affected-range result`, inverseId));
        } else if ((affectedRanges.mode ?? 'exact') === 'exact' && !rangesEqual(item.affectedRanges, declared)) {
          issues.push(issue('invalid-affected-ranges', ownerMutationId, `Mutation ${item.id} affected ranges differ from its declaration`, inverseId));
        }
      } catch {
        issues.push(issue('invalid-affected-ranges', ownerMutationId, `Mutation ${item.id} affected-range resolver failed`, inverseId));
      }
    }
    return issues;
  }
}

export interface HistoryEntry {
  operationId: string;
  baseRevision: number;
  committedRevision?: number;
  semanticCommandDescriptor: {
    id: string;
    params: unknown;
  };
  forwardMutations: MutationInfo[];
  inversePlan: MutationInfo[];
  affectedRanges: RangeRef[];
  status: 'active' | 'invalid';
  invalidReason?: string;
  /** Existing consumers read these exact same arrays; they are not a second state. */
  undo: MutationInfo[];
  redo: MutationInfo[];
  description?: string;
  timestamp: number;
}

export interface CommandRuntimeOptions {
  readonly getRevision?: () => number;
}

export interface RemoteMutationContext {
  readonly operationId?: string;
  readonly baseRevision?: number;
  readonly revision?: number;
}

interface StructuralDelta {
  readonly axis: 'row' | 'column';
  readonly at: number;
  readonly count: number;
  readonly direction: 1 | -1;
  readonly sheetId: string;
}

interface TransformValueResult {
  readonly value: unknown;
  readonly safe: boolean;
}

interface TransformedHistoryEntry {
  readonly ok: true;
  readonly inversePlan: MutationInfo[];
  readonly forwardMutations: MutationInfo[];
  readonly affectedRanges: RangeRef[];
}

interface InvalidHistoryTransform {
  readonly ok: false;
  readonly reason: string;
}

type HistoryTransformResult = TransformedHistoryEntry | InvalidHistoryTransform;

function structuralDelta(mutation: MutationInfo): StructuralDelta | undefined {
  const axis = mutation.id.includes('column') ? 'column' : mutation.id.includes('row') ? 'row' : undefined;
  if (!axis) return undefined;
  const direction = mutation.id.includes('insert') ? 1 : mutation.id.includes('delete') ? -1 : undefined;
  if (!direction || !isRecord(mutation.params)) return undefined;
  const at = mutation.params.at;
  const count = mutation.params.count;
  if (typeof at !== 'number' || typeof count !== 'number' || !Number.isSafeInteger(at) || !Number.isSafeInteger(count) || at < 0 || count < 1) return undefined;
  return { axis, at, count, direction, sheetId: mutation.sheetId };
}

function transformIndex(index: number, delta: StructuralDelta): number | undefined {
  if (!Number.isSafeInteger(index) || index < 0) return undefined;
  if (delta.direction === 1) return index >= delta.at ? index + delta.count : index;
  const deletedEnd = delta.at + delta.count - 1;
  if (index >= delta.at && index <= deletedEnd) return undefined;
  return index > deletedEnd ? index - delta.count : index;
}

function transformRange(range: RangeRef, delta: StructuralDelta): RangeRef | undefined {
  if (range.sheetId !== delta.sheetId) return structuredClone(range);
  const start = delta.axis === 'row' ? range.startRow : range.startColumn;
  const end = delta.axis === 'row' ? range.endRow : range.endColumn;
  if (delta.direction === -1) {
    const deletedEnd = delta.at + delta.count - 1;
    if (start <= deletedEnd && end >= delta.at) return undefined;
  }
  const nextStart = transformIndex(start, delta);
  const nextEnd = transformIndex(end, delta);
  if (nextStart === undefined || nextEnd === undefined) return undefined;
  let mappedStart = nextStart;
  let mappedEnd = nextEnd;
  if (delta.direction === 1 && start < delta.at && delta.at <= end) mappedEnd += delta.count;
  return delta.axis === 'row'
    ? { ...range, startRow: mappedStart, endRow: mappedEnd }
    : { ...range, startColumn: mappedStart, endColumn: mappedEnd };
}

function mutationAxis(id: string): 'row' | 'column' | undefined {
  if (id.includes('column')) return 'column';
  if (id.includes('row')) return 'row';
  return undefined;
}

function isCoordinateKey(key: string, axis: 'row' | 'column'): boolean {
  const normalized = key.toLowerCase();
  if (axis === 'row') {
    return normalized === 'row'
      || normalized === 'rowindex'
      || /^(start|end|top|bottom|anchor|target|source)row$/.test(normalized);
  }
  return normalized === 'column'
    || normalized === 'columnindex'
    || /^(start|end|top|bottom|anchor|target|source)column$/.test(normalized);
}

function isCoordinateListKey(key: string, axis: 'row' | 'column'): boolean {
  const normalized = key.toLowerCase();
  return axis === 'row'
    ? normalized === 'rows' || normalized === 'sourcerows' || normalized === 'rowindices' || normalized === 'rowindexes'
    : normalized === 'columns' || normalized === 'sourcecolumns' || normalized === 'columnindices' || normalized === 'columnindexes';
}

function transformPayload(value: unknown, delta: StructuralDelta, id: string, keyHint = ''): TransformValueResult {
  if (Array.isArray(value)) {
    const values: unknown[] = [];
    for (const item of value) {
      if (typeof item === 'number' && isCoordinateListKey(keyHint, delta.axis)) {
        const mapped = transformIndex(item, delta);
        if (mapped === undefined) return { value, safe: false };
        values.push(mapped);
        continue;
      }
      const transformed = transformPayload(item, delta, id, keyHint);
      if (!transformed.safe) return transformed;
      values.push(transformed.value);
    }
    return { value: values, safe: true };
  }
  if (!isRecord(value)) return { value, safe: true };
  if (isValidRangeRef(value)) {
    const mapped = transformRange(value, delta);
    return mapped ? { value: mapped, safe: true } : { value, safe: false };
  }

  const result: Record<string, unknown> = {};
  const structuralAxis = mutationAxis(id);
  for (const [key, child] of Object.entries(value)) {
    if (isCoordinateKey(key, delta.axis) && typeof child === 'number') {
      const mapped = transformIndex(child, delta);
      if (mapped === undefined) return { value, safe: false };
      result[key] = mapped;
      continue;
    }
    if (key === 'at' && typeof child === 'number' && (structuralAxis === delta.axis || 'count' in value)) {
      const mapped = transformIndex(child, delta);
      if (mapped === undefined) return { value, safe: false };
      result[key] = mapped;
      continue;
    }
    const transformed = transformPayload(child, delta, id, key);
    if (!transformed.safe) return transformed;
    result[key] = transformed.value;
  }
  return { value: result, safe: true };
}

function transformMutation(item: MutationInfo, delta: StructuralDelta): MutationInfo | undefined {
  const affectedRanges: RangeRef[] = [];
  for (const range of item.affectedRanges) {
    const mapped = transformRange(range, delta);
    if (!mapped) return undefined;
    affectedRanges.push(mapped);
  }
  if (item.sheetId !== delta.sheetId) return { ...item, affectedRanges };
  const params = transformPayload(item.params, delta, item.id);
  if (!params.safe) return undefined;
  return { ...item, params: params.value, affectedRanges };
}

function transformHistoryEntry(entry: HistoryEntry, remote: MutationInfo): HistoryTransformResult {
  const delta = structuralDelta(remote);
  if (!delta) return {
    ok: true,
    inversePlan: [...entry.inversePlan],
    forwardMutations: [...entry.forwardMutations],
    affectedRanges: [...entry.affectedRanges],
  };
  const inversePlan: MutationInfo[] = [];
  const forwardMutations: MutationInfo[] = [];
  for (const mutation of entry.inversePlan) {
    const transformed = transformMutation(mutation, delta);
    if (!transformed) return { ok: false, reason: `History ${entry.operationId} cannot be safely transformed across ${remote.id}` };
    inversePlan.push(transformed);
  }
  for (const mutation of entry.forwardMutations) {
    const transformed = transformMutation(mutation, delta);
    if (!transformed) return { ok: false, reason: `History ${entry.operationId} cannot be safely transformed across ${remote.id}` };
    forwardMutations.push(transformed);
  }
  const affectedRanges: RangeRef[] = [];
  for (const range of entry.affectedRanges) {
    const transformed = transformRange(range, delta);
    if (!transformed) return { ok: false, reason: `History ${entry.operationId} affected range intersects ${remote.id}` };
    affectedRanges.push(transformed);
  }
  return { ok: true, inversePlan, forwardMutations, affectedRanges };
}

/** 变更来源:正向命令、本地撤销、本地重做、远端协同重放 */
export type MutationSource = 'command' | 'undo' | 'redo' | 'remote';

export type MutationListener = (mutation: MutationInfo, source: MutationSource) => void;
export type MutationGuard = (mutation: MutationInfo, source: MutationSource) => void;
export type CommandListener = (commandId: string, params: unknown, result: CommandResult) => void;
export type HistoryReplayListener = (source: 'undo' | 'redo', entry: HistoryEntry) => void;

export class CommandRuntime {
  private readonly undoStack: HistoryEntry[] = [];
  private readonly redoStack: HistoryEntry[] = [];
  private activeEntry: HistoryEntry | null = null;
  private transactionDepth = 0;
  private readonly mutationListeners: MutationListener[] = [];
  private readonly commandListeners: CommandListener[] = [];
  private readonly historyReplayListeners: HistoryReplayListener[] = [];
  private cellValueResolver?: (sheet: WorksheetModel, row: number, column: number) => unknown;
  private mutationGuard?: MutationGuard;
  private revisionProvider?: () => number;
  private currentRevision = 0;
  private readonly invalidHistory: HistoryEntry[] = [];

  constructor(
    readonly workbook: WorkbookModel,
    readonly registry = new CommandRegistry(),
    options: CommandRuntimeOptions = {},
  ) {
    this.revisionProvider = options.getRevision;
  }

  setCellValueResolver(resolver: ((sheet: WorksheetModel, row: number, column: number) => unknown) | undefined): void {
    this.cellValueResolver = resolver;
  }

  /** Guard every local, undo/redo, and remote mutation at one boundary. */
  setMutationGuard(guard: MutationGuard | undefined): void {
    this.mutationGuard = guard;
  }

  setRevisionProvider(provider: (() => number) | undefined): void {
    this.revisionProvider = provider;
  }

  setRevision(revision: number): void {
    if (!Number.isSafeInteger(revision) || revision < 0) throw new Error('Revision must be a non-negative safe integer');
    this.currentRevision = revision;
  }

  onMutation(listener: MutationListener): () => void {
    this.mutationListeners.push(listener);
    return () => {
      const idx = this.mutationListeners.indexOf(listener);
      if (idx >= 0) this.mutationListeners.splice(idx, 1);
    };
  }

  onCommand(listener: CommandListener): () => void {
    this.commandListeners.push(listener);
    return () => {
      const idx = this.commandListeners.indexOf(listener);
      if (idx >= 0) this.commandListeners.splice(idx, 1);
    };
  }

  onHistoryReplay(listener: HistoryReplayListener): () => void {
    this.historyReplayListeners.push(listener);
    return () => {
      const idx = this.historyReplayListeners.indexOf(listener);
      if (idx >= 0) this.historyReplayListeners.splice(idx, 1);
    };
  }

  execute<P>(id: string, params: P): CommandResult {
    // Resolve the command before opening a transaction. An unknown command is
    // a protocol error and must not create an empty history entry or invoke a
    // host fallback.
    const command = this.registry.getCommand<P>(id);
    this.registry.assertComplete();
    const operationId = createOperationId();
    const mutations: MutationInfo[] = [];
    const isRootTransaction = this.transactionDepth === 0;

    if (isRootTransaction) {
      const inversePlan: MutationInfo[] = [];
      const forwardMutations: MutationInfo[] = [];
      this.activeEntry = {
        operationId,
        baseRevision: this.readRevision(),
        semanticCommandDescriptor: { id, params: structuredClone(params) },
        forwardMutations,
        inversePlan,
        affectedRanges: [],
        status: 'active',
        // Keep old public field names as references to the canonical arrays.
        undo: inversePlan,
        redo: forwardMutations,
        description: id,
        timestamp: Date.now(),
      };
    }
    this.transactionDepth += 1;

    const context: CommandContext = {
      workbook: this.workbook,
      operationId,
      resolveCellValue: (sheet, row, column) => this.cellValueResolver?.(sheet, row, column),
      applyMutation: (mutation) => {
        if (mutation.unitId !== this.workbook.unitId) {
          throw new Error(`Mutation unit mismatch: expected ${this.workbook.unitId}, received ${mutation.unitId}`);
        }
        // Registration and inverse validation happen before the mutation's
        // callback is allowed to touch the workbook. This makes both local
        // execution and every replay path fail closed on protocol drift.
        this.registry.assertMutation(mutation);
        this.mutationGuard?.(mutation, 'command');
        mutation.apply(context);
        const info: MutationInfo = {
          id: mutation.id,
          unitId: mutation.unitId,
          sheetId: mutation.sheetId,
          params: mutation.params,
          affectedRanges: mutation.affectedRanges,
          ...(mutation.permission ? { permission: structuredClone(mutation.permission) } : {}),
        };
        mutations.push(info);
        this.activeEntry?.inversePlan.unshift(...mutation.inverse);
        this.activeEntry?.forwardMutations.push(info);
        if (this.activeEntry) {
          this.activeEntry.affectedRanges.push(...mutation.affectedRanges.map((range) => structuredClone(range)));
        }

        for (const listener of this.mutationListeners) {
          listener(info, 'command');
        }
      },
      recordOperation: (operation, operationParams) => {
        const registered = this.registry.getOperation(operation.id);
        return registered.execute(operationParams, context);
      },
    };

    try {
      const commandResult = command.execute(params, context);
      this.transactionDepth -= 1;

      if (isRootTransaction) {
        if (this.activeEntry && (this.activeEntry.inversePlan.length > 0 || this.activeEntry.forwardMutations.length > 0)) {
          this.undoStack.push(this.activeEntry);
          if (this.undoStack.length > 200) this.undoStack.shift();
          this.redoStack.length = 0;
        }
        this.activeEntry = null;
      }

      const result: CommandResult = {
        ...commandResult,
        operationId,
        mutationCount: mutations.length,
      };

      for (const listener of this.commandListeners) {
        listener(id, params, result);
      }

      return result;
    } catch (err) {
      this.transactionDepth -= 1;
      if (isRootTransaction) {
        // Rollback applied mutations in this transaction if failed
        if (this.activeEntry && this.activeEntry.inversePlan.length > 0) {
          this.applyHistory(this.activeEntry.inversePlan, 'undo');
        }
        this.activeEntry = null;
      }
      throw err;
    }
  }

  undo(): boolean {
    this.registry.assertComplete();
    const entry = this.undoStack.pop();
    if (!entry) return false;
    if (entry.status !== 'active') return false;
    this.applyHistory(entry.inversePlan, 'undo');
    this.redoStack.push(entry);
    for (const listener of this.historyReplayListeners) listener('undo', entry);
    return true;
  }

  redo(): boolean {
    this.registry.assertComplete();
    const entry = this.redoStack.pop();
    if (!entry) return false;
    if (entry.status !== 'active') return false;
    this.applyHistory(entry.forwardMutations, 'redo');
    this.undoStack.push(entry);
    for (const listener of this.historyReplayListeners) listener('redo', entry);
    return true;
  }

  /**
   * 应用来自远端协同的变更序列:执行已注册的 mutation 处理器,
   * 以 'remote' 来源通知监听器(用于引擎同步/视图刷新),但不进入本地撤销栈。
   */
  applyRemoteMutations(items: readonly MutationInfo[], remoteContext: RemoteMutationContext = {}): void {
    this.registry.assertComplete();
    // A committed operation may contain several dependent mutations. Replay
    // them against an isolated snapshot first so a later rejection cannot
    // leave the live workbook partially changed.
    const preview = new CommandRuntime(WorkbookModel.fromSnapshot(this.workbook.snapshot()), this.registry);
    preview.applyHistory(items, 'remote');
    this.applyHistory(items, 'remote');
    for (const item of items) this.transformHistoryAgainstRemote(item);
    if (remoteContext.revision !== undefined) {
      if (!Number.isSafeInteger(remoteContext.revision) || remoteContext.revision < 1) throw new Error('Remote revision is invalid');
      this.currentRevision = Math.max(this.currentRevision, remoteContext.revision);
    }
  }

  markOperationCommitted(operationId: string, revision: number): void {
    if (!Number.isSafeInteger(revision) || revision < 1) throw new Error('Committed revision must be a positive safe integer');
    for (const entry of [...this.undoStack, ...this.redoStack, ...this.invalidHistory]) {
      if (entry.operationId === operationId) entry.committedRevision = revision;
    }
    this.currentRevision = Math.max(this.currentRevision, revision);
  }

  getInvalidHistoryEntries(): readonly HistoryEntry[] {
    return [...this.invalidHistory];
  }

  /** 当前事务嵌套深度(workspace 用以判断根事务冲刷协同队列) */
  get activeDepth(): number {
    return this.transactionDepth;
  }

  getHistoryDepth(): { undo: number; redo: number } {
    return { undo: this.undoStack.length, redo: this.redoStack.length };
  }

  getUndoEntries(): readonly HistoryEntry[] {
    return [...this.undoStack];
  }

  /** Read-only projection used by hosts to preflight a permission-safe redo. */
  getRedoEntries(): readonly HistoryEntry[] {
    return [...this.redoStack];
  }

  clearHistory(): void {
    this.undoStack.length = 0;
    this.redoStack.length = 0;
    this.invalidHistory.length = 0;
  }

  private readRevision(): number {
    const revision = this.revisionProvider?.() ?? this.currentRevision;
    if (!Number.isSafeInteger(revision) || revision < 0) throw new Error('History base revision is invalid');
    this.currentRevision = Math.max(this.currentRevision, revision);
    return revision;
  }

  private transformHistoryAgainstRemote(remote: MutationInfo): void {
    const stacks = [this.undoStack, this.redoStack];
    for (const stack of stacks) {
      for (let index = stack.length - 1; index >= 0; index -= 1) {
        const entry = stack[index]!;
        const transformed = transformHistoryEntry(entry, remote);
        if (!transformed.ok) {
          stack.splice(index, 1);
          entry.status = 'invalid';
          entry.invalidReason = transformed.reason;
          this.invalidHistory.push(entry);
          continue;
        }
        entry.inversePlan.splice(0, entry.inversePlan.length, ...transformed.inversePlan);
        entry.forwardMutations.splice(0, entry.forwardMutations.length, ...transformed.forwardMutations);
        entry.affectedRanges = transformed.affectedRanges;
      }
    }
  }

  private applyHistory(items: readonly MutationInfo[], source: MutationSource): void {
    const issues: MutationRegistryIssue[] = [];
    for (const item of items) {
      if (item.unitId !== this.workbook.unitId) {
        throw new Error(`Mutation unit mismatch: expected ${this.workbook.unitId}, received ${item.unitId}`);
      }
      this.registry.validateMutationInfo(item, issues);
    }
    if (issues.length > 0) {
      throw new Error(`Invalid mutation history: ${formatIssues(issues)}`);
    }
    for (const item of items) this.mutationGuard?.(item, source);
    for (const item of items) {
      const handler = this.registry.getMutation(item.id);
      const replayContext: CommandContext = {
        workbook: this.workbook,
        operationId: createOperationId(),
        resolveCellValue: (sheet, row, column) => this.cellValueResolver?.(sheet, row, column),
        applyMutation: () => {
          throw new Error('Nested mutation application is not allowed during mutation replay');
        },
        recordOperation: (operation, operationParams) => {
          const registered = this.registry.getOperation(operation.id);
          return registered.execute(operationParams, replayContext);
        },
      };
      handler(item, {
        ...replayContext,
      });
      for (const listener of this.mutationListeners) {
        listener(item, source);
      }
    }
  }
}
