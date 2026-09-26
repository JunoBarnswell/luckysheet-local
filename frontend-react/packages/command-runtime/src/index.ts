import { WorkbookModel, isWorkbookCalculationContextEffect, normalizeCellDataForStorage, normalizeDefinedNameModel, readChartTextFormula, structuralRuleFormulaFields, writeChartTextFormula, type CellData, type ConditionalFormatRule, type DataValidationRule, type ProtectionAction, type RangeRef, type StructuralDefinedNameOwnerDelta, type StructuralFormulaOwnerDelta, type StructuralFormulaOwnerState, type StructuralFormulaRule, type StructuralReferenceOwnerIndex, type WorkbookCalculationContextEffect, type WorksheetModel } from '@react-sheets/core-model';
import { collectFormulaDependencies, collectFormulaReferenceNodes, formatFormula, mapAstStructuralReferences, parseFormula, RangeIndex, ReferenceTransformDomain, MAX_COLUMN_INDEX, MAX_ROW_INDEX, type FormulaRuleReferenceFailureReason, type FormulaRuleReferenceOwnerIdentity } from '@react-sheets/formula-engine';

export interface MutationInfo<P = unknown> {
  id: string;
  unitId: string;
  sheetId: string;
  params: P;
  affectedRanges: RangeRef[];
  /** Local history payload; stripped from the client-authored operation envelope. */
  structuralFormulaOwnerDeltas?: StructuralFormulaOwnerDelta[];
  /** Exact defined-name states used by local undo and structural-history preconditions. */
  structuralDefinedNameOwnerDeltas?: StructuralDefinedNameOwnerDelta[];
  /** Derived local history/OT scope; never replaces a mutation's declared range contract. */
  structuralImpactRanges?: RangeRef[];
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
  /** How remote application transforms or invalidates existing local undo/redo entries. */
  readonly historyRebase?: MutationHistoryRebasePolicy;
  /** Calculation context transition emitted when the handler has no more specific effect. */
  readonly calculationContextEffect?: WorkbookCalculationContextEffect;
}

/** Short public name for feature packages that expose a mutation contract. */
export type MutationMetadata<P = unknown> = MutationRegistrationMetadata<P>;

export interface MutationInversePolicy {
  readonly allowedMutationIds: readonly string[];
  readonly minCount: number;
  readonly maxCount?: number;
}

export type MutationHistoryRebasePolicy =
  | { readonly kind: 'axis'; readonly axis: 'row' | 'column'; readonly direction: 1 | -1 }
  | { readonly kind: 'invalidate'; readonly reason: string; readonly when?: (mutation: MutationInfo) => boolean };

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
  /**
   * Commands driven by view geometry may persist canonical local state while
   * intentionally staying out of the user's edit history.  Rollback still
   * uses the mutation inverse plan if execution fails.
   */
  history?: 'record' | 'none';
  execute(params: P, context: CommandContext): CommandResult;
}

export interface Operation<P = unknown> {
  id: string;
  execute(params: P, context: CommandContext): OperationResult;
}

export interface Mutation<P = unknown> extends MutationInfo<P> {
  apply(context: CommandContext): unknown;
  inverse: MutationInfo[];
}

export interface CommandContext {
  readonly workbook: WorkbookModel;
  readonly operationId: string;
  /** Indexed structural-reference owners from the canonical formula runtime. */
  readonly structuralReferenceOwners: StructuralReferenceOwnerIndex;
  /** Optional canonical worksheet-value authority supplied by the host runtime. */
  readonly resolveCellValue?: (sheet: WorksheetModel, row: number, column: number) => unknown;
  applyMutation<P>(mutation: Mutation<P>): void;
  recordOperation<P>(operation: Operation<P>, params: P): OperationResult;
}

export type MutationHandler<P = unknown> = (item: MutationInfo<P>, context: CommandContext) => unknown;

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
  if (!isValidHistoryRebasePolicy(metadata.historyRebase)) {
    issues.push(issue('invalid-registration', id, `Mutation ${id} declares an invalid history rebase policy`));
  }
  if (metadata.calculationContextEffect !== undefined
    && !isWorkbookCalculationContextEffect(metadata.calculationContextEffect)) {
    issues.push(issue('invalid-registration', id, `Mutation ${id} declares an invalid calculation context effect`));
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
    if (registration.metadata.historyRebase?.kind === 'axis'
      && isRecord(item.params)
      && typeof item.params.sheetId === 'string'
      && item.params.sheetId !== item.sheetId) {
      issues.push(issue('invalid-registration', ownerMutationId, `Structural mutation ${item.id} envelope sheetId differs from its target`, inverseId));
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

function isValidHistoryRebasePolicy(value: unknown): value is MutationHistoryRebasePolicy {
  if (value === undefined) return true;
  if (!isRecord(value)) return false;
  if (value.kind === 'axis') {
    return (value.axis === 'row' || value.axis === 'column')
      && (value.direction === 1 || value.direction === -1);
  }
  return value.kind === 'invalidate'
    && typeof value.reason === 'string'
    && value.reason.trim().length > 0
    && (value.when === undefined || typeof value.when === 'function');
}

function buildStructuralReferenceIndex(workbook: WorkbookModel): StructuralReferenceOwnerIndex {
  const sheetOrder = workbook.sheetOrder.map((id) => ({ id, name: workbook.getSheet(id).name }));
  const index = new RangeIndex(sheetOrder);
  const addFormulaOwner = (owner: { sheetId: string; row: number; column: number }, formula: string, sourceId?: string): void => {
    try {
      const normalized = formula.trimStart().startsWith('=') ? formula : `=${formula}`;
      const dependencies = collectFormulaDependencies(parseFormula(normalized), owner, { sheetOrder });
      if (sourceId === undefined) index.set(owner, dependencies);
      else index.setStructuralReference(owner, sourceId, dependencies);
    } catch {
      if (sourceId === undefined) index.set(owner, [], true);
      else index.setStructuralReference(owner, sourceId, [], true);
    }
  };
  for (const sheet of workbook.getSheets()) {
    sheet.cells.forEach((cell, row, column) => {
      const owner = { sheetId: sheet.id, row, column };
      if (cell.formula !== undefined) {
        const sourceId = cell.formulaMetadata?.preservedOnly
          ? 'structural:preserved-formula'
          : undefined;
        addFormulaOwner(owner, cell.formula, sourceId);
      }
      if (cell.formulaMetadata?.sourceFormula !== undefined) {
        addFormulaOwner(owner, cell.formulaMetadata.sourceFormula, 'structural:formula-provenance');
      }
      if (cell.presentation?.kind === 'barcode' && cell.presentation.source.kind === 'formula') {
        addFormulaOwner(owner, cell.presentation.source.formula, 'structural:barcode');
      }
    });
  }
  for (const entry of workbook.definedNameModels) {
    const owner = {
      scope: entry.scope,
      name: entry.name,
      ...(entry.sheetId ? { sheetId: entry.sheetId } : {}),
    };
    const context = entry.anchor ?? (entry.scope === 'sheet'
      ? { sheetId: entry.sheetId!, row: 0, column: 0 }
      : undefined);
    let references: ReturnType<typeof collectFormulaReferenceNodes>;
    try {
      const normalized = entry.formula.trimStart().startsWith('=') ? entry.formula : `=${entry.formula}`;
      references = collectFormulaReferenceNodes(parseFormula(normalized));
    } catch {
      index.setDefinedNameReference(owner, [], context, entry.anchor, 'invalid-formula');
      continue;
    }
    index.setDefinedNameReference(owner, references, context, entry.anchor);
  }
  indexStructuralFormulaRules(workbook, index);
  return index;
}

function indexStructuralFormulaRules(
  workbook: WorkbookModel,
  index: StructuralReferenceOwnerIndex,
  previous?: Map<string, StructuralFormulaRuleReferenceIndexState>,
): void {
  const sheetOrder = workbook.sheetOrder.map((id) => ({ id, name: workbook.getSheet(id).name }));
  const entries: Array<{
    owner: FormulaRuleReferenceOwnerIdentity;
    formula: string;
    context: { sheetId: string; row: number; column: number };
    failure?: FormulaRuleReferenceFailureReason;
    signature: string;
  }> = [];
  for (const sheet of workbook.getSheets()) {
    for (const [ruleKind, rules] of [
      ['conditional-format', sheet.conditionalFormats],
      ['data-validation', sheet.dataValidations],
    ] as const) {
      const idCounts = new Map<string, number>();
      for (const rule of rules) {
        if (typeof rule.id === 'string') idCounts.set(rule.id, (idCounts.get(rule.id) ?? 0) + 1);
      }
      for (const [ruleIndex, rawRule] of rules.entries()) {
        const rule = rawRule as StructuralFormulaRule;
        const formulas = structuralRuleFormulaFields(rule);
        if (formulas.size === 0) continue;
        const ranges = Array.isArray(rule.ranges) ? rule.ranges : [];
        const ruleId = typeof rule.id === 'string' ? rule.id : '';
        const validIdentity = rule.sheetId === sheet.id
          && ruleId.trim().length > 0
          && !ruleId.includes('\u0000')
          && idCounts.get(ruleId) === 1;
        const validRanges = ranges.length > 0 && ranges.every((range) => typeof range === 'object' && range !== null
          && range.sheetId === sheet.id
          && Number.isSafeInteger(range.startRow) && range.startRow >= 0 && range.startRow <= MAX_ROW_INDEX
          && Number.isSafeInteger(range.endRow) && range.endRow >= range.startRow && range.endRow <= MAX_ROW_INDEX
          && Number.isSafeInteger(range.startColumn) && range.startColumn >= 0 && range.startColumn <= MAX_COLUMN_INDEX
          && Number.isSafeInteger(range.endColumn) && range.endColumn >= range.startColumn && range.endColumn <= MAX_COLUMN_INDEX);
        const firstRange = ranges[0] && typeof ranges[0] === 'object' ? ranges[0] : undefined;
        const context = rule.formulaAnchor ?? (firstRange ? {
          sheetId: firstRange.sheetId,
          row: firstRange.startRow,
          column: firstRange.startColumn,
        } : undefined);
        const validContext = context !== undefined
          && sheetOrder.some((identity) => identity.id === context.sheetId)
          && Number.isSafeInteger(context.row) && context.row >= 0 && context.row <= MAX_ROW_INDEX
          && Number.isSafeInteger(context.column) && context.column >= 0 && context.column <= MAX_COLUMN_INDEX;
        const failure: FormulaRuleReferenceFailureReason | undefined = !validIdentity
          ? 'invalid-owner'
          : !validRanges
            ? 'invalid-range'
            : !validContext
              ? 'unresolved-context'
              : undefined;
        const safeContext = validContext ? context! : { sheetId: sheet.id, row: 0, column: 0 };
        for (const [field, formula] of formulas) {
          const owner = {
            sheetId: sheet.id,
            ruleKind,
            ruleId: ruleId.trim() && !ruleId.includes('\u0000') ? ruleId : `\u0000invalid-rule-${ruleIndex}`,
            field,
          };
          const signature = JSON.stringify([
            owner,
            formula,
            failure ?? null,
            rule.formulaAnchor ?? null,
            ranges,
            sheetOrder.map(({ id, name }) => [id, name]),
          ]);
          entries.push({ owner, formula, context: safeContext, ...(failure ? { failure } : {}), signature });
        }
      }
    }
  }

  const current = new Map<string, StructuralFormulaRuleReferenceIndexState>();
  for (const entry of entries) {
    const key = formulaRuleReferenceIndexKey(entry.owner);
    current.set(key, { owner: entry.owner, signature: entry.signature });
    if (previous?.get(key)?.signature === entry.signature && index.hasFormulaRuleReference(entry.owner)) continue;
    let references: ReturnType<typeof collectFormulaReferenceNodes> = [];
    let failure = entry.failure;
    if (!failure) {
      try {
        const normalized = entry.formula.trimStart().startsWith('=') ? entry.formula : `=${entry.formula}`;
        references = collectFormulaReferenceNodes(parseFormula(normalized));
      } catch {
        failure = 'invalid-formula';
      }
    }
    index.setFormulaRuleReference(entry.owner, references, entry.context, failure);
  }
  if (!previous) return;
  for (const [key, entry] of previous) {
    if (!current.has(key)) index.removeFormulaRuleReference(entry.owner);
  }
  previous.clear();
  for (const [key, entry] of current) previous.set(key, entry);
}

function formulaRuleReferenceIndexKey(owner: FormulaRuleReferenceOwnerIdentity): string {
  return JSON.stringify([owner.sheetId, owner.ruleKind, owner.ruleId, owner.field]);
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

function structuralDelta(mutation: MutationInfo, policy: MutationHistoryRebasePolicy | undefined): StructuralDelta | undefined {
  if (policy?.kind !== 'axis' || !isRecord(mutation.params)) return undefined;
  const at = mutation.params.at;
  const count = mutation.params.count;
  if (typeof at !== 'number' || typeof count !== 'number' || !Number.isSafeInteger(at) || !Number.isSafeInteger(count) || at < 0 || count < 1) return undefined;
  return { axis: policy.axis, at, count, direction: policy.direction, sheetId: mutation.sheetId };
}

function transformIndex(index: number, delta: StructuralDelta): number | undefined {
  const maximum = delta.axis === 'row' ? MAX_ROW_INDEX : MAX_COLUMN_INDEX;
  if (!Number.isSafeInteger(index) || index < 0 || index > maximum) return undefined;
  const mapped = ReferenceTransformDomain.mapPoint(index, delta.at, delta.count, delta.direction, maximum);
  return mapped.kind === 'mapped'
    ? mapped.position
    : undefined;
}

function transformRange(range: RangeRef, delta: StructuralDelta): RangeRef | undefined {
  if (!isValidRangeRef(range)
    || range.startRow < 0 || range.endRow < range.startRow || range.endRow > MAX_ROW_INDEX
    || range.startColumn < 0 || range.endColumn < range.startColumn || range.endColumn > MAX_COLUMN_INDEX) return undefined;
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
  const mappedStart = nextStart;
  const mappedEnd = nextEnd;
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

function isFormulaSourceKey(key: string): boolean {
  const normalized = key.toLowerCase();
  return normalized === 'formula'
    || normalized === 'formulatext'
    || /^formula\d+$/.test(normalized)
    || normalized.endsWith('formula');
}

function transformPayload(
  value: unknown,
  delta: StructuralDelta,
  id: string,
  sheetOrder: readonly { readonly id: string; readonly name: string }[],
  keyHint = '',
  ownerSheetId = delta.sheetId,
): TransformValueResult {
  if (typeof value === 'string' && isFormulaSourceKey(keyHint) && value.trimStart().startsWith('=')) {
    try {
      const sourceAst = parseFormula(value);
      const mapped = formatFormula(mapAstStructuralReferences(sourceAst, {
        shift: { axis: delta.axis, at: delta.at, count: delta.count, op: delta.direction === 1 ? 'insert' : 'delete' },
        ownerSheetId,
        targetSheetId: delta.sheetId,
        sheetOrder,
      }));
      return { value: mapped === formatFormula(sourceAst) ? value : mapped, safe: true };
    } catch {
      return { value, safe: false };
    }
  }
  if (Array.isArray(value)) {
    const values: unknown[] = [];
    for (const item of value) {
      if (ownerSheetId === delta.sheetId && typeof item === 'number' && isCoordinateListKey(keyHint, delta.axis)) {
        const mapped = transformIndex(item, delta);
        if (mapped === undefined) return { value, safe: false };
        values.push(mapped);
        continue;
      }
      const transformed = transformPayload(item, delta, id, sheetOrder, keyHint, ownerSheetId);
      if (!transformed.safe) return transformed;
      values.push(transformed.value);
    }
    return { value: values, safe: true };
  }
  if (!isRecord(value)) return { value, safe: true };
  const valueSheetId = typeof value.sheetId === 'string'
    ? value.sheetId
    : value.scope === 'workbook' ? delta.sheetId : ownerSheetId;
  if (isValidRangeRef(value)) {
    const mapped = transformRange(value, delta);
    return mapped ? { value: mapped, safe: true } : { value, safe: false };
  }

  const result: Record<string, unknown> = {};
  const structuralAxis = mutationAxis(id);
  let formulaChanged = false;
  for (const [key, child] of Object.entries(value)) {
    if (valueSheetId === delta.sheetId && isCoordinateKey(key, delta.axis) && typeof child === 'number') {
      const mapped = transformIndex(child, delta);
      if (mapped === undefined) return { value, safe: false };
      result[key] = mapped;
      continue;
    }
    if (valueSheetId === delta.sheetId && key === 'at' && typeof child === 'number' && (structuralAxis === delta.axis || 'count' in value)) {
      const mapped = transformIndex(child, delta);
      if (mapped === undefined) return { value, safe: false };
      result[key] = mapped;
      continue;
    }
    const transformed = transformPayload(child, delta, id, sheetOrder, key, valueSheetId);
    if (!transformed.safe) return transformed;
    result[key] = transformed.value;
    if (isFormulaSourceKey(key) && typeof child === 'string' && transformed.value !== child) formulaChanged = true;
  }
  if (formulaChanged) {
    delete result.formulaValue;
    delete result.displayValue;
  }
  return { value: result, safe: true };
}

function transformMutation(
  item: MutationInfo,
  delta: StructuralDelta,
  sheetOrder: readonly { readonly id: string; readonly name: string }[],
): MutationInfo | undefined {
  const affectedRanges: RangeRef[] = [];
  for (const range of item.affectedRanges) {
    const mapped = transformRange(range, delta);
    if (!mapped) return undefined;
    affectedRanges.push(mapped);
  }
  const params = transformPayload(item.params, delta, item.id, sheetOrder, '', item.sheetId);
  if (!params.safe) return undefined;
  return { ...item, params: params.value, affectedRanges };
}

function transformHistoryEntry(
  entry: HistoryEntry,
  remote: MutationInfo,
  sheetOrder: readonly { readonly id: string; readonly name: string }[],
  policy: MutationHistoryRebasePolicy | undefined,
): HistoryTransformResult {
  const definedNameDeltas = entry.inversePlan.flatMap((mutation) => mutation.structuralDefinedNameOwnerDeltas ?? []);
  const hasStructuralOwnerPatch = entry.inversePlan.some((mutation) => (
    mutation.structuralFormulaOwnerDeltas?.length || mutation.structuralDefinedNameOwnerDeltas?.length
  ));
  if (policy?.kind === 'axis' && hasStructuralOwnerPatch) return {
    ok: false,
    reason: `History ${entry.operationId} contains a structural owner patch that cannot be safely rebased across ${remote.id}`,
  };
  if (definedNameDeltas.some((delta) => remoteChangesDefinedNameOwner(remote, delta.owner))) return {
    ok: false,
    reason: `History ${entry.operationId} contains a defined-name owner patch that conflicts with ${remote.id}`,
  };
  if (policy?.kind === 'invalidate' && (!policy.when || policy.when(remote))) return {
    ok: false,
    reason: `History ${entry.operationId} cannot be safely rebased across ${remote.id}: ${policy.reason}`,
  };
  const activePolicy = policy?.kind === 'axis' ? policy : undefined;
  const delta = structuralDelta(remote, activePolicy);
  if (activePolicy?.kind === 'axis' && !delta) return {
    ok: false,
    reason: `History ${entry.operationId} cannot be safely rebased across ${remote.id}: its axis transform payload is invalid`,
  };
  if (!delta) {
    if (entry.affectedRanges.length === 0 && remote.affectedRanges.length === 0) return {
      ok: false,
      reason: `History ${entry.operationId} and remote mutation ${remote.id} have no canonical conflict scope`,
    };
    const overlapsRemote = entry.affectedRanges.some((left) => remote.affectedRanges.some((right) => (
      left.sheetId === right.sheetId
      && left.startRow <= right.endRow && left.endRow >= right.startRow
      && left.startColumn <= right.endColumn && left.endColumn >= right.startColumn
    )));
    if (overlapsRemote) return {
      ok: false,
      reason: `History ${entry.operationId} overlaps remote mutation ${remote.id} and has no canonical rebase`,
    };
    return {
      ok: true,
      inversePlan: [...entry.inversePlan],
      forwardMutations: [...entry.forwardMutations],
      affectedRanges: [...entry.affectedRanges],
    };
  }
  const inversePlan: MutationInfo[] = [];
  const forwardMutations: MutationInfo[] = [];
  for (const mutation of entry.inversePlan) {
    const transformed = transformMutation(mutation, delta, sheetOrder);
    if (!transformed) return { ok: false, reason: `History ${entry.operationId} cannot be safely transformed across ${remote.id}` };
    inversePlan.push(transformed);
  }
  for (const mutation of entry.forwardMutations) {
    const transformed = transformMutation(mutation, delta, sheetOrder);
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

export type MutationListener = (mutation: MutationInfo, source: MutationSource, effect?: unknown) => void;
export type MutationGuard = (mutation: MutationInfo, source: MutationSource) => void;
export type CommandListener = (commandId: string, params: unknown, result: CommandResult) => void;
export type CommandAbortListener = (commandId: string, params: unknown, operationId: string) => void;
export type HistoryReplayListener = (source: 'undo' | 'redo', entry: HistoryEntry) => void;

interface StructuralFormulaRuleReferenceIndexState {
  readonly owner: FormulaRuleReferenceOwnerIdentity;
  readonly signature: string;
}

export class CommandRuntime {
  private readonly undoStack: HistoryEntry[] = [];
  private readonly redoStack: HistoryEntry[] = [];
  private activeEntry: HistoryEntry | null = null;
  private transactionDepth = 0;
  private readonly mutationListeners: MutationListener[] = [];
  private readonly commandListeners: CommandListener[] = [];
  private readonly commandAbortListeners: CommandAbortListener[] = [];
  private readonly historyReplayListeners: HistoryReplayListener[] = [];
  private cellValueResolver?: (sheet: WorksheetModel, row: number, column: number) => unknown;
  private mutationGuard?: MutationGuard;
  private revisionProvider?: () => number;
  private structuralReferenceOwnersProvider?: (workbook: WorkbookModel) => StructuralReferenceOwnerIndex;
  private readonly structuralFormulaRuleReferenceIndex = new Map<string, StructuralFormulaRuleReferenceIndexState>();
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

  getMutationGuard(): MutationGuard | undefined { return this.mutationGuard; }

  setRevisionProvider(provider: (() => number) | undefined): void {
    this.revisionProvider = provider;
  }

  setStructuralReferenceOwnersProvider(
    provider: ((workbook: WorkbookModel) => StructuralReferenceOwnerIndex) | undefined,
  ): void {
    this.structuralReferenceOwnersProvider = provider;
    this.structuralFormulaRuleReferenceIndex.clear();
  }

  private resolveStructuralReferenceOwners(): StructuralReferenceOwnerIndex {
    const provided = this.structuralReferenceOwnersProvider?.(this.workbook);
    if (!provided) {
      this.structuralFormulaRuleReferenceIndex.clear();
      return buildStructuralReferenceIndex(this.workbook);
    }
    indexStructuralFormulaRules(this.workbook, provided, this.structuralFormulaRuleReferenceIndex);
    return provided;
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

  onCommandAbort(listener: CommandAbortListener): () => void {
    this.commandAbortListeners.push(listener);
    return () => {
      const idx = this.commandAbortListeners.indexOf(listener);
      if (idx >= 0) this.commandAbortListeners.splice(idx, 1);
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

    const commandRuntime = this;
    const context: CommandContext = {
      workbook: this.workbook,
      operationId,
      get structuralReferenceOwners() { return commandRuntime.resolveStructuralReferenceOwners(); },
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
        const effect = mutation.apply(context) ?? this.registry.getMutationMetadata(mutation.id).calculationContextEffect;
        const formulaOwnerDeltas = isRecord(effect) && Array.isArray(effect.formulaOwnerDeltas)
          ? effect.formulaOwnerDeltas as StructuralFormulaOwnerDelta[]
          : [];
        const definedNameOwnerDeltas = isRecord(effect) && Array.isArray(effect.definedNameOwnerDeltas)
          ? effect.definedNameOwnerDeltas as StructuralDefinedNameOwnerDelta[]
          : [];
        const structuralImpactRanges = formulaOwnerDeltasRanges(formulaOwnerDeltas);
        const info: MutationInfo = {
          id: mutation.id,
          unitId: mutation.unitId,
          sheetId: mutation.sheetId,
          params: mutation.params,
          affectedRanges: mutation.affectedRanges,
          ...(formulaOwnerDeltas.length > 0
            ? { structuralFormulaOwnerDeltas: structuredClone(formulaOwnerDeltas) }
            : {}),
          ...(structuralImpactRanges.length > 0
            ? { structuralImpactRanges }
            : {}),
          ...(definedNameOwnerDeltas.length > 0
            ? { structuralDefinedNameOwnerDeltas: structuredClone(definedNameOwnerDeltas) }
            : {}),
          ...(mutation.permission ? { permission: structuredClone(mutation.permission) } : {}),
        };
        mutations.push(info);
        const inverse = mutation.inverse.map((item, index) => ({
          ...item,
          ...(index === 0 && formulaOwnerDeltas.length > 0
            ? { structuralFormulaOwnerDeltas: structuredClone(formulaOwnerDeltas) }
            : {}),
          ...(index === 0 && definedNameOwnerDeltas.length > 0
            ? { structuralDefinedNameOwnerDeltas: structuredClone(definedNameOwnerDeltas) }
            : {}),
        }));
        this.activeEntry?.inversePlan.unshift(...inverse);
        this.activeEntry?.forwardMutations.push(info);
        if (this.activeEntry) {
          this.activeEntry.affectedRanges.push(...mutation.affectedRanges.map((range) => structuredClone(range)));
          this.activeEntry.affectedRanges.push(...structuralImpactRanges.map((range) => structuredClone(range)));
        }

        for (const listener of this.mutationListeners) {
          listener(info, 'command', effect);
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
        if (command.history !== 'none' && this.activeEntry && (this.activeEntry.inversePlan.length > 0 || this.activeEntry.forwardMutations.length > 0)) {
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
        let rollbackError: unknown;
        try {
          if (this.activeEntry && this.activeEntry.inversePlan.length > 0) {
            this.applyHistory(this.activeEntry.inversePlan, 'undo');
          }
        } catch (error) {
          rollbackError = error;
        } finally {
          this.activeEntry = null;
          for (const listener of this.commandAbortListeners) listener(id, params, operationId);
        }
        if (rollbackError !== undefined) throw rollbackError;
      }
      throw err;
    }
  }

  undo(): boolean {
    this.registry.assertComplete();
    const entry = this.undoStack[this.undoStack.length - 1];
    if (!entry) return false;
    if (entry.status !== 'active') return false;
    this.preflightHistory(entry.inversePlan, 'undo');
    this.applyHistory(entry.inversePlan, 'undo');
    this.undoStack.pop();
    this.redoStack.push(entry);
    for (const listener of this.historyReplayListeners) listener('undo', entry);
    return true;
  }

  redo(): boolean {
    this.registry.assertComplete();
    const entry = this.redoStack[this.redoStack.length - 1];
    if (!entry) return false;
    if (entry.status !== 'active') return false;
    this.preflightHistory(entry.forwardMutations, 'redo');
    this.applyHistory(entry.forwardMutations, 'redo');
    this.redoStack.pop();
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
    if (remoteContext.revision !== undefined
      && (!Number.isSafeInteger(remoteContext.revision) || remoteContext.revision < 1)) {
      throw new Error('Remote revision is invalid');
    }
    // A committed operation may contain several dependent mutations. Replay
    // them against an isolated snapshot first so a later rejection cannot
    // leave the live workbook partially changed.
    this.preflightHistory(items, 'remote');
    this.applyHistory(items, 'remote');
    for (const item of items) {
      const remote = item.structuralImpactRanges?.length
        ? { ...item, affectedRanges: [...item.affectedRanges, ...item.structuralImpactRanges] }
        : item;
      this.transformHistoryAgainstRemote(remote);
    }
    if (remoteContext.revision !== undefined) {
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

  applyCommittedStructuralPatches(operationId: string, items: readonly MutationInfo[], revision: number): void {
    if (!Number.isSafeInteger(revision) || revision < 1) throw new Error('Committed revision must be a positive safe integer');
    const patched = items.filter((item) => (item.structuralFormulaOwnerDeltas?.length ?? 0) > 0
      || (item.structuralDefinedNameOwnerDeltas?.length ?? 0) > 0);
    if (patched.length === 0) {
      this.setRevision(Math.max(this.currentRevision, revision));
      return;
    }

    const entry = [...this.undoStack, ...this.redoStack].find((candidate) => candidate.operationId === operationId);
    if (entry) {
      const local = entry.inversePlan.flatMap((mutation) => mutation.structuralFormulaOwnerDeltas ?? []);
      const authoritative = patched.flatMap((item) => item.structuralFormulaOwnerDeltas ?? []);
      const ordered = (deltas: readonly StructuralFormulaOwnerDelta[]) => deltas.map((delta) => JSON.stringify(delta)).sort();
      const localNames = entry.inversePlan.flatMap((mutation) => mutation.structuralDefinedNameOwnerDeltas ?? []);
      const authoritativeNames = patched.flatMap((item) => item.structuralDefinedNameOwnerDeltas ?? []);
      const orderedNames = (deltas: readonly StructuralDefinedNameOwnerDelta[]) => deltas.map((delta) => JSON.stringify(delta)).sort();
      if (JSON.stringify(ordered(local)) !== JSON.stringify(ordered(authoritative))
        || JSON.stringify(orderedNames(localNames)) !== JSON.stringify(orderedNames(authoritativeNames))) {
        const stack = this.undoStack.includes(entry) ? this.undoStack : this.redoStack;
        const index = stack.indexOf(entry);
        if (index >= 0) stack.splice(index, 1);
        entry.status = 'invalid';
        entry.invalidReason = 'Server-derived structural owners differ from the local history patch';
        this.invalidHistory.push(entry);
      }
    }

    preflightCommittedStructuralPatches(this.workbook, patched);

    for (const item of patched) {
      const formulaDeltas = item.structuralFormulaOwnerDeltas ?? [];
      const definedNameDeltas = item.structuralDefinedNameOwnerDeltas ?? [];
      if (formulaDeltas.length === 0 && definedNameDeltas.length === 0) continue;
      for (const delta of formulaDeltas) applyFormulaOwnerDelta(this.workbook, delta, 'forward');
      for (const delta of definedNameDeltas) applyDefinedNameOwnerDelta(this.workbook, delta, 'forward');
      const effect = {
        kind: 'structural-transform' as const,
        removedCells: [],
        clearInputRanges: [],
        populateInputRanges: [],
        rewrittenFormulaOwners: formulaDeltas.flatMap((delta) => delta.kind === 'formula-cell' ? [delta.afterAddress] : []),
        formulaOwnerDeltas: formulaDeltas,
        definedNameOwnerDeltas: definedNameDeltas,
      };
      for (const listener of this.mutationListeners) listener(item, 'remote', effect);
    }

    this.setRevision(Math.max(this.currentRevision, revision));
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
    const sheetOrder = this.workbook.sheetOrder.map((id) => ({ id, name: this.workbook.getSheet(id).name }));
    const policy = this.registry.getMutationMetadata(remote.id).historyRebase;
    for (const stack of stacks) {
      for (let index = stack.length - 1; index >= 0; index -= 1) {
        const entry = stack[index]!;
        const transformed = transformHistoryEntry(entry, remote, sheetOrder, policy);
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
      const commandRuntime = this;
      const replayContext: CommandContext = {
        workbook: this.workbook,
        operationId: createOperationId(),
        get structuralReferenceOwners() { return commandRuntime.resolveStructuralReferenceOwners(); },
        resolveCellValue: (sheet, row, column) => this.cellValueResolver?.(sheet, row, column),
        applyMutation: () => {
          throw new Error('Nested mutation application is not allowed during mutation replay');
        },
        recordOperation: (operation, operationParams) => {
          const registered = this.registry.getOperation(operation.id);
          return registered.execute(operationParams, replayContext);
        },
      };
      const effect = handler(item, {
        ...replayContext,
      }) ?? this.registry.getMutationMetadata(item.id).calculationContextEffect;
      let notificationEffect: unknown = effect;
      const replaysOwnerFacts = source === 'undo' || source === 'redo' || source === 'remote';
      const formulaOwnerDeltas = replaysOwnerFacts ? item.structuralFormulaOwnerDeltas ?? [] : [];
      if (formulaOwnerDeltas.length > 0) {
        const direction = source === 'undo' ? 'undo' : 'forward';
        for (const delta of formulaOwnerDeltas) applyFormulaOwnerDelta(this.workbook, delta, direction);
      }
      let notificationFormulaOwnerDeltas = formulaOwnerDeltas;
      if (source === 'undo') notificationFormulaOwnerDeltas = formulaOwnerDeltas.map(inverseFormulaOwnerDelta);
      let notificationDefinedNameOwnerDeltas: readonly StructuralDefinedNameOwnerDelta[] = [];
      if (item.structuralDefinedNameOwnerDeltas && (source === 'undo' || source === 'redo' || source === 'remote')) {
        const direction = source === 'undo' ? 'undo' : 'forward';
        for (const delta of item.structuralDefinedNameOwnerDeltas) applyDefinedNameOwnerDelta(this.workbook, delta, direction);
        notificationDefinedNameOwnerDeltas = source === 'undo'
          ? item.structuralDefinedNameOwnerDeltas.map(inverseDefinedNameOwnerDelta)
          : item.structuralDefinedNameOwnerDeltas;
      }
      if (notificationFormulaOwnerDeltas.length > 0 || notificationDefinedNameOwnerDeltas.length > 0) {
        notificationEffect = structuralOwnerPatchReplayEffect(
          effect,
          notificationFormulaOwnerDeltas,
          notificationDefinedNameOwnerDeltas,
        );
      }
      for (const listener of this.mutationListeners) {
        listener(item, source, notificationEffect);
      }
    }
  }

  private preflightHistory(items: readonly MutationInfo[], source: MutationSource): void {
    const preview = new CommandRuntime(WorkbookModel.fromSnapshot(this.workbook.snapshot()), this.registry);
    preview.setStructuralReferenceOwnersProvider((workbook) => buildStructuralReferenceIndex(workbook));
    preview.applyHistory(items, source);
  }
}

function remoteChangesDefinedNameOwner(
  remote: MutationInfo,
  owner: StructuralDefinedNameOwnerDelta['owner'],
): boolean {
  if (remote.id === 'name.set' || remote.id === 'name.remove') {
    const params = isRecord(remote.params) ? remote.params : undefined;
    const candidate = remote.id === 'name.set'
      ? params && isRecord(params.model) ? params.model : undefined
      : params;
    if (!candidate || typeof candidate.name !== 'string') return true;
    const scope = candidate.scope ?? 'workbook';
    if (scope !== 'workbook' && scope !== 'sheet') return true;
    if (candidate.sheetId !== undefined && typeof candidate.sheetId !== 'string') return true;
    const sheetId = candidate.sheetId as string | undefined;
    if ((scope === 'sheet' && !sheetId) || (scope === 'workbook' && sheetId !== undefined)) return true;
    return owner.scope === scope
      && owner.name.trim().toUpperCase() === candidate.name.trim().toUpperCase()
      && (scope !== 'sheet' || owner.sheetId === sheetId);
  }
  return ['sheet.rename', 'sheet.remove', 'sheet.restore', 'sheet.duplicated', 'workbook.restore'].includes(remote.id);
}

function inverseDefinedNameOwnerDelta(delta: StructuralDefinedNameOwnerDelta): StructuralDefinedNameOwnerDelta {
  return {
    owner: structuredClone(delta.owner),
    before: structuredClone(delta.after),
    after: structuredClone(delta.before),
  };
}

function inverseFormulaOwnerDelta(delta: StructuralFormulaOwnerDelta): StructuralFormulaOwnerDelta {
  if (delta.kind === 'formula-cell') {
    return {
      ...delta,
      beforeAddress: structuredClone(delta.afterAddress),
      afterAddress: structuredClone(delta.beforeAddress),
      before: structuredClone(delta.after),
      after: structuredClone(delta.before),
    };
  }
  if (delta.kind === 'formula-rule') {
    return {
      ...delta,
      beforeFormula: delta.afterFormula,
      afterFormula: delta.beforeFormula,
      beforeRanges: structuredClone(delta.afterRanges),
      afterRanges: structuredClone(delta.beforeRanges),
    };
  }
  return { ...delta, beforeFormula: delta.afterFormula, afterFormula: delta.beforeFormula };
}

function structuralOwnerPatchReplayEffect(
  effect: unknown,
  formulaDeltas: readonly StructuralFormulaOwnerDelta[],
  definedNameDeltas: readonly StructuralDefinedNameOwnerDelta[],
): unknown {
  const calculationContextEffect = isWorkbookCalculationContextEffect(effect)
    ? effect
    : isRecord(effect) && isWorkbookCalculationContextEffect(effect.calculationContextEffect)
      ? effect.calculationContextEffect
      : undefined;
  const existing = isRecord(effect) && !isWorkbookCalculationContextEffect(effect) ? effect : {};
  const existingFormulaDeltas = Array.isArray(existing.formulaOwnerDeltas)
    ? existing.formulaOwnerDeltas as StructuralFormulaOwnerDelta[]
    : [];
  const replayFormulaDeltas = formulaDeltas.length > 0 ? formulaDeltas : existingFormulaDeltas;
  const existingDefinedNameDeltas = Array.isArray(existing.definedNameOwnerDeltas)
    ? existing.definedNameOwnerDeltas as StructuralDefinedNameOwnerDelta[]
    : [];
  const replayDefinedNameDeltas = definedNameDeltas.length > 0 ? definedNameDeltas : existingDefinedNameDeltas;
  const rewrittenFormulaOwners = [
    ...(Array.isArray(existing.rewrittenFormulaOwners) ? existing.rewrittenFormulaOwners : []),
    ...replayFormulaDeltas.flatMap((delta) => delta.kind === 'formula-cell' ? [structuredClone(delta.afterAddress)] : []),
  ];
  const uniqueRewrittenFormulaOwners = new Map<string, unknown>();
  for (const owner of rewrittenFormulaOwners) {
    const key = JSON.stringify(owner);
    if (key === undefined) throw new Error('STRUCTURAL_PATCH_INVARIANT: formula owner address is not serializable');
    uniqueRewrittenFormulaOwners.set(key, owner);
  }
  return {
    ...existing,
    kind: 'structural-transform' as const,
    removedCells: Array.isArray(existing.removedCells) ? existing.removedCells : [],
    clearInputRanges: Array.isArray(existing.clearInputRanges) ? existing.clearInputRanges : [],
    populateInputRanges: Array.isArray(existing.populateInputRanges) ? existing.populateInputRanges : [],
    rewrittenFormulaOwners: [...uniqueRewrittenFormulaOwners.values()],
    ...(replayFormulaDeltas.length > 0 ? { formulaOwnerDeltas: structuredClone(replayFormulaDeltas) } : {}),
    ...(replayDefinedNameDeltas.length > 0 ? { definedNameOwnerDeltas: structuredClone(replayDefinedNameDeltas) } : {}),
    ...(calculationContextEffect ? { calculationContextEffect } : {}),
  };
}

function formulaOwnerDeltasRanges(deltas: readonly StructuralFormulaOwnerDelta[]): RangeRef[] {
  const ranges = new Map<string, RangeRef>();
  for (const delta of deltas) {
    const affected = delta.kind === 'formula-rule'
      ? [...delta.beforeRanges, ...delta.afterRanges]
      : delta.kind === 'formula-cell'
        ? [delta.beforeAddress, delta.afterAddress].map((address) => ({
          sheetId: address.sheetId,
          startRow: address.row,
          endRow: address.row,
          startColumn: address.column,
          endColumn: address.column,
        }))
        : [];
    for (const range of affected) {
      ranges.set(JSON.stringify([range.sheetId, range.startRow, range.endRow, range.startColumn, range.endColumn]), structuredClone(range));
    }
  }
  return [...ranges.values()];
}

function formulaOwnerState(cell: CellData): StructuralFormulaOwnerState {
  return {
    formula: cell.formula ?? null,
    sourceFormula: cell.formulaMetadata?.sourceFormula ?? null,
    barcodeFormula: cell.presentation?.kind === 'barcode' && cell.presentation.source.kind === 'formula'
      ? cell.presentation.source.formula
      : null,
  };
}

function sameFormulaOwnerState(left: StructuralFormulaOwnerState, right: StructuralFormulaOwnerState): boolean {
  return left.formula === right.formula
    && left.sourceFormula === right.sourceFormula
    && left.barcodeFormula === right.barcodeFormula;
}

function prepareFormulaCellOwnerUpdate(
  cell: CellData,
  delta: Extract<StructuralFormulaOwnerDelta, { kind: 'formula-cell' }>,
  direction: 'undo' | 'forward',
): CellData | undefined {
  const expected = direction === 'undo' ? delta.after : delta.before;
  const target = direction === 'undo' ? delta.before : delta.after;
  const current = formulaOwnerState(cell);
  if (sameFormulaOwnerState(current, target)) {
    if (target.formula === null || cell.formulaValue === undefined) return undefined;
    const next = { ...cell };
    delete next.formulaValue;
    return normalizeCellDataForStorage(next);
  }
  const { sheetId, row, column } = direction === 'undo' ? delta.beforeAddress : delta.afterAddress;
  if (!sameFormulaOwnerState(current, expected)) {
    throw new Error(`STRUCTURAL_PATCH_PRECONDITION: formula owner ${sheetId}!${row}:${column} changed since the structural operation`);
  }

  const next: CellData = { ...cell };
  if (target.formula === null) delete next.formula;
  else next.formula = target.formula;
  if (target.formula !== null) delete next.formulaValue;
  if (target.sourceFormula === null) {
    if (next.formulaMetadata) {
      const metadata = { ...next.formulaMetadata };
      delete metadata.sourceFormula;
      next.formulaMetadata = metadata;
    }
  } else {
    if (!next.formulaMetadata) {
      throw new Error(`STRUCTURAL_PATCH_INVARIANT: formula provenance owner ${sheetId}!${row}:${column} changed type`);
    }
    next.formulaMetadata = { ...next.formulaMetadata, sourceFormula: target.sourceFormula };
  }
  if (target.barcodeFormula !== null) {
    if (next.presentation?.kind !== 'barcode' || next.presentation.source.kind !== 'formula') {
      throw new Error(`STRUCTURAL_PATCH_INVARIANT: barcode formula owner ${sheetId}!${row}:${column} changed type`);
    }
    next.presentation = { ...next.presentation, source: { ...next.presentation.source, formula: target.barcodeFormula } };
  } else if (current.barcodeFormula !== null && expected.barcodeFormula !== null) {
    throw new Error(`STRUCTURAL_PATCH_INVARIANT: barcode formula owner ${sheetId}!${row}:${column} cannot be removed by a reference delta`);
  }
  return normalizeCellDataForStorage(next);
}

type FormulaPatchState =
  | { readonly kind: 'formula-cell'; readonly cell: CellData }
  | { readonly kind: 'formula-rule'; readonly formula: string | number | undefined; readonly ranges: readonly RangeRef[] }
  | { readonly kind: 'formula-object'; readonly formula: string | undefined };

function formulaOwnerPatchKey(delta: StructuralFormulaOwnerDelta): string {
  if (delta.kind === 'formula-cell') {
    const { sheetId, row, column } = delta.afterAddress;
    return JSON.stringify(['formula-cell', sheetId, row, column]);
  }
  if (delta.kind === 'formula-rule') return JSON.stringify(['formula-rule', delta.sheetId, delta.ruleKind, delta.ruleId, delta.field]);
  switch (delta.ownerKind) {
    case 'chart-text': return JSON.stringify([delta.kind, delta.ownerKind, delta.sheetId, delta.payloadId, delta.field]);
    case 'shape-property': return JSON.stringify([delta.kind, delta.ownerKind, delta.sheetId, delta.payloadId]);
    case 'table-sheet-column': return JSON.stringify([delta.kind, delta.ownerKind, delta.sheetId, delta.fieldId]);
    case 'data-view-field': return JSON.stringify([delta.kind, delta.ownerKind, delta.viewId, delta.fieldId]);
    case 'cell-style-template': return JSON.stringify([delta.kind, delta.ownerKind, delta.templateId, delta.field]);
  }
}

function readFormulaRulePatchState(workbook: WorkbookModel, delta: Extract<StructuralFormulaOwnerDelta, { kind: 'formula-rule' }>): FormulaPatchState {
  const sheet = workbook.getSheet(delta.sheetId);
  const rules = delta.ruleKind === 'conditional-format' ? sheet.conditionalFormats : sheet.dataValidations;
  const matches = rules.filter((rule) => rule.id === delta.ruleId && rule.sheetId === delta.sheetId);
  if (matches.length !== 1) {
    throw new Error(`STRUCTURAL_PATCH_PRECONDITION: expected one ${delta.ruleKind} rule ${delta.sheetId}:${delta.ruleId}, found ${matches.length}`);
  }
  const rule = matches[0]!;
  const conditionalFormat = delta.ruleKind === 'conditional-format' ? rule as ConditionalFormatRule : undefined;
  const dataValidation = delta.ruleKind === 'data-validation' ? rule as DataValidationRule : undefined;
  const formula = delta.field === 'value1' ? conditionalFormat?.value1
    : delta.field === 'value2' ? conditionalFormat?.value2
      : delta.field === 'formula1' ? dataValidation?.formula1
        : delta.field === 'formula2' ? dataValidation?.formula2
          : dataValidation?.listSource?.kind === 'formula' ? dataValidation.listSource.formula : undefined;
  return { kind: 'formula-rule', formula, ranges: rule.ranges };
}

function readFormulaPatchState(workbook: WorkbookModel, delta: StructuralFormulaOwnerDelta): FormulaPatchState {
  if (delta.kind === 'formula-rule') return readFormulaRulePatchState(workbook, delta);
  if (delta.kind === 'formula-object') return { kind: 'formula-object', formula: readFormulaObjectOwner(workbook, delta) };
  const { sheetId, row, column } = delta.afterAddress;
  const cell = workbook.getSheet(sheetId).cells.getWithoutHydration(row, column);
  if (!cell) throw new Error(`STRUCTURAL_PATCH_PRECONDITION: formula owner ${sheetId}!${row}:${column} is missing after inverse`);
  return { kind: 'formula-cell', cell };
}

function preflightCommittedStructuralPatches(workbook: WorkbookModel, items: readonly MutationInfo[]): void {
  const formulaStates = new Map<string, FormulaPatchState>();
  const definedNameStates = new Map<string, ReturnType<WorkbookModel['getDefinedNameExact']>>();
  for (const item of items) {
    for (const delta of item.structuralFormulaOwnerDeltas ?? []) {
      const key = formulaOwnerPatchKey(delta);
      const state = formulaStates.get(key) ?? readFormulaPatchState(workbook, delta);
      if (delta.kind === 'formula-cell') {
        if (state.kind !== 'formula-cell') throw new Error('STRUCTURAL_PATCH_INVARIANT: formula-cell owner key collision');
        const nextCell = prepareFormulaCellOwnerUpdate(state.cell, delta, 'forward');
        formulaStates.set(key, { kind: 'formula-cell', cell: nextCell ?? state.cell });
        continue;
      }
      if (delta.kind === 'formula-rule') {
        if (state.kind !== 'formula-rule') throw new Error('STRUCTURAL_PATCH_INVARIANT: formula-rule owner key collision');
        if (JSON.stringify(state.ranges) !== JSON.stringify(delta.afterRanges)) {
          throw new Error(`STRUCTURAL_PATCH_PRECONDITION: ${delta.ruleKind} rule ${delta.sheetId}:${delta.ruleId}.${delta.field} changed since the structural operation`);
        }
        if (state.formula !== delta.afterFormula && state.formula !== delta.beforeFormula) {
          throw new Error(`STRUCTURAL_PATCH_PRECONDITION: ${delta.ruleKind} rule ${delta.sheetId}:${delta.ruleId}.${delta.field} changed since the structural operation`);
        }
        formulaStates.set(key, { kind: 'formula-rule', formula: delta.afterFormula, ranges: delta.afterRanges });
        continue;
      }
      if (state.kind !== 'formula-object') throw new Error('STRUCTURAL_PATCH_INVARIANT: formula-object owner key collision');
      if (state.formula !== delta.afterFormula && state.formula !== delta.beforeFormula) {
        const ownerIdentity = delta.ownerKind === 'data-view-field'
          ? `${delta.viewId}:${delta.fieldId}`
          : delta.ownerKind === 'cell-style-template'
            ? `${delta.templateId}.${delta.field}`
            : delta.ownerKind === 'table-sheet-column'
              ? `${delta.sheetId}:${delta.fieldId}`
              : `${delta.sheetId}:${delta.payloadId}`;
        throw new Error(`STRUCTURAL_PATCH_PRECONDITION: ${delta.ownerKind} formula owner ${ownerIdentity} changed since the structural operation`);
      }
      formulaStates.set(key, { kind: 'formula-object', formula: delta.afterFormula });
    }
    for (const delta of item.structuralDefinedNameOwnerDeltas ?? []) {
      const key = JSON.stringify([
        delta.owner.scope,
        delta.owner.name.trim().toUpperCase(),
        delta.owner.sheetId ?? null,
      ]);
      const current = definedNameStates.has(key)
        ? definedNameStates.get(key)
        : workbook.getDefinedNameExact(delta.owner.name, delta.owner.scope, delta.owner.sheetId);
      const next = prepareDefinedNameOwnerUpdate(current, delta, 'forward');
      definedNameStates.set(key, next ? normalizeDefinedNameModel(next) : current);
    }
  }
}

function definedNameStateMatches(
  current: ReturnType<WorkbookModel['getDefinedNameExact']>,
  owner: StructuralDefinedNameOwnerDelta['owner'],
  target: StructuralDefinedNameOwnerDelta['before'],
): boolean {
  if (!current
    || current.scope !== owner.scope
    || current.sheetId !== owner.sheetId
    || current.name.trim().toUpperCase() !== owner.name.trim().toUpperCase()
    || target.scope !== owner.scope
    || target.sheetId !== owner.sheetId
    || target.name.trim().toUpperCase() !== owner.name.trim().toUpperCase()
    || current.formula !== target.formula) return false;
  const currentAnchor = current.anchor;
  const targetAnchor = target.anchor;
  return currentAnchor === undefined
    ? targetAnchor === undefined
    : targetAnchor !== undefined
      && currentAnchor.sheetId === targetAnchor.sheetId
      && currentAnchor.row === targetAnchor.row
      && currentAnchor.column === targetAnchor.column;
}

function prepareDefinedNameOwnerUpdate(
  current: ReturnType<WorkbookModel['getDefinedNameExact']>,
  delta: StructuralDefinedNameOwnerDelta,
  direction: 'undo' | 'forward',
): ReturnType<WorkbookModel['getDefinedNameExact']> | undefined {
  const expected = direction === 'undo' ? delta.after : delta.before;
  const target = direction === 'undo' ? delta.before : delta.after;
  if (definedNameStateMatches(current, delta.owner, target)) return undefined;
  if (!definedNameStateMatches(current, delta.owner, expected)) {
    throw new Error(`STRUCTURAL_PATCH_PRECONDITION: defined-name owner ${delta.owner.scope}:${delta.owner.sheetId ?? '*'}:${delta.owner.name} changed since the structural operation`);
  }
  return {
    ...current!,
    formula: target.formula,
    anchor: target.anchor ? structuredClone(target.anchor) : undefined,
  };
}

function applyDefinedNameOwnerDelta(
  workbook: WorkbookModel,
  delta: StructuralDefinedNameOwnerDelta,
  direction: 'undo' | 'forward',
): void {
  const current = workbook.getDefinedNameExact(delta.owner.name, delta.owner.scope, delta.owner.sheetId);
  const next = prepareDefinedNameOwnerUpdate(current, delta, direction);
  if (next) workbook.setDefinedName(next);
}

function readFormulaObjectOwner(workbook: WorkbookModel, delta: Extract<StructuralFormulaOwnerDelta, { kind: 'formula-object' }>): string | undefined {
  switch (delta.ownerKind) {
    case 'chart-text': {
      const payload = workbook.getSheet(delta.sheetId).drawingPayloads.get(delta.payloadId);
      return payload?.kind === 'chart' ? readChartTextFormula(payload, delta.field) : undefined;
    }
    case 'shape-property': {
      const payload = workbook.getSheet(delta.sheetId).drawingPayloads.get(delta.payloadId);
      return payload?.kind === 'shape' ? payload.propertyFormula : undefined;
    }
    case 'table-sheet-column': {
      const columns = workbook.getSheet(delta.sheetId).tableSheet?.columns.filter((column) => column.fieldId === delta.fieldId) ?? [];
      return columns.length === 1 ? columns[0]!.formula : undefined;
    }
    case 'data-view-field': {
      const fields = workbook.dataModel.views.get(delta.viewId)?.fields.filter((field) => field.fieldId === delta.fieldId) ?? [];
      return fields.length === 1 ? fields[0]!.formula : undefined;
    }
    case 'cell-style-template': {
      const validation = workbook.cellStyleTemplates.get(delta.templateId)?.dataValidation;
      if (delta.field === 'formula1') return validation?.formula1;
      if (delta.field === 'formula2') return validation?.formula2;
      return validation?.listSource?.kind === 'formula' ? validation.listSource.formula : undefined;
    }
  }
}

function writeFormulaObjectOwner(
  workbook: WorkbookModel,
  delta: Extract<StructuralFormulaOwnerDelta, { kind: 'formula-object' }>,
  formula: string,
): boolean {
  switch (delta.ownerKind) {
    case 'chart-text': {
      const payload = workbook.getSheet(delta.sheetId).drawingPayloads.get(delta.payloadId);
      if (!payload || payload.kind !== 'chart') return false;
      writeChartTextFormula(payload, delta.field, formula);
      return true;
    }
    case 'shape-property': {
      const payload = workbook.getSheet(delta.sheetId).drawingPayloads.get(delta.payloadId);
      if (!payload || payload.kind !== 'shape') return false;
      payload.propertyFormula = formula;
      return true;
    }
    case 'table-sheet-column': {
      const columns = workbook.getSheet(delta.sheetId).tableSheet?.columns.filter((column) => column.fieldId === delta.fieldId) ?? [];
      if (columns.length !== 1) return false;
      columns[0]!.formula = formula;
      return true;
    }
    case 'data-view-field': {
      const fields = workbook.dataModel.views.get(delta.viewId)?.fields.filter((field) => field.fieldId === delta.fieldId) ?? [];
      if (fields.length !== 1) return false;
      fields[0]!.formula = formula;
      return true;
    }
    case 'cell-style-template': {
      const validation = workbook.cellStyleTemplates.get(delta.templateId)?.dataValidation;
      if (!validation) return false;
      if (delta.field === 'formula1') validation.formula1 = formula;
      else if (delta.field === 'formula2') validation.formula2 = formula;
      else if (validation.listSource?.kind === 'formula') validation.listSource.formula = formula;
      else return false;
      return true;
    }
  }
}

function applyFormulaOwnerDelta(
  workbook: WorkbookModel,
  delta: StructuralFormulaOwnerDelta,
  direction: 'undo' | 'forward',
): void {
  if (delta.kind === 'formula-object') {
    const expectedFormula = direction === 'undo' ? delta.afterFormula : delta.beforeFormula;
    const targetFormula = direction === 'undo' ? delta.beforeFormula : delta.afterFormula;
    const currentFormula = readFormulaObjectOwner(workbook, delta);
    const ownerIdentity = delta.ownerKind === 'data-view-field'
      ? `${delta.viewId}:${delta.fieldId}`
      : delta.ownerKind === 'cell-style-template'
        ? `${delta.templateId}.${delta.field}`
        : delta.ownerKind === 'table-sheet-column'
          ? `${delta.sheetId}:${delta.fieldId}`
          : `${delta.sheetId}:${delta.payloadId}`;
    if (currentFormula === targetFormula) return;
    if (currentFormula !== expectedFormula) {
      throw new Error(`STRUCTURAL_PATCH_PRECONDITION: ${delta.ownerKind} formula owner ${ownerIdentity} changed since the structural operation`);
    }
    if (!writeFormulaObjectOwner(workbook, delta, targetFormula)) {
      throw new Error(`STRUCTURAL_PATCH_PRECONDITION: ${delta.ownerKind} formula owner ${ownerIdentity} is missing or ambiguous`);
    }
    return;
  }
  if (delta.kind === 'formula-rule') {
    const sheet = workbook.getSheet(delta.sheetId);
    const rules = delta.ruleKind === 'conditional-format' ? sheet.conditionalFormats : sheet.dataValidations;
    const matches = rules.filter((rule) => rule.id === delta.ruleId && rule.sheetId === delta.sheetId);
    if (matches.length !== 1) {
      throw new Error(`STRUCTURAL_PATCH_PRECONDITION: expected one ${delta.ruleKind} rule ${delta.sheetId}:${delta.ruleId}, found ${matches.length}`);
    }
    const rule = matches[0]!;
    const conditionalFormat = delta.ruleKind === 'conditional-format' ? rule as ConditionalFormatRule : undefined;
    const dataValidation = delta.ruleKind === 'data-validation' ? rule as DataValidationRule : undefined;
    const currentFormula = delta.field === 'value1' ? conditionalFormat?.value1
      : delta.field === 'value2' ? conditionalFormat?.value2
        : delta.field === 'formula1' ? dataValidation?.formula1
          : delta.field === 'formula2' ? dataValidation?.formula2
            : dataValidation?.listSource?.kind === 'formula' ? dataValidation.listSource.formula : undefined;
    const expectedFormula = direction === 'undo' ? delta.afterFormula : delta.beforeFormula;
    const targetFormula = direction === 'undo' ? delta.beforeFormula : delta.afterFormula;
    const targetRanges = direction === 'undo' ? delta.beforeRanges : delta.afterRanges;
    const sameRanges = (left: readonly RangeRef[], right: readonly RangeRef[]) => JSON.stringify(left) === JSON.stringify(right);
    if (!sameRanges(rule.ranges, targetRanges)) {
      throw new Error(`STRUCTURAL_PATCH_PRECONDITION: ${delta.ruleKind} rule ${delta.sheetId}:${delta.ruleId}.${delta.field} changed since the structural operation`);
    }
    if (currentFormula === targetFormula) return;
    if (currentFormula !== expectedFormula) {
      throw new Error(`STRUCTURAL_PATCH_PRECONDITION: ${delta.ruleKind} rule ${delta.sheetId}:${delta.ruleId}.${delta.field} changed since the structural operation`);
    }
    if (delta.field === 'value1') {
      if (!conditionalFormat) throw new Error('STRUCTURAL_PATCH_INVARIANT: value1 formula field is not owned by data validation');
      conditionalFormat.value1 = targetFormula;
    } else if (delta.field === 'value2') {
      if (!conditionalFormat) throw new Error('STRUCTURAL_PATCH_INVARIANT: value2 formula field is not owned by data validation');
      conditionalFormat.value2 = targetFormula;
    } else if (!dataValidation) {
      throw new Error(`STRUCTURAL_PATCH_INVARIANT: ${delta.field} is not owned by conditional formatting`);
    } else if (delta.field === 'formula1') dataValidation.formula1 = targetFormula;
    else if (delta.field === 'formula2') dataValidation.formula2 = targetFormula;
    else if (delta.field === 'listSource.formula') {
      if (dataValidation.listSource?.kind !== 'formula') throw new Error('STRUCTURAL_PATCH_INVARIANT: data-validation list formula owner changed type');
      dataValidation.listSource = { ...dataValidation.listSource, formula: targetFormula };
    }
    return;
  }
  const address = direction === 'undo' ? delta.beforeAddress : delta.afterAddress;
  const { sheetId, row, column } = address;
  const sheet = workbook.getSheet(sheetId);
  const cell = sheet.cells.getWithoutHydration(row, column);
  if (!cell) throw new Error(`STRUCTURAL_PATCH_PRECONDITION: formula owner ${sheetId}!${row}:${column} is missing after inverse`);
  const next = prepareFormulaCellOwnerUpdate(cell, delta, direction);
  if (!next) return;
  if (!sheet.cells.replaceCellWithoutHydration(row, column, next)) {
    throw new Error(`STRUCTURAL_PATCH_PRECONDITION: formula owner ${sheetId}!${row}:${column} is missing after inverse`);
  }
}
