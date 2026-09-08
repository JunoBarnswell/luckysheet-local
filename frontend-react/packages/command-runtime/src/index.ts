import { WorkbookModel, type KernelReplicaManifest, type KernelReplicaPagePayload, type ProtectionAction, type RangeRef, type WorksheetModel } from '@react-sheets/core-model';

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
}

/** Short public name for feature packages that expose a mutation contract. */
export type MutationMetadata<P = unknown> = MutationRegistrationMetadata<P>;

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
   * Commands that only publish ephemeral UI events may stay out of edit
   * history. Any command with workbook mutations is committed atomically by
   * the cloud kernel before the browser replica changes.
   */
  history?: 'record' | 'none';
  execute(params: P, context: CommandContext): CommandResult;
}

export interface Operation<P = unknown> {
  id: string;
  execute(params: P, context: CommandContext): OperationResult;
}

export interface Mutation<P = unknown> extends MutationInfo<P> {
}

export interface CommandContext {
  readonly workbook: WorkbookModel;
  readonly operationId: string;
  /** Optional canonical worksheet-value authority supplied by the host runtime. */
  readonly resolveCellValue?: (sheet: WorksheetModel, row: number, column: number) => unknown;
  /**
   * Compose another command into the current plan. The nested command shares
   * this operation identity and may only append mutations to this transaction;
   * it never enters the serialized commit queue on its own.
   */
  executeCommand<P>(commandId: string, params: P): CommandResult;
  applyMutation<P>(mutation: Mutation<P>): void;
  recordOperation<P>(operation: Operation<P>, params: P): OperationResult;
}

export type MutationHandler<P = unknown> = (item: MutationInfo<P>, context: CommandContext) => void;

export interface MutationRegistration<P = unknown> {
  readonly id: string;
  readonly metadata: MutationRegistrationMetadata<P>;
}

/** Constructor options are intentionally non-permissive; metadata is always required. */
export interface CommandRegistryOptions {
  readonly requireMutationMetadata?: true;
}

export type CanonicalMutationMetadata<P = unknown> = MutationRegistrationMetadata<P>;

type RegisteredMutation<P = unknown> = Omit<MutationRegistration<P>, 'metadata'> & {
  readonly metadata: CanonicalMutationMetadata<P>;
};

export type MutationRegistryIssueCode =
  | 'invalid-registration'
  | 'missing-schema'
  | 'missing-permission'
  | 'missing-affected-ranges'
  | 'unknown-mutation'
  | 'invalid-params'
  | 'invalid-affected-ranges'
  ;

export interface MutationRegistryIssue {
  readonly code: MutationRegistryIssueCode;
  readonly mutationId: string;
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
): MutationRegistryIssue {
  return { code, mutationId, message };
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
  if (issues.length > 0) return { issues };
  return {
    metadata: metadata as unknown as CanonicalMutationMetadata,
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
    if (this.mutations.has(candidate.id)) throw new Error(`Duplicate mutation: ${candidate.id}`);
    this.mutations.set(candidate.id, {
      id: candidate.id,
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
      const { id, metadata } = registration;
      if (!id || typeof id !== 'string') {
        issues.push(issue('invalid-registration', id || '<empty>', `Invalid mutation registration: ${id || '<empty>'}`));
        continue;
      }
      const normalized = validateRegistrationMetadata(id, metadata);
      issues.push(...normalized.issues);
      if (!normalized.metadata) continue;
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
    return issues;
  }

  assertMutation<P>(mutation: Mutation<P>): void {
    const issues = this.validateMutation(mutation);
    if (issues.length > 0) throw new Error(`Invalid mutation ${mutation.id}: ${formatIssues(issues)}`);
  }

  validateMutationInfo<P>(
    item: MutationInfo<P>,
    issues: MutationRegistryIssue[] = [],
  ): MutationRegistryIssue[] {
    const registration = this.mutations.get(item.id);
    if (!registration) {
      issues.push(issue('unknown-mutation', item.id, `Unknown mutation: ${item.id}`));
      return issues;
    }

    if (!isRecord(item) || typeof item.unitId !== 'string' || !item.unitId || typeof item.sheetId !== 'string' || !item.sheetId) {
      issues.push(issue('invalid-registration', item.id, `Mutation ${item.id} has invalid unitId or sheetId`));
    }
    if (!Array.isArray(item.affectedRanges) || !item.affectedRanges.every(isValidRangeRef)) {
      issues.push(issue('invalid-affected-ranges', item.id, `Mutation ${item.id} has invalid affected ranges`));
    }

    const schema = registration.metadata.schema;
    let valid = false;
    try {
      valid = schema.validate(item.params);
    } catch {
      valid = false;
    }
    if (!valid) {
      issues.push(issue('invalid-params', item.id, `Mutation ${item.id} parameters do not match ${schema.name ?? 'its schema'}`));
    }

    const affectedRanges = registration.metadata.affectedRanges;
    if (Array.isArray(item.affectedRanges)) {
      try {
        const declared = affectedRanges.resolve(item.params as never);
        if (!Array.isArray(declared) || !declared.every(isValidRangeRef)) {
          issues.push(issue('invalid-affected-ranges', item.id, `Mutation ${item.id} declared an invalid affected-range result`));
        } else if ((affectedRanges.mode ?? 'exact') === 'exact' && !rangesEqual(item.affectedRanges, declared)) {
          issues.push(issue('invalid-affected-ranges', item.id, `Mutation ${item.id} affected ranges differ from its declaration`));
        }
      } catch {
        issues.push(issue('invalid-affected-ranges', item.id, `Mutation ${item.id} affected-range resolver failed`));
      }
    }
    return issues;
  }
}

export interface HistoryEntry {
  /** Identity of the original semantic command shown to the user. */
  operationId: string;
  baseRevision: number;
  committedRevision: number;
  /**
   * The immutable committed operation that the next undo/redo request targets.
   * Every replay is itself a committed operation, so this target advances after
   * each successful replay without reconstructing client-side inverse data.
   */
  replayTargetOperationId: string;
  replayTargetBaseRevision: number;
  replayTargetRevision: number;
  semanticCommandDescriptor: {
    id: string;
    params: unknown;
  };
  affectedRanges: RangeRef[];
  description?: string;
  timestamp: number;
}

export interface CommandRuntimeOptions {
  readonly getRevision?: () => number;
}

/** 变更来源:正向命令、本地撤销、本地重做、远端协同重放 */
export type MutationSource = 'command' | 'undo' | 'redo' | 'remote';

export type MutationListener = (mutation: MutationInfo, source: MutationSource) => void;
export type MutationGuard = (mutation: MutationInfo, source: MutationSource) => void;
export type CommandListener = (commandId: string, params: unknown, result: CommandResult) => void;
export type HistoryReplayListener = (source: 'undo' | 'redo', entry: HistoryEntry) => void;

export interface KernelCommitRequest {
  readonly operationId: string;
  readonly baseRevision: number;
  readonly mutations: readonly MutationInfo[];
  readonly intent?: {
    readonly type: 'undo';
    readonly targetOperationId: string;
    readonly targetBaseRevision: number;
  };
}
export interface KernelCommittedOperation {
  readonly operationId: string;
  readonly baseRevision: number;
  readonly revision: number;
  readonly manifest: KernelReplicaManifest;
  readonly pages: readonly KernelReplicaPagePayload[];
  readonly removedPages: readonly { sheetId: string; pageRow: number; pageColumn: number }[];
  readonly affectedRanges: readonly RangeRef[];
}
export type KernelCommitPort = (request: KernelCommitRequest) => Promise<KernelCommittedOperation>;

export class CommandCommitError extends Error {
  readonly recovery = 'reload-the-committed-cloud-revision';
  constructor(readonly code: 'CLOUD_COMMIT_UNAVAILABLE' | 'COMMITTED_REVISION_MISMATCH' | 'KERNEL_COMMIT_IDENTITY_MISMATCH' | 'MUTATION_WORKBOOK_MISMATCH', readonly object: string) {
    super(`${code}: ${object}`);
    this.name = 'CommandCommitError';
  }
}

/** Plans typed intents and publishes only server-acknowledged Rust transactions. */
export class CommandRuntime {
  private readonly undoStack: HistoryEntry[] = [];
  private readonly redoStack: HistoryEntry[] = [];
  private readonly mutationListeners: MutationListener[] = [];
  private readonly commandListeners: CommandListener[] = [];
  private readonly historyReplayListeners: HistoryReplayListener[] = [];
  private cellValueResolver?: (sheet: WorksheetModel, row: number, column: number) => unknown;
  private mutationGuard?: MutationGuard;
  private revisionProvider?: () => number;
  private commitPort?: KernelCommitPort;
  private serial: Promise<unknown> = Promise.resolve();
  private transactionDepth = 0;

  constructor(readonly workbook: WorkbookModel, readonly registry = new CommandRegistry(), options: CommandRuntimeOptions = {}) { this.revisionProvider = options.getRevision; }
  setCommitPort(port: KernelCommitPort): void { this.commitPort = port; }
  async whenIdle(): Promise<void> { await this.serial; }
  setCellValueResolver(resolver: ((sheet: WorksheetModel, row: number, column: number) => unknown) | undefined): void { this.cellValueResolver = resolver; }
  setMutationGuard(guard: MutationGuard | undefined): void { this.mutationGuard = guard; }
  setRevisionProvider(provider: (() => number) | undefined): void { this.revisionProvider = provider; }
  setRevision(revision: number): void { if (revision !== this.workbook.revision) throw new CommandCommitError('COMMITTED_REVISION_MISMATCH', this.workbook.unitId); }
  onMutation(listener: MutationListener): () => void { this.mutationListeners.push(listener); return () => { const i = this.mutationListeners.indexOf(listener); if (i >= 0) this.mutationListeners.splice(i, 1); }; }
  onCommand(listener: CommandListener): () => void { this.commandListeners.push(listener); return () => { const i = this.commandListeners.indexOf(listener); if (i >= 0) this.commandListeners.splice(i, 1); }; }
  onHistoryReplay(listener: HistoryReplayListener): () => void { this.historyReplayListeners.push(listener); return () => { const i = this.historyReplayListeners.indexOf(listener); if (i >= 0) this.historyReplayListeners.splice(i, 1); }; }

  execute<P>(id: string, params: P): Promise<CommandResult> {
    return this.enqueue(async () => {
      const command = this.registry.getCommand<P>(id);
      this.registry.assertComplete();
      const operationId = createOperationId();
      const baseRevision = this.readRevision();
      const forwardMutations: MutationInfo[] = [];
      const context: CommandContext = {
        workbook: this.workbook, operationId,
        resolveCellValue: (sheet, row, column) => this.cellValueResolver?.(sheet, row, column),
        executeCommand: (commandId, parameters) => this.registry.getCommand(commandId).execute(parameters, context),
        applyMutation: mutation => {
          if (mutation.unitId !== this.workbook.unitId) throw new CommandCommitError('MUTATION_WORKBOOK_MISMATCH', mutation.unitId);
          this.registry.assertMutation(mutation);
          this.mutationGuard?.(mutation, 'command');
          const info: MutationInfo = structuredClone({ id: mutation.id, unitId: mutation.unitId, sheetId: mutation.sheetId, params: mutation.params, affectedRanges: mutation.affectedRanges, ...(mutation.permission ? { permission: mutation.permission } : {}) });
          forwardMutations.push(info);
        },
        recordOperation: (operation, parameters) => this.registry.getOperation(operation.id).execute(parameters, context),
      };
      // Planning has no model writes. A rejected native/ACL transaction leaves this replica untouched.
      const planned = command.execute(params, context);
      let affectedRanges: RangeRef[] = [];
      if (forwardMutations.length) {
        const committed = await this.commit(operationId, forwardMutations, 'command');
        affectedRanges = [...committed.affectedRanges];
        if (command.history !== 'none') {
          const entry: HistoryEntry = {
            operationId,
            baseRevision,
            committedRevision: committed.revision,
            replayTargetOperationId: operationId,
            replayTargetBaseRevision: baseRevision,
            replayTargetRevision: committed.revision,
            semanticCommandDescriptor: { id, params: structuredClone(params) },
            affectedRanges,
            description: id,
            timestamp: Date.now(),
          };
          this.undoStack.push(entry);
          if (this.undoStack.length > 200) this.undoStack.shift();
          this.redoStack.length = 0;
        }
      }
      const result = { ...planned, operationId, mutationCount: forwardMutations.length, affectedRanges };
      for (const listener of this.commandListeners) listener(id, params, result);
      return result;
    });
  }
  undo(): Promise<boolean> { return this.enqueue(async () => {
    const entry = this.undoStack.at(-1);
    if (!entry) return false;
    const replayOperationId = createOperationId();
    const committed = await this.commit(replayOperationId, [], 'undo', {
      type: 'undo',
      targetOperationId: entry.replayTargetOperationId,
      targetBaseRevision: entry.replayTargetBaseRevision,
    });
    this.undoStack.pop();
    entry.replayTargetOperationId = replayOperationId;
    entry.replayTargetBaseRevision = committed.baseRevision;
    entry.replayTargetRevision = committed.revision;
    entry.affectedRanges = [...committed.affectedRanges];
    this.redoStack.push(entry);
    for (const listener of this.historyReplayListeners) listener('undo', entry);
    return true;
  }); }
  redo(): Promise<boolean> { return this.enqueue(async () => {
    const entry = this.redoStack.at(-1);
    if (!entry) return false;
    const replayOperationId = createOperationId();
    const committed = await this.commit(replayOperationId, [], 'redo', {
      type: 'undo',
      targetOperationId: entry.replayTargetOperationId,
      targetBaseRevision: entry.replayTargetBaseRevision,
    });
    this.redoStack.pop();
    entry.replayTargetOperationId = replayOperationId;
    entry.replayTargetBaseRevision = committed.baseRevision;
    entry.replayTargetRevision = committed.revision;
    entry.affectedRanges = [...committed.affectedRanges];
    this.undoStack.push(entry);
    for (const listener of this.historyReplayListeners) listener('redo', entry);
    return true;
  }); }
  applyRemoteCommit(committed: KernelCommittedOperation, mutations: readonly MutationInfo[] = []): void {
    this.assertCommit(committed, committed.operationId, this.workbook.revision);
    this.workbook.applyCommittedManifest(committed.manifest, committed.pages);
    for (const item of mutations) for (const listener of this.mutationListeners) listener(item, 'remote');
  }
  markOperationCommitted(operationId: string, revision: number): void { if (revision !== this.workbook.revision) throw new Error('COMMITTED_REVISION_MISMATCH'); for (const entry of [...this.undoStack, ...this.redoStack]) if (entry.operationId === operationId) entry.committedRevision = revision; }
  get activeDepth(): number { return this.transactionDepth; }
  getHistoryDepth(): { undo: number; redo: number } { return { undo: this.undoStack.length, redo: this.redoStack.length }; }
  getUndoEntries(): readonly HistoryEntry[] { return [...this.undoStack]; }
  getRedoEntries(): readonly HistoryEntry[] { return [...this.redoStack]; }
  clearHistory(): void { this.undoStack.length = 0; this.redoStack.length = 0; }
  private readRevision(): number { const revision = this.revisionProvider?.() ?? this.workbook.revision; if (!Number.isSafeInteger(revision) || revision < 0 || revision !== this.workbook.revision) throw new Error('COMMITTED_REVISION_MISMATCH'); return revision; }
  private enqueue<T>(action: () => Promise<T>): Promise<T> {
    const result = this.serial.then(async () => { this.transactionDepth++; try { return await action(); } finally { this.transactionDepth--; } });
    this.serial = result.then(() => undefined, () => undefined);
    return result;
  }
  private async commit(operationId: string, mutations: readonly MutationInfo[], source: MutationSource, intent?: KernelCommitRequest['intent']): Promise<KernelCommittedOperation> {
    if (!this.commitPort) throw new CommandCommitError('CLOUD_COMMIT_UNAVAILABLE', this.workbook.unitId);
    if ((intent === undefined) === (mutations.length === 0)) throw new Error('KERNEL_COMMIT_PAYLOAD_INVALID');
    const issues: MutationRegistryIssue[] = [];
    for (const item of mutations) { this.registry.validateMutationInfo(item, issues); this.mutationGuard?.(item, source); }
    if (issues.length) throw new Error(formatIssues(issues));
    const baseRevision = this.readRevision();
    const committed = await this.commitPort({ operationId, baseRevision, mutations, ...(intent ? { intent } : {}) });
    this.assertCommit(committed, operationId, baseRevision);
    this.workbook.applyCommittedManifest(committed.manifest, committed.pages);
    for (const item of mutations) for (const listener of this.mutationListeners) listener(item, source);
    return committed;
  }
  private assertCommit(commit: KernelCommittedOperation, operationId: string, baseRevision: number): void {
    if (commit.operationId !== operationId || commit.baseRevision !== baseRevision || commit.revision !== baseRevision + 1 || commit.manifest.revision !== commit.revision || commit.manifest.unitId !== this.workbook.unitId) throw new CommandCommitError('KERNEL_COMMIT_IDENTITY_MISMATCH', operationId);
  }
}
