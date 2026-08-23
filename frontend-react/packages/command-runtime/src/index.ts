import type { RangeRef, WorkbookModel } from '@react-sheets/core-model';

export interface MutationInfo<P = unknown> {
  id: string;
  unitId: string;
  sheetId: string;
  params: P;
  affectedRanges: RangeRef[];
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
  readonly schema?: MutationParamsSchema<P>;
  readonly permission?: MutationPermissionMetadata;
  readonly affectedRanges?: MutationAffectedRangesMetadata<P>;
  /** Optional explicit inverse allow-list for this mutation. */
  readonly inverseIds?: readonly string[];
}

/** Short public name for feature packages that expose a mutation contract. */
export type MutationMetadata<P = unknown> = MutationRegistrationMetadata<P>;

export interface CommandResult {
  operationId: string;
  mutationCount: number;
  affectedRanges: RangeRef[];
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
  applyMutation<P>(mutation: Mutation<P>): void;
  recordOperation<P>(operation: Operation<P>, params: P): OperationResult;
}

export type MutationHandler<P = unknown> = (item: MutationInfo<P>, context: CommandContext) => void;

export interface MutationRegistration<P = unknown> {
  readonly id: string;
  readonly handler: MutationHandler<P>;
  readonly metadata?: MutationRegistrationMetadata<P>;
}

export interface CommandRegistryOptions {
  /**
   * Make `validateCompleteness()` require schema, permission and affected-range
   * declarations for every mutation. Existing feature registration remains
   * readable during the migration, while production gates can opt into the
   * strict contract and fail closed.
   */
  readonly requireMutationMetadata?: boolean;
}

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
  | 'inverse-not-allowed';

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

function createOperationId(): string {
  if (typeof crypto !== 'undefined' && crypto.randomUUID) {
    return crypto.randomUUID();
  }
  return 'op-' + Math.random().toString(36).substring(2, 9) + '-' + Date.now().toString(36);
}

export class CommandRegistry {
  private readonly commands = new Map<string, Command<unknown>>();
  private readonly operations = new Map<string, Operation<unknown>>();
  private readonly mutations = new Map<string, MutationRegistration<unknown>>();

  constructor(private readonly options: CommandRegistryOptions = {}) {}

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

  registerMutation<P>(registration: MutationRegistration<P>): void;
  registerMutation<P>(id: string, handler: MutationHandler<P>, metadata?: MutationRegistrationMetadata<P>): void;
  registerMutation<P>(
    idOrRegistration: string | MutationRegistration<P>,
    handler?: MutationHandler<P>,
    metadata?: MutationRegistrationMetadata<P>,
  ): void {
    const registration: MutationRegistration<P> =
      typeof idOrRegistration === 'string'
        ? { id: idOrRegistration, handler: handler as MutationHandler<P>, metadata }
        : idOrRegistration;
    if (!registration.id || typeof registration.id !== 'string') {
      throw new Error('Mutation registration requires a non-empty id');
    }
    if (typeof registration.handler !== 'function') {
      throw new Error(`Mutation registration requires a handler: ${registration.id}`);
    }
    if (this.mutations.has(registration.id)) throw new Error(`Duplicate mutation: ${registration.id}`);
    this.mutations.set(registration.id, registration as MutationRegistration<unknown>);
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

  getMutationMetadata<P>(id: string): MutationRegistrationMetadata<P> | undefined {
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
   * Validate all registered mutation contracts. This is intentionally
   * separate from registration so a migration can register legacy handlers,
   * run the gate, and receive an actionable list rather than a first-error
   * failure.
   */
  validateCompleteness(options: CommandRegistryOptions = {}): MutationRegistryCompletenessResult {
    const requireMetadata = options.requireMutationMetadata ?? this.options.requireMutationMetadata ?? true;
    const issues: MutationRegistryIssue[] = [];

    for (const registration of this.mutations.values()) {
      const { id, handler, metadata } = registration;
      if (!id || typeof id !== 'string' || typeof handler !== 'function') {
        issues.push(issue('invalid-registration', id || '<empty>', `Invalid mutation registration: ${id || '<empty>'}`));
        continue;
      }
      if (!requireMetadata) continue;
      if (!metadata?.schema) {
        issues.push(issue('missing-schema', id, `Mutation ${id} must declare a parameter schema`));
      } else if (typeof metadata.schema.validate !== 'function') {
        issues.push(issue('invalid-registration', id, `Mutation ${id} declares an invalid parameter schema`));
      }
      if (!metadata?.permission || typeof metadata.permission.capability !== 'string' || !metadata.permission.capability) {
        issues.push(issue('missing-permission', id, `Mutation ${id} must declare a permission capability`));
      }
      if (!metadata?.affectedRanges) {
        issues.push(issue('missing-affected-ranges', id, `Mutation ${id} must declare an affected-range resolver`));
      } else if (typeof metadata.affectedRanges.resolve !== 'function') {
        issues.push(issue('invalid-registration', id, `Mutation ${id} declares an invalid affected-range resolver`));
      }
      if (metadata?.inverseIds !== undefined && !Array.isArray(metadata.inverseIds)) {
        issues.push(issue('invalid-registration', id, `Mutation ${id} declares an invalid inverse allow-list`));
      }
      for (const inverseId of Array.isArray(metadata?.inverseIds) ? metadata.inverseIds : []) {
        if (!this.mutations.has(inverseId)) {
          issues.push(issue('unknown-inverse', id, `Mutation ${id} declares unknown inverse mutation ${inverseId}`, inverseId));
        }
      }
    }

    return { ok: issues.length === 0, issues };
  }

  assertComplete(options: CommandRegistryOptions = {}): void {
    const result = this.validateCompleteness(options);
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
    if (!Array.isArray(mutation.inverse) || mutation.inverse.length === 0) {
      issues.push(issue('invalid-inverse', mutation.id, `Mutation ${mutation.id} must provide at least one inverse mutation`));
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
        if (registration.metadata?.inverseIds && !registration.metadata.inverseIds.includes(inverse.id)) {
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

    const schema = registration.metadata?.schema;
    if (schema) {
      let valid = false;
      try {
        valid = schema.validate(item.params);
      } catch {
        valid = false;
      }
      if (!valid) {
        issues.push(issue('invalid-params', ownerMutationId, `Mutation ${item.id} parameters do not match ${schema.name ?? 'its schema'}`, inverseId));
      }
    }

    const affectedRanges = registration.metadata?.affectedRanges;
    if (affectedRanges && Array.isArray(item.affectedRanges)) {
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
  undo: MutationInfo[];
  redo: MutationInfo[];
  description?: string;
  timestamp: number;
}

/** 变更来源:正向命令、本地撤销、本地重做、远端协同重放 */
export type MutationSource = 'command' | 'undo' | 'redo' | 'remote';

export type MutationListener = (mutation: MutationInfo, source: MutationSource) => void;
export type CommandListener = (commandId: string, params: unknown, result: CommandResult) => void;

export class CommandRuntime {
  private readonly undoStack: HistoryEntry[] = [];
  private readonly redoStack: HistoryEntry[] = [];
  private activeEntry: HistoryEntry | null = null;
  private transactionDepth = 0;
  private readonly mutationListeners: MutationListener[] = [];
  private readonly commandListeners: CommandListener[] = [];

  constructor(
    readonly workbook: WorkbookModel,
    readonly registry = new CommandRegistry(),
  ) {}

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

  execute<P>(id: string, params: P): CommandResult {
    // Resolve the command before opening a transaction. An unknown command is
    // a protocol error and must not create an empty history entry or invoke a
    // host fallback.
    const command = this.registry.getCommand<P>(id);
    const operationId = createOperationId();
    const mutations: MutationInfo[] = [];
    const isRootTransaction = this.transactionDepth === 0;

    if (isRootTransaction) {
      this.activeEntry = {
        operationId,
        undo: [],
        redo: [],
        description: id,
        timestamp: Date.now(),
      };
    }
    this.transactionDepth += 1;

    const context: CommandContext = {
      workbook: this.workbook,
      operationId,
      applyMutation: (mutation) => {
        if (mutation.unitId !== this.workbook.unitId) {
          throw new Error(`Mutation unit mismatch: expected ${this.workbook.unitId}, received ${mutation.unitId}`);
        }
        // Registration and inverse validation happen before the mutation's
        // callback is allowed to touch the workbook. This makes both local
        // execution and every replay path fail closed on protocol drift.
        this.registry.assertMutation(mutation);
        mutation.apply(context);
        const info: MutationInfo = {
          id: mutation.id,
          unitId: mutation.unitId,
          sheetId: mutation.sheetId,
          params: mutation.params,
          affectedRanges: mutation.affectedRanges,
        };
        mutations.push(info);
        this.activeEntry?.undo.unshift(...mutation.inverse);
        this.activeEntry?.redo.push(info);

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
        if (this.activeEntry && (this.activeEntry.undo.length > 0 || this.activeEntry.redo.length > 0)) {
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
        if (this.activeEntry && this.activeEntry.undo.length > 0) {
          this.applyHistory(this.activeEntry.undo, 'undo');
        }
        this.activeEntry = null;
      }
      throw err;
    }
  }

  undo(): boolean {
    const entry = this.undoStack.pop();
    if (!entry) return false;
    this.applyHistory(entry.undo, 'undo');
    this.redoStack.push(entry);
    return true;
  }

  redo(): boolean {
    const entry = this.redoStack.pop();
    if (!entry) return false;
    this.applyHistory(entry.redo, 'redo');
    this.undoStack.push(entry);
    return true;
  }

  /**
   * 应用来自远端协同的变更序列:执行已注册的 mutation 处理器,
   * 以 'remote' 来源通知监听器(用于引擎同步/视图刷新),但不进入本地撤销栈。
   */
  applyRemoteMutations(items: readonly MutationInfo[]): void {
    this.applyHistory(items, 'remote');
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

  clearHistory(): void {
    this.undoStack.length = 0;
    this.redoStack.length = 0;
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
    for (const item of items) {
      const handler = this.registry.getMutation(item.id);
      const replayContext: CommandContext = {
        workbook: this.workbook,
        operationId: createOperationId(),
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
