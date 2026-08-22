import type { RangeRef, WorkbookModel } from '@react-sheets/core-model';

export interface MutationInfo<P = unknown> {
  id: string;
  unitId: string;
  sheetId: string;
  params: P;
  affectedRanges: RangeRef[];
}

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

function createOperationId(): string {
  if (typeof crypto !== 'undefined' && crypto.randomUUID) {
    return crypto.randomUUID();
  }
  return 'op-' + Math.random().toString(36).substring(2, 9) + '-' + Date.now().toString(36);
}

export class CommandRegistry {
  private readonly commands = new Map<string, Command<unknown>>();
  private readonly operations = new Map<string, Operation<unknown>>();
  private readonly mutations = new Map<string, MutationHandler<unknown>>();

  registerCommand<P>(command: Command<P>): void {
    if (this.commands.has(command.id)) throw new Error(`Duplicate command: ${command.id}`);
    this.commands.set(command.id, command as Command<unknown>);
  }

  registerOperation<P>(operation: Operation<P>): void {
    if (this.operations.has(operation.id)) throw new Error(`Duplicate operation: ${operation.id}`);
    this.operations.set(operation.id, operation as Operation<unknown>);
  }

  registerMutation<P>(id: string, handler: MutationHandler<P>): void {
    if (this.mutations.has(id)) throw new Error(`Duplicate mutation: ${id}`);
    this.mutations.set(id, handler as MutationHandler<unknown>);
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
    const mutation = this.mutations.get(id);
    if (!mutation) throw new Error(`Unknown mutation: ${id}`);
    return mutation as MutationHandler<P>;
  }

  hasCommand(id: string): boolean {
    return this.commands.has(id);
  }

  hasMutation(id: string): boolean {
    return this.mutations.has(id);
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
      recordOperation: (operation, operationParams) => operation.execute(operationParams, context),
    };

    try {
      const commandResult = this.registry.getCommand<P>(id).execute(params, context);
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
    for (const item of items) {
      const handler = this.registry.getMutation(item.id);
      handler(item, {
        workbook: this.workbook,
        operationId: createOperationId(),
        applyMutation: () => undefined,
        recordOperation: () => ({ operationId: createOperationId() }),
      });
      for (const listener of this.mutationListeners) {
        listener(item, source);
      }
    }
  }
}
