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
  return crypto.randomUUID();
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
}

interface HistoryEntry {
  operationId: string;
  undo: MutationInfo[];
  redo: MutationInfo[];
}

export class CommandRuntime {
  private readonly undoStack: HistoryEntry[] = [];
  private readonly redoStack: HistoryEntry[] = [];
  private activeEntry: HistoryEntry | null = null;

  constructor(
    readonly workbook: WorkbookModel,
    readonly registry = new CommandRegistry(),
  ) {}

  execute<P>(id: string, params: P): CommandResult {
    const operationId = createOperationId();
    const mutations: MutationInfo[] = [];
    this.activeEntry = { operationId, undo: [], redo: [] };
    const context: CommandContext = {
      workbook: this.workbook,
      operationId,
      applyMutation: (mutation) => {
        mutation.apply(context);
        mutations.push(mutation);
        this.activeEntry?.undo.unshift(...mutation.inverse);
        this.activeEntry?.redo.push({
          id: mutation.id,
          unitId: mutation.unitId,
          sheetId: mutation.sheetId,
          params: mutation.params,
          affectedRanges: mutation.affectedRanges,
        });
      },
      recordOperation: (operation, operationParams) => operation.execute(operationParams, context),
    };

    try {
      const commandResult = this.registry.getCommand<P>(id).execute(params, context);
      if (mutations.length > 0) {
        this.undoStack.push(this.activeEntry);
        this.redoStack.length = 0;
      }
      return { ...commandResult, operationId, mutationCount: mutations.length };
    } finally {
      this.activeEntry = null;
    }
  }

  undo(): boolean {
    const entry = this.undoStack.pop();
    if (!entry) return false;
    this.applyHistory(entry.undo);
    this.redoStack.push(entry);
    return true;
  }

  redo(): boolean {
    const entry = this.redoStack.pop();
    if (!entry) return false;
    this.applyHistory(entry.redo);
    this.undoStack.push(entry);
    return true;
  }

  getHistoryDepth(): { undo: number; redo: number } {
    return { undo: this.undoStack.length, redo: this.redoStack.length };
  }

  clearHistory(): void {
    this.undoStack.length = 0;
    this.redoStack.length = 0;
  }

  private applyHistory(items: MutationInfo[]): void {
    for (const item of items) {
      this.registry.getMutation(item.id)(item, {
        workbook: this.workbook,
        operationId: createOperationId(),
        applyMutation: () => undefined,
        recordOperation: () => ({ operationId: createOperationId() }),
      });
    }
  }
}
