import type { MutationInfo } from '@react-sheets/command-runtime';

export interface CollaborativeUndoEntry {
  operationId: string;
  actorId: string;
  undoMutations: MutationInfo[];
  timestamp: number;
}

/** 协同撤销 — 只撤销「我的最近操作」，生成 compensating mutations，不倒退服务器历史 */
export class CollaborativeUndoStack {
  private readonly stacks = new Map<string, CollaborativeUndoEntry[]>();

  push(actorId: string, entry: CollaborativeUndoEntry): void {
    const stack = this.stacks.get(actorId) ?? [];
    stack.push(entry);
    if (stack.length > 50) stack.shift();
    this.stacks.set(actorId, stack);
  }

  pop(actorId: string): CollaborativeUndoEntry | undefined {
    const stack = this.stacks.get(actorId);
    if (!stack || stack.length === 0) return undefined;
    return stack.pop();
  }

  /** 生成补偿命令 mutations — 远端 mutation 不进本地 undo */
  createCompensatingCommand(entry: CollaborativeUndoEntry): MutationInfo[] {
    return [...entry.undoMutations];
  }

  clear(actorId?: string): void {
    if (actorId) {
      this.stacks.delete(actorId);
    } else {
      this.stacks.clear();
    }
  }

  depth(actorId: string): number {
    return this.stacks.get(actorId)?.length ?? 0;
  }
}
