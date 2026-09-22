import type { CheckpointResponse } from '@react-sheets/protocol';

/** One in-flight server checkpoint; continuous editing is coalesced into bounded batches. */
export class CheckpointCoordinator {
  private requestedRevision = 0;
  private coveredRevision = 0;
  private timer: ReturnType<typeof setTimeout> | undefined;
  private inFlight: Promise<void> | undefined;
  private failed = false;
  private disposed = false;

  constructor(
    private readonly checkpoint: () => Promise<CheckpointResponse>,
    private readonly covered: (revision: number) => Promise<void>,
    private readonly onError: (error: Error) => void,
  ) {}

  get hasFailure(): boolean { return this.failed; }

  request(revision: number): void {
    if (!Number.isSafeInteger(revision) || revision < 0) throw new Error('CHECKPOINT_REVISION_INVALID');
    this.requestedRevision = Math.max(this.requestedRevision, revision);
    if (this.disposed || this.failed || this.timer !== undefined || this.inFlight || this.requestedRevision <= this.coveredRevision) return;
    this.timer = setTimeout(() => {
      this.timer = undefined;
      void this.flush().catch(error => this.onError(error instanceof Error ? error : new Error(String(error))));
    }, 1000);
  }

  /** Explicit save/reconnect can retry a failed checkpoint without resending committed requests. */
  async flush(): Promise<void> {
    if (this.timer !== undefined) { clearTimeout(this.timer); this.timer = undefined; }
    if (this.inFlight) { await this.inFlight; return this.flush(); }
    if (this.disposed || this.requestedRevision <= this.coveredRevision) return;
    const target = this.requestedRevision;
    const run = (async () => {
      const result = await this.checkpoint();
      if (!Number.isSafeInteger(result.revision) || result.revision < target) throw new Error('CHECKPOINT_COVERAGE_INVALID: 服务端检查点未覆盖已确认的版本，恢复日志已保留');
      if (this.disposed) return;
      await this.covered(result.revision);
      this.coveredRevision = Math.max(this.coveredRevision, result.revision);
      this.failed = false;
    })();
    this.inFlight = run;
    try { await run; }
    catch (error) { this.failed = true; throw error; }
    finally { if (this.inFlight === run) this.inFlight = undefined; }
    this.request(this.requestedRevision);
  }

  dispose(): void {
    this.disposed = true;
    if (this.timer !== undefined) clearTimeout(this.timer);
    this.timer = undefined;
  }
}
