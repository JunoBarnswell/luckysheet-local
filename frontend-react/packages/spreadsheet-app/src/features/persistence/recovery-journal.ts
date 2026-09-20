import { validateOperationEnvelope, type OperationEnvelope } from '@react-sheets/protocol';
import { computeChecksum } from './checksum';

interface RecoveryRecord { key: string; version: 1; payload: string; checksum: string; }

// A new page gets a new identity even when a browser duplicates sessionStorage.
const pendingReleases = new Map<string, Promise<void>>();

/** Unacknowledged intent only; Java remains the workbook authority. */
export class RecoveryJournal {
  readonly clientSessionId = crypto.randomUUID();
  private ownership: Promise<void> | null = null;
  private unlock: (() => void) | null = null;
  private readonly key: string;
  private readonly prefix: string;
  private readonly database: Promise<IDBDatabase>;
  private tail: Promise<void> = Promise.resolve();
  constructor(origin: string, subject: string, unitId: string) {
    if (!subject || !unitId) throw new Error('RECOVERY_IDENTITY_REQUIRED');
    this.prefix = JSON.stringify([origin, subject, unitId]).slice(0, -1) + ',';
    this.key = JSON.stringify([origin, subject, unitId, this.clientSessionId]);
    this.database = new Promise((resolve, reject) => {
      const request = indexedDB.open('react-sheets-recovery', 1);
      request.onupgradeneeded = () => request.result.createObjectStore('pending', { keyPath: 'key' });
      request.onerror = () => reject(new Error('RECOVERY_STORAGE_UNAVAILABLE', { cause: request.error }));
      request.onblocked = () => reject(new Error('RECOVERY_STORAGE_BLOCKED: 请关闭旧版本页面后重新打开'));
      request.onsuccess = () => resolve(request.result);
    });
  }
  private async ownPage(): Promise<void> {
    if (!navigator.locks) throw new Error('RECOVERY_LOCKS_UNAVAILABLE: 请使用安全来源下支持 Web Locks 的浏览器');
    let ready = this.ownership;
    if (!ready) {
      ready = new Promise<void>((resolve, reject) => {
        void navigator.locks.request(`recovery-owner:${this.key}`, async () => {
          resolve();
          // The browser releases this lock on navigation or process termination.
          await new Promise<void>(release => { this.unlock = release; });
        }).catch(reject);
      });
      this.ownership = ready;
    }
    await ready;
  }
  async load(): Promise<OperationEnvelope[]> {
    await pendingReleases.get(this.prefix);
    await this.ownPage();
    return navigator.locks.request(`recovery-catalog:${this.prefix}`, async () => {
      const db = await this.database;
      const records = await new Promise<RecoveryRecord[]>((resolve, reject) => {
        const request = db.transaction('pending').objectStore('pending').getAll(IDBKeyRange.bound(this.prefix, this.prefix + '\uffff'));
        request.onsuccess = () => resolve(request.result as RecoveryRecord[]);
        request.onerror = () => reject(new Error('RECOVERY_READ_FAILED', { cause: request.error }));
      });
      const claimed: RecoveryRecord[] = [];
      const releases: Array<() => void> = [];
      try {
        for (const record of records) {
          if (record.key === this.key) { claimed.push(record); continue; }
          // Never take pending intent from another live browser page.
          await new Promise<void>((resolve, reject) => {
            void navigator.locks.request(`recovery-owner:${record.key}`, { ifAvailable: true }, async lock => {
              if (!lock) { resolve(); return; }
              claimed.push(record);
              await new Promise<void>(release => { releases.push(release); resolve(); });
            }).catch(reject);
          });
        }
        const operations = claimed.flatMap(record => {
          if (record.version !== 1 || computeChecksum(record.payload) !== record.checksum) throw new Error('RECOVERY_CHECKSUM_INVALID: 恢复日志已保留，请导出备份后修复');
          const parsed: OperationEnvelope[] = JSON.parse(record.payload);
          if (!Array.isArray(parsed)) throw new Error('RECOVERY_SCHEMA_INVALID');
          return parsed.map(validateOperationEnvelope);
        });
        const ids = new Set<string>();
        for (const operation of operations) {
          if (ids.has(operation.operationId)) throw new Error('RECOVERY_DUPLICATE_OPERATION: 请核对恢复日志');
          ids.add(operation.operationId);
        }
        await this.write(db, operations, claimed.filter(record => record.key !== this.key).map(record => record.key));
        return operations;
      } finally { for (const release of releases) release(); }
    });
  }
  private write(db: IDBDatabase, operations: readonly OperationEnvelope[], retired: string[] = []): Promise<void> {
    const payload = JSON.stringify(operations);
    const record: RecoveryRecord = { key: this.key, version: 1, payload, checksum: computeChecksum(payload) };
    return new Promise<void>((resolve, reject) => {
      const transaction = db.transaction('pending', 'readwrite', { durability: 'strict' });
      const store = transaction.objectStore('pending');
      if (operations.length) store.put(record); else store.delete(this.key);
      for (const key of retired) store.delete(key);
      transaction.oncomplete = () => resolve();
      transaction.onerror = () => reject(new Error('RECOVERY_WRITE_FAILED: 暂停提交，请释放磁盘空间后重新打开', { cause: transaction.error }));
      transaction.onabort = () => reject(new Error('RECOVERY_WRITE_ABORTED', { cause: transaction.error }));
    });
  }
  persist(operations: readonly OperationEnvelope[]): Promise<void> {
    const captured = structuredClone(operations);
    this.tail = this.tail.then(async () => {
      await this.ownPage();
      await navigator.locks.request(`recovery-catalog:${this.prefix}`, async () => this.write(await this.database, captured));
    });
    return this.tail;
  }
  release(): void {
    const unlock = () => { this.unlock?.(); this.unlock = null; this.ownership = null; };
    const previous = pendingReleases.get(this.prefix) ?? Promise.resolve();
    const released = Promise.all([previous, this.tail.then(unlock, unlock)]).then(() => undefined);
    pendingReleases.set(this.prefix, released);
    void released.then(() => { if (pendingReleases.get(this.prefix) === released) pendingReleases.delete(this.prefix); });
  }
  async flushed(): Promise<void> { await this.tail; }
}
