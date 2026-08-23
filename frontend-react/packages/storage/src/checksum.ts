import { createHash } from 'node:crypto';

export function computeSnapshotChecksum(snapshotJson: string): string {
  return createHash('sha256').update(snapshotJson).digest('hex');
}

export function verifySnapshotChecksum(snapshotJson: string, expected: string): boolean {
  return computeSnapshotChecksum(snapshotJson) === expected;
}
