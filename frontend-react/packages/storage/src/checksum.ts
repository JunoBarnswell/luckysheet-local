import { sha256Hex } from './sha256';

export function computeSnapshotChecksum(snapshotJson: string): string {
  return sha256Hex(snapshotJson);
}

export function verifySnapshotChecksum(snapshotJson: string, expected: string): boolean {
  return computeSnapshotChecksum(snapshotJson) === expected;
}
