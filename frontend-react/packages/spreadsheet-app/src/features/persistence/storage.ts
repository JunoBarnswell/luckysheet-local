import type { KernelReplicaManifest } from '@react-sheets/core-model';
import { LocalDataBlockStore } from './data-block-store';
import { LocalSparseOverlayStore } from '../data-source/overlay-store';
import { computeChecksum } from './checksum';
import {
  WorkspaceMemoryCoordinator,
  type WorkspacePersistenceState,
} from './memory';

/**
 * Options for the browser cache coordinator.
 *
 * Workbook snapshots, operation journals, drafts, and catalog records are
 * server-owned. The browser may only retain cache data such as blocks,
 * sparse overlays, and assets.
 */
export interface WorkspacePersistenceOptions {
  unitId?: string | (() => string);
}

/**
 * Read-only metadata used by the session UI. This is a fingerprint of the
 * current model manifest, not a browser-owned workbook checkpoint.
 */
export interface PersistenceSnapshotMeta {
  unitId: string;
  revision: number;
  checksum: string;
  updatedAt: string;
}

function manifestPayload(manifest: KernelReplicaManifest): string {
  return JSON.stringify(manifest);
}

export function buildPersistenceMeta(
  manifest: KernelReplicaManifest,
  revision: number,
): PersistenceSnapshotMeta {
  return {
    unitId: manifest.unitId,
    revision,
    checksum: computeChecksum(manifestPayload(manifest)),
    updatedAt: new Date().toISOString(),
  };
}

/**
 * Browser persistence owns cache coordination only. The authoritative
 * workbook snapshot and operation history live behind the server API.
 */
export class WorkspacePersistence {
  readonly coordinator: WorkspaceMemoryCoordinator;
  readonly dataBlocks: LocalDataBlockStore;
  readonly sparseOverlays: LocalSparseOverlayStore;

  constructor(options: WorkspacePersistenceOptions = {}) {
    this.coordinator = new WorkspaceMemoryCoordinator();
    this.dataBlocks = new LocalDataBlockStore(this.coordinator, options.unitId);
    this.sparseOverlays = new LocalSparseOverlayStore({
      coordinator: this.coordinator,
      unitId: options.unitId,
    });
  }

  get state(): WorkspacePersistenceState {
    return this.coordinator.state;
  }

  async ensureReady(): Promise<void> {
    this.coordinator.ensureReady();
  }

  disposeAsync(): Promise<void> {
    return this.coordinator.disposeAsync();
  }
}
