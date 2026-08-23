import type { WorkbookSnapshot } from '@react-sheets/core-model';
import type {
  CompatibilityReport,
  XlsxExportOptions,
  XlsxImportOptions,
  XlsxSourceArtifact,
} from '@react-sheets/exchange-xlsx';
import type {
  WorkbookAclRole,
  ApiRequestOptions,
  WorkbookApiClient,
  CursorPage,
  WorkbookCatalogQuery as ProtocolWorkbookCatalogQuery,
  WorkbookCreateMetadata,
  WorkbookCopyRequest,
  WorkbookImportRequest,
  WorkbookMetadataPatch,
  UserPreferences,
  UserPreferencesPatch,
  WorkbookSourceArtifactMetadata,
  WorkbookSummary,
  WorkbookUserState as ProtocolWorkbookUserState,
  WorkspaceFolder as ProtocolWorkspaceFolder,
  WorkspaceSpace as ProtocolWorkspaceSpace,
} from '@react-sheets/protocol';
import type {
  WorkspaceRecord,
  WorkspaceRecordMetadata,
  WorkspaceRole,
  WorkspaceStorageLocation,
  WorkspaceUserState,
} from '../persistence/storage';

export type WorkbookRole = WorkbookAclRole | WorkspaceRole;
export type WorkbookCatalogView = 'all' | 'recent' | 'local' | 'owned' | 'shared' | 'trash';
export type WorkbookStorageLocation = WorkspaceStorageLocation;
export type WorkbookLifecycle = WorkspaceRecordMetadata['lifecycle'];
export type WorkbookSource = WorkspaceRecordMetadata['source'];
export type WorkbookSyncState = 'synced' | 'syncing' | 'pending' | 'offline' | 'conflict' | 'error';

export interface WorkbookCatalogQuery {
  view?: WorkbookCatalogView;
  query?: string;
  spaceId?: string;
  folderId?: string;
  cursor?: string;
  limit?: number;
}

export type WorkbookCatalogRequestOptions = ApiRequestOptions;

export interface WorkbookCatalogPage {
  entries: WorkbookCatalogEntry[];
  nextCursor: string | null;
}

export interface WorkbookCatalogEntry {
  unitId: string;
  name: string;
  revision: number;
  updatedAt: string;
  storage: WorkbookStorageLocation;
  syncState: WorkbookSyncState;
  role: WorkbookRole;
  lifecycle: WorkbookLifecycle;
  source: WorkbookSource;
  ownerId?: string;
  ownerName?: string;
  spaceId?: string;
  spaceName?: string;
  folderId?: string;
  locationPath: readonly string[];
  sourceFileName?: string;
  deletedAt?: string;
  favorite: boolean;
  lastOpenedAt?: string;
  pendingOperationCount: number;
  localRecord?: WorkspaceRecord;
}

export interface WorkbookCatalogCreateInput {
  snapshot: WorkbookSnapshot;
  destination?: 'local' | 'remote';
  metadata?: WorkbookCreateMetadata;
  role?: WorkbookRole;
  source?: WorkbookSource;
}

export interface WorkbookCatalogImportInput {
  fileName: string;
  buffer: ArrayBuffer;
  destination?: 'local' | 'remote';
  folderId?: string;
  spaceId?: string;
  options?: Partial<XlsxImportOptions>;
  execution?: 'worker' | 'inline-test';
  workerPort?: import('@react-sheets/exchange-xlsx').XlsxWorkerPort;
}

export interface WorkbookCatalogImportResult {
  entry: WorkbookCatalogEntry;
  snapshot: WorkbookSnapshot;
  report: CompatibilityReport;
  sourceArtifact?: XlsxSourceArtifact;
}

export interface WorkbookCatalogExportInput {
  fileName?: string;
  options?: Partial<XlsxExportOptions>;
  execution?: 'worker' | 'inline-test';
  workerPort?: import('@react-sheets/exchange-xlsx').XlsxWorkerPort;
}

export interface WorkbookCatalogExportResult {
  unitId: string;
  fileName: string;
  buffer: ArrayBuffer;
  report: CompatibilityReport;
}

export interface WorkbookCatalogOpenResult {
  entry: WorkbookCatalogEntry;
  snapshot: WorkbookSnapshot;
}

export interface WorkbookCatalogRemoteClient extends Pick<WorkbookApiClient,
  | 'getSnapshot'
  | 'listWorkbookAcl'
  | 'putWorkbookAcl'
  | 'deleteWorkbookAcl'
  | 'getAccess'
  | 'createWorkbook'
  | 'listWorkbookPage'
  | 'updateWorkbook'
  | 'copyWorkbook'
  | 'moveToTrash'
  | 'restoreFromTrash'
  | 'purgeWorkbook'
  | 'getWorkbookUserState'
  | 'putWorkbookUserState'
  | 'createWorkbookImport'
  | 'putWorkbookSourceArtifact'
  | 'getWorkbookSourceArtifact'
  | 'commitOperation'
  | 'checkpointWorkbook'
  | 'listSpaces'
  | 'getUserPreferences'
  | 'putUserPreferences'
  | 'listFolders'
  | 'createSpace'
  | 'createFolder'
  | 'updateFolder'
  | 'deleteFolder'
  | 'listSpaceMembers'
  | 'putSpaceMember'
  | 'deleteSpaceMember'
> {
  // This explicit protocol boundary keeps Catalog service
  // code never constructs requests or reads auth tokens directly.
}

export type WorkbookCatalogProtocolQuery = ProtocolWorkbookCatalogQuery;
export type WorkbookCatalogProtocolSummary = WorkbookSummary;
export type WorkbookCatalogProtocolPage = CursorPage<WorkbookSummary>;
export type WorkbookCatalogProtocolUserState = ProtocolWorkbookUserState;
export type WorkbookCatalogProtocolImport = WorkbookImportRequest;
export type WorkbookCatalogProtocolMetadataPatch = WorkbookMetadataPatch;
export type WorkbookCatalogProtocolCopy = WorkbookCopyRequest;
export type WorkbookCatalogProtocolArtifactMetadata = WorkbookSourceArtifactMetadata;
export type WorkbookCatalogProtocolSpace = ProtocolWorkspaceSpace;
export type WorkbookCatalogProtocolFolder = ProtocolWorkspaceFolder;
export type WorkbookCatalogProtocolRole = WorkbookAclRole;
export type WorkbookCatalogProtocolUserStateInput = Omit<ProtocolWorkbookUserState, 'unitId'>;

export type WorkbookCatalogLocation = Pick<WorkspaceRecordMetadata, 'spaceId' | 'folderId' | 'locationPath'>;
export type WorkbookCatalogUserState = WorkspaceUserState;
