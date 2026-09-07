import type {
  CompatibilityReport,
  NativeDocumentExportOptions,
  NativeDocumentImportOptions,
  NativeDocumentArtifact,
} from '@react-sheets/exchange-excel-ooxml';
import type {
  WorkbookAclRole,
  ApiRequestOptions,
  WorkbookAccessResponse,
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
  WorkbookManifest,
  KernelWorkbookCreateRequest,
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
export type WorkbookCatalogView = 'all' | 'recent' | 'owned' | 'shared' | 'trash';
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
}

export interface WorkbookCatalogCreateInput {
  plan: Pick<KernelWorkbookCreateRequest, 'unitId' | 'name' | 'sheets' | 'initialMutations'>;
  destination?: 'remote';
  metadata?: WorkbookCreateMetadata;
  role?: WorkbookRole;
  source?: WorkbookSource;
}

export interface WorkbookCatalogImportInput {
  fileName: string;
  buffer: ArrayBuffer;
  destination?: 'remote';
  folderId?: string;
  spaceId?: string;
  options?: Partial<NativeDocumentImportOptions>;
}

export interface WorkbookCatalogImportResult {
  entry: WorkbookCatalogEntry;
  manifest: WorkbookManifest;
  report: CompatibilityReport;
  artifact: NativeDocumentArtifact;
}

export interface WorkbookCatalogExportInput {
  fileName?: string;
  options?: Partial<NativeDocumentExportOptions>;
}

export interface WorkbookCatalogExportResult {
  unitId: string;
  fileName: string;
  buffer: ArrayBuffer;
  report: CompatibilityReport;
}

export interface WorkbookResolutionBinding {
  location: WorkbookStorageLocation;
  syncMode: WorkspaceRecord['syncMode'];
}

/**
 * Pure workbook identity/access resolution.  Resolution is read-only; the
 * editor is responsible for consuming it and the ready callback is the only
 * place that records a successful open.
 */
export interface WorkbookResolution {
  schema: 'WorkbookResolution';
  unitId: string;
  /** Resolution source identifies the authoritative owner, never a cache. */
  source: 'remote' | 'shared';
  mode: 'remote';
  lifecycle: 'active';
  binding: WorkbookResolutionBinding;
  manifest: WorkbookManifest;
  revision: number;
  access: WorkbookAccessResponse | null;
}

export interface WorkbookCatalogRemoteClient extends Pick<WorkbookApiClient,
  | 'getManifest'
  | 'createKernelWorkbook'
  | 'commitKernelOperation'
  | 'listWorkbookAcl'
  | 'putWorkbookAcl'
  | 'deleteWorkbookAcl'
  | 'getAccess'
  | 'listWorkbookPage'
  | 'updateWorkbook'
  | 'copyWorkbook'
  | 'moveToTrash'
  | 'restoreFromTrash'
  | 'purgeWorkbook'
  | 'getWorkbookUserState'
  | 'putWorkbookUserState'
  | 'createWorkbookImport'
  | 'getWorkbookSourceArtifact'
  | 'saveNativeDocumentArtifact'
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
