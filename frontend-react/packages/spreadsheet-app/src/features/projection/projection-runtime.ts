import type { ChartDrawingPayload, RangeRef, WorksheetModel } from '@react-sheets/core-model';
import type { MutationInfo } from '@react-sheets/command-runtime';
import { chartSourceRanges } from '../chart/data';
import { buildCanvasSheetSnapshot, type CanvasSheetSnapshot } from '../../ui-snapshot';
import type { SpreadsheetRuntime } from '../../runtime';

type SheetProjectionDomain = 'content' | 'dimensions' | 'formulaResults' | 'dataRules' | 'drawings' | 'review' | 'structure';

interface SheetProjectionRevision {
  content: number;
  dimensions: number;
  formulaResults: number;
  dataRules: number;
  drawings: number;
  review: number;
  structure: number;
}

interface CachedSheetProjection {
  revision: string;
  snapshot: CanvasSheetSnapshot;
}

type ChartSourceBinding =
  | { kind: 'range'; ownerId: string; range: RangeRef }
  | { kind: 'table'; ownerId: string }
  | { kind: 'pivot'; ownerId: string };

interface IndexedChartSourceBinding {
  readonly key: string;
  readonly binding: ChartSourceBinding;
}

interface ChartSourceOwnerIdentity {
  readonly sheetId: string;
  readonly payloadId: string;
  readonly drawingId: string;
}

interface ChartFormulaSheetIdentity {
  readonly id: string;
  readonly name: string;
}

function chartSourceIndexKey(kind: 'sheet' | 'table' | 'pivot', id: string): string {
  return `${kind}:${id}`;
}

function chartSourceOwnerKey(sheetId: string, payloadId: string): string {
  return JSON.stringify([sheetId, payloadId]);
}

function chartSourceDrawingKey(sheetId: string, drawingId: string): string {
  return JSON.stringify([sheetId, drawingId]);
}

const PROJECTION_DOMAINS: readonly SheetProjectionDomain[] = [
  'content', 'dimensions', 'formulaResults', 'dataRules', 'drawings', 'review', 'structure',
];

const STRUCTURAL_REFERENCE_MUTATIONS = new Set([
  'rows.inserted', 'rows.deleted', 'columns.inserted', 'columns.deleted',
  'cells.inserted', 'cells.deleted', 'cells.inserted.restore', 'cells.deleted.restore',
  'rows.permuted', 'range.move',
  'sheet.add', 'sheet.remove', 'sheet.restore', 'sheet.rename', 'sheet.reordered', 'sheet.duplicated',
]);

const NON_CELL_CHART_SOURCE_MUTATION_PREFIXES = [
  'drawing.', 'shape.', 'chart.', 'image.', 'camera.', 'formControl.',
  'comment.', 'note.', 'hyperlink.',
] as const;

/** Keep the active projection and one recently used sheet; all other sheets stay lazy. */
const MAX_SHEET_PROJECTION_CACHE = 2;

function createSheetProjectionRevision(): SheetProjectionRevision {
  return { content: 0, dimensions: 0, formulaResults: 0, dataRules: 0, drawings: 0, review: 0, structure: 0 };
}

function projectionDomainsForMutation(mutation: MutationInfo): readonly SheetProjectionDomain[] {
  const id = mutation.id;
  if (id.startsWith('drawing.') || id.startsWith('shape.') || id.startsWith('chart.') || id.startsWith('image.') || id.startsWith('camera.') || id.startsWith('formControl.')) return ['drawings'];
  if (id.startsWith('comment.') || id.startsWith('note.') || id.startsWith('hyperlink.')) return ['review'];
  if (STRUCTURAL_REFERENCE_MUTATIONS.has(id) || id.startsWith('sheet.rows.') || id.startsWith('sheet.columns.') || id.startsWith('sheet.cellShift.')) return PROJECTION_DOMAINS;
  if (id.startsWith('sheet.row.') || id.startsWith('sheet.column.') || id.startsWith('sheet.dimension.') || id.startsWith('sheet.visibility.') || id.startsWith('sheet.freeze.')) return ['dimensions'];
  if (id.startsWith('filter.') || id.startsWith('sheetTable.') || id.startsWith('dataRegion.') || id.startsWith('dataSource.') || id.startsWith('validation.') || id.startsWith('conditionalFormat.') || id.startsWith('outline.')) return ['dataRules', 'content'];
  if (id.startsWith('pivot.')) return ['content', 'formulaResults', 'dataRules'];
  if (id.startsWith('formula.')) return ['content', 'formulaResults'];
  return ['content'];
}

function rangesIntersect(left: RangeRef, right: RangeRef): boolean {
  return left.sheetId === right.sheetId
    && left.startRow <= right.endRow
    && left.endRow >= right.startRow
    && left.startColumn <= right.endColumn
    && left.endColumn >= right.startColumn;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object' && !Array.isArray(value);
}

/**
 * Owns derived worksheet projections and their dependency invalidation. The
 * workbook session keeps UI intent; this runtime keeps only canonical-model
 * views, revisions, and the bounded lazy cache.
 */
export class ProjectionRuntime {
  private workbookProjectionEpoch = 0;
  private readonly sheetProjectionRevisions = new Map<string, SheetProjectionRevision>();
  private readonly sheetProjectionCache = new Map<string, CachedSheetProjection>();
  private readonly sheetProjectionAccessOrder: string[] = [];
  private chartSourceIndexDirty = true;
  private readonly chartSourceIndex = new Map<string, Set<ChartSourceBinding>>();
  private readonly chartSourceIndexByOwner = new Map<string, IndexedChartSourceBinding[]>();
  private readonly chartSourceOwnerIdentities = new Map<string, ChartSourceOwnerIdentity>();
  private readonly chartSourceOwnersByDrawing = new Map<string, string>();
  private chartFormulaSheetOrder: ChartFormulaSheetIdentity[] = [];
  private readonly chartFormulaSheetOrderIndex = new Map<string, number>();

  constructor(
    private readonly runtime: SpreadsheetRuntime,
    private readonly getDateSystem: () => '1900' | '1904',
  ) {}

  get cache(): ReadonlyMap<string, unknown> {
    return this.sheetProjectionCache;
  }

  invalidateSheetProjection(sheetId: string, domains: readonly SheetProjectionDomain[]): void {
    if (!this.runtime.model.sheets.has(sheetId)) return;
    const revision = this.sheetProjectionRevisions.get(sheetId) ?? createSheetProjectionRevision();
    for (const domain of domains) revision[domain] += 1;
    this.sheetProjectionRevisions.set(sheetId, revision);
  }

  invalidateProjectionMutations(mutations: readonly MutationInfo[]): void {
    for (const mutation of mutations) this.invalidateSheetProjection(mutation.sheetId, projectionDomainsForMutation(mutation));
    const chartIndexWasDirty = this.chartSourceIndexDirty;
    const chartOwnerSheets = new Set<string>();
    const chartOwnersToRefresh = new Set(this.invalidateDependentChartProjections(mutations));
    const newChartOwnerIdentities = new Map<string, ChartSourceOwnerIdentity>();
    for (const mutation of mutations) {
      for (const delta of mutation.structuralFormulaOwnerDeltas ?? []) {
        if (delta.kind === 'formula-object' && delta.ownerKind === 'chart-text') {
          chartOwnerSheets.add(delta.sheetId);
          chartOwnersToRefresh.add(chartSourceOwnerKey(delta.sheetId, delta.payloadId));
        }
      }
      if (mutation.id === 'drawing.add' && isRecord(mutation.params)) {
        const drawing = isRecord(mutation.params.drawing) ? mutation.params.drawing : undefined;
        const payload = isRecord(mutation.params.payload) ? mutation.params.payload : undefined;
        if (payload?.kind === 'chart') {
          if (typeof drawing?.id !== 'string' || typeof drawing.payloadId !== 'string') {
            throw new Error('CHART_SOURCE_INDEX_INVARIANT: drawing.add chart owner identity is incomplete');
          }
          const ownerId = chartSourceOwnerKey(mutation.sheetId, drawing.payloadId);
          chartOwnersToRefresh.add(ownerId);
          newChartOwnerIdentities.set(ownerId, { sheetId: mutation.sheetId, payloadId: drawing.payloadId, drawingId: drawing.id });
        }
      } else if (mutation.id === 'drawing.remove' && isRecord(mutation.params)) {
        const drawingId = mutation.params.drawingId;
        if (typeof drawingId === 'string') {
          const ownerId = this.chartSourceOwnersByDrawing.get(chartSourceDrawingKey(mutation.sheetId, drawingId));
          if (ownerId) chartOwnersToRefresh.add(ownerId);
        }
      } else if (mutation.id === 'drawing.payload.update' && isRecord(mutation.params)) {
        const payloadId = mutation.params.payloadId;
        const before = isRecord(mutation.params.before) ? mutation.params.before : undefined;
        const after = isRecord(mutation.params.after) ? mutation.params.after : undefined;
        if (typeof payloadId === 'string' && (before?.kind === 'chart' || after?.kind === 'chart')) {
          chartOwnersToRefresh.add(chartSourceOwnerKey(mutation.sheetId, payloadId));
        }
      }
    }
    for (const sheetId of chartOwnerSheets) this.invalidateSheetProjection(sheetId, ['drawings']);
    const renamedSheetIds = new Set<string>();
    let workbookSheetIdentityChanged = false;
    for (const mutation of mutations) {
      if (mutation.id === 'sheet.rename') renamedSheetIds.add(mutation.sheetId);
      if (mutation.id === 'sheet.add' || mutation.id === 'sheet.remove' || mutation.id === 'sheet.restore'
        || mutation.id === 'sheet.reordered' || mutation.id === 'sheet.duplicated') {
        workbookSheetIdentityChanged = true;
      }
    }
    if (!chartIndexWasDirty) {
      if (workbookSheetIdentityChanged) this.rebuildChartSourceIndex();
      else this.refreshChartSourceOwners(chartOwnersToRefresh, newChartOwnerIdentities, renamedSheetIds);
    }
  }

  invalidateDependentChartProjections(mutations: readonly MutationInfo[]): ReadonlySet<string> {
    const structurallyChangedSheets = new Set(
      mutations.filter((mutation) => STRUCTURAL_REFERENCE_MUTATIONS.has(mutation.id)).map((mutation) => mutation.sheetId),
    );
    const tableIds = new Set<string>();
    for (const mutation of mutations) {
      if (mutation.id !== 'table.add' && mutation.id !== 'table.remove') continue;
      if (mutation.params === null || typeof mutation.params !== 'object' || Array.isArray(mutation.params)) continue;
      const params = mutation.params as Record<string, unknown>;
      const tableId = mutation.id === 'table.add' ? params.id : params.tableId;
      if (typeof tableId === 'string') tableIds.add(tableId);
    }
    return this.invalidateDependentChartProjectionsForRanges(
      mutations
        .filter((mutation) => !NON_CELL_CHART_SOURCE_MUTATION_PREFIXES.some((prefix) => mutation.id.startsWith(prefix)))
        .flatMap((mutation) => mutation.affectedRanges),
      structurallyChangedSheets,
      tableIds,
    );
  }

  private invalidateDependentChartProjectionsForRanges(
    ranges: readonly RangeRef[],
    structurallyChangedSheets?: ReadonlySet<string>,
    tableIds: ReadonlySet<string> = new Set(),
  ): ReadonlySet<string> {
    this.ensureChartSourceIndex();
    const owners = new Set<string>();
    for (const sheetId of structurallyChangedSheets ?? []) {
      for (const binding of this.chartSourceIndex.get(chartSourceIndexKey('sheet', sheetId)) ?? []) {
        if (binding.kind === 'range') owners.add(binding.ownerId);
      }
    }
    for (const affectedRange of ranges) {
      for (const binding of this.chartSourceIndex.get(chartSourceIndexKey('sheet', affectedRange.sheetId)) ?? []) {
        if (binding.kind === 'range' && rangesIntersect(binding.range, affectedRange)) owners.add(binding.ownerId);
      }
    }
    for (const tableId of tableIds) {
      for (const binding of this.chartSourceIndex.get(chartSourceIndexKey('table', tableId)) ?? []) {
        if (binding.kind === 'table') owners.add(binding.ownerId);
      }
    }
    this.invalidateChartOwnerProjections(owners);
    return owners;
  }

  invalidateChartProjectionsForPivot(pivotId: string): void {
    this.ensureChartSourceIndex();
    const owners = new Set<string>();
    for (const binding of this.chartSourceIndex.get(chartSourceIndexKey('pivot', pivotId)) ?? []) {
      if (binding.kind === 'pivot') owners.add(binding.ownerId);
    }
    this.invalidateChartOwnerProjections(owners);
  }

  invalidateFormulaResultProjections(addresses?: readonly { sheetId: string; row: number; column: number }[]): void {
    if (!addresses) {
      for (const sheet of this.runtime.model.getSheets()) this.invalidateSheetProjection(sheet.id, ['formulaResults']);
      return;
    }
    if (addresses.length === 0) return;
    const ranges: RangeRef[] = [];
    for (const address of addresses) {
      this.invalidateSheetProjection(address.sheetId, ['formulaResults']);
      ranges.push({
        sheetId: address.sheetId,
        startRow: address.row,
        endRow: address.row,
        startColumn: address.column,
        endColumn: address.column,
      });
    }
    this.invalidateDependentChartProjectionsForRanges(ranges);
  }

  invalidateDataSourceProjection(sourceId: string): void {
    for (const sheet of this.runtime.model.getSheets()) {
      if (sheet.dataRegions.some((region) => region.sourceId === sourceId)) this.invalidateSheetProjection(sheet.id, ['content', 'dataRules', 'formulaResults']);
    }
  }

  invalidateAllSheetProjections(): void {
    this.workbookProjectionEpoch += 1;
    this.sheetProjectionRevisions.clear();
    this.sheetProjectionCache.clear();
    this.sheetProjectionAccessOrder.length = 0;
    this.chartSourceIndexDirty = true;
    this.chartFormulaSheetOrder = [];
    this.chartFormulaSheetOrderIndex.clear();
  }

  getCanvasProjection(sheet: WorksheetModel): CanvasSheetSnapshot {
    const cached = this.sheetProjectionCache.get(sheet.id);
    const revision = this.projectionRevisionForSheet(sheet.id);
    if (cached?.revision === revision) {
      this.touchSheetProjection(sheet.id);
      return cached.snapshot;
    }
    const referenceDate = this.runtime.formula.getCanonicalReferenceDate();
    const snapshot = buildCanvasSheetSnapshot(
      this.runtime.model,
      sheet,
      this.runtime.formula,
      true,
      this.runtime.pivotResults,
      this.runtime.dataContent,
      this.getDateSystem(),
      this.runtime.pivotErrors,
      referenceDate ? { referenceDate } : undefined,
    );
    this.sheetProjectionCache.set(sheet.id, { revision, snapshot });
    this.touchSheetProjection(sheet.id);
    return snapshot;
  }

  getActiveProjectionSheetIds(activeSheet: WorksheetModel): ReadonlySet<string> {
    const ids = new Set<string>([activeSheet.id]);
    const sheetOrder = this.runtime.model.getSheets().map(({ id, name }) => ({ id, name }));
    const addRange = (range: RangeRef | undefined): void => {
      if (range && this.runtime.model.sheets.has(range.sheetId)) ids.add(range.sheetId);
    };
    for (const sparkline of activeSheet.sparklines) addRange(sparkline.sourceRange);
    for (const payload of activeSheet.drawingPayloads.values()) {
      switch (payload.kind) {
        case 'camera':
          addRange(payload.sourceRange);
          break;
        case 'chart':
          if ('pivotId' in payload.source) {
            const pivotId = payload.source.pivotId;
            const owner = this.runtime.model.getSheets().find((sheet) => sheet.pivots.some((pivot) => pivot.id === pivotId));
            if (owner) ids.add(owner.id);
          }
          for (const range of chartSourceRanges(payload, [...this.runtime.model.dataModel.tables.values()], { ownerSheetId: activeSheet.id, sheetOrder })) addRange(range);
          break;
        case 'form-control':
          if ('inputRange' in payload) addRange(payload.inputRange);
          if ('cellLink' in payload && payload.cellLink) ids.add(payload.cellLink.sheetId);
          break;
        default:
          break;
      }
    }
    if (activeSheet.reportSheet) {
      const tableId = activeSheet.reportSheet.tableId;
      if (tableId) addRange(this.runtime.model.dataModel.tables.get(tableId)?.sourceRange);
    }
    return ids;
  }

  prune(requiredProjectionIds: ReadonlySet<string>): void {
    const liveSheetIds = new Set(this.runtime.model.getSheets().map((sheet) => sheet.id));
    for (const sheetId of this.sheetProjectionCache.keys()) {
      if (!liveSheetIds.has(sheetId)) this.sheetProjectionCache.delete(sheetId);
    }
    for (let index = this.sheetProjectionAccessOrder.length - 1; index >= 0; index -= 1) {
      if (!this.sheetProjectionCache.has(this.sheetProjectionAccessOrder[index]!)) this.sheetProjectionAccessOrder.splice(index, 1);
    }
    const keep = new Set(requiredProjectionIds);
    for (let index = this.sheetProjectionAccessOrder.length - 1; index >= 0 && keep.size < MAX_SHEET_PROJECTION_CACHE; index -= 1) {
      keep.add(this.sheetProjectionAccessOrder[index]!);
    }
    for (const sheetId of this.sheetProjectionCache.keys()) {
      if (!keep.has(sheetId)) this.sheetProjectionCache.delete(sheetId);
    }
    for (let index = this.sheetProjectionAccessOrder.length - 1; index >= 0; index -= 1) {
      if (!this.sheetProjectionCache.has(this.sheetProjectionAccessOrder[index]!)) this.sheetProjectionAccessOrder.splice(index, 1);
    }
  }

  private projectionRevisionForSheet(sheetId: string): string {
    const revision = this.sheetProjectionRevisions.get(sheetId) ?? createSheetProjectionRevision();
    return `${this.workbookProjectionEpoch}:${PROJECTION_DOMAINS.map((domain) => revision[domain]).join(':')}`;
  }

  private touchSheetProjection(sheetId: string): void {
    const index = this.sheetProjectionAccessOrder.indexOf(sheetId);
    if (index >= 0) this.sheetProjectionAccessOrder.splice(index, 1);
    this.sheetProjectionAccessOrder.push(sheetId);
  }

  private ensureChartSourceIndex(): void {
    if (!this.chartSourceIndexDirty) return;
    this.rebuildChartSourceIndex();
  }

  private rebuildChartSourceIndex(): void {
    this.chartSourceIndexDirty = true;
    this.chartSourceIndex.clear();
    this.chartSourceIndexByOwner.clear();
    this.chartSourceOwnerIdentities.clear();
    this.chartSourceOwnersByDrawing.clear();
    const sheets = this.runtime.model.getSheets();
    this.chartFormulaSheetOrder = sheets.map(({ id, name }) => ({ id, name }));
    this.chartFormulaSheetOrderIndex.clear();
    this.chartFormulaSheetOrder.forEach((sheet, index) => this.chartFormulaSheetOrderIndex.set(sheet.id, index));
    for (const owner of sheets) {
      const drawingByPayloadId = new Map<string, { id: string; kind: string; sheetId: string }>();
      for (const drawing of owner.drawings) {
        if (drawingByPayloadId.has(drawing.payloadId)) {
          throw new Error(`CHART_SOURCE_INDEX_INVARIANT: duplicate drawing owner for payload ${drawing.payloadId}`);
        }
        drawingByPayloadId.set(drawing.payloadId, { id: drawing.id, kind: drawing.kind, sheetId: drawing.sheetId });
      }
      for (const [payloadId, payload] of owner.drawingPayloads) {
        if (payload.kind !== 'chart') continue;
        const ownerId = chartSourceOwnerKey(owner.id, payloadId);
        const drawing = drawingByPayloadId.get(payloadId);
        if (!drawing || drawing.kind !== 'chart' || drawing.sheetId !== owner.id || payload.chartId !== payloadId) {
          throw new Error(`CHART_SOURCE_INDEX_INVARIANT: chart payload ${payloadId} has no matching drawing owner`);
        }
        const identity = { sheetId: owner.id, payloadId, drawingId: drawing.id };
        this.installChartSourceBindings(
          ownerId,
          identity,
          this.buildChartSourceBindings(owner.id, ownerId, payload, this.chartFormulaSheetOrder),
        );
      }
    }
    this.chartSourceIndexDirty = false;
  }

  private refreshChartSourceOwners(
    ownerIds: ReadonlySet<string>,
    newIdentities: ReadonlyMap<string, ChartSourceOwnerIdentity>,
    renamedSheetIds: ReadonlySet<string>,
  ): void {
    if (ownerIds.size === 0 && renamedSheetIds.size === 0) return;
    this.ensureChartSourceIndex();
    this.chartSourceIndexDirty = true;
    let nextSheetOrder = this.chartFormulaSheetOrder;
    if (renamedSheetIds.size > 0) {
      nextSheetOrder = [...this.chartFormulaSheetOrder];
      for (const renamedSheetId of renamedSheetIds) {
        const sheet = this.runtime.model.sheets.get(renamedSheetId);
        const index = this.chartFormulaSheetOrderIndex.get(renamedSheetId);
        if (!sheet || index === undefined) throw new Error(`CHART_SOURCE_INDEX_INVARIANT: renamed sheet ${renamedSheetId} is absent from formula sheet order`);
        nextSheetOrder[index] = { id: renamedSheetId, name: sheet.name };
      }
    }
    if (ownerIds.size === 0) {
      this.chartFormulaSheetOrder = nextSheetOrder;
      this.chartSourceIndexDirty = false;
      return;
    }

    const replacements = new Map<string, { identity: ChartSourceOwnerIdentity; entries?: IndexedChartSourceBinding[] }>();
    for (const ownerId of ownerIds) {
      const identity = newIdentities.get(ownerId) ?? this.chartSourceOwnerIdentities.get(ownerId);
      if (!identity) throw new Error(`CHART_SOURCE_INDEX_INVARIANT: chart owner ${ownerId} is not registered`);
      const payload = this.runtime.model.sheets.get(identity.sheetId)?.drawingPayloads.get(identity.payloadId);
      replacements.set(ownerId, {
        identity,
        ...(payload?.kind === 'chart'
          ? { entries: this.buildChartSourceBindings(identity.sheetId, ownerId, payload, nextSheetOrder) }
          : {}),
      });
    }

    for (const ownerId of ownerIds) {
      for (const { key, binding } of this.chartSourceIndexByOwner.get(ownerId) ?? []) {
        if (!this.chartSourceIndex.get(key)?.has(binding)) {
          throw new Error(`CHART_SOURCE_INDEX_INVARIANT: owner ${ownerId} binding is missing from source key ${key}`);
        }
      }
    }

    for (const ownerId of ownerIds) this.removeChartSourceBindings(ownerId);
    for (const [ownerId, replacement] of replacements) {
      if (!replacement.entries) continue;
      this.installChartSourceBindings(ownerId, replacement.identity, replacement.entries);
    }
    this.chartFormulaSheetOrder = nextSheetOrder;
    this.chartSourceIndexDirty = false;
  }

  private buildChartSourceBindings(
    ownerSheetId: string,
    ownerId: string,
    payload: ChartDrawingPayload,
    sheetOrder: readonly ChartFormulaSheetIdentity[],
  ): IndexedChartSourceBinding[] {
    const entries: IndexedChartSourceBinding[] = [];
    if (payload.source.kind === 'table') {
      entries.push({
        key: chartSourceIndexKey('table', payload.source.tableId),
        binding: { kind: 'table', ownerId },
      });
    } else if (payload.source.kind === 'pivot') {
      entries.push({
        key: chartSourceIndexKey('pivot', payload.source.pivotId),
        binding: { kind: 'pivot', ownerId },
      });
    }
    const sourceTable = payload.source.kind === 'table'
      ? this.runtime.model.dataModel.tables.get(payload.source.tableId)
      : undefined;
    for (const range of chartSourceRanges(payload, sourceTable ? [sourceTable] : [], { ownerSheetId, sheetOrder })) {
      entries.push({
        key: chartSourceIndexKey('sheet', range.sheetId),
        binding: { kind: 'range', ownerId, range: structuredClone(range) },
      });
    }
    return entries;
  }

  private installChartSourceBindings(ownerId: string, identity: ChartSourceOwnerIdentity, entries: IndexedChartSourceBinding[]): void {
    for (const { key, binding } of entries) {
      const bindings = this.chartSourceIndex.get(key) ?? new Set<ChartSourceBinding>();
      bindings.add(binding);
      this.chartSourceIndex.set(key, bindings);
    }
    if (entries.length > 0) this.chartSourceIndexByOwner.set(ownerId, entries);
    this.chartSourceOwnerIdentities.set(ownerId, identity);
    this.chartSourceOwnersByDrawing.set(chartSourceDrawingKey(identity.sheetId, identity.drawingId), ownerId);
  }

  private removeChartSourceBindings(ownerId: string): void {
    const entries = this.chartSourceIndexByOwner.get(ownerId) ?? [];
    for (const { key, binding } of entries) {
      const bindings = this.chartSourceIndex.get(key);
      if (!bindings || !bindings.delete(binding)) throw new Error(`CHART_SOURCE_INDEX_INVARIANT: owner ${ownerId} binding is missing from source key ${key}`);
      if (bindings.size === 0) this.chartSourceIndex.delete(key);
    }
    this.chartSourceIndexByOwner.delete(ownerId);
    const identity = this.chartSourceOwnerIdentities.get(ownerId);
    if (identity) this.chartSourceOwnersByDrawing.delete(chartSourceDrawingKey(identity.sheetId, identity.drawingId));
    this.chartSourceOwnerIdentities.delete(ownerId);
  }

  private invalidateChartOwnerProjections(ownerIds: ReadonlySet<string>): void {
    for (const ownerId of ownerIds) {
      const identity = this.chartSourceOwnerIdentities.get(ownerId);
      if (!identity) throw new Error(`CHART_SOURCE_INDEX_INVARIANT: owner ${ownerId} has no worksheet identity`);
      this.invalidateSheetProjection(identity.sheetId, ['content', 'formulaResults', 'dataRules']);
    }
  }
}
