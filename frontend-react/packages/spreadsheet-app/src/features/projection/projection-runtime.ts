import type { RangeRef, WorksheetModel } from '@react-sheets/core-model';
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

function chartSourceIndexKey(kind: 'sheet' | 'table' | 'pivot', id: string): string {
  return `${kind}:${id}`;
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
  private readonly chartSourceIndex = new Map<string, ChartSourceBinding[]>();

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
    const chartOwnerSheets = new Set<string>();
    for (const mutation of mutations) {
      for (const delta of mutation.structuralFormulaOwnerDeltas ?? []) {
        if (delta.kind === 'formula-object' && delta.ownerKind === 'chart-text') chartOwnerSheets.add(delta.sheetId);
      }
    }
    for (const sheetId of chartOwnerSheets) this.invalidateSheetProjection(sheetId, ['drawings']);
    if (mutations.some((mutation) => this.rebuildChartIndexForMutation(mutation))) this.chartSourceIndexDirty = true;
    this.invalidateDependentChartProjections(mutations);
  }

  invalidateDependentChartProjections(mutations: readonly MutationInfo[]): void {
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
    this.invalidateDependentChartProjectionsForRanges(
      mutations.flatMap((mutation) => mutation.affectedRanges),
      structurallyChangedSheets,
      tableIds,
    );
  }

  private invalidateDependentChartProjectionsForRanges(
    ranges: readonly RangeRef[],
    structurallyChangedSheets?: ReadonlySet<string>,
    tableIds: ReadonlySet<string> = new Set(),
  ): void {
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
    for (const ownerId of owners) this.invalidateSheetProjection(ownerId, ['content', 'formulaResults', 'dataRules']);
  }

  invalidateChartProjectionsForPivot(pivotId: string): void {
    this.ensureChartSourceIndex();
    for (const binding of this.chartSourceIndex.get(chartSourceIndexKey('pivot', pivotId)) ?? []) {
      if (binding.kind === 'pivot') this.invalidateSheetProjection(binding.ownerId, ['content', 'formulaResults', 'dataRules']);
    }
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
    this.chartSourceIndex.clear();
    const tables = [...this.runtime.model.dataModel.tables.values()];
    const sheetOrder = this.runtime.model.getSheets().map(({ id, name }) => ({ id, name }));
    for (const owner of this.runtime.model.getSheets()) {
      for (const payload of owner.drawingPayloads.values()) {
        if (payload.kind !== 'chart') continue;
        if (payload.source.kind === 'table') {
          const key = chartSourceIndexKey('table', payload.source.tableId);
          const bindings = this.chartSourceIndex.get(key) ?? [];
          bindings.push({ kind: 'table', ownerId: owner.id });
          this.chartSourceIndex.set(key, bindings);
        } else if (payload.source.kind === 'pivot') {
          const key = chartSourceIndexKey('pivot', payload.source.pivotId);
          const bindings = this.chartSourceIndex.get(key) ?? [];
          bindings.push({ kind: 'pivot', ownerId: owner.id });
          this.chartSourceIndex.set(key, bindings);
        }
        for (const range of chartSourceRanges(payload, tables, { ownerSheetId: owner.id, sheetOrder })) {
          const key = chartSourceIndexKey('sheet', range.sheetId);
          const bindings = this.chartSourceIndex.get(key) ?? [];
          bindings.push({ kind: 'range', ownerId: owner.id, range: structuredClone(range) });
          this.chartSourceIndex.set(key, bindings);
        }
      }
    }
    this.chartSourceIndexDirty = false;
  }

  private rebuildChartIndexForMutation(mutation: MutationInfo): boolean {
    return STRUCTURAL_REFERENCE_MUTATIONS.has(mutation.id)
      || mutation.id.startsWith('drawing.')
      || mutation.id.startsWith('chart.')
      || mutation.id.startsWith('sheetTable.')
      || mutation.id === 'table.add'
      || mutation.id === 'table.remove';
  }
}
