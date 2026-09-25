import type { CellAddress, FormulaAst, FormulaReferenceNode } from './ast';
import { assertCellAddress, cellAddressKey, compareCellAddresses } from './address';
import { FormulaReferenceError } from './errors';
import type { FormulaDependency } from './range-index';
import { resolveFormulaSheetId, type FormulaSheetIdentity } from './sheet-reference';

type Axis = 'row' | 'column';
type TreeKey = 'start' | 'end';

interface ReferenceGeometry {
  readonly sheetId: string;
  readonly startRow: number;
  readonly endRow: number;
  readonly startColumn: number;
  readonly endColumn: number;
  readonly rowStructural: boolean;
  readonly columnStructural: boolean;
}

interface IndexedReference {
  readonly id: string;
  readonly ownerKey: string;
  readonly sourceId: string;
  readonly owner?: CellAddress;
  readonly definedNameOwner?: DefinedNameReferenceOwnerIdentity;
  readonly sheetId: string;
  readonly axis: Axis;
  readonly start: number;
  readonly end: number;
  readonly crossStart: number;
  readonly crossEnd: number;
  readonly structural: boolean;
  readonly point: boolean;
}

interface IntervalNode {
  readonly value: IndexedReference;
  readonly left?: IntervalNode;
  readonly right?: IntervalNode;
  readonly height: number;
  readonly maxEnd: number;
  readonly minStart: number;
}

interface AxisTrees {
  byStart?: IntervalNode;
  byEnd?: IntervalNode;
}

interface OwnerReferences {
  readonly address?: CellAddress;
  readonly definedNameOwner?: DefinedNameReferenceOwnerIdentity;
  readonly sourceId: string;
  readonly postings: readonly IndexedReference[];
  readonly position?: IndexedReference;
  readonly anchorPosition?: IndexedReference;
  readonly failure?: DefinedNameReferenceFailureReason;
}

export interface IndexedReferenceOwnerSource {
  readonly address: CellAddress;
  readonly sourceId: string;
}

export interface DefinedNameReferenceOwnerIdentity {
  readonly scope: 'workbook' | 'sheet';
  readonly name: string;
  readonly sheetId?: string;
}

export type DefinedNameReferenceFailureReason = 'invalid-formula' | 'unresolved-context' | 'invalid-reference';

export interface DefinedNameReferenceFailure {
  readonly owner: DefinedNameReferenceOwnerIdentity;
  readonly reason: DefinedNameReferenceFailureReason;
}

export interface DefinedNameReferenceIndexUpdate {
  readonly owner: DefinedNameReferenceOwnerIdentity;
  readonly remove?: boolean;
  readonly references?: readonly FormulaReferenceNode[];
  readonly context?: CellAddress;
  readonly anchor?: CellAddress;
  readonly failure?: DefinedNameReferenceFailureReason;
}

/**
 * Incremental spatial index for formula-reference owners. Calculation and
 * structural-only sources have separate identities; point queries expose only
 * calculation formulas while structural queries include every source.
 */
export class ReferenceIndex {
  private readonly owners = new Map<string, OwnerReferences>();
  private readonly sheets = new Map<string, Map<Axis, AxisTrees>>();
  private readonly ownerPositions = new Map<string, IntervalNode>();
  private readonly definedNameAnchors = new Map<string, Map<Axis, AxisTrees>>();
  private readonly definedNameFailures = new Map<string, DefinedNameReferenceFailure>();

  constructor(private sheetOrder: readonly FormulaSheetIdentity[] = []) {}

  setSheetOrder(sheetOrder: readonly FormulaSheetIdentity[]): void {
    this.sheetOrder = sheetOrder;
  }

  set(owner: CellAddress, dependencies: readonly FormulaDependency[], sourceId = 'formula'): void {
    assertCellAddress(owner);
    assertSourceId(sourceId);
    const ownerKey = cellAddressKey(owner);
    const storageKey = ownerSourceKey(ownerKey, sourceId);
    const postings: IndexedReference[] = [];
    let sequence = 0;

    for (const dependency of dependencies) {
      const geometries = dependencyGeometries(dependency, owner, this.sheetOrder);
      const point = dependency.kind !== 'name';
      for (const geometry of geometries) {
        for (const axis of ['row', 'column'] as const) {
          const structural = axis === 'row' ? geometry.rowStructural : geometry.columnStructural;
          const posting: IndexedReference = {
            id: `${storageKey}\u0000${sequence++}`,
            ownerKey,
            sourceId,
            owner: copyAddress(owner),
            sheetId: geometry.sheetId,
            axis,
            start: axis === 'row' ? geometry.startRow : geometry.startColumn,
            end: axis === 'row' ? geometry.endRow : geometry.endColumn,
            crossStart: axis === 'row' ? geometry.startColumn : geometry.startRow,
            crossEnd: axis === 'row' ? geometry.endColumn : geometry.endRow,
            structural,
            point,
          };
          postings.push(posting);
        }
      }
    }

    const position: IndexedReference | undefined = sourceId.startsWith('structural:') ? {
      id: `${storageKey}\u0000position`,
      ownerKey,
      sourceId,
      owner: copyAddress(owner),
      sheetId: owner.sheetId,
      axis: 'row',
      start: owner.row,
      end: owner.row,
      crossStart: owner.column,
      crossEnd: owner.column,
      structural: false,
      point: false,
    } : undefined;
    const previous = this.owners.get(storageKey);
    if (previous) this.remove(owner, sourceId);
    const inserted: IndexedReference[] = [];
    let positionInserted = false;
    try {
      for (const posting of postings) {
        this.insert(posting);
        inserted.push(posting);
      }
      if (position) {
        this.insertOwnerPosition(position);
        positionInserted = true;
      }
      this.owners.set(storageKey, { address: copyAddress(owner), sourceId, postings, position });
    } catch (error) {
      if (positionInserted && position) this.eraseOwnerPosition(position);
      for (const posting of inserted.reverse()) this.erase(posting);
      if (previous) {
        for (const posting of previous.postings) this.insert(posting);
        if (previous.position) this.insertOwnerPosition(previous.position);
        this.owners.set(storageKey, previous);
      }
      throw error;
    }
  }

  remove(owner: CellAddress, sourceId = 'formula'): boolean {
    assertCellAddress(owner);
    assertSourceId(sourceId);
    const storageKey = ownerSourceKey(cellAddressKey(owner), sourceId);
    const entry = this.owners.get(storageKey);
    if (!entry) return false;
    for (const posting of entry.postings) this.erase(posting);
    if (entry.position) this.eraseOwnerPosition(entry.position);
    if (entry.anchorPosition) this.eraseDefinedNameAnchor(entry.anchorPosition);
    this.owners.delete(storageKey);
    return true;
  }

  getDependents(address: CellAddress): readonly CellAddress[] {
    assertCellAddress(address);
    const tree = this.getTrees(address.sheetId, 'row', false)?.byStart;
    const matches: IndexedReference[] = [];
    queryPoint(tree, address.row, matches);
    const owners = new Map<string, CellAddress>();
    for (const posting of matches) {
      if (posting.sourceId !== 'formula' || !posting.point || address.column < posting.crossStart || address.column > posting.crossEnd) continue;
      if (posting.owner) owners.set(posting.ownerKey, posting.owner);
    }
    return [...owners.values()].map(copyAddress).sort(compareCellAddresses);
  }

  /** Return formulas whose references may change at or after a structural index. */
  getStructuralDependents(sheetId: string, axis: Axis, at: number): readonly CellAddress[] {
    if (!sheetId.trim() || !Number.isSafeInteger(at) || at < 0) {
      throw new FormulaReferenceError('Structural reference index query bounds are invalid');
    }
    const tree = this.getTrees(sheetId, axis, false)?.byEnd;
    const matches: IndexedReference[] = [];
    queryEndAtLeast(tree, at, matches);
    const owners = new Map<string, CellAddress>();
    for (const posting of matches) {
      if (posting.structural && posting.owner) owners.set(posting.ownerKey, posting.owner);
    }
    return [...owners.values()].map(copyAddress).sort(compareCellAddresses);
  }

  setDefinedName(
    ownerInput: DefinedNameReferenceOwnerIdentity,
    references: readonly FormulaReferenceNode[],
    context?: CellAddress,
    anchor?: CellAddress,
    initialFailure?: DefinedNameReferenceFailureReason,
  ): void {
    const owner = normalizeDefinedNameOwner(ownerInput);
    if (context) assertCellAddress(context);
    if (anchor) assertCellAddress(anchor);
    const ownerKey = definedNameOwnerKey(owner);
    const storageKey = `${ownerKey}\u0000defined-name`;
    const postings: IndexedReference[] = [];
    let failure = initialFailure;
    let sequence = 0;

    if (!failure) {
      for (const reference of references) {
        let geometries: ReferenceGeometry[];
        try {
          geometries = referenceGeometries(reference, context, this.sheetOrder);
        } catch (error) {
          if (!(error instanceof FormulaReferenceError)) throw error;
          failure = error.message.includes('requires a worksheet context') ? 'unresolved-context' : 'invalid-reference';
          postings.length = 0;
          break;
        }
        for (const geometry of geometries) {
          for (const axis of ['row', 'column'] as const) {
            postings.push({
              id: `${storageKey}\u0000${sequence++}`,
              ownerKey,
              sourceId: 'defined-name',
              definedNameOwner: owner,
              sheetId: geometry.sheetId,
              axis,
              start: axis === 'row' ? geometry.startRow : geometry.startColumn,
              end: axis === 'row' ? geometry.endRow : geometry.endColumn,
              crossStart: axis === 'row' ? geometry.startColumn : geometry.startRow,
              crossEnd: axis === 'row' ? geometry.endColumn : geometry.endRow,
              structural: axis === 'row' ? geometry.rowStructural : geometry.columnStructural,
              point: false,
            });
          }
        }
      }
    }

    const anchorPosition: IndexedReference | undefined = anchor ? {
      id: `${storageKey}\u0000anchor`,
      ownerKey,
      sourceId: 'defined-name-anchor',
      definedNameOwner: owner,
      sheetId: anchor.sheetId,
      axis: 'row',
      start: anchor.row,
      end: anchor.row,
      crossStart: anchor.column,
      crossEnd: anchor.column,
      structural: false,
      point: false,
    } : undefined;
    const previous = this.owners.get(storageKey);
    if (previous) this.removeOwnerEntry(storageKey, previous);
    const inserted: IndexedReference[] = [];
    let anchorInserted = false;
    try {
      for (const posting of postings) {
        this.insert(posting);
        inserted.push(posting);
      }
      if (anchorPosition) {
        this.insertDefinedNameAnchor(anchorPosition);
        anchorInserted = true;
      }
      this.owners.set(storageKey, {
        definedNameOwner: owner,
        sourceId: 'defined-name',
        postings,
        anchorPosition,
        ...(failure ? { failure } : {}),
      });
      if (failure) this.definedNameFailures.set(storageKey, { owner, reason: failure });
      else this.definedNameFailures.delete(storageKey);
    } catch (error) {
      if (anchorInserted && anchorPosition) this.eraseDefinedNameAnchor(anchorPosition);
      for (const posting of inserted.reverse()) this.erase(posting);
      if (previous) this.restoreOwnerEntry(storageKey, previous);
      throw error;
    }
  }

  removeDefinedName(ownerInput: DefinedNameReferenceOwnerIdentity): boolean {
    const owner = normalizeDefinedNameOwner(ownerInput);
    const storageKey = `${definedNameOwnerKey(owner)}\u0000defined-name`;
    const entry = this.owners.get(storageKey);
    if (!entry) return false;
    this.removeOwnerEntry(storageKey, entry);
    return true;
  }

  updateDefinedNames(updates: readonly DefinedNameReferenceIndexUpdate[]): void {
    const staged = updates.map((update) => {
      const owner = normalizeDefinedNameOwner(update.owner);
      return {
        update: { ...update, owner },
        key: `${definedNameOwnerKey(owner)}\u0000defined-name`,
        previous: this.owners.get(`${definedNameOwnerKey(owner)}\u0000defined-name`),
      };
    });
    const keys = new Set<string>();
    for (const item of staged) {
      if (keys.has(item.key)) throw new FormulaReferenceError('Defined-name reference batch contains duplicate owners');
      keys.add(item.key);
    }
    let applied = 0;
    try {
      for (const { update } of staged) {
        if (update.remove) this.removeDefinedName(update.owner);
        else this.setDefinedName(update.owner, update.references ?? [], update.context, update.anchor, update.failure);
        applied += 1;
      }
    } catch (error) {
      for (let index = applied - 1; index >= 0; index -= 1) {
        const item = staged[index]!;
        const current = this.owners.get(item.key);
        if (current) this.removeOwnerEntry(item.key, current);
        if (item.previous) this.restoreOwnerEntry(item.key, item.previous);
      }
      throw error;
    }
  }

  getStructuralDefinedNameDependents(sheetId: string, axis: Axis, at: number): readonly DefinedNameReferenceOwnerIdentity[] {
    if (!sheetId.trim() || !Number.isSafeInteger(at) || at < 0) {
      throw new FormulaReferenceError('Structural defined-name query bounds are invalid');
    }
    const matches: IndexedReference[] = [];
    queryEndAtLeast(this.getTrees(sheetId, axis, false)?.byEnd, at, matches);
    return uniqueDefinedNameOwners(matches.filter((posting) => posting.structural && posting.definedNameOwner));
  }

  getRangeDefinedNameDependents(
    sheetId: string,
    range: { readonly startRow: number; readonly endRow: number; readonly startColumn: number; readonly endColumn: number },
  ): readonly DefinedNameReferenceOwnerIdentity[] {
    assertReferenceRange(sheetId, range);
    const matches: IndexedReference[] = [];
    queryOverlap(this.getTrees(sheetId, 'row', false)?.byStart, range.startRow, range.endRow, matches);
    return uniqueDefinedNameOwners(matches.filter((posting) => posting.definedNameOwner
      && posting.crossStart <= range.endColumn && posting.crossEnd >= range.startColumn));
  }

  getDefinedNamesAnchoredInRange(
    sheetId: string,
    range: { readonly startRow: number; readonly endRow: number; readonly startColumn: number; readonly endColumn: number },
  ): readonly DefinedNameReferenceOwnerIdentity[] {
    assertReferenceRange(sheetId, range);
    const matches: IndexedReference[] = [];
    queryOverlap(this.getDefinedNameAnchorTrees(sheetId, 'row', false)?.byStart, range.startRow, range.endRow, matches);
    return uniqueDefinedNameOwners(matches.filter((posting) => posting.definedNameOwner
      && posting.crossStart >= range.startColumn && posting.crossStart <= range.endColumn));
  }

  getDefinedNamesAnchoredAtOrAfter(sheetId: string, axis: Axis, at: number): readonly DefinedNameReferenceOwnerIdentity[] {
    if (!sheetId.trim() || !Number.isSafeInteger(at) || at < 0) {
      throw new FormulaReferenceError('Defined-name anchor query bounds are invalid');
    }
    const matches: IndexedReference[] = [];
    queryEndAtLeast(this.getDefinedNameAnchorTrees(sheetId, axis, false)?.byEnd, at, matches);
    return uniqueDefinedNameOwners(matches.filter((posting) => posting.definedNameOwner));
  }

  getDefinedNameReferenceFailures(): readonly DefinedNameReferenceFailure[] {
    return [...this.definedNameFailures.values()]
      .map((entry) => ({ owner: { ...entry.owner }, reason: entry.reason }))
      .sort((left, right) => compareDefinedNameOwners(left.owner, right.owner));
  }

  getRangeDependents(
    sheetId: string,
    range: { readonly startRow: number; readonly endRow: number; readonly startColumn: number; readonly endColumn: number },
  ): readonly CellAddress[] {
    if (!sheetId.trim()
      || !Number.isSafeInteger(range.startRow) || range.startRow < 0
      || !Number.isSafeInteger(range.endRow) || range.endRow < range.startRow
      || !Number.isSafeInteger(range.startColumn) || range.startColumn < 0
      || !Number.isSafeInteger(range.endColumn) || range.endColumn < range.startColumn) {
      throw new FormulaReferenceError('Reference range query bounds are invalid');
    }
    const tree = this.getTrees(sheetId, 'row', false)?.byStart;
    const matches: IndexedReference[] = [];
    queryOverlap(tree, range.startRow, range.endRow, matches);
    const owners = new Map<string, CellAddress>();
    for (const posting of matches) {
      if (posting.crossStart > range.endColumn || posting.crossEnd < range.startColumn) continue;
      if (posting.owner) owners.set(posting.ownerKey, posting.owner);
    }
    return [...owners.values()].map(copyAddress).sort(compareCellAddresses);
  }

  getOwnersInRange(
    sheetId: string,
    range: { readonly startRow: number; readonly endRow: number; readonly startColumn: number; readonly endColumn: number },
  ): readonly IndexedReferenceOwnerSource[] {
    if (!sheetId.trim()
      || !Number.isSafeInteger(range.startRow) || range.startRow < 0
      || !Number.isSafeInteger(range.endRow) || range.endRow < range.startRow
      || !Number.isSafeInteger(range.startColumn) || range.startColumn < 0
      || !Number.isSafeInteger(range.endColumn) || range.endColumn < range.startColumn) {
      throw new FormulaReferenceError('Reference owner range query bounds are invalid');
    }
    const matches: IndexedReference[] = [];
    queryOverlap(this.ownerPositions.get(sheetId), range.startRow, range.endRow, matches);
    return matches
      .flatMap((posting) => {
        const owner = posting.owner;
        return owner && posting.crossStart >= range.startColumn && posting.crossStart <= range.endColumn
          ? [{ address: copyAddress(owner), sourceId: posting.sourceId }]
          : [];
      })
      .sort((left, right) => compareCellAddresses(left.address, right.address)
        || (left.sourceId < right.sourceId ? -1 : left.sourceId > right.sourceId ? 1 : 0));
  }

  clear(): void {
    this.owners.clear();
    this.sheets.clear();
    this.ownerPositions.clear();
    this.definedNameAnchors.clear();
    this.definedNameFailures.clear();
  }

  private removeOwnerEntry(storageKey: string, entry: OwnerReferences): void {
    const removedPostings: IndexedReference[] = [];
    let positionRemoved = false;
    let anchorRemoved = false;
    try {
      for (const posting of entry.postings) {
        this.erase(posting);
        removedPostings.push(posting);
      }
      if (entry.position) {
        this.eraseOwnerPosition(entry.position);
        positionRemoved = true;
      }
      if (entry.anchorPosition) {
        this.eraseDefinedNameAnchor(entry.anchorPosition);
        anchorRemoved = true;
      }
    } catch (error) {
      for (const posting of removedPostings.reverse()) this.insert(posting);
      if (positionRemoved && entry.position) this.insertOwnerPosition(entry.position);
      if (anchorRemoved && entry.anchorPosition) this.insertDefinedNameAnchor(entry.anchorPosition);
      this.owners.set(storageKey, entry);
      if (entry.definedNameOwner && entry.failure) {
        this.definedNameFailures.set(storageKey, { owner: entry.definedNameOwner, reason: entry.failure });
      }
      throw error;
    }
    this.owners.delete(storageKey);
    this.definedNameFailures.delete(storageKey);
  }

  private restoreOwnerEntry(storageKey: string, entry: OwnerReferences): void {
    for (const posting of entry.postings) this.insert(posting);
    if (entry.position) this.insertOwnerPosition(entry.position);
    if (entry.anchorPosition) this.insertDefinedNameAnchor(entry.anchorPosition);
    this.owners.set(storageKey, entry);
    if (entry.definedNameOwner && entry.failure) {
      this.definedNameFailures.set(storageKey, { owner: entry.definedNameOwner, reason: entry.failure });
    }
  }

  private insertOwnerPosition(position: IndexedReference): void {
    this.ownerPositions.set(position.sheetId, insertNode(this.ownerPositions.get(position.sheetId), position, 'start'));
  }

  private eraseOwnerPosition(position: IndexedReference): void {
    const root = this.ownerPositions.get(position.sheetId);
    if (!root) throw new Error('REFERENCE_INDEX_INVARIANT: owner position has no sheet index');
    const next = removeNode(root, position, 'start');
    if (next) this.ownerPositions.set(position.sheetId, next);
    else this.ownerPositions.delete(position.sheetId);
  }

  private insertDefinedNameAnchor(position: IndexedReference): void {
    const committed: Array<{ axis: Axis; indexed: IndexedReference }> = [];
    try {
      for (const axis of ['row', 'column'] as const) {
        const indexed = axis === 'row'
          ? position
          : { ...position, axis, start: position.crossStart, end: position.crossStart, crossStart: position.start, crossEnd: position.start };
        const trees = this.getDefinedNameAnchorTrees(indexed.sheetId, axis, true)!;
        const byStart = insertNode(trees.byStart, indexed, 'start');
        const byEnd = insertNode(trees.byEnd, indexed, 'end');
        trees.byStart = byStart;
        trees.byEnd = byEnd;
        committed.push({ axis, indexed });
      }
    } catch (error) {
      for (const { axis, indexed } of committed.reverse()) {
        const trees = this.getDefinedNameAnchorTrees(indexed.sheetId, axis, false)!;
        trees.byStart = removeNode(trees.byStart!, indexed, 'start');
        trees.byEnd = removeNode(trees.byEnd!, indexed, 'end');
        if (!trees.byStart && !trees.byEnd) this.definedNameAnchors.get(indexed.sheetId)?.delete(axis);
        if (this.definedNameAnchors.get(indexed.sheetId)?.size === 0) this.definedNameAnchors.delete(indexed.sheetId);
      }
      throw error;
    }
  }

  private eraseDefinedNameAnchor(position: IndexedReference): void {
    const staged: Array<{ axis: Axis; indexed: IndexedReference; trees: AxisTrees; byStart?: IntervalNode; byEnd?: IntervalNode }> = [];
    for (const axis of ['row', 'column'] as const) {
      const indexed = axis === 'row'
        ? position
        : { ...position, axis, start: position.crossStart, end: position.crossStart, crossStart: position.start, crossEnd: position.start };
      const trees = this.getDefinedNameAnchorTrees(indexed.sheetId, axis, false);
      if (!trees?.byStart || !trees.byEnd) throw new Error('REFERENCE_INDEX_INVARIANT: defined-name anchor has no axis index');
      staged.push({
        axis,
        indexed,
        trees,
        byStart: removeNode(trees.byStart, indexed, 'start'),
        byEnd: removeNode(trees.byEnd, indexed, 'end'),
      });
    }
    for (const { axis, indexed, trees, byStart, byEnd } of staged) {
      trees.byStart = byStart;
      trees.byEnd = byEnd;
      if (!byStart && !byEnd) this.definedNameAnchors.get(indexed.sheetId)?.delete(axis);
      if (this.definedNameAnchors.get(indexed.sheetId)?.size === 0) this.definedNameAnchors.delete(indexed.sheetId);
    }
  }

  private insert(posting: IndexedReference): void {
    const trees = this.getTrees(posting.sheetId, posting.axis, true)!;
    const byStart = insertNode(trees.byStart, posting, 'start');
    const byEnd = insertNode(trees.byEnd, posting, 'end');
    trees.byStart = byStart;
    trees.byEnd = byEnd;
  }

  private erase(posting: IndexedReference): void {
    const trees = this.getTrees(posting.sheetId, posting.axis, false);
    if (!trees) throw new Error('REFERENCE_INDEX_INVARIANT: owner posting has no sheet index');
    const byStart = removeNode(trees.byStart, posting, 'start');
    const byEnd = removeNode(trees.byEnd, posting, 'end');
    trees.byStart = byStart;
    trees.byEnd = byEnd;
    if (!trees.byStart && !trees.byEnd) {
      const axes = this.sheets.get(posting.sheetId);
      axes?.delete(posting.axis);
      if (axes?.size === 0) this.sheets.delete(posting.sheetId);
    }
  }

  private getTrees(sheetId: string, axis: Axis, create: boolean): AxisTrees | undefined {
    let axes = this.sheets.get(sheetId);
    if (!axes && create) {
      axes = new Map<Axis, AxisTrees>();
      this.sheets.set(sheetId, axes);
    }
    if (!axes) return undefined;
    let trees = axes.get(axis);
    if (!trees && create) {
      trees = {};
      axes.set(axis, trees);
    }
    return trees;
  }

  private getDefinedNameAnchorTrees(sheetId: string, axis: Axis, create: boolean): AxisTrees | undefined {
    let axes = this.definedNameAnchors.get(sheetId);
    if (!axes && create) {
      axes = new Map<Axis, AxisTrees>();
      this.definedNameAnchors.set(sheetId, axes);
    }
    if (!axes) return undefined;
    let trees = axes.get(axis);
    if (!trees && create) {
      trees = {};
      axes.set(axis, trees);
    }
    return trees;
  }
}

function assertSourceId(sourceId: string): void {
  if (typeof sourceId !== 'string' || sourceId.trim().length === 0 || sourceId.includes('\u0000')) {
    throw new FormulaReferenceError('Reference owner source identity is invalid');
  }
}

function ownerSourceKey(ownerKey: string, sourceId: string): string {
  return `${ownerKey}\u0000${sourceId}`;
}

function normalizeDefinedNameOwner(owner: DefinedNameReferenceOwnerIdentity): DefinedNameReferenceOwnerIdentity {
  const name = owner.name.trim();
  if (!name || (owner.scope !== 'workbook' && owner.scope !== 'sheet')) {
    throw new FormulaReferenceError('Defined-name reference owner identity is invalid');
  }
  if (owner.scope === 'sheet' && !owner.sheetId?.trim()) {
    throw new FormulaReferenceError('Sheet-scoped defined-name reference owner requires a worksheet identity');
  }
  if (owner.scope === 'workbook' && owner.sheetId !== undefined) {
    throw new FormulaReferenceError('Workbook-scoped defined-name reference owner cannot specify a worksheet identity');
  }
  return { scope: owner.scope, name, ...(owner.sheetId ? { sheetId: owner.sheetId.trim() } : {}) };
}

function definedNameOwnerKey(owner: DefinedNameReferenceOwnerIdentity): string {
  return `defined-name:${JSON.stringify([owner.scope, owner.scope === 'sheet' ? owner.sheetId : null, owner.name.trim().toUpperCase()])}`;
}

function compareDefinedNameOwners(left: DefinedNameReferenceOwnerIdentity, right: DefinedNameReferenceOwnerIdentity): number {
  const leftKey = definedNameOwnerKey(left);
  const rightKey = definedNameOwnerKey(right);
  return leftKey < rightKey ? -1 : leftKey > rightKey ? 1 : 0;
}

function uniqueDefinedNameOwners(postings: readonly IndexedReference[]): DefinedNameReferenceOwnerIdentity[] {
  const owners = new Map<string, DefinedNameReferenceOwnerIdentity>();
  for (const posting of postings) {
    const owner = posting.definedNameOwner;
    if (owner) owners.set(definedNameOwnerKey(owner), owner);
  }
  return [...owners.values()].map((owner) => ({ ...owner })).sort(compareDefinedNameOwners);
}

function assertReferenceRange(
  sheetId: string,
  range: { readonly startRow: number; readonly endRow: number; readonly startColumn: number; readonly endColumn: number },
): void {
  if (!sheetId.trim()
    || !Number.isSafeInteger(range.startRow) || range.startRow < 0
    || !Number.isSafeInteger(range.endRow) || range.endRow < range.startRow
    || !Number.isSafeInteger(range.startColumn) || range.startColumn < 0
    || !Number.isSafeInteger(range.endColumn) || range.endColumn < range.startColumn) {
    throw new FormulaReferenceError('Reference range query bounds are invalid');
  }
}

function dependencyGeometries(
  dependency: FormulaDependency,
  owner: CellAddress,
  sheetOrder: readonly FormulaSheetIdentity[],
): ReferenceGeometry[] {
  switch (dependency.kind) {
    case 'cell':
      return [cellGeometry(
        dependency.address.sheetId,
        dependency.address.row,
        dependency.address.column,
      )];
    case 'range': {
      const sheetId = dependency.start.sheetId;
      const endSheetId = dependency.end.sheetId;
      if (sheetId !== endSheetId) throw new FormulaReferenceError('A range cannot cross worksheets');
      return [rectangleGeometry(sheetId, dependency.start.row, dependency.end.row, dependency.start.column, dependency.end.column, true, true)];
    }
    case 'reference':
      return referenceGeometries(dependency.reference, owner, sheetOrder);
    case 'name':
      return [];
  }
}

function referenceGeometries(
  reference: FormulaReferenceNode,
  owner: CellAddress | undefined,
  sheetOrder: readonly FormulaSheetIdentity[],
): ReferenceGeometry[] {
  switch (reference.type) {
    case 'cell-reference': {
      const { sheetId, row, column } = reference.reference;
      return [cellGeometry(resolveReferenceSheetId(sheetId, owner, sheetOrder), row, column)];
    }
    case 'range-reference': {
      const first = reference.start.reference;
      const second = reference.end.reference;
      const sheetId = resolveReferenceSheetId(first.sheetId, owner, sheetOrder);
      const endSheetId = resolveReferenceSheetId(second.sheetId ?? first.sheetId, owner, sheetOrder);
      if (sheetId !== endSheetId) throw new FormulaReferenceError('A range cannot cross worksheets');
      return [rectangleGeometry(
        sheetId,
        first.row,
        second.row,
        first.column,
        second.column,
        true,
        true,
      )];
    }
    case 'whole-row-reference':
      return [rectangleGeometry(resolveReferenceSheetId(reference.sheetId, owner, sheetOrder), reference.startRow, reference.endRow, 0, Number.MAX_SAFE_INTEGER, true, false)];
    case 'whole-column-reference':
      return [rectangleGeometry(resolveReferenceSheetId(reference.sheetId, owner, sheetOrder), 0, Number.MAX_SAFE_INTEGER, reference.startColumn, reference.endColumn, false, true)];
    case 'reference-union':
      return reference.references.flatMap((item) => referenceGeometries(item, owner, sheetOrder));
    case 'reference-intersection': {
      const left = referenceGeometries(reference.left, owner, sheetOrder);
      const right = referenceGeometries(reference.right, owner, sheetOrder);
      const intersections: ReferenceGeometry[] = [];
      for (const leftRange of left) {
        for (const rightRange of right) {
          const intersection = intersectGeometry(leftRange, rightRange);
          if (intersection) intersections.push(intersection);
        }
      }
      return intersections;
    }
    case 'sheet-range-reference': {
      const startSheetId = resolveFormulaSheetId(reference.qualifier.startSheetId, owner?.sheetId ?? '', sheetOrder);
      const endSheetId = resolveFormulaSheetId(reference.qualifier.endSheetId, owner?.sheetId ?? '', sheetOrder);
      const start = sheetOrder.findIndex((sheet) => sheet.id === startSheetId);
      const end = sheetOrder.findIndex((sheet) => sheet.id === endSheetId);
      if (start < 0 || end < 0) throw new FormulaReferenceError('3-D reference sheet boundary is unresolved');
      return sheetOrder
        .slice(Math.min(start, end), Math.max(start, end) + 1)
        .flatMap((sheet) => referenceGeometries(reference.reference, owner ? { ...owner, sheetId: sheet.id } : { sheetId: sheet.id, row: 0, column: 0 }, sheetOrder));
    }
    case 'spill-reference':
      return isReferenceNode(reference.operand) ? referenceGeometries(reference.operand, owner, sheetOrder) : [];
    case 'external-reference':
    case 'table-reference':
    case 'invalid-reference':
      return [];
  }
}

function resolveReferenceSheetId(
  reference: string | undefined,
  owner: CellAddress | undefined,
  sheetOrder: readonly FormulaSheetIdentity[],
): string {
  if (reference === undefined && !owner) {
    throw new FormulaReferenceError('A defined-name reference requires a worksheet context');
  }
  return resolveFormulaSheetId(reference, owner?.sheetId ?? '', sheetOrder);
}

function isReferenceNode(node: FormulaAst): node is FormulaReferenceNode {
  switch (node.type) {
    case 'cell-reference':
    case 'invalid-reference':
    case 'range-reference':
    case 'whole-column-reference':
    case 'whole-row-reference':
    case 'spill-reference':
    case 'table-reference':
    case 'reference-union':
    case 'reference-intersection':
    case 'sheet-range-reference':
    case 'external-reference':
      return true;
    default:
      return false;
  }
}

function cellGeometry(sheetId: string, row: number, column: number): ReferenceGeometry {
  return rectangleGeometry(sheetId, row, row, column, column, true, true);
}

function rectangleGeometry(
  sheetId: string,
  startRow: number,
  endRow: number,
  startColumn: number,
  endColumn: number,
  rowStructural: boolean,
  columnStructural: boolean,
): ReferenceGeometry {
  if (!sheetId.trim()
    || !Number.isSafeInteger(startRow) || startRow < 0
    || !Number.isSafeInteger(endRow) || endRow < 0
    || !Number.isSafeInteger(startColumn) || startColumn < 0
    || !Number.isSafeInteger(endColumn) || endColumn < 0) {
    throw new FormulaReferenceError('Reference index geometry is outside valid worksheet coordinates');
  }
  return {
    sheetId,
    startRow: Math.min(startRow, endRow),
    endRow: Math.max(startRow, endRow),
    startColumn: Math.min(startColumn, endColumn),
    endColumn: Math.max(startColumn, endColumn),
    rowStructural,
    columnStructural,
  };
}

function intersectGeometry(left: ReferenceGeometry, right: ReferenceGeometry): ReferenceGeometry | undefined {
  if (left.sheetId !== right.sheetId) return undefined;
  const startRow = Math.max(left.startRow, right.startRow);
  const endRow = Math.min(left.endRow, right.endRow);
  const startColumn = Math.max(left.startColumn, right.startColumn);
  const endColumn = Math.min(left.endColumn, right.endColumn);
  if (startRow > endRow || startColumn > endColumn) return undefined;
  return rectangleGeometry(
    left.sheetId,
    startRow,
    endRow,
    startColumn,
    endColumn,
    left.rowStructural || right.rowStructural,
    left.columnStructural || right.columnStructural,
  );
}

function copyAddress(address: CellAddress): CellAddress {
  return { sheetId: address.sheetId, row: address.row, column: address.column };
}

function compareNumber(left: number, right: number): number {
  return left < right ? -1 : left > right ? 1 : 0;
}

function comparePosting(left: IndexedReference, right: IndexedReference, key: TreeKey): number {
  const coordinate = key === 'start'
    ? compareNumber(left.start, right.start)
    : compareNumber(left.end, right.end);
  return coordinate
    || compareNumber(left.start, right.start)
    || compareNumber(left.end, right.end)
    || left.id.localeCompare(right.id);
}

function node(value: IndexedReference, left?: IntervalNode, right?: IntervalNode): IntervalNode {
  return {
    value,
    ...(left ? { left } : {}),
    ...(right ? { right } : {}),
    height: 1 + Math.max(left?.height ?? 0, right?.height ?? 0),
    maxEnd: Math.max(value.end, left?.maxEnd ?? Number.MIN_SAFE_INTEGER, right?.maxEnd ?? Number.MIN_SAFE_INTEGER),
    minStart: Math.min(value.start, left?.minStart ?? Number.MAX_SAFE_INTEGER, right?.minStart ?? Number.MAX_SAFE_INTEGER),
  };
}

function balance(root: IntervalNode): IntervalNode {
  const factor = (root.left?.height ?? 0) - (root.right?.height ?? 0);
  if (factor > 1 && root.left) {
    if ((root.left.left?.height ?? 0) < (root.left.right?.height ?? 0)) {
      const rotatedLeft = rotateLeft(root.left);
      return rotateRight(node(root.value, rotatedLeft, root.right));
    }
    return rotateRight(root);
  }
  if (factor < -1 && root.right) {
    if ((root.right.right?.height ?? 0) < (root.right.left?.height ?? 0)) {
      const rotatedRight = rotateRight(root.right);
      return rotateLeft(node(root.value, root.left, rotatedRight));
    }
    return rotateLeft(root);
  }
  return root;
}

function rotateLeft(root: IntervalNode): IntervalNode {
  const pivot = root.right;
  if (!pivot) return root;
  return node(pivot.value, node(root.value, root.left, pivot.left), pivot.right);
}

function rotateRight(root: IntervalNode): IntervalNode {
  const pivot = root.left;
  if (!pivot) return root;
  return node(pivot.value, pivot.left, node(root.value, pivot.right, root.right));
}

function insertNode(root: IntervalNode | undefined, value: IndexedReference, key: TreeKey): IntervalNode {
  if (!root) return node(value);
  const order = comparePosting(value, root.value, key);
  if (order === 0) throw new Error('REFERENCE_INDEX_INVARIANT: duplicate interval posting');
  return balance(order < 0
    ? node(root.value, insertNode(root.left, value, key), root.right)
    : node(root.value, root.left, insertNode(root.right, value, key)));
}

function removeNode(root: IntervalNode | undefined, value: IndexedReference, key: TreeKey): IntervalNode | undefined {
  if (!root) throw new Error('REFERENCE_INDEX_INVARIANT: interval posting was not indexed');
  const order = comparePosting(value, root.value, key);
  if (order < 0) return balance(node(root.value, removeNode(root.left, value, key), root.right));
  if (order > 0) return balance(node(root.value, root.left, removeNode(root.right, value, key)));
  if (!root.left) return root.right;
  if (!root.right) return root.left;
  const successor = minimum(root.right);
  return balance(node(successor.value, root.left, removeNode(root.right, successor.value, key)));
}

function minimum(root: IntervalNode): IntervalNode {
  return root.left ? minimum(root.left) : root;
}

function queryPoint(root: IntervalNode | undefined, point: number, result: IndexedReference[]): void {
  if (!root || root.maxEnd < point || root.minStart > point) return;
  if (root.left && root.left.maxEnd >= point && root.left.minStart <= point) queryPoint(root.left, point, result);
  if (root.value.start <= point && root.value.end >= point) result.push(root.value);
  if (root.value.start <= point) queryPoint(root.right, point, result);
}

function queryEndAtLeast(root: IntervalNode | undefined, threshold: number, result: IndexedReference[]): void {
  if (!root) return;
  if (root.value.end < threshold) {
    queryEndAtLeast(root.right, threshold, result);
    return;
  }
  queryEndAtLeast(root.left, threshold, result);
  result.push(root.value);
  collectAll(root.right, result);
}

function queryOverlap(root: IntervalNode | undefined, start: number, end: number, result: IndexedReference[]): void {
  if (!root || root.maxEnd < start || root.minStart > end) return;
  queryOverlap(root.left, start, end, result);
  if (root.value.start <= end && root.value.end >= start) result.push(root.value);
  queryOverlap(root.right, start, end, result);
}

function collectAll(root: IntervalNode | undefined, result: IndexedReference[]): void {
  if (!root) return;
  result.push(root.value);
  collectAll(root.left, result);
  collectAll(root.right, result);
}
