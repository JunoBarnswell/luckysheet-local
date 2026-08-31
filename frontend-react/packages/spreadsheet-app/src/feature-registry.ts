import type { CommandRuntime } from '@react-sheets/command-runtime';
import { registerSheetCommands } from '@react-sheets/sheet-features';
import { registerPlatformFeatures } from './platform-features';
import { registerChartFeature } from './features/chart';
import { registerDrawingFeature, type DrawingRuntime } from './features/drawing';
import { registerDataSourceFeature } from './features/data-source';
import { registerEditingFeatures } from './features/editing';
import { registerPivotFeature } from './features/pivot';
import { registerPivotControlFeature } from './features/pivot-controls';
import { registerReviewFeature } from './features/review/commands';
import { registerSparklineFeature } from './features/sparkline';
import { registerInsertCommands } from './features/insert';
import { registerFindReplaceFeature } from './features/find-replace/commands';
import { buildExcelParityReport, createExcelFeatureRegistry, type ExcelParityReport } from './excel-parity';
import {
  RIBBON_COMMAND_CATALOG,
  RIBBON_TAB_SURFACES,
  type RibbonCommandId,
  type RibbonSurfaceDefinition,
} from './ui-command-catalog';

/**
 * A feature manifest declares ownership of a visual surface.  `surfaceCommandId`
 * is the catalog identity consumed by the Ribbon, while `requiredCommandId`
 * is the executable command capability used to resolve the feature owner.
 * Keeping the two names distinct prevents a UI identity from being mistaken for
 * a model command (the old manifest field conflated those contracts).
 */
export interface SpreadsheetFeatureSurface {
  readonly id: string;
  readonly tab: RibbonSurfaceDefinition['tab'];
  readonly group: RibbonSurfaceDefinition['group'];
  readonly label: string;
  readonly surfaceCommandId?: RibbonCommandId;
  readonly requiredCommandId?: string;
  readonly controlId?: RibbonSurfaceDefinition['controlId'];
  readonly icon: string;
  readonly order?: number;
  readonly appearance?: RibbonSurfaceDefinition['appearance'];
  readonly breakpoints?: RibbonSurfaceDefinition['breakpoints'];
  readonly overflowTarget?: RibbonSurfaceDefinition['overflowTarget'];
  readonly menuId?: RibbonSurfaceDefinition['menuId'];
  readonly ariaLabel?: RibbonSurfaceDefinition['ariaLabel'];
}

export interface SpreadsheetFeatureManifest {
  id: string;
  version: string;
  /** Document/environment gates are part of capability compilation, not UI guesses. */
  documentTypes?: readonly string[];
  environments?: readonly string[];
  dependencies?: string[];
  commandIds: string[];
  mutationIds?: string[];
  ribbon?: ReadonlyArray<SpreadsheetFeatureSurface>;
  contextualTabs?: ReadonlyArray<SpreadsheetFeatureSurface>;
  permissions?: string[];
  parity?: ExcelParityReport;
  lifecycle?: FeatureLifecycleHooks;
}

export type FeatureLifecyclePhase = 'starting' | 'ready' | 'rendered' | 'steady' | 'failed' | 'disposed';

export interface FeatureLifecycleContext {
  readonly phase: FeatureLifecyclePhase;
  readonly documentType: string;
  readonly environment: string;
  readonly manifests: readonly SpreadsheetFeatureManifest[];
}

export interface FeatureLifecycleHooks {
  readonly starting?: (context: FeatureLifecycleContext) => void;
  readonly ready?: (context: FeatureLifecycleContext) => void;
  readonly rendered?: (context: FeatureLifecycleContext) => void;
  readonly steady?: (context: FeatureLifecycleContext) => void;
  readonly dispose?: (context: FeatureLifecycleContext) => void;
}

export interface FeatureRuntimeActivation {
  readonly documentType?: string;
  readonly environment?: string;
}

export interface CompiledFeatureSurfaceEntry {
  readonly featureId: string;
  readonly id: string;
  readonly tab: RibbonSurfaceDefinition['tab'];
  readonly group: RibbonSurfaceDefinition['group'];
  readonly label: string;
  /** Catalog command identity. It is intentionally not the executable command id. */
  readonly commandId?: RibbonCommandId;
  readonly controlId?: RibbonSurfaceDefinition['controlId'];
  readonly order: number;
  readonly appearance: RibbonSurfaceDefinition['appearance'];
  readonly breakpoints: RibbonSurfaceDefinition['breakpoints'];
  readonly overflowTarget?: RibbonSurfaceDefinition['overflowTarget'];
  readonly menuId?: RibbonSurfaceDefinition['menuId'];
  readonly ariaLabel?: RibbonSurfaceDefinition['ariaLabel'];
  /** Executable capability required by this surface, when one exists. */
  readonly requiredCommandId?: string;
  readonly icon: string;
}

export interface CompiledFeatureSurfaceSchema {
  readonly ribbon: readonly CompiledFeatureSurfaceEntry[];
  readonly contextualTabs: readonly CompiledFeatureSurfaceEntry[];
}

/**
 * Dependency-closed feature lifecycle. Registration is pure (no DOM/model
 * access); activation moves all selected features through one ordered state
 * machine and fails before exposing any surface when a dependency is absent.
 */
export class SpreadsheetFeatureRuntime {
  private manifests = new Map<string, SpreadsheetFeatureManifest>();
  private ordered: SpreadsheetFeatureManifest[] = [];
  private activeContext: FeatureLifecycleContext | null = null;
  private currentPhase: FeatureLifecyclePhase = 'disposed';
  private parityReport: ExcelParityReport | undefined;

  register(manifest: SpreadsheetFeatureManifest): void {
    const normalized = normalizeManifest(manifest);
    if (this.manifests.has(normalized.id)) throw new Error(`FEATURE_DUPLICATE: ${normalized.id}`);
    this.manifests.set(normalized.id, normalized);
  }

  load(manifests: readonly SpreadsheetFeatureManifest[]): void {
    this.dispose();
    this.manifests.clear();
    for (const manifest of manifests) this.register(manifest);
    this.ordered = topologicallyOrder([...this.manifests.values()]);
    this.currentPhase = 'starting';
  }

  activate(options: FeatureRuntimeActivation = {}): readonly SpreadsheetFeatureManifest[] {
    if (this.currentPhase !== 'starting' && this.currentPhase !== 'disposed') throw new Error(`FEATURE_LIFECYCLE_INVALID: cannot activate from ${this.currentPhase}`);
    const documentType = options.documentType ?? 'spreadsheet';
    const environment = options.environment ?? (typeof window === 'undefined' ? 'worker' : 'browser');
    const selected = this.ordered.filter((manifest) => {
      const documents = manifest.documentTypes;
      const environments = manifest.environments;
      return (!documents || documents.includes(documentType)) && (!environments || environments.includes(environment));
    });
    const selectedIds = new Set(selected.map((manifest) => manifest.id));
    for (const manifest of selected) {
      for (const dependency of manifest.dependencies ?? []) {
        if (!selectedIds.has(dependency)) throw new Error(`FEATURE_DEPENDENCY_UNAVAILABLE: ${manifest.id} requires ${dependency}`);
      }
    }
    const context = { phase: 'starting' as const, documentType, environment, manifests: selected };
    this.activeContext = context;
    try {
      this.invoke('starting', context);
    } catch (error) {
      this.currentPhase = 'failed';
      throw new Error(`FEATURE_LIFECYCLE_FAILED: ${error instanceof Error ? error.message : 'feature lifecycle hook failed'}`);
    }
    return selected;
  }

  /** Advance one lifecycle boundary after its owning host has completed. */
  advance(phase: Extract<FeatureLifecyclePhase, 'ready' | 'rendered' | 'steady'>): void {
    const order: readonly FeatureLifecyclePhase[] = ['starting', 'ready', 'rendered', 'steady'];
    if (!this.activeContext || this.currentPhase === 'disposed' || this.currentPhase === 'failed') throw new Error(`FEATURE_LIFECYCLE_INVALID: cannot advance from ${this.currentPhase}`);
    const current = order.indexOf(this.currentPhase);
    const next = order.indexOf(phase);
    if (next !== current + 1) throw new Error(`FEATURE_LIFECYCLE_INVALID: expected ${order[current + 1]}, received ${phase}`);
    try {
      this.currentPhase = phase;
      this.invoke(phase, { ...this.activeContext, phase });
    } catch (error) {
      this.currentPhase = 'failed';
      throw new Error(`FEATURE_LIFECYCLE_FAILED: ${error instanceof Error ? error.message : 'feature lifecycle hook failed'}`);
    }
  }

  dispose(): void {
    if (this.activeContext) this.invoke('dispose', { ...this.activeContext, phase: 'disposed' });
    this.activeContext = null;
    this.currentPhase = 'disposed';
  }

  getPhase(): FeatureLifecyclePhase { return this.currentPhase; }

  getActiveManifests(): readonly SpreadsheetFeatureManifest[] { return this.activeContext?.manifests ?? []; }

  getSurfaceSchema(): CompiledFeatureSurfaceSchema {
    return compileFeatureSurfaceSchema(this.getActiveManifests());
  }

  getManifests(): readonly SpreadsheetFeatureManifest[] { return this.ordered.map((manifest) => ({ ...manifest, commandIds: [...manifest.commandIds] })); }

  setParityReport(report: ExcelParityReport): void { this.parityReport = structuredClone(report); }

  getParityReport(): ExcelParityReport {
    if (!this.parityReport) throw new Error('FEATURE_REGISTRY_NOT_COMPILED: parity report is unavailable before registration');
    return structuredClone(this.parityReport);
  }

  private invoke(phase: keyof FeatureLifecycleHooks, context: FeatureLifecycleContext): void {
    for (const manifest of context.manifests) manifest.lifecycle?.[phase]?.(context);
  }
}

function normalizeManifest(manifest: SpreadsheetFeatureManifest): SpreadsheetFeatureManifest {
  if (!manifest.id.trim() || !manifest.version.trim()) throw new Error('FEATURE_MANIFEST_INVALID: id and version are required');
  const commands = [...new Set(manifest.commandIds)];
  const normalizeSurface = (entry: SpreadsheetFeatureSurface): SpreadsheetFeatureSurface => {
    if (!entry.id.trim() || !entry.tab.trim() || !entry.group.trim()) throw new Error(`FEATURE_SURFACE_INVALID: ${manifest.id} has an incomplete surface`);
    if (!entry.surfaceCommandId && !entry.controlId) {
      throw new Error(`FEATURE_SURFACE_INVALID: ${manifest.id}/${entry.id} requires a catalog command or control identity`);
    }
    if (entry.requiredCommandId !== undefined && !entry.requiredCommandId.trim()) throw new Error(`FEATURE_SURFACE_INVALID: ${manifest.id}/${entry.id} has an empty required command`);
    return {
      ...entry,
      id: entry.id.trim(),
      tab: entry.tab,
      group: entry.group,
      label: entry.label.trim(),
      ...(entry.requiredCommandId ? { requiredCommandId: entry.requiredCommandId.trim() } : {}),
      ...(entry.breakpoints ? { breakpoints: [...entry.breakpoints] } : {}),
    };
  };
  return {
    ...manifest,
    id: manifest.id.trim(),
    version: manifest.version.trim(),
    commandIds: commands,
    ...(manifest.dependencies ? { dependencies: [...new Set(manifest.dependencies)] } : {}),
    ...(manifest.documentTypes ? { documentTypes: [...manifest.documentTypes] } : {}),
    ...(manifest.environments ? { environments: [...manifest.environments] } : {}),
    ...(manifest.ribbon ? { ribbon: manifest.ribbon.map(normalizeSurface) } : {}),
    ...(manifest.contextualTabs ? { contextualTabs: manifest.contextualTabs.map(normalizeSurface) } : {}),
  };
}

function topologicallyOrder(manifests: readonly SpreadsheetFeatureManifest[]): SpreadsheetFeatureManifest[] {
  const byId = new Map(manifests.map((manifest) => [manifest.id, manifest]));
  const temporary = new Set<string>();
  const permanent = new Set<string>();
  const result: SpreadsheetFeatureManifest[] = [];
  const visit = (id: string, owner?: string): void => {
    if (permanent.has(id)) return;
    if (temporary.has(id)) throw new Error(`FEATURE_DEPENDENCY_CYCLE: ${id}`);
    const manifest = byId.get(id);
    if (!manifest) throw new Error(`FEATURE_DEPENDENCY_UNAVAILABLE: ${owner ?? 'feature'} requires ${id}`);
    temporary.add(id);
    for (const dependency of manifest.dependencies ?? []) visit(dependency, manifest.id);
    temporary.delete(id);
    permanent.add(id);
    result.push(manifest);
  };
  for (const manifest of manifests) visit(manifest.id);
  return result;
}

function withCoreDependency(manifest: SpreadsheetFeatureManifest): SpreadsheetFeatureManifest {
  return { ...manifest, dependencies: [...new Set(['sheet-features', ...(manifest.dependencies ?? [])])] };
}

const PRIMARY_RIBBON_TABS = new Set(['file', 'home', 'insert', 'pageLayout', 'formulas', 'data', 'review', 'view', 'settings']);
const CANONICAL_SURFACE_TABS = new Set(RIBBON_TAB_SURFACES.map((surface) => surface.tab));
const ALL_RIBBON_BREAKPOINTS: RibbonSurfaceDefinition['breakpoints'] = ['wide', 'compact', 'narrow'];

/**
 * Compile the visual schema into feature ownership once. Home, Insert and the
 * shape-format contextual tab are defined by `RIBBON_TAB_SURFACES`; every other
 * tab is projected from command placements only when no feature supplied an
 * explicit surface for that tab. The UI never reads either source directly.
 */
function attachCanonicalRibbonSurfaces(manifests: readonly SpreadsheetFeatureManifest[]): SpreadsheetFeatureManifest[] {
  const owners = new Map<string, string>();
  for (const manifest of manifests) for (const commandId of manifest.commandIds) owners.set(commandId, manifest.id);
  const core = manifests.find((manifest) => manifest.id === 'sheet-features');
  const next = manifests.map((manifest) => ({
    ...manifest,
    // A stale Home/Insert surface can never shadow the canonical visual schema.
    ribbon: [...(manifest.ribbon ?? []).filter((surface) => !CANONICAL_SURFACE_TABS.has(surface.tab))],
    contextualTabs: [...(manifest.contextualTabs ?? []).filter((surface) => !CANONICAL_SURFACE_TABS.has(surface.tab))],
  }));
  const byId = new Map(next.map((manifest) => [manifest.id, manifest]));

  const catalogById = new Map(RIBBON_COMMAND_CATALOG.map((definition) => [definition.id, definition] as const));
  const explicitTabs = new Set(next.flatMap((manifest) => [
    ...(manifest.ribbon ?? []).map((surface) => surface.tab),
    ...(manifest.contextualTabs ?? []).map((surface) => surface.tab),
  ]));

  for (const surface of RIBBON_TAB_SURFACES) {
    const definition = surface.commandId ? catalogById.get(surface.commandId) : undefined;
    if (surface.commandId && !definition) throw new Error(`FEATURE_SURFACE_UNAVAILABLE: canonical surface ${surface.id} references unknown command ${surface.commandId}`);
    const requiredCommandId = definition?.commandId;
    const ownerId = requiredCommandId ? owners.get(requiredCommandId) : core?.id;
    if (!ownerId) throw new Error(`FEATURE_SURFACE_UNAVAILABLE: canonical surface ${surface.id} requires ${requiredCommandId}`);
    const owner = byId.get(ownerId);
    if (!owner) throw new Error(`FEATURE_SURFACE_UNAVAILABLE: canonical surface ${surface.id} owner ${ownerId} is not registered`);
    const entry: SpreadsheetFeatureSurface = {
      ...surface,
      label: definition?.labelKey ?? surface.id,
      icon: definition?.icon ?? 'command',
      ...(surface.commandId ? { surfaceCommandId: surface.commandId } : {}),
      ...(requiredCommandId ? { requiredCommandId } : {}),
    };
    const target = PRIMARY_RIBBON_TABS.has(surface.tab) ? owner.ribbon! : owner.contextualTabs!;
    if (!target.some((candidate) => candidate.id === entry.id && candidate.tab === entry.tab && candidate.group === entry.group)) target.push(entry);
  }

  for (const definition of RIBBON_COMMAND_CATALOG) {
    for (const placement of definition.placements) {
      // Canonical surfaces own the complete Home/Insert/shape-format visual
      // membership. A feature-provided surface similarly owns its tab; only
      // otherwise-unclaimed tabs may be projected from command placements.
      if (CANONICAL_SURFACE_TABS.has(placement.tab) || explicitTabs.has(placement.tab)) continue;
      const ownerId = owners.get(definition.commandId ?? '') ?? core?.id;
      if (!ownerId) throw new Error(`FEATURE_SURFACE_UNAVAILABLE: ${definition.id} has no owning feature`);
      const owner = byId.get(ownerId);
      if (!owner) throw new Error(`FEATURE_SURFACE_UNAVAILABLE: ${definition.id} owner ${ownerId} is not registered`);
      const entry: SpreadsheetFeatureSurface = {
        id: definition.id,
        tab: placement.tab,
        group: placement.group,
        label: definition.labelKey,
        surfaceCommandId: definition.id,
        ...(definition.commandId ? { requiredCommandId: definition.commandId } : {}),
        icon: definition.icon ?? 'command',
        order: definition.priority,
        appearance: definition.display === 'large' ? 'large' : 'small',
        breakpoints: ALL_RIBBON_BREAKPOINTS,
      };
      const target = PRIMARY_RIBBON_TABS.has(placement.tab) ? owner.ribbon! : owner.contextualTabs!;
      if (!target.some((candidate) => candidate.id === entry.id && candidate.tab === entry.tab && candidate.group === entry.group)) target.push(entry);
    }
  }
  return next;
}

/** Compile Home/Insert/contextual surfaces from the same feature ownership map. */
export function compileFeatureSurfaceSchema(manifests: readonly SpreadsheetFeatureManifest[]): CompiledFeatureSurfaceSchema {
  const sourceManifests = manifests;
  const commands = new Map<string, string>();
  for (const manifest of sourceManifests) for (const commandId of manifest.commandIds) {
    const owner = commands.get(commandId);
    if (owner && owner !== manifest.id) throw new Error(`FEATURE_COMMAND_OWNER_AMBIGUOUS: ${commandId} belongs to ${owner} and ${manifest.id}`);
    commands.set(commandId, manifest.id);
  }
  const compile = (source: SpreadsheetFeatureManifest['ribbon'] | SpreadsheetFeatureManifest['contextualTabs']): CompiledFeatureSurfaceEntry[] => (source ?? []).map((entry) => {
    const featureId = entry.requiredCommandId ? commands.get(entry.requiredCommandId) : undefined;
    if (!featureId && entry.requiredCommandId) throw new Error(`FEATURE_SURFACE_UNAVAILABLE: ${entry.requiredCommandId} has no registered feature command`);
    const commandId = entry.surfaceCommandId;
    return {
      featureId: featureId ?? sourceManifests.find((manifest) => manifest.id === 'sheet-features')?.id ?? 'sheet-features',
      id: entry.id,
      tab: entry.tab as RibbonSurfaceDefinition['tab'],
      group: entry.group as RibbonSurfaceDefinition['group'],
      label: entry.label,
      ...(commandId ? { commandId } : {}),
      ...(entry.controlId ? { controlId: entry.controlId } : {}),
      order: entry.order ?? 0,
      appearance: entry.appearance ?? (entry.controlId ? 'state-control' : 'small'),
      breakpoints: entry.breakpoints ?? ALL_RIBBON_BREAKPOINTS,
      ...(entry.overflowTarget ? { overflowTarget: entry.overflowTarget } : {}),
      ...(entry.menuId ? { menuId: entry.menuId } : {}),
      ...(entry.ariaLabel ? { ariaLabel: entry.ariaLabel } : {}),
      ...(entry.requiredCommandId ? { requiredCommandId: entry.requiredCommandId } : {}),
      icon: entry.icon,
    };
  });
  const ribbon = sourceManifests.flatMap((manifest) => compile(manifest.ribbon));
  const contextualTabs = sourceManifests.flatMap((manifest) => compile(manifest.contextualTabs));
  return { ribbon, contextualTabs };
}

/**
 * Register every spreadsheet feature against one CommandRuntime.
 *
 * Chart, Pivot and Sparkline own their command and mutation registrations in
 * spreadsheet-app. There is deliberately no compatibility namespace or
 * command forwarding layer: callers must use the canonical feature command.
 */
export function registerSpreadsheetFeatures(
  runtime: CommandRuntime,
  drawingRuntime: DrawingRuntime,
  featureRuntime?: SpreadsheetFeatureRuntime,
): SpreadsheetFeatureManifest[] {
  registerSheetCommands(runtime);
  registerEditingFeatures(runtime);

  const drawingManifest = registerDrawingFeature(runtime, drawingRuntime);
  const pivotControlManifest = registerPivotControlFeature(runtime);
  const dataSourceManifest = registerDataSourceFeature(runtime);
  const chartManifest = registerChartFeature(runtime);
  const pivotManifest = registerPivotFeature(runtime);
  const sparklineManifest = registerSparklineFeature(runtime);
  const reviewManifest = registerReviewFeature(runtime);
  const findReplaceManifest = registerFindReplaceFeature(runtime);
  const insertCommandIds = registerInsertCommands(runtime);
  const platformCommandIds = registerPlatformFeatures(runtime);
  const parityReport = buildExcelParityReport(createExcelFeatureRegistry().features);
  const coreManifest: SpreadsheetFeatureManifest = {
    id: 'sheet-features',
    version: '1.0.0',
    commandIds: [],
    documentTypes: ['spreadsheet'],
  };

  const featureOwnedCommandIds = new Set([
    ...drawingManifest.commandIds,
    ...pivotControlManifest.commandIds,
    ...dataSourceManifest.commandIds,
    ...chartManifest.commandIds,
    ...pivotManifest.commandIds,
    ...sparklineManifest.commandIds,
    ...reviewManifest.commandIds,
    ...findReplaceManifest.commandIds,
    ...insertCommandIds,
    ...platformCommandIds,
  ]);
  coreManifest.commandIds = runtime.registry.listCommandIds().filter((commandId) => !featureOwnedCommandIds.has(commandId));
  const manifests: SpreadsheetFeatureManifest[] = [
    coreManifest,
    withCoreDependency(drawingManifest),
    withCoreDependency(pivotControlManifest),
    withCoreDependency(dataSourceManifest),
    withCoreDependency(chartManifest),
    withCoreDependency(pivotManifest),
    withCoreDependency(sparklineManifest),
    withCoreDependency(reviewManifest),
    withCoreDependency(findReplaceManifest),
    withCoreDependency({ id: 'insert', version: '1.0.0', commandIds: insertCommandIds, permissions: ['sheet.structure.write', 'sheet.format.write'] }),
    {
      id: 'platform',
      version: '1.0.0',
      dependencies: ['sheet-features'],
      commandIds: platformCommandIds,
      permissions: ['history.restore', 'persistence.write', 'print.export', 'document.exchange', 'query.execute'],
    },
    withCoreDependency({ id: 'excel-parity', version: '1.0.0', commandIds: [], parity: parityReport }),
  ];
  const surfacedManifests = attachCanonicalRibbonSurfaces(manifests);
  featureRuntime?.load(surfacedManifests);
  featureRuntime?.setParityReport(parityReport);
  return surfacedManifests;
}

/** Activate the already-compiled registry once the workbook instance exists. */
export function activateSpreadsheetFeatures(
  featureRuntime: SpreadsheetFeatureRuntime,
  options: FeatureRuntimeActivation = {},
): readonly SpreadsheetFeatureManifest[] {
  return featureRuntime.activate(options);
}

export function advanceSpreadsheetFeatures(
  featureRuntime: SpreadsheetFeatureRuntime,
  phase: Extract<FeatureLifecyclePhase, 'ready' | 'rendered' | 'steady'>,
): void {
  featureRuntime.advance(phase);
}

export function createSpreadsheetFeatureRuntime(): SpreadsheetFeatureRuntime {
  return new SpreadsheetFeatureRuntime();
}

export function getFeatureRegistry(featureRuntime: SpreadsheetFeatureRuntime): SpreadsheetFeatureManifest[] {
  return featureRuntime.getManifests().map((manifest) => ({
    ...manifest,
    commandIds: [...manifest.commandIds],
    ...(manifest.dependencies ? { dependencies: [...manifest.dependencies] } : {}),
    ...(manifest.documentTypes ? { documentTypes: [...manifest.documentTypes] } : {}),
    ...(manifest.environments ? { environments: [...manifest.environments] } : {}),
    ...(manifest.ribbon ? { ribbon: manifest.ribbon.map((entry) => ({ ...entry })) } : {}),
    ...(manifest.contextualTabs ? { contextualTabs: manifest.contextualTabs.map((entry) => ({ ...entry })) } : {}),
  }));
}

export function getExcelParityReport(featureRuntime: SpreadsheetFeatureRuntime): ExcelParityReport {
  return featureRuntime.getParityReport();
}
