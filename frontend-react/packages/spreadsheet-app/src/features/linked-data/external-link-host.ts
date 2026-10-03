import { synchronizeRecordCalculations } from './record-calculation';
import { WorkbookModel, type ExternalLinkBinding } from '@react-sheets/core-model';
import { FormulaEngine, type ExternalCalculationLink } from '@react-sheets/formula-engine';
import { validateExternalCalculationGraph, type ExternalCalculationGraph, type ExternalCalculationNode } from '@react-sheets/protocol';
import { configureWorkbookSpillEnvironments, syncWorkbookSheetTables } from '../../formula-spill-sync';
import type { SpreadsheetRuntime } from '../../runtime';

interface Refresh { model: WorkbookModel; formula: FormulaEngine; revision: number; bindings: string; dirty: boolean; promise: Promise<void> }
const pending = new WeakMap<SpreadsheetRuntime, Refresh>();
function bindings(model: WorkbookModel): string { return JSON.stringify([...model.dataModel.externalLinks.values()]); }

/** Server events invalidate inputs before another graph can be consumed. */
export function requestExternalRecalculation(runtime: SpreadsheetRuntime): Promise<void> {
  if (runtime.disposed) return Promise.resolve();
  const current = runtime.formula.getExternalCalculationLinks();
  if (current.length) runtime.formula.applyExternalCalculationLinks(current.map(link => ({ ...link, state: 'refreshing', cells: [] })));
  const active = pending.get(runtime);
  if (active?.model === runtime.model && active.formula === runtime.formula) active.dirty = true;
  return refreshExternalLinks(runtime);
}

/** One version-pinned authorized DAG, evaluated with the canonical AST/Worker for every source. */
export function refreshExternalLinks(runtime: SpreadsheetRuntime): Promise<void> {
  const active = pending.get(runtime), model = runtime.model, formula = runtime.formula;
  if (active?.model === model && active.formula === formula) {
    if (active.bindings !== bindings(model) || active.revision !== runtime.remoteRevision) active.dirty = true;
    return active.promise;
  }
  const entry: Refresh = { model, formula, revision: runtime.remoteRevision, bindings: bindings(model), dirty: false, promise: Promise.resolve() };
  const refresh = (async () => {
    do {
      entry.dirty = false; entry.revision = runtime.remoteRevision; entry.bindings = bindings(model);
      const declared = [...model.dataModel.externalLinks.values()];
      let updates: ExternalCalculationLink[];
      try {
        if (!declared.length) updates = [];
        else {
          const graph = validateExternalCalculationGraph(await runtime.api.getExternalCalculationGraph(model.unitId), model.unitId);
          const root = graph.nodes.find(node => node.unitId === model.unitId);
          if (root?.state !== 'connected' || JSON.stringify(root.snapshot.dataModel.externalLinks) !== entry.bindings) throw Object.assign(new Error('External binding definitions changed; reload the target'), { code: 'EXTERNAL_LINK_BINDING_CHANGED' });
          updates = await calculateExternalGraph(graph, declared);
          runtime.collab?.send({ type: 'calculation.subscribe', unitId: model.unitId });
        }
      } catch (cause) {
        const error = cause as { status?: number; code?: string; message?: string };
        const state = error.status === 401 || error.status === 403 ? 'denied' : error.status === 404 || error.code === 'CIRCULAR_DEPENDENCY' ? 'broken' : 'unavailable';
        updates = declared.map(binding => {
          const previous = formula.getExternalCalculationLinks().find(link => link.id === binding.id);
          return ({ id: binding.id, token: binding.token, sourceUnitId: binding.sourceUnitId,
          subject: previous?.subject ?? 'unresolved', sourceRevision: previous?.sourceRevision ?? 0, accessRevision: previous?.accessRevision ?? 0, state, sheets: [], cells: [],
          error: { code: error.code ?? 'EXTERNAL_CALCULATION_GRAPH_FAILED', message: error.message ?? String(cause) } }); });
        if (error.status === 401 || error.status === 403) runtime.handlers.onPhaseChange?.('error');
      }
      if (runtime.disposed || runtime.model !== model || runtime.formula !== formula) return;
      if (entry.bindings !== bindings(model) || entry.revision !== runtime.remoteRevision) entry.dirty = true;
      if (entry.dirty) continue;
      formula.applyExternalCalculationLinks(updates);
      const report = await formula.recalculateAsync();
      if (runtime.disposed || runtime.model !== model || runtime.formula !== formula || entry.dirty) continue;
      runtime.handlers.onCalculationApplied?.(report.changedAddresses ?? report.recalculated);
      runtime.handlers.onMutationsApplied?.();
    } while (entry.dirty && !runtime.disposed && runtime.model === model && runtime.formula === formula);
  })().finally(() => { if (pending.get(runtime) === entry) pending.delete(runtime); });
  entry.promise = refresh; pending.set(runtime, entry);
  return refresh;
}

export async function calculateExternalGraph(graph: ExternalCalculationGraph, declared: readonly ExternalLinkBinding[]): Promise<ExternalCalculationLink[]> {
  graph = validateExternalCalculationGraph(graph, graph.rootUnitId);
  const nodes = new Map(graph.nodes.map(node => [node.unitId, node]));
  const engines = new Map<string, { source: WorkbookModel; engine: FormulaEngine }>();
  const evaluate = async (unitId: string): Promise<void> => {
    if (engines.has(unitId)) return;
    const node = nodes.get(unitId);
    if (node?.state !== 'connected') return;
    for (const binding of node.snapshot.dataModel.externalLinks) await evaluate(binding.sourceUnitId);
    const source = WorkbookModel.fromSnapshot(node.snapshot);
    const engine = new FormulaEngine({ defaultSheetId: source.primarySheetId,
      sheetOrder: source.getSheets().map(({ id, name }) => ({ id, name })), blockedRanges: node.blockedRanges,
      collationContext: source.collationContext, calculationSettings: source.calculationSettings });
    engines.set(unitId, { source, engine });
    configureWorkbookSpillEnvironments(engine, source);
    syncWorkbookSheetTables(engine, source, false);
    engine.setDefinedNameModels(source.definedNameModels, false);
    for (const sheet of source.getSheets()) sheet.cells.forEach((cell, row, column) => {
      engine.synchronizeInputs([{ address: { sheetId: sheet.id, row, column }, input: cell.formula ? { kind: 'formula', formula: cell.formula } : { kind: 'value', value: cell.value ?? null } }]);
    });
    synchronizeRecordCalculations(engine, source);
    engine.applyExternalCalculationLinks([...source.dataModel.externalLinks.values()].map(link => materialize(link)));
    await engine.recalculateAsync(undefined, undefined, true);
  };
  const materialize = (binding: ExternalLinkBinding): ExternalCalculationLink => {
    const node: ExternalCalculationNode | undefined = nodes.get(binding.sourceUnitId);
    if (!node) throw new Error('EXTERNAL_CALCULATION_GRAPH_NODE_MISSING');
    if (node.state !== 'connected') return { id: binding.id, token: binding.token, sourceUnitId: binding.sourceUnitId,
      subject: graph.subject, sourceRevision: 0, accessRevision: 0, state: node.state, sheets: [], cells: [], error: node.error };
    const evaluated = engines.get(binding.sourceUnitId);
    if (!evaluated) throw new Error('EXTERNAL_CALCULATION_GRAPH_SOURCE_NOT_EVALUATED');
    const { source, engine } = evaluated;
    const cells: ExternalCalculationLink['cells'][number][] = [];
    const recordResults = engine.getRecordFormulaOwners(), cached = new Set<string>();
    const sheets = binding.sheets.flatMap(boundSheet => {
      const sheet = source.sheets.get(boundSheet.sheetId);
      if (!sheet) return [];
      sheet.cells.forEach((cell, row, column) => {
        const address = { sheetId: sheet.id, row, column };
        const spill = engine.getSpillValueAt(sheet.id, row, column), result = engine.getCellResult(address);
        if (cell.formula && spill === undefined && result === undefined) throw new Error('EXTERNAL_CALCULATION_RESULT_MISSING');
        cells.push({ address, value: spill !== undefined ? spill : cell.formula ? result!.value : cell.value ?? null }); cached.add(JSON.stringify(address));
      });
      for (const owner of recordResults) if (owner.address.sheetId === sheet.id && !cached.has(JSON.stringify(owner.address))) {
        const result = engine.getCellResult(owner.address);
        if (!result) throw new Error('RECORD_CALCULATION_RESULT_MISSING');
        cells.push({ address: owner.address, value: result.value }); cached.add(JSON.stringify(owner.address));
      }
      for (const spill of engine.getSpillsForSheet(sheet.id)) if (spill.state === 'ok') {
        for (let row = spill.range.startRow; row <= spill.range.endRow; row++) for (let column = spill.range.startColumn; column <= spill.range.endColumn; column++) {
          if (row === spill.anchor.row && column === spill.anchor.column || cached.has(JSON.stringify({ sheetId: sheet.id, row, column }))) continue;
          const value = engine.getSpillValueAt(sheet.id, row, column);
          if (value !== undefined && !Array.isArray(value)) cells.push({ address: { sheetId: sheet.id, row, column }, value });
        }
      }
      return [{ id: sheet.id, name: boundSheet.token, rowCount: sheet.rowCount, columnCount: sheet.columnCount }];
    });
    return { id: binding.id, token: binding.token, sourceUnitId: binding.sourceUnitId, subject: graph.subject,
      sourceRevision: node.revision, accessRevision: node.accessRevision,
      state: sheets.length === binding.sheets.length ? 'connected' : 'broken', sheets,
      cells: sheets.length === binding.sheets.length ? cells : [], blockedRanges: node.blockedRanges };
  };
  try {
    for (const binding of declared) await evaluate(binding.sourceUnitId);
    return declared.map(binding => materialize(binding));
  } finally { for (const { engine } of engines.values()) engine.disposeCalculationTasks(); }
}
