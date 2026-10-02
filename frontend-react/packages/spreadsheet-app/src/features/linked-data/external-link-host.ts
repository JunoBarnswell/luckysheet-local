import { synchronizeRecordCalculations } from './record-calculation';
import { WorkbookModel } from '@react-sheets/core-model';
import { FormulaEngine, type ExternalCalculationLink } from '@react-sheets/formula-engine';
import { configureWorkbookSpillEnvironments, syncWorkbookSheetTables } from '../../formula-spill-sync';
import type { SpreadsheetRuntime } from '../../runtime';

const pending = new WeakMap<SpreadsheetRuntime, { model: WorkbookModel; formula: FormulaEngine; promise: Promise<void> }>();

/** Coalesced host refresh. Source and destination evaluators share the same AST and Worker. */
export function refreshExternalLinks(runtime: SpreadsheetRuntime): Promise<void> {
  const active = pending.get(runtime);
  const model = runtime.model, formula = runtime.formula;
  if (active?.model === model && active.formula === formula) return active.promise;
  const refresh = (async () => {
    const updates = await Promise.all([...model.dataModel.externalLinks.values()].map(async binding => {
      const previous = formula.getExternalCalculationLinks().find(link => link.id === binding.id);
      try {
        const input = await runtime.api.getExternalLinkInputs(model.unitId, binding.id);
        if (JSON.stringify(input.binding) !== JSON.stringify(binding)) throw new Error('EXTERNAL_LINK_BINDING_CHANGED');
        const source = WorkbookModel.fromSnapshot(input.snapshot);
        const engine = new FormulaEngine({ defaultSheetId: source.primarySheetId, sheetOrder: source.getSheets().map(({ id, name }) => ({ id, name })), blockedRanges: input.blockedRanges, collationContext: source.collationContext, calculationSettings: source.calculationSettings });
        configureWorkbookSpillEnvironments(engine, source);
        syncWorkbookSheetTables(engine, source, false);
        engine.setDefinedNameModels(source.definedNameModels, false);
        for (const sheet of source.getSheets()) sheet.cells.forEach((cell, row, column) => {
          engine.synchronizeInputs([{ address: { sheetId: sheet.id, row, column }, input: cell.formula ? { kind: 'formula', formula: cell.formula } : { kind: 'value', value: cell.value ?? null } }]);
        });
        synchronizeRecordCalculations(engine, source);
        try {
          await engine.recalculateAsync(undefined, undefined, true);
          const cells: ExternalCalculationLink['cells'][number][] = [];
          const recordResults = engine.getRecordFormulaOwners();
          const cachedAddresses = new Set<string>();
          const sheets = binding.sheets.flatMap(boundSheet => {
            const sheet = source.sheets.get(boundSheet.sheetId);
            if (!sheet) return [];
            sheet.cells.forEach((cell, row, column) => {
              const value = engine.getCellResult({ sheetId: sheet.id, row, column })?.value ?? cell.value ?? null;
              cells.push({ address: { sheetId: sheet.id, row, column }, value });
              cachedAddresses.add(JSON.stringify({ sheetId: sheet.id, row, column }));
            });
            for (const owner of recordResults) if (owner.address.sheetId === sheet.id && !cachedAddresses.has(JSON.stringify(owner.address))) {
              const result = engine.getCellResult(owner.address);
              if (!result) throw new Error('RECORD_CALCULATION_RESULT_MISSING');
              cells.push({ address: owner.address, value: result.value });
            }
            for (const spill of engine.getSpillsForSheet(sheet.id)) if (spill.state === 'ok') {
              for (let row = spill.range.startRow; row <= spill.range.endRow; row++) for (let column = spill.range.startColumn; column <= spill.range.endColumn; column++) {
                if (row === spill.anchor.row && column === spill.anchor.column) continue;
                const value = engine.getSpillValueAt(sheet.id, row, column);
                if (value !== undefined && !Array.isArray(value)) cells.push({ address: { sheetId: sheet.id, row, column }, value });
              }
            }
            return [{ id: sheet.id, name: boundSheet.token, rowCount: sheet.rowCount, columnCount: sheet.columnCount }];
          });
          return { id: binding.id, token: binding.token, sourceUnitId: binding.sourceUnitId, subject: input.subject, sourceRevision: input.sourceRevision, accessRevision: input.accessRevision, state: sheets.length === binding.sheets.length ? 'connected' : 'broken', sheets, cells: sheets.length === binding.sheets.length ? cells : [], blockedRanges: input.blockedRanges } satisfies ExternalCalculationLink;
        } finally { engine.disposeCalculationTasks(); }
      } catch (error) {
        const status = (error as { status?: number }).status;
        const state = status === 401 || status === 403 ? 'denied' : status === 404 ? 'broken' : 'unavailable';
        return { id: binding.id, token: binding.token, sourceUnitId: binding.sourceUnitId, subject: previous?.subject ?? 'unresolved', sourceRevision: previous?.sourceRevision ?? 0, accessRevision: previous?.accessRevision ?? 0, state, sheets: [], cells: [], error: { code: (error as { code?: string }).code ?? 'EXTERNAL_LINK_REFRESH_FAILED', message: error instanceof Error ? error.message : String(error) } } satisfies ExternalCalculationLink;
      }
    }));
    if (runtime.disposed || runtime.model !== model || runtime.formula !== formula) return;
    formula.applyExternalCalculationLinks(updates);
    const report = await formula.recalculateAsync();
    if (runtime.disposed || runtime.model !== model || runtime.formula !== formula) return;
    runtime.handlers.onCalculationApplied?.(report.changedAddresses ?? report.recalculated);
    runtime.handlers.onMutationsApplied?.();
  })().finally(() => { if (pending.get(runtime)?.promise === refresh) pending.delete(runtime); });
  pending.set(runtime, { model, formula, promise: refresh });
  return refresh;
}
