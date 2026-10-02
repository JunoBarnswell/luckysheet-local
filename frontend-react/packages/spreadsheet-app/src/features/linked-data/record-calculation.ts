import { assertRecordTable, assertRecordRelationship, recordRows, type WorkbookModel, type WorkbookTableModel } from '@react-sheets/core-model';
import { formatFormula, parseFormula, type FormulaAst, type FormulaEngine, type CellAddress } from '@react-sheets/formula-engine';

interface RelationIndex {
  readonly forward: ReadonlyMap<string, readonly string[]>;
  readonly reverse: ReadonlyMap<string, readonly string[]>;
}

/** Derived relation membership indexes; they hold identities, never a copy of record values. */
export function buildRecordRelationIndexes(workbook: WorkbookModel): ReadonlyMap<string, RelationIndex> {
  const indexes = new Map<string, RelationIndex>();
  for (const relation of workbook.dataModel.relationships.values()) {
    if (!workbook.getTable(relation.fromTableId).recordIdFieldId) continue;
    assertRecordRelationship(workbook, relation);
    const from = workbook.getTable(relation.fromTableId), field = from.fields.find(field => field.id === relation.fromFieldId)!;
    const forward = new Map<string, string[]>(), reverse = new Map<string, string[]>();
    for (const [id, row] of recordRows(workbook, from)) {
      const target = workbook.getSheet(from.sourceRange!.sheetId).cells.getWithoutHydration(row, from.sourceRange!.startColumn + field.ordinal)?.value;
      const targets = typeof target === 'string' && target ? [target] : [];
      forward.set(id, targets);
      for (const target of targets) { const owners = reverse.get(target) ?? []; owners.push(id); reverse.set(target, owners); }
    }
    indexes.set(relation.id, { forward, reverse });
  }
  return indexes;
}

const span = { start: 0, end: 0 };
function reference(address: CellAddress): FormulaAst {
  return { type: 'cell-reference', reference: { ...address, absoluteRow: true, absoluteColumn: true }, span };
}

/** Field definitions project into the existing dependency graph and Worker input journal. */
export function synchronizeRecordCalculations(engine: FormulaEngine, workbook: WorkbookModel): readonly CellAddress[] {
  const tables = [...workbook.dataModel.tables.values()].filter(table => table.recordIdFieldId);
  const rows = new Map<string, ReadonlyMap<string, number>>();
  for (const table of tables) { assertRecordTable(workbook, table); rows.set(table.id, recordRows(workbook, table)); }
  const relations = buildRecordRelationIndexes(workbook);
  const owners: { tableId: string; recordId: string; fieldId: string; address: CellAddress }[] = [];
  const inputs: { address: CellAddress; input: { kind: 'formula'; formula: string } }[] = [];
  for (const table of tables) {
    const range = table.sourceRange!;
    for (const [recordId, row] of rows.get(table.id)!) for (const field of table.fields) {
      if (!field.calculation) continue;
      const address = { sheetId: range.sheetId, row, column: range.startColumn + field.ordinal };
      const calculation = field.calculation;
      let ast: FormulaAst;
      if (calculation.kind === 'formula') ast = parseFormula(calculation.formula);
      else {
        const relation = workbook.dataModel.relationships.get(calculation.relationshipId);
        if (!relation) throw new Error(`RECORD_RELATION_MISSING: ${calculation.relationshipId}`);
        const forward = calculation.direction === 'forward';
        if (table.id !== (forward ? relation.fromTableId : relation.toTableId)) throw new Error('RECORD_RELATION_OWNER_INVALID');
        const target = workbook.getTable(forward ? relation.toTableId : relation.fromTableId);
        const targetField = target.fields.find(field => field.id === calculation.targetFieldId);
        if (!targetField || !target.sourceRange) throw new Error('RECORD_RELATION_FIELD_MISSING');
        const members = (forward ? relations.get(relation.id)!.forward : relations.get(relation.id)!.reverse).get(recordId) ?? [];
        const argumentsList = members.map(id => {
          const targetRow = rows.get(target.id)?.get(id);
          if (targetRow === undefined) throw new Error('RECORD_RELATION_TARGET_MISSING');
          return reference({ sheetId: target.sourceRange!.sheetId, row: targetRow, column: target.sourceRange!.startColumn + targetField.ordinal });
        });
        ast = calculation.kind === 'rollup' ? { type: 'function-call', name: calculation.aggregate, arguments: argumentsList, span }
          : argumentsList.length ? { type: 'function-call', name: 'VSTACK', arguments: argumentsList, span } : { type: 'string-literal', value: '', span };
      }
      owners.push({ tableId: table.id, recordId, fieldId: field.id, address });
      const formula = formatFormula(ast);
      if (engine.getCellResult(address)?.formula !== formula) inputs.push({ address, input: { kind: 'formula', formula } });
    }
  }
  const nextKeys = new Set(owners.map(owner => JSON.stringify(owner.address)));
  const cleared = engine.getRecordFormulaOwners().filter(owner => !nextKeys.has(JSON.stringify(owner.address))).map(owner => ({ address: owner.address, input: null }));
  engine.setRecordFormulaOwners(owners);
  return engine.synchronizeInputs([...cleared, ...inputs]);
}

export function recordSheetTables(workbook: WorkbookModel): import('@react-sheets/formula-engine').SheetTableRef[] {
  return [...workbook.dataModel.tables.values()].filter(table => table.recordIdFieldId && table.sourceRange).map(table => ({
    id: table.id, name: table.name, sheetId: table.sourceRange!.sheetId, range: table.sourceRange!, recordIdFieldId: table.recordIdFieldId,
    hasHeaderRow: true, hasTotalRow: false, columns: [...table.fields].sort((a, b) => a.ordinal - b.ordinal).map(field => ({ id: field.id, name: field.name })),
  }));
}
