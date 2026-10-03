import { createCellSetMutationParams } from '@react-sheets/sheet-features';
import { canonicalRecordFieldFormula } from '@react-sheets/core-model';
import { assertRecordFieldWrite, assertRecordTable, assertRecordRelationship, resolveRecordField } from '@react-sheets/core-model';
import { CALCULATION_CONTEXT_EFFECTS, type ExternalLinkBinding } from '@react-sheets/core-model';
import { assertExternalLinkBinding, type WorkbookModel } from '@react-sheets/core-model';
import type { CommandRuntime } from '@react-sheets/command-runtime';

export function registerLinkedDataCommands(runtime: CommandRuntime): void {
  const validSet = (value: unknown): value is { link: ExternalLinkBinding } => {
    try { if (!value || typeof value !== 'object' || Object.keys(value).length !== 1) return false; assertExternalLinkBinding((value as { link: ExternalLinkBinding }).link); return true; } catch { return false; }
  };
  const validRemove = (value: unknown): value is { linkId: string } => Boolean(value && typeof value === 'object' && Object.keys(value).length === 1 && typeof (value as { linkId?: unknown }).linkId === 'string');
  const set = (workbook: WorkbookModel, link: ExternalLinkBinding): void => {
    assertExternalLinkBinding(link);
    if (link.sourceUnitId === workbook.unitId) throw new Error('EXTERNAL_LINK_SOURCE_INVALID');
    for (const other of workbook.dataModel.externalLinks.values()) if (other.id !== link.id && other.token.toUpperCase() === link.token.toUpperCase()) throw new Error('EXTERNAL_LINK_TOKEN_CONFLICT');
    workbook.dataModel.externalLinks.set(link.id, structuredClone(link));
  };
  runtime.registry.registerMutation<{ link: ExternalLinkBinding }>({ id: 'externalLink.set', handler: (item, context) => { if (!validSet(item.params)) throw new Error('Invalid external link'); set(context.workbook, item.params.link); }, metadata: { calculation: { inputs: 'none' as const, visibility: false, spillBlockers: 'none' as const, mode: false, context: CALCULATION_CONTEXT_EFFECTS.rebuild },
    schema: { name: 'ExternalLinkSet', validate: validSet }, permission: { capability: 'workbook.external-link.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: () => [], mode: 'exact' },  inverseIds: ['externalLink.set', 'externalLink.remove'],
  } });
  runtime.registry.registerMutation<{ linkId: string }>({ id: 'externalLink.remove', handler: (item, context) => { if (!validRemove(item.params) || !context.workbook.dataModel.externalLinks.delete(item.params.linkId)) throw new Error('External link not found'); }, metadata: { calculation: { inputs: 'none' as const, visibility: false, spillBlockers: 'none' as const, mode: false, context: CALCULATION_CONTEXT_EFFECTS.rebuild },
    schema: { name: 'ExternalLinkRemove', validate: validRemove }, permission: { capability: 'workbook.external-link.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: () => [], mode: 'exact' },  inverseIds: ['externalLink.set'],
  } });
  runtime.registry.registerCommand<{ link: ExternalLinkBinding }>({ id: 'externalLink.set', execute: (params, context) => {
    if (!validSet(params)) throw new Error('Invalid external link');
    const previous = context.workbook.dataModel.externalLinks.get(params.link.id);
    const base = { unitId: context.workbook.unitId, sheetId: context.workbook.primarySheetId, affectedRanges: [] };
    context.applyMutation({ id: 'externalLink.set', ...base, params, inverse: [previous ? { id: 'externalLink.set', ...base, params: { link: structuredClone(previous) } } : { id: 'externalLink.remove', ...base, params: { linkId: params.link.id } }], apply: () => set(context.workbook, params.link) });
    return { operationId: context.operationId, mutationCount: 1, affectedRanges: [] };
  } });
  runtime.registry.registerCommand<{ linkId: string }>({ id: 'externalLink.remove', execute: (params, context) => {
    const previous = context.workbook.dataModel.externalLinks.get(params.linkId);
    if (!previous) throw new Error('External link not found');
    const base = { unitId: context.workbook.unitId, sheetId: context.workbook.primarySheetId, affectedRanges: [] };
    context.applyMutation({ id: 'externalLink.remove', ...base, params, inverse: [{ id: 'externalLink.set', ...base, params: { link: structuredClone(previous) } }], apply: () => { context.workbook.dataModel.externalLinks.delete(params.linkId); } });
    return { operationId: context.operationId, mutationCount: 1, affectedRanges: [] };
  } });
}

export function registerRecordCommands(runtime: CommandRuntime): void {
  const object = (value: unknown): value is Record<string, unknown> => Boolean(value && typeof value === 'object' && !Array.isArray(value));
  const validTable = (value: unknown): value is { table: import('@react-sheets/core-model').WorkbookTableModel } => object(value) && object(value.table) && typeof value.table.id === 'string' && Array.isArray(value.table.fields);
  const validRelation = (value: unknown): value is { relationship: import('@react-sheets/core-model').DataRelationship } => object(value) && object(value.relationship) && typeof value.relationship.id === 'string';
  const validRemove = (value: unknown): value is { relationshipId: string } => object(value) && typeof value.relationshipId === 'string';
  const validSet = (value: unknown): value is import('@react-sheets/core-model').RecordFieldAddress & { value: import('@react-sheets/core-model').TableScalar } => object(value) && ['tableId', 'recordId', 'fieldId'].every(key => typeof value[key] === 'string') && ('value' in value) && (value.value === null || ['string', 'number', 'boolean'].includes(typeof value.value));
  const rangesForRelation = (workbook: WorkbookModel, relation: import('@react-sheets/core-model').DataRelationship) => [workbook.getTable(relation.fromTableId).sourceRange!, workbook.getTable(relation.toTableId).sourceRange!];
  const setTable = (workbook: WorkbookModel, table: import('@react-sheets/core-model').WorkbookTableModel) => {
    const previous = workbook.getTable(table.id);
    if (JSON.stringify(previous.sourceRange) !== JSON.stringify(table.sourceRange) || previous.sourceId !== table.sourceId) throw new Error('RECORD_TABLE_SOURCE_CHANGED');
    if (previous.sourceSheetId !== table.sourceSheetId || previous.rowCount !== table.rowCount || previous.fields.length !== table.fields.length || previous.fields.some(field => table.fields.find(next => next.id === field.id)?.ordinal !== field.ordinal)) throw new Error('RECORD_FIELD_IDENTITY_IMMUTABLE');
    assertRecordTable(workbook, table);
    workbook.dataModel.tables.set(table.id, structuredClone(table));
  };
  runtime.registry.registerMutation<{ table: import('@react-sheets/core-model').WorkbookTableModel }>({ id: 'table.configure', handler: (item, context) => { if (!validTable(item.params)) throw new Error('Invalid record table'); setTable(context.workbook, item.params.table); }, metadata: { calculation: { inputs: 'none' as const, visibility: false, spillBlockers: 'none' as const, mode: false, context: CALCULATION_CONTEXT_EFFECTS.rebuild },
    schema: { name: 'RecordTableConfigure', validate: validTable }, permission: { capability: 'workbook.table.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: params => params.table.sourceRange ? [params.table.sourceRange] : [], mode: 'declared' },  inverseIds: ['table.configure'],
  } });
  runtime.registry.registerCommand<{ table: import('@react-sheets/core-model').WorkbookTableModel }>({ id: 'table.configure', execute: (params, context) => {
    if (!validTable(params)) throw new Error('Invalid record table');
    const currentIdentity = context.workbook.getTable(params.table.id).recordIdFieldId;
    if (currentIdentity && params.table.recordIdFieldId !== currentIdentity) throw new Error('RECORD_IDENTITY_IMMUTABLE');
    const previous = structuredClone(context.workbook.getTable(params.table.id)), ranges = params.table.sourceRange ? [params.table.sourceRange] : [];
    const base = { unitId: context.workbook.unitId, sheetId: params.table.sourceRange!.sheetId, affectedRanges: ranges };
    context.applyMutation({ id: 'table.configure', ...base, params, inverse: [{ id: 'table.configure', ...base, params: { table: previous } }], apply: () => setTable(context.workbook, params.table) });
    return { operationId: context.operationId, mutationCount: 1, affectedRanges: ranges };
  } });
  runtime.registry.registerMutation<{ relationship: import('@react-sheets/core-model').DataRelationship }>({ id: 'relationship.set', handler: (item, context) => { if (!validRelation(item.params)) throw new Error('Invalid record relation'); assertRecordRelationship(context.workbook, item.params.relationship); context.workbook.dataModel.relationships.set(item.params.relationship.id, structuredClone(item.params.relationship)); }, metadata: { calculation: { inputs: 'none' as const, visibility: false, spillBlockers: 'none' as const, mode: false, context: CALCULATION_CONTEXT_EFFECTS.rebuild },
    schema: { name: 'RecordRelationshipSet', validate: validRelation }, permission: { capability: 'workbook.table.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: () => [], mode: 'declared' },  inverseIds: ['relationship.set', 'relationship.remove'],
  } });
  runtime.registry.registerMutation<{ relationshipId: string }>({ id: 'relationship.remove', handler: (item, context) => {
    if (!validRemove(item.params)) throw new Error('Invalid record relation');
    for (const table of context.workbook.dataModel.tables.values()) for (const field of table.fields) if (field.calculation && field.calculation.kind !== 'formula' && field.calculation.relationshipId === item.params.relationshipId) throw new Error('RECORD_RELATION_IN_USE');
    if (!context.workbook.dataModel.relationships.delete(item.params.relationshipId)) throw new Error('Record relation not found');
  }, metadata: { calculation: { inputs: 'none' as const, visibility: false, spillBlockers: 'none' as const, mode: false, context: CALCULATION_CONTEXT_EFFECTS.rebuild },
    schema: { name: 'RecordRelationshipRemove', validate: validRemove }, permission: { capability: 'workbook.table.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: () => [], mode: 'declared' },  inverseIds: ['relationship.set'],
  } });
  runtime.registry.registerCommand<{ relationship: import('@react-sheets/core-model').DataRelationship }>({ id: 'relationship.set', execute: (params, context) => {
    assertRecordRelationship(context.workbook, params.relationship);
    const previous = context.workbook.dataModel.relationships.get(params.relationship.id), ranges = rangesForRelation(context.workbook, params.relationship);
    const base = { unitId: context.workbook.unitId, sheetId: ranges[0]!.sheetId, affectedRanges: ranges };
    context.applyMutation({ id: 'relationship.set', ...base, params, inverse: [previous ? { id: 'relationship.set', ...base, params: { relationship: structuredClone(previous) } } : { id: 'relationship.remove', ...base, params: { relationshipId: params.relationship.id } }], apply: () => { context.workbook.dataModel.relationships.set(params.relationship.id, structuredClone(params.relationship)); } });
    return { operationId: context.operationId, mutationCount: 1, affectedRanges: ranges };
  } });
  type RecordRestore = import('@react-sheets/core-model').RecordFieldAddress & { previous: import('@react-sheets/core-model').CellData | null };
  const validRestore = (value: unknown): value is RecordRestore => object(value) && ['tableId', 'recordId', 'fieldId'].every(key => typeof value[key] === 'string') && ('previous' in value) && (value.previous === null || object(value.previous) && value.previous.formula === undefined);
  runtime.registry.registerMutation<RecordRestore>({ id: 'record.restore', handler: (item, context) => {
    if (!validRestore(item.params)) throw new Error('Invalid Record field restore');
    assertRecordFieldWrite(context.workbook, item.params, item.params.previous?.value ?? null);
    const target = resolveRecordField(context.workbook, item.params), sheet = context.workbook.getSheet(target.sheetId);
    if (item.params.previous === null) sheet.cells.delete(target.row, target.column);
    else sheet.cells.set(target.row, target.column, structuredClone(item.params.previous));
  }, metadata: { calculation: { inputs: 'cells' as const, visibility: false, spillBlockers: 'none' as const, mode: false },
    schema: { name: 'RecordFieldRestore', validate: validRestore }, permission: { capability: 'sheet.cell.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: () => [], mode: 'declared' }, inverseIds: ['record.set'],
  } });
  runtime.registry.registerMutation<import('@react-sheets/core-model').RecordFieldAddress & { value: import('@react-sheets/core-model').TableScalar }>({ id: 'record.set', handler: (item, context) => {
    if (!validSet(item.params)) throw new Error('Invalid record field write');
    assertRecordFieldWrite(context.workbook, item.params, item.params.value);
    const target = resolveRecordField(context.workbook, item.params);
    if (!target.writable) throw new Error('RECORD_FIELD_READ_ONLY');
    context.workbook.getSheet(target.sheetId).cells.set(target.row, target.column, { ...target.cell, value: item.params.value, formula: undefined, formulaValue: undefined });

  }, metadata: { calculation: { inputs: 'cells' as const, visibility: false, spillBlockers: 'none' as const, mode: false },
    schema: { name: 'RecordFieldSet', validate: validSet }, permission: { capability: 'sheet.cell.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: () => [], mode: 'declared' }, inverseIds: ['record.restore'],
  } });
  runtime.registry.registerCommand<import('@react-sheets/core-model').RecordFieldAddress & { value: import('@react-sheets/core-model').TableScalar }>({ id: 'record.set', execute: (params, context) => {
    if (!validSet(params)) throw new Error('Invalid record field write');
    assertRecordFieldWrite(context.workbook, params, params.value);
    const target = resolveRecordField(context.workbook, params), table = context.workbook.getTable(params.tableId);
    if (!target.writable || target.cell?.formula) throw new Error('RECORD_FIELD_READ_ONLY');
    const ranges = [{ sheetId: target.sheetId, startRow: target.row, endRow: target.row, startColumn: target.column, endColumn: target.column }];
    const base = { unitId: context.workbook.unitId, sheetId: target.sheetId, affectedRanges: ranges };
    const sheet = context.workbook.getSheet(target.sheetId);
    const authority = (value: import('@react-sheets/core-model').TableScalar) => createCellSetMutationParams(sheet, { sheetId: target.sheetId, row: target.row, column: target.column, value: { ...target.cell, value } }, 'direct-entry', false, id => context.workbook.getSheet(id)).writeAuthority;
    const mutationParams = { ...params, writeAuthority: authority(params.value) };
    const inverseParams = { tableId: params.tableId, recordId: params.recordId, fieldId: params.fieldId, previous: target.cell ? structuredClone(target.cell) : null };
    context.applyMutation({ id: 'record.set', ...base, params: mutationParams, inverse: [{ id: 'record.restore', ...base, params: inverseParams }], apply: () => {
      context.workbook.getSheet(target.sheetId).cells.set(target.row, target.column, { ...target.cell, value: params.value, formula: undefined, formulaValue: undefined });
  
    } });
    return { operationId: context.operationId, mutationCount: 1, affectedRanges: ranges };
  } });
}

/** Explicit worksheet-to-record migration. IDs and cleared computed inputs commit atomically with field definitions. */
export function registerRecordPromotionCommand(runtime: CommandRuntime): void {
  runtime.registry.registerCommand<{ tableId: string; identityFieldId: string; fieldId?: string; calculation?: import('@react-sheets/core-model').RecordFieldCalculation }>({ id: 'table.promoteRecords', execute: (params, context) => {
    const previous = context.workbook.getTable(params.tableId), table = structuredClone(previous), range = table.sourceRange;
    if (!range) throw new Error('UNSUPPORTED_FEATURE: Record promotion needs a worksheet source');
    const identity = table.fields.find(field => field.id === params.identityFieldId);
    if (!identity || identity.calculation || identity.type !== 'text') throw new Error('Record ID requires a stored text field');
    if (table.recordIdFieldId && table.recordIdFieldId !== identity.id) throw new Error('Record identity field is immutable');
    table.recordIdFieldId = identity.id;
    if (params.fieldId) {
      const field = table.fields.find(field => field.id === params.fieldId);
      if (!field || field.id === identity.id || !params.calculation) throw new Error('Computed field is invalid');
      field.calculation = structuredClone(params.calculation);
      if (field.calculation.kind === 'formula') field.calculation.formula = canonicalRecordFieldFormula(table, field.calculation.formula);
    }
    const sheet = context.workbook.getSheet(range.sheetId), ranges: import('@react-sheets/core-model').RangeRef[] = [], ids = new Set<string>();
    let count = 0;
    const write = (row: number, column: number, value: import('@react-sheets/core-model').CellData): void => {
      const previous = sheet.cells.getWithoutHydration(row, column), affected = [{ sheetId: sheet.id, startRow: row, endRow: row, startColumn: column, endColumn: column }];
      context.applyMutation({ id: 'cell.set', unitId: context.workbook.unitId, sheetId: sheet.id, params: createCellSetMutationParams(sheet, { sheetId: sheet.id, row, column, value }, 'script', false, id => context.workbook.getSheet(id)), affectedRanges: affected,
        inverse: [{ id: 'cell.restore', unitId: context.workbook.unitId, sheetId: sheet.id, params: { sheetId: sheet.id, row, column, previous: previous ? structuredClone(previous) : undefined }, affectedRanges: affected }], apply: () => { sheet.cells.set(row, column, value); } });
      count++; ranges.push(...affected);
    };
    for (let row = range.startRow + 1; row <= range.endRow; row++) {
      const column = range.startColumn + identity.ordinal, cell = sheet.cells.getWithoutHydration(row, column);
      if (cell?.formula) throw new Error('Record identities cannot be formulas');
      let id = cell?.value;
      if (id === undefined || id === null || id === '') { id = `record-${crypto.randomUUID()}`; write(row, column, { ...cell, value: id }); }
      if (typeof id !== 'string' || !id.trim() || ids.has(id)) throw new Error('Record identities must be unique non-empty strings');
      ids.add(id);
      for (const field of table.fields) if (field.calculation && !previous.fields.find(old => old.id === field.id)?.calculation) {
        const column = range.startColumn + field.ordinal, cell = sheet.cells.getWithoutHydration(row, column);
        if (cell?.formula) write(row, column, { ...cell, formula: undefined, formulaValue: undefined, value: null });
        else if (cell?.value != null && cell.value !== '') throw new Error('Computed fields cannot replace stored values');
      }
    }
    assertRecordTable(context.workbook, table);
    const base = { unitId: context.workbook.unitId, sheetId: sheet.id, affectedRanges: [range] };
    context.applyMutation({ id: 'table.configure', ...base, params: { table }, inverse: [{ id: 'table.configure', ...base, params: { table: structuredClone(previous) } }], apply: () => { context.workbook.dataModel.tables.set(table.id, table); } });
    return { operationId: context.operationId, mutationCount: count + 1, affectedRanges: [...ranges, range] };
  } });
}
