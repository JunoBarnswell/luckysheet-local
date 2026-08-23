import type {
  PivotAggregateFunction, PivotDataSource, PivotFieldCatalog, PivotFieldDataType, PivotFieldPlacement,
  PivotFilter, PivotGroup, PivotModel, PivotResultCell, PivotResultNode, PivotResultTree, PivotScalar,
  PivotSourceRowPath, PivotValueField, WorkbookModel, WorksheetModel,
  PivotCalculatedField, PivotCalculatedItem,
} from '@react-sheets/core-model';
import { FormulaEngine, type FormulaValue } from '@react-sheets/formula-engine';

interface SourceRow { values: Record<string, PivotScalar>; paths: PivotSourceRowPath[]; }
interface AxisGroup { values: PivotScalar[]; rows: SourceRow[]; }

export interface PivotResultTable {
  headers: string[];
  rows: Array<{ keys: string[]; values: PivotScalar[] }>;
  grandTotal: PivotScalar[];
  tree: PivotResultTree;
}

const jsonKey = (values: PivotScalar[]): string => JSON.stringify(values);
const same = (left: PivotScalar, right: PivotScalar): boolean => left === right || (left == null && right == null);
const display = (value: PivotScalar): string => value == null || value === '' ? '(blank)' : String(value);

function toNumber(value: PivotScalar): number | null {
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  if (typeof value !== 'string' || value.trim() === '') return null;
  const parsed = Number(value.replace(/[$,%]/g, ''));
  return Number.isFinite(parsed) ? parsed : null;
}

function compare(left: PivotScalar, right: PivotScalar): number {
  if (same(left, right)) return 0;
  if (left == null) return -1;
  if (right == null) return 1;
  const leftNumber = toNumber(left); const rightNumber = toNumber(right);
  return leftNumber != null && rightNumber != null ? leftNumber - rightNumber : String(left).localeCompare(String(right));
}

function inferType(values: PivotScalar[]): PivotFieldDataType {
  const present = values.filter((value) => value != null && value !== '');
  if (!present.length) return 'mixed';
  if (present.every((value) => typeof value === 'boolean')) return 'boolean';
  if (present.every((value) => typeof value === 'number' && Number.isFinite(value))) return 'number';
  const dateLike = present.every((value) => typeof value === 'string' && /^\d{4}-\d{2}-\d{2}(?:[T ].*)?$/.test(value) && !Number.isNaN(Date.parse(value)));
  if (dateLike) return 'date';
  if (present.every((value) => typeof value === 'string')) return 'text';
  return 'mixed';
}

function sourceRange(pivot: PivotModel): PivotDataSource {
  return pivot.dataSource ?? { kind: 'worksheet-range', range: pivot.sourceRange };
}

function readTable(sheet: WorksheetModel, range: { sheetId: string; startRow: number; endRow: number; startColumn: number; endColumn: number }): SourceRow[] {
  const names: string[] = [];
  for (let column = range.startColumn; column <= range.endColumn; column++) names.push(String(sheet.cells.get(range.startRow, column)?.value ?? `Col${column - range.startColumn + 1}`));
  const rows: SourceRow[] = [];
  for (let row = range.startRow + 1; row <= range.endRow; row++) {
    const values: Record<string, PivotScalar> = {};
    names.forEach((name, index) => { values[name] = sheet.cells.get(row, range.startColumn + index)?.value ?? null; });
    rows.push({ values, paths: [{ sheetId: range.sheetId, row }] });
  }
  return rows;
}

const formulaFunctions = new Set(['SUM', 'COUNT', 'AVERAGE', 'MIN', 'MAX', 'IF', 'AND', 'OR', 'NOT', 'ROUND', 'ABS', 'CONCAT', 'LEFT', 'RIGHT', 'LEN']);

function columnLabel(index: number): string {
  let value = '';
  let current = index + 1;
  while (current > 0) {
    const remainder = (current - 1) % 26;
    value = String.fromCharCode(65 + remainder) + value;
    current = Math.floor((current - 1) / 26);
  }
  return value;
}

function formulaScalar(value: FormulaValue): PivotScalar | null {
  return typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean' || value === null ? value : null;
}

function rewriteCalculatedFormula(formula: string, fields: string[]): string {
  let rewritten = formula.trim().replace(/^=/, '');
  fields
    .map((field, index) => ({ field, index }))
    .sort((left, right) => right.field.length - left.field.length)
    .forEach(({ field, index }) => {
      const escaped = field.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      const reference = `${columnLabel(index)}1`;
      rewritten = rewritten.replace(new RegExp(`\\[${escaped}\\]`, 'g'), reference);
      if (!formulaFunctions.has(field.toUpperCase())) {
        rewritten = rewritten.replace(new RegExp(`(?<![A-Za-z0-9_])${escaped}(?![A-Za-z0-9_])`, 'g'), reference);
      }
    });
  return '=' + rewritten;
}

function calculateRowFormula(row: SourceRow, formula: string, fields: string[]): PivotScalar | null {
  const engine = new FormulaEngine({ defaultSheetId: 'pivot' });
  fields.forEach((field, index) => engine.setValue({ sheetId: 'pivot', row: 0, column: index }, row.values[field] ?? null));
  engine.setFormula({ sheetId: 'pivot', row: 1, column: 0 }, rewriteCalculatedFormula(formula, fields));
  return formulaScalar(engine.getCellValue({ sheetId: 'pivot', row: 1, column: 0 }));
}

function applyCalculatedData(rows: SourceRow[], fields: PivotCalculatedField[] = [], items: PivotCalculatedItem[] = []): SourceRow[] {
  if (fields.length === 0 && items.length === 0) return rows;
  return rows.map((row) => {
    const values = { ...row.values };
    for (const field of fields) {
      values[field.name] = calculateRowFormula({ ...row, values }, field.formula, Object.keys(values));
    }
    for (const item of items) {
      if (same(values[item.field] ?? null, item.name)) values[item.name] = calculateRowFormula({ ...row, values }, item.formula, Object.keys(values));
    }
    return { ...row, values };
  });
}

function joinTables(workbook: WorkbookModel, source: Extract<PivotDataSource, { kind: 'worksheet-ranges' }>): SourceRow[] {
  if (!source.ranges.length) return [];
  let current = readTable(workbook.getSheet(source.ranges[0]!.sheetId), source.ranges[0]!);
  for (let index = 1; index < source.ranges.length; index++) {
    const range = source.ranges[index]!; const right = readTable(workbook.getSheet(range.sheetId), range);
    const relationship = source.relationships.find((candidate) => candidate.left.sheetId === source.ranges[index - 1]!.sheetId && candidate.right.sheetId === range.sheetId);
    if (!relationship) throw new Error('Every local worksheet range must have an adjacent typed relationship');
    const joined: SourceRow[] = [];
    for (const left of current) {
      const matches = right.filter((row) => same(left.values[relationship.left.field] ?? null, row.values[relationship.right.field] ?? null));
      if (!matches.length && relationship.join === 'left') joined.push(left);
      for (const match of matches) joined.push({ values: { ...left.values, ...match.values }, paths: [...left.paths, ...match.paths] });
    }
    current = joined;
  }
  return current;
}

function readSource(workbook: WorkbookModel, pivot: PivotModel): SourceRow[] {
  const source = sourceRange(pivot);
  const rows = source.kind === 'worksheet-range' ? readTable(workbook.getSheet(source.range.sheetId), source.range) : joinTables(workbook, source);
  return applyCalculatedData(rows, pivot.layout.calculatedFields, pivot.layout.calculatedItems);
}

export function getPivotFieldCatalog(workbook: WorkbookModel, pivot: PivotModel): PivotFieldCatalog {
  if (pivot.fieldCatalog) return structuredClone(pivot.fieldCatalog);
  const rows = readSource(workbook, pivot); const names = new Set<string>();
  rows.forEach((row) => Object.keys(row.values).forEach((name) => names.add(name)));
  return { fields: [...names].map((name, ordinal) => {
    const values = rows.map((row) => row.values[name] ?? null);
    return { id: name, name, ordinal, dataType: inferType(values), values: [...new Map(values.filter((value) => value != null).map((value) => [JSON.stringify(value), value])).values()] };
  }) };
}

function matchesFilter(row: SourceRow, filter: PivotFilter): boolean {
  const value = row.values[filter.field] ?? null;
  if (filter.kind === 'top-items') throw new Error('top-items filters must be applied by the topItems stage');
  if (filter.kind === 'manual') return filter.exclude ? !filter.selected.some((item) => same(item, value)) : filter.selected.some((item) => same(item, value));
  const leftNumber = toNumber(value); const rightNumber = toNumber(filter.value); const order = leftNumber != null && rightNumber != null ? leftNumber - rightNumber : compare(value, filter.value);
  switch (filter.operator) {
    case 'equals': return same(value, filter.value); case 'not-equals': return !same(value, filter.value); case 'contains': return String(value ?? '').includes(String(filter.value ?? ''));
    case 'greater-than': return order > 0; case 'greater-or-equal': return order >= 0; case 'less-than': return order < 0; case 'less-or-equal': return order <= 0;
  }
}

function aggregate(rows: SourceRow[], field: string, operation: PivotAggregateFunction): number | null {
  let count = 0;
  let sum = 0;
  let min = Number.POSITIVE_INFINITY;
  let max = Number.NEGATIVE_INFINITY;
  let product = 1;
  let mean = 0;
  let squaredDelta = 0;
  const distinct = new Set<string>();
  for (const row of rows) {
    const raw = row.values[field] ?? null;
    if (operation === 'count' && raw != null && raw !== '') count += 1;
    if (operation === 'distinct-count' && raw != null) distinct.add(JSON.stringify(raw));
    const number = toNumber(raw);
    if (number == null) continue;
    count += operation === 'count' ? 0 : 1;
    sum += number;
    min = Math.min(min, number);
    max = Math.max(max, number);
    product *= number;
    const delta = number - mean;
    mean += delta / count;
    squaredDelta += delta * (number - mean);
  }
  switch (operation) {
    case 'count': return count;
    case 'count-numbers': return count;
    case 'sum': return count ? sum : 0;
    case 'average': return count ? sum / count : null;
    case 'min': return count ? min : null;
    case 'max': return count ? max : null;
    case 'product': return count ? product : null;
    case 'distinct-count': return distinct.size;
    case 'stdev': return count < 2 ? null : Math.sqrt(squaredDelta / (count - 1));
    case 'stdevp': return count ? Math.sqrt(squaredDelta / count) : null;
    case 'var': return count < 2 ? null : squaredDelta / (count - 1);
    case 'varp': return count ? squaredDelta / count : null;
  }
}

function grouped(value: PivotScalar, group?: PivotGroup): PivotScalar {
  if (!group || value == null) return value;
  if (group.kind === 'manual') return group.groups.find((candidate) => candidate.items.some((item) => same(item, value)))?.name ?? value;
  if (group.kind === 'number') { const number = toNumber(value); if (number == null || group.interval <= 0) return value; const start = group.start ?? 0; return start + Math.floor((number - start) / group.interval) * group.interval; }
  const date = new Date(String(value)); if (Number.isNaN(date.getTime())) return value;
  if (group.unit === 'year') return date.getFullYear(); if (group.unit === 'quarter') return `${date.getFullYear()} Q${Math.floor(date.getMonth() / 3) + 1}`; if (group.unit === 'month') return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
  if (group.unit === 'week') return Math.ceil((((date.getTime() - new Date(date.getFullYear(), 0, 1).getTime()) / 86400000) + date.getDay() + 1) / 7);
  return date.toISOString().slice(0, 10);
}

function axisGroups(rows: SourceRow[], placements: PivotFieldPlacement[]): AxisGroup[] {
  const map = new Map<string, AxisGroup>();
  for (const row of rows) { const values = placements.map((placement) => grouped(row.values[placement.field] ?? null, placement.group)); const group = map.get(jsonKey(values)) ?? { values, rows: [] }; group.rows.push(row); map.set(jsonKey(values), group); }
  const placement = placements[placements.length - 1]; const result = [...map.values()].sort((left, right) => {
    if (placement?.sort?.by === 'value' && placement.sort.valueField) return (aggregate(left.rows, placement.sort.valueField, 'sum') ?? 0) - (aggregate(right.rows, placement.sort.valueField, 'sum') ?? 0);
    for (let index = 0; index < left.values.length; index++) { const order = compare(left.values[index] ?? null, right.values[index] ?? null); if (order) return order; } return 0;
  });
  if (placement?.sort?.direction === 'descending') result.reverse(); return result;
}

function topItems(rows: SourceRow[], filters: PivotFilter[]): SourceRow[] {
  let result = rows;
  for (const filter of filters) {
    if (filter.kind !== 'top-items' || filter.count < 1) continue;
    const buckets = new Map<string, SourceRow[]>(); for (const row of result) { const bucketKey = JSON.stringify(row.values[filter.field] ?? null); const bucket = buckets.get(bucketKey) ?? []; bucket.push(row); buckets.set(bucketKey, bucket); }
    const ranked = [...buckets.values()].sort((left, right) => (aggregate(left, filter.valueField, 'sum') ?? 0) - (aggregate(right, filter.valueField, 'sum') ?? 0)); if (filter.direction === 'top') ranked.reverse(); result = ranked.slice(0, filter.count).flat();
  }
  return result;
}

function matchesSlicersAndTimelines(workbook: WorkbookModel, rows: SourceRow[], pivot: PivotModel): SourceRow[] {
  const linked = workbook.getSheet(pivot.sheetId).pivots.filter((candidate) => candidate.id !== pivot.id);
  const slicers = [...(pivot.slicers ?? []), ...linked.flatMap((candidate) => (candidate.slicers ?? []).filter((slicer) => slicer.connectedPivotIds?.includes(pivot.id)))];
  const timelines = [...(pivot.timelines ?? []), ...linked.flatMap((candidate) => (candidate.timelines ?? []).filter((timeline) => timeline.connectedPivotIds?.includes(pivot.id)))];
  return rows.filter((row) => slicers.every((slicer) => slicer.selected.length === 0 || slicer.selected.some((value) => same(value, row.values[slicer.field] ?? null))) && timelines.every((timeline) => {
    const raw = row.values[timeline.field];
    if (raw == null) return false;
    const date = new Date(String(raw));
    if (Number.isNaN(date.getTime())) return false;
    const start = timeline.start ? new Date(timeline.start).getTime() : Number.NEGATIVE_INFINITY;
    const end = timeline.end ? new Date(timeline.end).getTime() : Number.POSITIVE_INFINITY;
    return date.getTime() >= start && date.getTime() <= end;
  }));
}

function resultCells(rows: SourceRow[], columns: AxisGroup[], values: PivotValueField[]): PivotResultCell[] {
  return columns.map((column) => { const columnRows = rows.filter((row) => column.rows.includes(row)); return { kind: 'detail', columnPath: column.values, sourceRowPaths: columnRows.flatMap((row) => row.paths), values: values.map((value) => aggregate(columnRows, value.field, value.summarizeBy)) }; });
}
function resultNodes(rows: SourceRow[], placements: PivotFieldPlacement[], depth: number, columns: AxisGroup[], values: PivotValueField[], showSubtotals: boolean): PivotResultNode[] {
  if (depth >= placements.length) return [];
  return axisGroups(rows, [placements[depth]!]).map((group) => { const children = resultNodes(group.rows, placements, depth + 1, columns, values, showSubtotals); const leaf = children.length === 0; return { kind: leaf ? 'leaf' : 'subtotal', field: placements[depth]!.field, key: group.values[0] ?? null, label: display(group.values[0] ?? null), depth, children, values: resultCells(group.rows, columns, values), subtotal: showSubtotals && !leaf, sourceRowPaths: group.rows.flatMap((row) => row.paths) }; });
}

function applyShowAs(tree: PivotResultTree, fields: PivotValueField[]): void {
  const leaves: PivotResultNode[] = []; const collect = (nodes: PivotResultNode[]) => nodes.forEach((node) => node.children.length ? collect(node.children) : leaves.push(node)); collect(tree.rows);
  const raw = new Map<PivotResultCell, PivotScalar[]>(); const snapshot = (nodes: PivotResultNode[]) => nodes.forEach((node) => { node.values.forEach((cell) => raw.set(cell, [...cell.values])); snapshot(node.children); }); snapshot(tree.rows);
  const rawValues = (cell: PivotResultCell | undefined, index: number): PivotScalar | null => cell ? raw.get(cell)?.[index] ?? null : null;
  const visit = (nodes: PivotResultNode[], parent?: PivotResultNode) => nodes.forEach((node) => {
    node.values.forEach((cell, columnIndex) => fields.forEach((field, valueIndex) => {
      const spec = field.showAs ?? { kind: 'normal' as const }; const current = toNumber(rawValues(cell, valueIndex)); if (current == null || spec.kind === 'normal') return;
      const grand = toNumber(tree.grandTotal?.values[valueIndex] ?? null); const rowTotal = node.values.reduce((sum, item) => sum + (toNumber(rawValues(item, valueIndex)) ?? 0), 0); const columnTotal = leaves.reduce((sum, item) => sum + (toNumber(rawValues(item.values[columnIndex], valueIndex)) ?? 0), 0); const parentTotal = parent ? toNumber(rawValues(parent.values[columnIndex], valueIndex)) : null;
      if (spec.kind === 'grand-percentage') cell.values[valueIndex] = grand ? current / grand : null;
      else if (spec.kind === 'row-percentage') cell.values[valueIndex] = rowTotal ? current / rowTotal : null;
      else if (spec.kind === 'column-percentage') cell.values[valueIndex] = columnTotal ? current / columnTotal : null;
      else if (spec.kind === 'parent-percentage') cell.values[valueIndex] = parentTotal ? current / parentTotal : null;
      else if (spec.kind === 'difference' || spec.kind === 'percentage-difference') { const base = spec.base === 'grand' ? grand : spec.base === 'row' ? rowTotal : spec.base === 'column' ? columnTotal : parentTotal; cell.values[valueIndex] = base == null ? null : spec.kind === 'difference' ? current - base : base ? (current - base) / base : null; }
      else if (spec.kind === 'running-total') { const end = spec.axis === 'row' ? leaves.indexOf(node) : columnIndex; if (spec.axis === 'row') cell.values[valueIndex] = leaves.slice(0, end + 1).reduce((sum, item) => sum + (toNumber(rawValues(item.values[columnIndex], valueIndex)) ?? 0), 0); else cell.values[valueIndex] = node.values.slice(0, end + 1).reduce((sum, item) => sum + (toNumber(rawValues(item, valueIndex)) ?? 0), 0); }
      else if (spec.kind === 'rank') { const series = spec.axis === 'row' ? leaves.map((item) => toNumber(rawValues(item.values[columnIndex], valueIndex))) : node.values.map((item) => toNumber(rawValues(item, valueIndex))); const ranked = series.filter((value): value is number => value != null).sort((left, right) => spec.direction === 'ascending' ? left - right : right - left); cell.values[valueIndex] = ranked.indexOf(current) + 1; }
      else if (spec.kind === 'index') cell.values[valueIndex] = grand != null && rowTotal && columnTotal ? current * grand / rowTotal / columnTotal : null;
    }));
    visit(node.children, node);
  }); visit(tree.rows);
}

export function computePivotResult(workbook: WorkbookModel, pivot: PivotModel): PivotResultTree {
  const definition = pivot; let rows = matchesSlicersAndTimelines(workbook, readSource(workbook, definition), definition); rows = rows.filter((row) => definition.layout.filters.filter((filter) => filter.kind !== 'top-items').every((filter) => matchesFilter(row, filter))); rows = topItems(rows, definition.layout.filters);
  const catalog = getPivotFieldCatalog(workbook, definition); const columns = definition.layout.columns.length ? axisGroups(rows, definition.layout.columns) : [{ values: [], rows }]; const grandTotal: PivotResultCell | null = definition.layout.showGrandTotals ? { kind: 'grand-total', columnPath: [], values: definition.layout.values.map((field) => aggregate(rows, field.field, field.summarizeBy)), sourceRowPaths: rows.flatMap((row) => row.paths) } : null;
  const tree: PivotResultTree = { schema: 'PivotResultTreeV1', pivotId: definition.id, fields: catalog, columnPaths: columns.map((column) => column.values), rows: resultNodes(rows, definition.layout.rows, 0, columns, definition.layout.values, definition.layout.showSubtotals), grandTotal, sourceRowPaths: rows.flatMap((row) => row.paths) }; applyShowAs(tree, definition.layout.values); return tree;
}

export function computePivotTable(workbook: WorkbookModel, pivot: PivotModel): PivotResultTable {
  const tree = computePivotResult(workbook, pivot); const definition = pivot; const rows = tree.rows.map((node) => ({ keys: [node.label], values: node.values.flatMap((cell) => cell.values) })); const headers = [...definition.layout.rows.map((field) => field.field), ...tree.columnPaths.flatMap((path) => definition.layout.values.map((field) => path.length ? `${path.map(display).join(' / ')} ${field.displayName ?? `${field.summarizeBy.toUpperCase()} of ${field.field}`}` : field.displayName ?? `${field.summarizeBy.toUpperCase()} of ${field.field}`))]; return { headers, rows, grandTotal: tree.grandTotal?.values ?? [], tree };
}
