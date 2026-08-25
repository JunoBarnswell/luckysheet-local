import type { GanttSheetDefinition, WorkbookTableModel } from '@react-sheets/core-model';
import type { CanvasSheetSnapshot } from '../../ui-snapshot';

export interface GanttTaskProjection {
  id: string;
  title: string;
  row: number;
  start: string;
  end: string;
  startMs: number;
  endMs: number;
  progress: number;
  level: number;
  parentId?: string;
  dependencies: string[];
}

export interface GanttProjection {
  status: 'ready' | 'error';
  error?: string;
  tasks: GanttTaskProjection[];
  timelineStartMs: number;
  timelineEndMs: number;
  unit: GanttSheetDefinition['timeline']['unit'];
}

const DAY_MS = 24 * 60 * 60 * 1000;

function excelSerialToMs(value: number): number {
  return Date.UTC(1899, 11, 30) + value * DAY_MS;
}

function parseDate(value: string): number | null {
  const trimmed = value.trim();
  if (!trimmed) return null;
  const numeric = Number(trimmed);
  if (Number.isFinite(numeric) && /^-?\d+(\.\d+)?$/.test(trimmed)) {
    const serial = excelSerialToMs(numeric);
    return Number.isFinite(serial) ? serial : null;
  }
  const parsed = Date.parse(trimmed);
  return Number.isFinite(parsed) ? parsed : null;
}

function dependencyList(value: string): string[] {
  return [...new Set(value.split(/[;,\s]+/).map((item) => item.trim()).filter(Boolean))];
}

function fail(error: string, unit: GanttSheetDefinition['timeline']['unit']): GanttProjection {
  return { status: 'error', error, tasks: [], timelineStartMs: 0, timelineEndMs: 0, unit };
}

function readField(sheet: CanvasSheetSnapshot, row: number, column: number): string {
  return sheet.getCell(row, column)?.value?.trim() ?? '';
}

function validateAcyclic(tasks: readonly GanttTaskProjection[]): string | undefined {
  const byId = new Map(tasks.map((task) => [task.id, task]));
  const visiting = new Set<string>();
  const visited = new Set<string>();
  const visit = (id: string): string | undefined => {
    if (visiting.has(id)) return `Gantt hierarchy contains a cycle at ${id}`;
    if (visited.has(id)) return undefined;
    visiting.add(id);
    const task = byId.get(id);
    if (task?.parentId) {
      const error = visit(task.parentId);
      if (error) return error;
    }
    for (const dependency of task?.dependencies ?? []) {
      if (!byId.has(dependency)) return `Task ${id} references missing dependency ${dependency}`;
    }
    visiting.delete(id);
    visited.add(id);
    return undefined;
  };
  for (const task of tasks) {
    if (task.parentId && !byId.has(task.parentId)) return `Task ${task.id} references missing parent ${task.parentId}`;
    const error = visit(task.id);
    if (error) return error;
  }
  // Dependency cycles are checked independently from the parent hierarchy.
  const dependencyVisit = (id: string, path: Set<string>): string | undefined => {
    if (path.has(id)) return `Gantt dependencies contain a cycle at ${id}`;
    const task = byId.get(id);
    if (!task) return undefined;
    const next = new Set(path).add(id);
    for (const dependency of task.dependencies) {
      const error = dependencyVisit(dependency, next);
      if (error) return error;
    }
    return undefined;
  };
  for (const task of tasks) {
    const error = dependencyVisit(task.id, new Set());
    if (error) return error;
  }
  return undefined;
}

export function buildGanttProjection(
  sheet: CanvasSheetSnapshot,
  tables: readonly WorkbookTableModel[],
): GanttProjection {
  const definition = sheet.ganttSheet;
  const unit = definition?.timeline.unit ?? 'week';
  if (!definition) return fail('GanttSheet definition is unavailable', unit);
  const table = tables.find((candidate) => candidate.id === definition.viewId);
  if (!table) return fail(`Binding table ${definition.viewId} is unavailable`, unit);
  const fieldIds = new Set(table.fields.map((field) => field.id));
  const required = ['id', 'title', 'start', 'end', 'progress'] as const;
  for (const key of required) {
    if (!fieldIds.has(definition.fieldMap[key])) return fail(`Gantt field mapping ${key} is invalid`, unit);
  }
  for (const key of ['parentId', 'dependencies'] as const) {
    const fieldId = definition.fieldMap[key];
    if (fieldId && !fieldIds.has(fieldId)) return fail(`Gantt field mapping ${key} is invalid`, unit);
  }
  if (definition.calendar.workingDays.some((day) => !Number.isInteger(day) || day < 0 || day > 6)
    || definition.calendar.dayStartHour < 0 || definition.calendar.dayEndHour > 24
    || definition.calendar.dayStartHour >= definition.calendar.dayEndHour) {
    return fail('Gantt calendar definition is invalid', unit);
  }
  const ordinal = (fieldId: string) => table.fields.find((field) => field.id === fieldId)?.ordinal;
  const columnOf = (fieldId: string) => ordinal(fieldId) ?? -1;
  const tasks: GanttTaskProjection[] = [];
  const ids = new Set<string>();
  const rowCount = table.sourceRange ? Math.max(0, table.rowCount) : Math.max(0, sheet.rowCount - 1);
  for (let index = 0; index < rowCount; index += 1) {
    const row = index + 1;
    const id = readField(sheet, row, columnOf(definition.fieldMap.id));
    if (!id) continue;
    if (!ids.add(id)) return fail(`Gantt task id ${id} is duplicated`, unit);
    const title = readField(sheet, row, columnOf(definition.fieldMap.title));
    const start = readField(sheet, row, columnOf(definition.fieldMap.start));
    const end = readField(sheet, row, columnOf(definition.fieldMap.end));
    const startMs = parseDate(start);
    const endMs = parseDate(end);
    if (startMs === null || endMs === null || endMs < startMs) return fail(`Task ${id} has an invalid date range`, unit);
    const progressValue = Number(readField(sheet, row, columnOf(definition.fieldMap.progress)) || '0');
    if (!Number.isFinite(progressValue) || progressValue < 0 || progressValue > 100) return fail(`Task ${id} has invalid progress`, unit);
    const parentId = definition.fieldMap.parentId ? readField(sheet, row, columnOf(definition.fieldMap.parentId)) || undefined : undefined;
    const dependencies = definition.fieldMap.dependencies ? dependencyList(readField(sheet, row, columnOf(definition.fieldMap.dependencies))) : [];
    tasks.push({ id, title, row, start, end, startMs, endMs, progress: progressValue, level: 0, parentId, dependencies });
  }
  const cycleError = validateAcyclic(tasks);
  if (cycleError) return fail(cycleError, unit);
  const byId = new Map(tasks.map((task) => [task.id, task]));
  const levels = new Map<string, number>();
  const levelOf = (task: GanttTaskProjection): number => {
    const cached = levels.get(task.id);
    if (cached !== undefined) return cached;
    const level = task.parentId && byId.has(task.parentId) ? levelOf(byId.get(task.parentId)!) + 1 : 0;
    levels.set(task.id, level);
    return level;
  };
  for (const task of tasks) task.level = levelOf(task);
  const taskStart = tasks.length > 0 ? Math.min(...tasks.map((task) => task.startMs)) : Date.now();
  const taskEnd = tasks.length > 0 ? Math.max(...tasks.map((task) => task.endMs)) : taskStart + DAY_MS;
  const configuredStart = definition.timeline.start ? parseDate(definition.timeline.start) : null;
  const configuredEnd = definition.timeline.end ? parseDate(definition.timeline.end) : null;
  if (definition.timeline.start && configuredStart === null) return fail('Gantt timeline start is invalid', unit);
  if (definition.timeline.end && configuredEnd === null) return fail('Gantt timeline end is invalid', unit);
  const timelineStartMs = configuredStart ?? taskStart;
  const timelineEndMs = configuredEnd ?? taskEnd;
  if (timelineEndMs < timelineStartMs) return fail('Gantt timeline bounds are invalid', unit);
  return { status: 'ready', tasks, timelineStartMs, timelineEndMs, unit };
}
