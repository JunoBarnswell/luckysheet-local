import type {
  ConnectorDrawingPayload,
  DrawingConnectionEndpoint,
  DrawingConnectionPoint,
  DrawingObject,
  DrawingPayload,
  DrawingTransform,
  DrawingGroup,
  WorksheetSnapSettings,
} from './domain';
import {
  DEFAULT_WORKSHEET_SNAP_SETTINGS,
  isDrawingConnectorPayload,
  isDrawingGroup,
  isWorksheetSnapSettings,
} from './domain';

export interface DrawingGraphSheet {
  id: string;
  drawings: readonly DrawingObject[];
  drawingPayloads: ReadonlyMap<string, DrawingPayload> | Readonly<Record<string, DrawingPayload>>;
  drawingGroups?: readonly DrawingGroup[];
  snapSettings?: WorksheetSnapSettings;
}

export interface ConnectorRoutePlan {
  payload: ConnectorDrawingPayload;
  transform: DrawingTransform;
}

export interface ConnectorTransformOverride {
  drawingId: string;
  transform: DrawingTransform;
}

function payloadAt(sheet: DrawingGraphSheet, payloadId: string): DrawingPayload | undefined {
  const source = sheet.drawingPayloads as ReadonlyMap<string, DrawingPayload> & Readonly<Record<string, DrawingPayload>>;
  return typeof source.get === 'function' ? source.get(payloadId) : source[payloadId];
}

function finiteTransform(transform: DrawingTransform): boolean {
  return Number.isFinite(transform.x) && Number.isFinite(transform.y)
    && Number.isFinite(transform.width) && transform.width >= 0
    && Number.isFinite(transform.height) && transform.height >= 0
    && (transform.rotation === undefined || Number.isFinite(transform.rotation));
}

function endpointPosition(drawing: DrawingObject, point: DrawingConnectionPoint): { x: number; y: number } {
  const transform = drawing.transform;
  const local = {
    x: point === 'right' ? transform.width : point === 'left' ? 0 : transform.width / 2,
    y: point === 'bottom' ? transform.height : point === 'top' ? 0 : transform.height / 2,
  };
  const angle = (transform.rotation ?? 0) * Math.PI / 180;
  if (angle === 0) return { x: transform.x + local.x, y: transform.y + local.y };
  const center = { x: transform.width / 2, y: transform.height / 2 };
  const dx = local.x - center.x;
  const dy = local.y - center.y;
  return {
    x: transform.x + center.x + dx * Math.cos(angle) - dy * Math.sin(angle),
    y: transform.y + center.y + dx * Math.sin(angle) + dy * Math.cos(angle),
  };
}

function deduplicateRoute(points: Array<{ x: number; y: number }>): Array<{ x: number; y: number }> {
  const result: Array<{ x: number; y: number }> = [];
  for (const point of points) {
    const previous = result.at(-1);
    if (!previous || previous.x !== point.x || previous.y !== point.y) result.push(point);
  }
  return result;
}

function routePoints(type: ConnectorDrawingPayload['connectorType'], start: { x: number; y: number }, end: { x: number; y: number }): Array<{ x: number; y: number }> {
  if (type === 'straight') return deduplicateRoute([start, end]);
  if (type === 'elbow') {
    const horizontalFirst = Math.abs(end.x - start.x) >= Math.abs(end.y - start.y);
    const bend = horizontalFirst ? { x: end.x, y: start.y } : { x: start.x, y: end.y };
    return deduplicateRoute([start, bend, end]);
  }
  const dx = end.x - start.x;
  const dy = end.y - start.y;
  const offset = Math.max(16, Math.hypot(dx, dy) * 0.2);
  const bend = { x: (start.x + end.x) / 2, y: (start.y + end.y) / 2 - offset };
  return deduplicateRoute([start, bend, end]);
}

function routeTransform(points: readonly { x: number; y: number }[]): DrawingTransform {
  if (points.length === 0) throw new Error('Connector route must contain points');
  const xs = points.map((point) => point.x);
  const ys = points.map((point) => point.y);
  const x = Math.min(...xs);
  const y = Math.min(...ys);
  return { x, y, width: Math.max(...xs) - x, height: Math.max(...ys) - y, rotation: 0 };
}

function drawingAt(sheet: DrawingGraphSheet, drawingId: string, overrides: ReadonlyMap<string, DrawingTransform>): DrawingObject {
  const drawing = sheet.drawings.find((entry) => entry.id === drawingId);
  if (!drawing) throw new Error(`Connector endpoint drawing is missing: ${drawingId}`);
  const transform = overrides.get(drawingId);
  return transform ? { ...drawing, transform } : drawing;
}

function endpointAt(sheet: DrawingGraphSheet, endpoint: DrawingConnectionEndpoint, overrides: ReadonlyMap<string, DrawingTransform>): { x: number; y: number } {
  const drawing = drawingAt(sheet, endpoint.drawingId, overrides);
  if (drawing.kind === 'connector') throw new Error(`Connector endpoint cannot target another connector: ${endpoint.drawingId}`);
  return endpointPosition(drawing, endpoint.connectionPoint);
}

/** Compute the only valid route for a connector from the current endpoint graph. */
export function planConnectorRoute(
  sheet: DrawingGraphSheet,
  connectorDrawing: DrawingObject,
  payload: ConnectorDrawingPayload,
  overrides: readonly ConnectorTransformOverride[] = [],
): ConnectorRoutePlan {
  if (connectorDrawing.kind !== 'connector' || connectorDrawing.sheetId !== sheet.id) throw new Error(`Connector drawing sheet/kind is invalid: ${connectorDrawing.id}`);
  if (!isDrawingConnectorPayload(payload)) throw new Error(`Connector payload is invalid: ${connectorDrawing.payloadId}`);
  if (payload.start.drawingId === payload.end.drawingId) throw new Error(`Connector cannot connect a drawing to itself: ${connectorDrawing.id}`);
  const transformOverrides = new Map(overrides.map((entry) => [entry.drawingId, entry.transform]));
  const start = endpointAt(sheet, payload.start, transformOverrides);
  const end = endpointAt(sheet, payload.end, transformOverrides);
  const points = routePoints(payload.connectorType, start, end);
  if (points.length < 2) throw new Error(`Connector route is degenerate: ${connectorDrawing.id}`);
  return {
    payload: { ...structuredClone(payload), route: { points: structuredClone(points) } },
    transform: routeTransform(points),
  };
}

function sameJson(left: unknown, right: unknown): boolean {
  return JSON.stringify(left) === JSON.stringify(right);
}

/** Validate that persisted route and transform are planner output, not client-supplied geometry. */
export function assertCanonicalConnector(sheet: DrawingGraphSheet, drawing: DrawingObject, payload: ConnectorDrawingPayload): void {
  const planned = planConnectorRoute(sheet, drawing, payload);
  if (!sameJson(payload.route, planned.payload.route)) throw new Error(`Connector route is not canonical: ${drawing.id}`);
  if (!sameJson(drawing.transform, planned.transform)) throw new Error(`Connector transform is not canonical: ${drawing.id}`);
}

/** Validate all drawing graph invariants before a snapshot or mutation enters the model. */
export function validateDrawingGraph(sheet: DrawingGraphSheet): void {
  if (!sheet.id.trim()) throw new Error('Drawing graph sheet identity is required');
  const ids = new Set<string>();
  const payloadIds = new Set<string>();
  for (const drawing of sheet.drawings) {
    if (!drawing.id.trim() || ids.has(drawing.id)) throw new Error(`Drawing identity is duplicated or empty: ${drawing.id}`);
    if (drawing.sheetId !== sheet.id || !finiteTransform(drawing.transform)) throw new Error(`Drawing graph contains an invalid drawing: ${drawing.id}`);
    const payload = payloadAt(sheet, drawing.payloadId);
    if (!payload || payload.kind !== drawing.kind) throw new Error(`Drawing payload is missing or mismatched: ${drawing.id}`);
    if (payloadIds.has(drawing.payloadId)) throw new Error(`Drawing payload identity is duplicated: ${drawing.payloadId}`);
    payloadIds.add(drawing.payloadId);
    ids.add(drawing.id);
  }
  const payloadEntries = sheet.drawingPayloads instanceof Map ? [...sheet.drawingPayloads.entries()] : Object.entries(sheet.drawingPayloads);
  const orphanConnector = payloadEntries.find(([payloadId, payload]) => !payloadIds.has(payloadId) && payload.kind === 'connector');
  if (orphanConnector) throw new Error(`Connector payload is orphaned: ${orphanConnector[0]}`);
  for (const drawing of sheet.drawings) {
    const payload = payloadAt(sheet, drawing.payloadId);
    if (payload?.kind === 'connector') assertCanonicalConnector(sheet, drawing, payload);
  }
  const owned = new Set<string>();
  const groups = sheet.drawingGroups ?? [];
  for (const group of groups) {
    if (!isDrawingGroup(group) || group.sheetId !== sheet.id) throw new Error(`Drawing group is invalid: ${group.id}`);
    if (groups.filter((entry) => entry.id === group.id).length > 1) throw new Error(`Drawing group identity is duplicated: ${group.id}`);
    const members = new Set(group.memberDrawingIds);
    if (members.size !== group.memberDrawingIds.length) throw new Error(`Drawing group contains duplicate members: ${group.id}`);
    for (const memberId of members) {
      if (!ids.has(memberId)) throw new Error(`Drawing group references missing drawing: ${memberId}`);
      if (owned.has(memberId)) throw new Error(`Drawing belongs to multiple groups: ${memberId}`);
      owned.add(memberId);
    }
  }
  if (sheet.snapSettings !== undefined && !isWorksheetSnapSettings(sheet.snapSettings)) throw new Error('Worksheet snap settings are invalid');
}

export function canonicalSnapSettings(value: WorksheetSnapSettings | undefined): WorksheetSnapSettings {
  const settings = value ?? DEFAULT_WORKSHEET_SNAP_SETTINGS;
  if (!isWorksheetSnapSettings(settings)) throw new Error('Worksheet snap settings are invalid');
  return structuredClone(settings);
}

export function recomputeConnectorRoutes(
  sheet: DrawingGraphSheet,
  transforms: readonly ConnectorTransformOverride[],
): Array<{ drawingId: string; before: ConnectorRoutePlan; after: ConnectorRoutePlan }> {
  const changedIds = new Set(transforms.map((entry) => entry.drawingId));
  const result: Array<{ drawingId: string; before: ConnectorRoutePlan; after: ConnectorRoutePlan }> = [];
  for (const drawing of sheet.drawings) {
    const payload = payloadAt(sheet, drawing.payloadId);
    if (payload?.kind !== 'connector') continue;
    if (!changedIds.has(payload.start.drawingId) && !changedIds.has(payload.end.drawingId)) continue;
    const before = { payload: structuredClone(payload), transform: structuredClone(drawing.transform) };
    const after = planConnectorRoute(sheet, drawing, payload, transforms);
    result.push({ drawingId: drawing.id, before, after });
  }
  return result;
}
