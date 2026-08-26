import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import {
  DEFAULT_WORKSHEET_SNAP_SETTINGS,
  planConnectorRoute,
  validateDrawingGraph,
  type ConnectorDrawingPayload,
  type DrawingGraphSheet,
  type DrawingObject,
} from './index';

const sheetId = 'sheet-1';

function shape(id: string, x: number, y: number): DrawingObject {
  return {
    id,
    sheetId,
    kind: 'shape',
    payloadId: `${id}-payload`,
    anchor: { kind: 'absolute' },
    transform: { x, y, width: 40, height: 20, rotation: 0 },
    zIndex: 1,
  };
}

function connectorPayload(): ConnectorDrawingPayload {
  return {
    kind: 'connector',
    connectorType: 'elbow',
    start: { drawingId: 'shape-a', connectionPoint: 'right' },
    end: { drawingId: 'shape-b', connectionPoint: 'left' },
    stroke: '#111827',
    startArrowhead: 'none',
    endArrowhead: 'triangle',
    route: { points: [{ x: 0, y: 0 }, { x: 0, y: 0 }] },
  };
}

function graph(): DrawingGraphSheet {
  const a = shape('shape-a', 20, 20);
  const b = shape('shape-b', 180, 80);
  const connector: DrawingObject = {
    id: 'connector-1',
    sheetId,
    kind: 'connector',
    payloadId: 'connector-1-payload',
    anchor: { kind: 'absolute' },
    transform: { x: 0, y: 0, width: 0, height: 0, rotation: 0 },
    zIndex: 2,
  };
  const payloads = new Map<string, any>([
    [a.payloadId, { kind: 'shape', type: 'rectangle', fill: '#fff', stroke: '#000' }],
    [b.payloadId, { kind: 'shape', type: 'rectangle', fill: '#fff', stroke: '#000' }],
  ]);
  const payload = connectorPayload();
  const planned = planConnectorRoute({ id: sheetId, drawings: [a, b, connector], drawingPayloads: payloads }, connector, payload);
  connector.transform = planned.transform;
  payloads.set(connector.payloadId, planned.payload);
  return {
    id: sheetId,
    drawings: [a, b, connector],
    drawingPayloads: payloads,
    snapSettings: structuredClone(DEFAULT_WORKSHEET_SNAP_SETTINGS),
  };
}

function payloadAt(graph: DrawingGraphSheet, payloadId: string) {
  const payloads = graph.drawingPayloads;
  return payloads instanceof Map ? payloads.get(payloadId) : (payloads as Readonly<Record<string, unknown>>)[payloadId];
}

describe('canonical drawing graph planner', () => {
  it('plans and validates a connector route from typed endpoints', () => {
    const value = graph();
    validateDrawingGraph(value);
    const connector = value.drawings.find((drawing) => drawing.kind === 'connector')!;
    const payload = payloadAt(value, connector.payloadId) as ConnectorDrawingPayload;
    assert.deepEqual(payload.route.points[0], { x: 60, y: 30 });
    assert.deepEqual(payload.route.points.at(-1), { x: 180, y: 90 });
  });

  it('rejects an unknown connector endpoint and illegal worksheet grid', () => {
    const value = graph();
    const connector = value.drawings.find((drawing) => drawing.kind === 'connector')!;
    const payload = payloadAt(value, connector.payloadId) as ConnectorDrawingPayload;
    assert.throws(() => planConnectorRoute(value, connector, { ...payload, end: { drawingId: 'missing', connectionPoint: 'left' } }), /missing/);
    assert.throws(() => validateDrawingGraph({ ...value, snapSettings: { enabled: true, snapToGrid: true, gridSize: 0 } }), /snap settings/);
  });

  it('rejects duplicate group ownership and non-canonical route data', () => {
    const value = graph();
    const connector = value.drawings.find((drawing) => drawing.kind === 'connector')!;
    const payload = payloadAt(value, connector.payloadId) as ConnectorDrawingPayload;
    assert.throws(() => validateDrawingGraph({ ...value, drawingGroups: [
      { id: 'group-a', sheetId, memberDrawingIds: ['shape-a', 'shape-b'] },
      { id: 'group-b', sheetId, memberDrawingIds: ['shape-b', 'connector-1'] },
    ] }), /multiple groups/);
    const broken = new Map(value.drawingPayloads instanceof Map ? value.drawingPayloads : Object.entries(value.drawingPayloads));
    broken.set(connector.payloadId, { ...payload, route: { points: [{ x: 0, y: 0 }, { x: 1, y: 1 }] } });
    assert.throws(() => validateDrawingGraph({ ...value, drawingPayloads: broken }), /not canonical/);
  });
});
