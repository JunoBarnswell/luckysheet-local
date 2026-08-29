import {
  DEFAULT_WORKSHEET_SNAP_SETTINGS,
  planConnectorRoute,
  sha256Hex,
  type DrawingAnchor,
  type DrawingObject,
  type DrawingPayload,
  type DrawingTransform,
  type ConnectorDrawingPayload,
  type ImageDrawingPayload,
  type ShapeDrawingPayload,
  type TextBoxDrawingPayload,
  type DrawingConnectionPoint,
} from '@react-sheets/core-model';
import { child, children, descendants, encodeXml, localName, parseXml, serializeXml, textContent, type XmlNode } from './xml';
import type { NativeDrawingGraph, NativeDrawingNodeOwnership, NativeRelationship } from './types';

const EMU_PER_PIXEL = 9_525;
const DRAWING_NS = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const DRAWING_MAIN_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const OFFICE_REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CHART_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const REL_DRAWING_IMAGE = `${OFFICE_REL_NS}/image`;

export interface DrawingGeometryContext {
  defaultRowHeightPx: number;
  defaultColumnWidthPx: number;
  rowHeightsPx?: Readonly<Record<number, number>>;
  columnWidthsPx?: Readonly<Record<number, number>>;
}

export interface ParsedDrawingPart {
  drawings: DrawingObject[];
  drawingPayloads: Record<string, DrawingPayload>;
  ownership: NativeDrawingNodeOwnership[];
  assetPartById: Record<string, string>;
}

interface ParsedAnchor {
  anchor: DrawingAnchor;
  transform: DrawingTransform;
  node: XmlNode;
  object: XmlNode;
  nativeObjectId: number;
}

interface PendingConnector {
  parsed: ParsedAnchor;
  connectorType: ConnectorDrawingPayload['connectorType'];
  startNativeId: number;
  endNativeId: number;
  startPoint: DrawingConnectionPoint;
  endPoint: DrawingConnectionPoint;
  stroke: string;
  strokeWidth?: number;
  startArrowhead: ConnectorDrawingPayload['startArrowhead'];
  endArrowhead: ConnectorDrawingPayload['endArrowhead'];
}

interface DrawingOwnerLookup {
  nativeObjectId: number;
  drawingId?: string;
  kind: NativeDrawingNodeOwnership['kind'];
}

const SHAPE_TYPES: Record<string, ShapeDrawingPayload['type']> = {
  rect: 'rectangle',
  roundRect: 'rounded-rectangle',
  ellipse: 'ellipse',
  triangle: 'triangle',
  rtTriangle: 'right-triangle',
  diamond: 'diamond',
  parallelogram: 'parallelogram',
  trapezoid: 'trapezoid',
  hexagon: 'hexagon',
  octagon: 'octagon',
  plus: 'plus',
  homePlate: 'home-plate',
  cube: 'cube',
  can: 'cylinder',
  sun: 'sun',
  moon: 'moon',
  heart: 'heart',
  lightningBolt: 'lightning',
  cloud: 'cloud',
  frame: 'frame',
  line: 'line',
  rightArrow: 'arrow',
  leftRightArrow: 'left-right-arrow',
  upDownArrow: 'up-down-arrow',
  quadArrow: 'quad-arrow',
  bentUpArrow: 'bent-arrow',
  uturnArrow: 'u-turn-arrow',
  leftBrace: 'left-brace',
  rightBrace: 'right-brace',
  leftRightBrace: 'left-right-brace',
  leftBracket: 'left-bracket',
  rightBracket: 'right-bracket',
  leftRightBracket: 'left-right-bracket',
  wedgeRoundRectCallout: 'callout',
  wedgeRectCallout: 'wedge-rect-callout',
  cloudCallout: 'cloud-callout',
  star: 'star',
  star4: 'star4',
  star5: 'star5',
  star6: 'star6',
  star8: 'star8',
  star16: 'star16',
  explosion1: 'explosion1',
  explosion2: 'explosion2',
};

const SHAPE_PRESETS_BY_TYPE = Object.fromEntries(Object.entries(SHAPE_TYPES).map(([preset, type]) => [type, preset])) as Record<ShapeDrawingPayload['type'], string>;

export function emptyNativeDrawingGraph(): NativeDrawingGraph {
  return { schema: 'NativeDrawingGraph', nodes: [] };
}

/** Index every anchor before canonical projection so unknown nodes retain ownership. */
export function readNativeDrawingGraph(
  files: Record<string, Uint8Array>,
  relationships: Record<string, NativeRelationship[]>,
  sheetPartById: Record<string, string>,
): NativeDrawingGraph {
  const nodes: NativeDrawingNodeOwnership[] = [];
  for (const [sheetId, sheetPart] of Object.entries(sheetPartById)) {
    for (const relation of relationships[sheetPart] ?? []) {
      if (relationshipKind(relation.type) !== 'drawing') continue;
      const drawingPart = resolveTarget(sheetPart, relation.target);
      const bytes = files[drawingPart];
      if (!bytes) throw new Error(`NATIVE_DRAWING_RELATIONSHIP_MISSING: ${drawingPart}`);
      const root = firstDrawingRoot(parseXml(decodeUtf8(bytes)), drawingPart);
      for (const anchor of drawingAnchors(root)) {
        const cNvPr = descendants(anchor, 'cNvPr')[0];
        const nativeObjectId = parseNativeId(cNvPr?.attrs.id, `${drawingPart} anchor`);
        const object = drawingObjectNode(anchor);
        const kind = object ? drawingNodeKind(object) : 'unknown';
        const name = cNvPr?.attrs.name?.trim();
        const editable = kind === 'image' || kind === 'shape' || kind === 'textbox' || kind === 'connector';
        nodes.push({
          drawingPart,
          nativeObjectId,
          kind,
          ...(editable ? { drawingId: name?.startsWith('drawing-') ? name : `drawing-${sheetId}-${nativeObjectId}` } : name?.startsWith('drawing-') ? { drawingId: name } : {}),
          // Standard picture/shape/textbox/connector nodes have a complete
          // canonical projection and can be replaced as a set on save.
          // Charts and all other DrawingML nodes remain preserved-owned for
          // the chart/opaque owners.
          ownership: editable ? 'editable-owned' : 'preserved-owned',
        });
      }
    }
  }
  return { schema: 'NativeDrawingGraph', nodes };
}

export function parseNativeDrawingPart(input: {
  drawingPart: string;
  sheetId: string;
  xml: string;
  relationships: readonly NativeRelationship[];
  files: Record<string, Uint8Array>;
  geometry: DrawingGeometryContext;
  existingOwnership?: readonly NativeDrawingNodeOwnership[];
}): ParsedDrawingPart {
  const root = firstDrawingRoot(parseXml(input.xml), input.drawingPart);
  const drawings: DrawingObject[] = [];
  const drawingPayloads: Record<string, DrawingPayload> = {};
  const ownership: NativeDrawingNodeOwnership[] = [];
  const pendingConnectors: PendingConnector[] = [];
  const ownerByNativeId = new Map<number, DrawingOwnerLookup>();
  const assetPartById: Record<string, string> = {};
  const seenNativeIds = new Set<number>();

  for (const anchorNode of drawingAnchors(root)) {
    const object = drawingObjectNode(anchorNode);
    const cNvPr = descendants(anchorNode, 'cNvPr')[0];
    if (!object || !cNvPr) {
      const nativeObjectId = parseOptionalNativeId(cNvPr?.attrs.id);
      if (nativeObjectId !== undefined) ownership.push({ drawingPart: input.drawingPart, nativeObjectId, kind: 'unknown', ownership: 'preserved-owned' });
      continue;
    }
    const nativeObjectId = parseNativeId(cNvPr.attrs.id, `${input.drawingPart} anchor`);
    if (seenNativeIds.has(nativeObjectId)) throw new Error(`NATIVE_DRAWING_ID_DUPLICATE: ${input.drawingPart} cNvPr ${nativeObjectId}`);
    seenNativeIds.add(nativeObjectId);
    const parsedAnchor = parseAnchor(anchorNode, object, nativeObjectId, input.geometry, input.drawingPart);
    const kind = drawingNodeKind(object);
    if (kind === 'chart') {
      ownership.push({ drawingPart: input.drawingPart, nativeObjectId, kind, ownership: 'preserved-owned' });
      continue;
    }
    if (kind === 'unknown') {
      ownership.push({ drawingPart: input.drawingPart, nativeObjectId, kind, ownership: 'preserved-owned' });
      continue;
    }
    const drawingId = `drawing-${input.sheetId}-${nativeObjectId}`;
    const payloadId = `${drawingId}-payload`;
    if (kind === 'image') {
      const parsedImage = parseImagePayload(object, cNvPr, input, payloadId, parsedAnchor.transform);
      const payload = parsedImage.payload;
      const previousPart = assetPartById[payload.asset.assetId];
      if (previousPart && previousPart !== parsedImage.mediaPart) throw new Error(`NATIVE_ASSET_PART_CONFLICT: ${payload.asset.assetId}`);
      assetPartById[payload.asset.assetId] = parsedImage.mediaPart;
      const drawing = createDrawingObject(drawingId, input.sheetId, 'image', payloadId, parsedAnchor);
      drawings.push(drawing);
      drawingPayloads[payloadId] = payload;
      ownerByNativeId.set(nativeObjectId, { nativeObjectId, drawingId, kind });
      ownership.push({ drawingPart: input.drawingPart, nativeObjectId, kind, drawingId, ownership: 'editable-owned' });
      continue;
    }
    if (kind === 'shape' || kind === 'textbox') {
      const payload = kind === 'textbox'
        ? parseTextBoxPayload(object, payloadId)
        : parseShapePayload(object, payloadId);
      if (!payload) {
        ownership.push({ drawingPart: input.drawingPart, nativeObjectId, kind, ownership: 'preserved-owned' });
        continue;
      }
      const drawing = createDrawingObject(drawingId, input.sheetId, kind, payloadId, parsedAnchor, cNvPr.attrs.name);
      drawings.push(drawing);
      drawingPayloads[payloadId] = payload;
      ownerByNativeId.set(nativeObjectId, { nativeObjectId, drawingId, kind });
      ownership.push({ drawingPart: input.drawingPart, nativeObjectId, kind, drawingId, ownership: 'editable-owned' });
      continue;
    }
    if (kind === 'connector') {
      const connection = parseConnectorEndpoints(object, parsedAnchor, input.drawingPart);
      if (!connection) {
        ownership.push({ drawingPart: input.drawingPart, nativeObjectId, kind, ownership: 'preserved-owned' });
        continue;
      }
      pendingConnectors.push({ parsed: parsedAnchor, connectorType: connection.connectorType, startNativeId: connection.startNativeId, endNativeId: connection.endNativeId, startPoint: connection.startPoint, endPoint: connection.endPoint, stroke: connection.stroke, ...(connection.strokeWidth === undefined ? {} : { strokeWidth: connection.strokeWidth }), startArrowhead: connection.startArrowhead, endArrowhead: connection.endArrowhead });
      // The connector itself is added in the second pass after endpoint IDs
      // are resolved.  This avoids fabricating a dangling endpoint.
      continue;
    }
  }

  for (const pending of pendingConnectors) {
    const nativeObjectId = pending.parsed.nativeObjectId;
    const start = ownerByNativeId.get(pending.startNativeId);
    const end = ownerByNativeId.get(pending.endNativeId);
    if (!start?.drawingId || !end?.drawingId || start.kind === 'connector' || end.kind === 'connector') {
      ownership.push({ drawingPart: input.drawingPart, nativeObjectId, kind: 'connector', ownership: 'preserved-owned' });
      continue;
    }
    const drawingId = `drawing-${input.sheetId}-${nativeObjectId}`;
    const payloadId = `${drawingId}-payload`;
    const payload: ConnectorDrawingPayload = {
      kind: 'connector',
      connectorType: pending.connectorType,
      start: { drawingId: start.drawingId, connectionPoint: pending.startPoint },
      end: { drawingId: end.drawingId, connectionPoint: pending.endPoint },
      stroke: pending.stroke,
      ...(pending.strokeWidth === undefined ? {} : { strokeWidth: pending.strokeWidth }),
      startArrowhead: pending.startArrowhead,
      endArrowhead: pending.endArrowhead,
      route: { points: [{ x: 0, y: 0 }, { x: 1, y: 1 }] },
    };
    const connector: DrawingObject = createDrawingObject(drawingId, input.sheetId, 'connector', payloadId, pending.parsed);
    drawings.push(connector);
    drawingPayloads[payloadId] = payload;
    const graph = { id: input.sheetId, drawings, drawingPayloads, snapSettings: DEFAULT_WORKSHEET_SNAP_SETTINGS };
    const planned = planConnectorRoute(graph, connector, payload);
    connector.transform = planned.transform;
    drawingPayloads[payloadId] = planned.payload;
    ownerByNativeId.set(nativeObjectId, { nativeObjectId, drawingId, kind: 'connector' });
    ownership.push({ drawingPart: input.drawingPart, nativeObjectId, kind: 'connector', drawingId, ownership: 'editable-owned' });
  }

  // An imported package may have been indexed before its canonical names were
  // known. Merge the fresh ownership facts instead of trusting stale IDs.
  const previous = new Map((input.existingOwnership ?? []).map((entry) => [entry.nativeObjectId, entry]));
  for (const entry of ownership) {
    const existing = previous.get(entry.nativeObjectId);
    if (existing && existing.drawingPart !== entry.drawingPart) throw new Error(`NATIVE_DRAWING_OWNERSHIP_CONFLICT: ${entry.nativeObjectId}`);
  }
  return { drawings, drawingPayloads, ownership, assetPartById };
}

export interface DrawingWriterInput {
  drawingPart: string;
  originalXml?: string;
  originalRelationships: readonly NativeRelationship[];
  drawings: readonly DrawingObject[];
  drawingPayloads: Readonly<Record<string, DrawingPayload>> | ReadonlyMap<string, DrawingPayload>;
  geometry: DrawingGeometryContext;
  assetBytes?: Record<string, Uint8Array>;
  ownership?: readonly NativeDrawingNodeOwnership[];
}

export interface DrawingWriterOutput {
  xml: string;
  relationships: NativeRelationship[];
  media: Record<string, Uint8Array>;
  ownership: NativeDrawingNodeOwnership[];
}

/** Write canonical image/shape/textbox/connector nodes while preserving all unknown anchors. */
export function writeNativeDrawingPart(input: DrawingWriterInput): DrawingWriterOutput {
  const originalXml = input.originalXml ?? emptyDrawingXml();
  const root = firstDrawingRoot(parseXml(originalXml), input.drawingPart);
  const originalAnchors = drawingAnchors(root);
  const editableIds = new Set((input.ownership ?? []).filter((entry) => entry.ownership === 'editable-owned').map((entry) => entry.nativeObjectId));
  const currentIds = new Set(input.drawings.map((drawing) => drawing.id));
  const preservedChildren = root.children.filter((node) => {
    if (!isDrawingAnchor(node)) return true;
    const cNvPr = descendants(node, 'cNvPr')[0];
    const id = parseOptionalNativeId(cNvPr?.attrs.id);
    if (id === undefined || !editableIds.has(id)) return true;
    const owner = (input.ownership ?? []).find((entry) => entry.nativeObjectId === id);
    // A mapped editable node is replaced from the canonical collection.  An
    // ownership entry without a canonical identity is intentionally retained
    // because removing it would risk deleting an unknown native object.
    return !(owner?.ownership === 'editable-owned' && owner.drawingId !== undefined);
  });
  const existingNativeIds = originalAnchors.map((anchor) => parseOptionalNativeId(descendants(anchor, 'cNvPr')[0]?.attrs.id)).filter((value): value is number => value !== undefined);
  let nextNativeId = Math.max(0, ...existingNativeIds, ...[...(input.ownership ?? [])].map((entry) => entry.nativeObjectId)) + 1;
  const nativeIdByDrawingId = new Map<string, number>();
  for (const owner of input.ownership ?? []) if (owner.drawingId && currentIds.has(owner.drawingId)) nativeIdByDrawingId.set(owner.drawingId, owner.nativeObjectId);
  const media: Record<string, Uint8Array> = {};
  let relationships = input.originalRelationships.map((relation) => ({ ...relation }));
  const anchors: string[] = [];
  const ownership: NativeDrawingNodeOwnership[] = preservedChildren.flatMap((node) => {
    if (!isDrawingAnchor(node)) return [];
    const cNvPr = descendants(node, 'cNvPr')[0];
    const nativeObjectId = parseOptionalNativeId(cNvPr?.attrs.id);
    if (nativeObjectId === undefined) return [];
    const existing = (input.ownership ?? []).find((entry) => entry.nativeObjectId === nativeObjectId);
    return [existing ?? { drawingPart: input.drawingPart, nativeObjectId, kind: drawingNodeKind(drawingObjectNode(node) ?? node), ownership: 'preserved-owned' }];
  });
  const canonicalDrawings = [...input.drawings]
    .filter((drawing) => ['image', 'shape', 'textbox', 'connector'].includes(drawing.kind))
    .sort((left, right) => left.zIndex - right.zIndex || left.id.localeCompare(right.id));
  // Allocate every native identity before serializing a connector.  Z-order
  // must not decide whether an endpoint can be represented.
  for (const drawing of canonicalDrawings) {
    const nativeObjectId = nativeIdByDrawingId.get(drawing.id) ?? nextNativeId++;
    nativeIdByDrawingId.set(drawing.id, nativeObjectId);
  }
  for (const drawing of canonicalDrawings) {
    const payload = payloadFrom(input.drawingPayloads, drawing.payloadId);
    if (!payload) throw new Error(`NATIVE_DRAWING_PAYLOAD_MISSING: ${drawing.payloadId}`);
    if (payload.kind !== 'image' && payload.kind !== 'shape' && payload.kind !== 'textbox' && payload.kind !== 'connector') continue;
    const nativeObjectId = nativeIdByDrawingId.get(drawing.id)!;
    const owner = (input.ownership ?? []).find((entry) => entry.nativeObjectId === nativeObjectId && entry.drawingId === drawing.id && entry.ownership === 'editable-owned');
    if (owner) {
      const originalAnchor = originalAnchors.find((anchor) => parseOptionalNativeId(descendants(anchor, 'cNvPr')[0]?.attrs.id) === nativeObjectId);
      if (originalAnchor) assertNoUnknownDrawingChildren(originalAnchor, drawingPartForError(input.drawingPart, nativeObjectId));
    }
    let objectXml: string;
    if (payload.kind === 'image') {
      const mediaPart = mediaPartFor(payload);
      const bytes = input.assetBytes?.[payload.asset.assetId];
      if (!bytes || bytes.byteLength !== payload.asset.byteLength) throw new Error(`ASSET_EXPORT_MISSING: ${payload.asset.assetId}`);
      if (sha256Hex(bytes) !== payload.asset.contentHash) throw new Error(`ASSET_EXPORT_HASH_MISMATCH: ${payload.asset.assetId}`);
      media[mediaPart] = bytes.slice();
      relationships = addRelationship(relationships, { type: REL_DRAWING_IMAGE, target: relativeTarget(input.drawingPart, mediaPart) });
      const relationship = relationships.find((entry) => relationshipKind(entry.type) === 'image' && resolveTarget(input.drawingPart, entry.target) === mediaPart);
      if (!relationship) throw new Error(`NATIVE_DRAWING_RELATIONSHIP_MISSING: ${payload.asset.assetId}`);
      objectXml = serializePicture(drawing, payload, nativeObjectId, relationship.id);
    } else if (payload.kind === 'shape') {
      objectXml = serializeShape(drawing, payload, nativeObjectId);
    } else if (payload.kind === 'textbox') {
      objectXml = serializeTextBox(drawing, payload, nativeObjectId);
    } else {
      objectXml = serializeConnector(drawing, payload, nativeObjectId, nativeIdByDrawingId);
    }
    anchors.push(serializeAnchor(drawing, objectXml, input.geometry));
    ownership.push({ drawingPart: input.drawingPart, nativeObjectId, kind: payload.kind, drawingId: drawing.id, ownership: 'editable-owned' });
  }
  const namespaceMap = new Map<string, string>([
    ['xmlns:xdr', DRAWING_NS], ['xmlns:a', DRAWING_MAIN_NS], ['xmlns:r', OFFICE_REL_NS], ['xmlns:c', CHART_NS],
  ]);
  for (const [key, value] of Object.entries(root.attrs)) if (key === 'xmlns' || key.startsWith('xmlns:')) namespaceMap.set(key, value);
  const namespace = [...namespaceMap.entries()].map(([key, value]) => ` ${key}="${encodeXml(value)}"`).join('');
  const body = preservedChildren.map(serializeXml).join('') + anchors.join('');
  return {
    xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr${namespace}>${body}</xdr:wsDr>`,
    relationships,
    media,
    ownership,
  };
}

function drawingPartForError(part: string, nativeObjectId: number): string {
  return `${part} cNvPr ${nativeObjectId}`;
}

function assertNoUnknownDrawingChildren(anchor: XmlNode, label: string): void {
  const object = drawingObjectNode(anchor);
  const allowedAnchorChildren = new Set(['from', 'to', 'pos', 'ext', 'pic', 'sp', 'cxnSp', 'graphicFrame', 'clientData']);
  const unknownAnchorChild = anchor.children.find((node) => !allowedAnchorChildren.has(localName(node.name)));
  if (unknownAnchorChild) throw new Error(`NATIVE_DRAWING_UNKNOWN_NODE_EDIT_UNSAFE: ${label} anchor child ${localName(unknownAnchorChild.name)}`);
  if (!object) return;
  const allowedObjectChildren = localName(object.name) === 'pic'
    ? new Set(['nvPicPr', 'blipFill', 'spPr', 'style'])
    : localName(object.name) === 'cxnSp'
      ? new Set(['nvCxnSpPr', 'spPr', 'style', 'txBody'])
      : new Set(['nvSpPr', 'spPr', 'style', 'txBody']);
  const unknownObjectChild = object.children.find((node) => !allowedObjectChildren.has(localName(node.name)));
  if (unknownObjectChild) throw new Error(`NATIVE_DRAWING_UNKNOWN_NODE_EDIT_UNSAFE: ${label} object child ${localName(unknownObjectChild.name)}`);
  const extension = descendants(object, 'extLst')[0] ?? descendants(object, 'ext')[0];
  if (extension) throw new Error(`NATIVE_DRAWING_UNKNOWN_NODE_EDIT_UNSAFE: ${label} nested ${localName(extension.name)}`);
}

function parseAnchor(anchorNode: XmlNode, object: XmlNode, nativeObjectId: number, geometry: DrawingGeometryContext, part: string): ParsedAnchor {
  const kind = localName(anchorNode.name);
  if (kind === 'absoluteAnchor') {
    const pos = child(anchorNode, 'pos');
    const ext = child(anchorNode, 'ext');
    const x = emu(pos?.attrs.x, `${part} absolute x`);
    const y = emu(pos?.attrs.y, `${part} absolute y`);
    const width = emu(ext?.attrs.cx, `${part} absolute width`);
    const height = emu(ext?.attrs.cy, `${part} absolute height`);
    return { anchor: { kind: 'absolute' }, transform: positiveTransform({ x, y, width, height }, part), node: anchorNode, object, nativeObjectId };
  }
  const from = marker(child(anchorNode, 'from'), geometry, `${part} from`);
  if (kind === 'oneCellAnchor') {
    const ext = child(anchorNode, 'ext');
    const width = emu(ext?.attrs.cx, `${part} one-cell width`);
    const height = emu(ext?.attrs.cy, `${part} one-cell height`);
    return { anchor: { kind: 'one-cell', row: from.row, column: from.column }, transform: positiveTransform({ x: from.x, y: from.y, width, height }, part), node: anchorNode, object, nativeObjectId };
  }
  if (kind === 'twoCellAnchor') {
    const to = marker(child(anchorNode, 'to'), geometry, `${part} to`);
    return { anchor: { kind: 'two-cell', row: from.row, column: from.column, endRow: to.row, endColumn: to.column }, transform: positiveTransform({ x: from.x, y: from.y, width: to.x - from.x, height: to.y - from.y }, part), node: anchorNode, object, nativeObjectId };
  }
  throw new Error(`NATIVE_DRAWING_ANCHOR_UNSUPPORTED: ${part} ${kind}`);
}

function parseImagePayload(object: XmlNode, cNvPr: XmlNode, input: Parameters<typeof parseNativeDrawingPart>[0], payloadId: string, transform: DrawingTransform): { payload: ImageDrawingPayload; mediaPart: string } {
  const blip = descendants(object, 'blip')[0];
  const relationId = blip?.attrs['r:embed'] ?? blip?.attrs.embed;
  const relation = relationId ? input.relationships.find((entry) => entry.id === relationId && relationshipKind(entry.type) === 'image') : undefined;
  if (!relation) throw new Error(`NATIVE_DRAWING_IMAGE_RELATIONSHIP_MISSING: ${input.drawingPart} ${relationId ?? ''}`);
  const mediaPart = resolveTarget(input.drawingPart, relation.target);
  const bytes = input.files[mediaPart];
  if (!bytes) throw new Error(`NATIVE_DRAWING_IMAGE_RESOURCE_MISSING: ${mediaPart}`);
  const mimeType = mimeTypeForPart(mediaPart);
  const hash = sha256Hex(bytes);
  const sourceRect = descendants(object, 'srcRect')[0];
  const crop = sourceRect ? {
    left: fraction(sourceRect.attrs.l, `${payloadId} crop left`),
    top: fraction(sourceRect.attrs.t, `${payloadId} crop top`),
    right: fraction(sourceRect.attrs.r, `${payloadId} crop right`),
    bottom: fraction(sourceRect.attrs.b, `${payloadId} crop bottom`),
  } : undefined;
  return { mediaPart, payload: {
    kind: 'image',
    asset: { schema: 'AssetRef', assetId: `asset-${hash}`, contentHash: hash, mimeType, byteLength: bytes.byteLength, width: Math.max(1, Math.round(transform.width)), height: Math.max(1, Math.round(transform.height)) },
    ...(cNvPr.attrs.descr ? { altText: cNvPr.attrs.descr } : {}),
    ...(crop ? { crop } : {}),
  } };
}

function parseShapePayload(object: XmlNode, payloadId: string): ShapeDrawingPayload | undefined {
  const preset = descendants(object, 'prstGeom')[0]?.attrs.prst;
  const type = preset ? SHAPE_TYPES[preset] : undefined;
  if (!type) return undefined;
  const spPr = child(object, 'spPr') ?? descendants(object, 'spPr')[0];
  const fill = drawingColor(child(spPr, 'solidFill')) ?? '#ffffff';
  const line = child(spPr, 'ln');
  const stroke = drawingColor(child(line, 'solidFill')) ?? '#334155';
  const text = textBodyString(child(object, 'txBody'));
  const frame = textBodyFrame(child(object, 'txBody'));
  return {
    kind: 'shape', type, fill, stroke,
    ...(line?.attrs.w ? { strokeWidth: Math.max(0.1, Number(line.attrs.w) / 12_700) } : {}),
    ...(text ? { text } : {}),
    ...(frame ? { textColor: frame.textColor, fontSize: frame.fontSize, textDirection: frame.direction, textAlignment: frame.horizontalAlignment, textVerticalAlignment: frame.verticalAlignment } : {}),
  };
}

function parseTextBoxPayload(object: XmlNode, payloadId: string): TextBoxDrawingPayload | undefined {
  const body = child(object, 'txBody');
  if (!body) return undefined;
  const frame = textBodyFrame(body);
  return {
    kind: 'textbox',
    text: textBodyString(body),
    textFrame: frame ?? {
      fontFamily: 'Aptos', fontSize: 14, bold: false, italic: false, underline: false,
      textColor: '#1f2937', horizontalAlignment: 'left', verticalAlignment: 'top', direction: 'horizontal',
      margin: { top: 8, right: 8, bottom: 8, left: 8 }, wrap: true, autofit: 'none',
    },
  };
}

function parseConnectorEndpoints(object: XmlNode, parsed: ParsedAnchor, part: string): PendingConnector | { connectorType: ConnectorDrawingPayload['connectorType']; startNativeId: number; endNativeId: number; startPoint: DrawingConnectionPoint; endPoint: DrawingConnectionPoint; stroke: string; strokeWidth?: number; startArrowhead: ConnectorDrawingPayload['startArrowhead']; endArrowhead: ConnectorDrawingPayload['endArrowhead'] } | undefined {
  const cNvCxnSpPr = child(object, 'nvCxnSpPr') ? child(child(object, 'nvCxnSpPr'), 'cNvCxnSpPr') : descendants(object, 'cNvCxnSpPr')[0];
  const start = cNvCxnSpPr ? child(cNvCxnSpPr, 'stCxn') : undefined;
  const end = cNvCxnSpPr ? child(cNvCxnSpPr, 'endCxn') : undefined;
  const startNativeId = parseOptionalNativeId(start?.attrs.id);
  const endNativeId = parseOptionalNativeId(end?.attrs.id);
  if (startNativeId === undefined || endNativeId === undefined) return undefined;
  const preset = descendants(object, 'prstGeom')[0]?.attrs.prst ?? '';
  const connectorType: ConnectorDrawingPayload['connectorType'] = preset.startsWith('curvedConnector') ? 'curved' : preset.startsWith('bentConnector') ? 'elbow' : 'straight';
  const line = descendants(object, 'ln')[0];
  return { connectorType, startNativeId, endNativeId, startPoint: connectionPoint(start!.attrs.idx), endPoint: connectionPoint(end!.attrs.idx), stroke: drawingColor(child(line, 'solidFill')) ?? '#334155', ...(line?.attrs.w ? { strokeWidth: Math.max(0.1, Number(line.attrs.w) / 12_700) } : {}), startArrowhead: arrowhead(child(line, 'headEnd')?.attrs.type), endArrowhead: arrowhead(child(line, 'tailEnd')?.attrs.type) };
}

function createDrawingObject(id: string, sheetId: string, kind: DrawingObject['kind'], payloadId: string, parsed: ParsedAnchor, name?: string): DrawingObject {
  return { id, sheetId, kind, ...(name ? { name } : {}), anchor: structuredClone(parsed.anchor), transform: structuredClone(parsed.transform), zIndex: parsed.nativeObjectId, payloadId };
}

function textBodyString(body: XmlNode | undefined): string {
  if (!body) return '';
  const paragraphs = children(body, 'p');
  return paragraphs.length ? paragraphs.map((paragraph) => descendants(paragraph, 't').map(textContent).join('')).join('\n') : descendants(body, 't').map(textContent).join('');
}

function textBodyFrame(body: XmlNode | undefined): TextBoxDrawingPayload['textFrame'] | undefined {
  if (!body) return undefined;
  const bodyPr = child(body, 'bodyPr');
  const paragraph = child(body, 'p');
  const pPr = child(paragraph, 'pPr');
  const runProperties = descendants(body, 'rPr')[0];
  const textColor = drawingColor(child(runProperties, 'solidFill')) ?? '#1f2937';
  const size = runProperties?.attrs.sz ? Number(runProperties.attrs.sz) / 100 * (96 / 72) : 14;
  const alignment = pPr?.attrs.algn === 'ctr' ? 'center' : pPr?.attrs.algn === 'r' ? 'right' : 'left';
  const vertical = bodyPr?.attrs.anchor === 'ctr' ? 'middle' : bodyPr?.attrs.anchor === 'b' ? 'bottom' : 'top';
  const direction = bodyPr?.attrs.vert && bodyPr.attrs.vert !== 'horz' ? 'vertical' : 'horizontal';
  const margin = {
    left: bodyPr?.attrs.lIns ? emu(bodyPr.attrs.lIns, 'text left margin') : 8,
    right: bodyPr?.attrs.rIns ? emu(bodyPr.attrs.rIns, 'text right margin') : 8,
    top: bodyPr?.attrs.tIns ? emu(bodyPr.attrs.tIns, 'text top margin') : 8,
    bottom: bodyPr?.attrs.bIns ? emu(bodyPr.attrs.bIns, 'text bottom margin') : 8,
  };
  return {
    fontFamily: runProperties?.attrs.typeface ?? 'Aptos',
    fontSize: Number.isFinite(size) && size > 0 ? size : 14,
    bold: runProperties?.attrs.b === '1' || runProperties?.attrs.b === 'true',
    italic: runProperties?.attrs.i === '1' || runProperties?.attrs.i === 'true',
    underline: runProperties?.attrs.u !== undefined && runProperties.attrs.u !== 'none',
    textColor,
    horizontalAlignment: alignment,
    verticalAlignment: vertical,
    direction,
    margin,
    wrap: bodyPr?.attrs.wrap !== 'none',
    autofit: child(body, 'spAutoFit') ? 'resize-shape' : child(body, 'normAutofit') ? 'shrink-text' : 'none',
  };
}

function serializePicture(drawing: DrawingObject, payload: ImageDrawingPayload, nativeObjectId: number, relationshipId: string): string {
  const crop = payload.crop;
  const srcRect = crop ? `<a:srcRect l="${fractionToXml(crop.left)}" t="${fractionToXml(crop.top)}" r="${fractionToXml(crop.right)}" b="${fractionToXml(crop.bottom)}"/>` : '';
  return `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${nativeObjectId}" name="${encodeXml(drawing.id)}"${payload.altText ? ` descr="${encodeXml(payload.altText)}"` : ''}/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="${encodeXml(relationshipId)}"/>${srcRect}<a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>`;
}

function serializeShape(drawing: DrawingObject, payload: ShapeDrawingPayload, nativeObjectId: number): string {
  const text = payload.text ? serializeTextBody(payload.text, {
    fontFamily: 'Aptos', fontSize: payload.fontSize ?? 13, bold: false, italic: false, underline: false,
    textColor: payload.textColor ?? '#1e293b', horizontalAlignment: payload.textAlignment ?? 'center', verticalAlignment: payload.textVerticalAlignment ?? 'middle', direction: payload.textDirection ?? 'horizontal', margin: { top: 8, right: 8, bottom: 8, left: 8 }, wrap: true, autofit: 'none',
  }) : '';
  return `<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="${nativeObjectId}" name="${encodeXml(drawing.id)}"/><xdr:cNvSpPr/><xdr:spLocks noGrp="1"/></xdr:nvSpPr><xdr:spPr><a:solidFill>${serializeColor(payload.fill)}</a:solidFill><a:ln w="${Math.round((payload.strokeWidth ?? 1.5) * 12_700)}"><a:solidFill>${serializeColor(payload.stroke)}</a:solidFill><a:prstDash val="solid"/></a:ln><a:prstGeom prst="${encodeXml(SHAPE_PRESETS_BY_TYPE[payload.type])}"><a:avLst/></a:prstGeom></xdr:spPr>${text}</xdr:sp>`;
}

function serializeTextBox(drawing: DrawingObject, payload: TextBoxDrawingPayload, nativeObjectId: number): string {
  return `<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="${nativeObjectId}" name="${encodeXml(drawing.id)}"/><xdr:cNvSpPr txBox="1"/><xdr:spLocks noGrp="1"/></xdr:nvSpPr><xdr:spPr><a:noFill/><a:ln><a:noFill/></a:ln></xdr:spPr>${serializeTextBody(payload.text, payload.textFrame)}</xdr:sp>`;
}

function serializeConnector(drawing: DrawingObject, payload: ConnectorDrawingPayload, nativeObjectId: number, nativeIdByDrawingId: ReadonlyMap<string, number>): string {
  const startId = nativeIdByDrawingId.get(payload.start.drawingId);
  const endId = nativeIdByDrawingId.get(payload.end.drawingId);
  if (startId === undefined || endId === undefined) throw new Error(`UNSUPPORTED_FEATURE: connector endpoint ownership is unavailable for ${drawing.id}`);
  const preset = payload.connectorType === 'curved' ? 'curvedConnector2' : payload.connectorType === 'elbow' ? 'bentConnector2' : 'line';
  return `<xdr:cxnSp><xdr:nvCxnSpPr><xdr:cNvPr id="${nativeObjectId}" name="${encodeXml(drawing.id)}"/><xdr:cNvCxnSpPr><a:stCxn id="${startId}" idx="${connectionIndex(payload.start.connectionPoint)}"/><a:endCxn id="${endId}" idx="${connectionIndex(payload.end.connectionPoint)}"/></xdr:cNvCxnSpPr></xdr:nvCxnSpPr><xdr:spPr><a:ln w="${Math.round((payload.strokeWidth ?? 1.5) * 12_700)}"><a:solidFill>${serializeColor(payload.stroke)}</a:solidFill>${serializeArrow('headEnd', payload.startArrowhead)}${serializeArrow('tailEnd', payload.endArrowhead)}</a:ln><a:prstGeom prst="${preset}"><a:avLst/></a:prstGeom></xdr:spPr></xdr:cxnSp>`;
}

function serializeTextBody(text: string, frame: TextBoxDrawingPayload['textFrame']): string {
  const paragraphs = text.split(/\r?\n/).map((line) => `<a:p><a:pPr algn="${frame.horizontalAlignment === 'center' ? 'ctr' : frame.horizontalAlignment === 'right' ? 'r' : 'l'}"/><a:r><a:rPr lang="en-US" sz="${Math.max(1, Math.round(frame.fontSize * 72 / 96 * 100))}"${frame.bold ? ' b="1"' : ''}${frame.italic ? ' i="1"' : ''}${frame.underline ? ' u="sng"' : ''} typeface="${encodeXml(frame.fontFamily)}">${serializeColorNode(frame.textColor)}</a:rPr><a:t>${encodeXml(line)}</a:t></a:r></a:p>`).join('');
  const body = `<a:bodyPr wrap="${frame.wrap ? 'square' : 'none'}" anchor="${frame.verticalAlignment === 'middle' ? 'ctr' : frame.verticalAlignment === 'bottom' ? 'b' : 't'}" vert="${frame.direction === 'vertical' ? 'vert270' : 'horz'}" lIns="${Math.round(frame.margin.left * EMU_PER_PIXEL)}" rIns="${Math.round(frame.margin.right * EMU_PER_PIXEL)}" tIns="${Math.round(frame.margin.top * EMU_PER_PIXEL)}" bIns="${Math.round(frame.margin.bottom * EMU_PER_PIXEL)}"/>`;
  return `<xdr:txBody>${body}<a:lstStyle/>${paragraphs}</xdr:txBody>`;
}

function serializeAnchor(drawing: DrawingObject, objectXml: string, geometry: DrawingGeometryContext): string {
  const anchor = drawing.anchor;
  if (anchor.kind === 'two-cell' && anchor.row !== undefined && anchor.column !== undefined && anchor.endRow !== undefined && anchor.endColumn !== undefined) {
    const from = markerXml(anchor.row, anchor.column, 'from');
    const to = markerXml(anchor.endRow, anchor.endColumn, 'to');
    return `<xdr:twoCellAnchor editAs="twoCell">${from}${to}${objectXml}<xdr:clientData/></xdr:twoCellAnchor>`;
  }
  if (anchor.kind === 'one-cell' && anchor.row !== undefined && anchor.column !== undefined) {
    const from = markerXml(anchor.row, anchor.column, 'from');
    return `<xdr:oneCellAnchor>${from}<xdr:ext cx="${Math.round(Math.max(0, drawing.transform.width) * EMU_PER_PIXEL)}" cy="${Math.round(Math.max(0, drawing.transform.height) * EMU_PER_PIXEL)}"/>${objectXml}<xdr:clientData/></xdr:oneCellAnchor>`;
  }
  return `<xdr:absoluteAnchor><xdr:pos x="${Math.round(drawing.transform.x * EMU_PER_PIXEL)}" y="${Math.round(drawing.transform.y * EMU_PER_PIXEL)}"/><xdr:ext cx="${Math.round(Math.max(0, drawing.transform.width) * EMU_PER_PIXEL)}" cy="${Math.round(Math.max(0, drawing.transform.height) * EMU_PER_PIXEL)}"/>${objectXml}<xdr:clientData/></xdr:absoluteAnchor>`;
}

function markerXml(row: number, column: number, kind: 'from' | 'to'): string {
  return `<xdr:${kind}><xdr:col>${column}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:${kind}>`;
}

function marker(node: XmlNode | undefined, geometry: DrawingGeometryContext, label: string): { row: number; column: number; x: number; y: number } {
  if (!node) throw new Error(`NATIVE_DRAWING_ANCHOR_INVALID: ${label} marker is missing`);
  const column = integer(node, 'col', label);
  const row = integer(node, 'row', label);
  const colOff = emu(child(node, 'colOff')?.text || textContent(child(node, 'colOff')), `${label} colOff`);
  const rowOff = emu(child(node, 'rowOff')?.text || textContent(child(node, 'rowOff')), `${label} rowOff`);
  return { row, column, x: columnOffset(column, geometry) + colOff, y: rowOffset(row, geometry) + rowOff };
}

function columnOffset(column: number, geometry: DrawingGeometryContext): number {
  if (column < 0 || !Number.isSafeInteger(column)) throw new Error(`NATIVE_DRAWING_COORDINATE_INVALID: column ${column}`);
  let total = 0;
  for (let index = 0; index < column; index += 1) total += geometry.columnWidthsPx?.[index] ?? geometry.defaultColumnWidthPx;
  return total;
}

function rowOffset(row: number, geometry: DrawingGeometryContext): number {
  if (row < 0 || !Number.isSafeInteger(row)) throw new Error(`NATIVE_DRAWING_COORDINATE_INVALID: row ${row}`);
  let total = 0;
  for (let index = 0; index < row; index += 1) total += geometry.rowHeightsPx?.[index] ?? geometry.defaultRowHeightPx;
  return total;
}

function emu(value: string | undefined, label: string): number {
  const number = Number(value);
  if (!Number.isFinite(number) || number < 0) throw new Error(`NATIVE_DRAWING_METRIC_INVALID: ${label}`);
  return number / EMU_PER_PIXEL;
}

function positiveTransform(transform: DrawingTransform, part: string): DrawingTransform {
  if (![transform.x, transform.y, transform.width, transform.height].every(Number.isFinite) || transform.width <= 0 || transform.height <= 0) throw new Error(`NATIVE_DRAWING_TRANSFORM_INVALID: ${part}`);
  return transform;
}

function fraction(value: string | undefined, label: string): number {
  const numeric = Number(value ?? 0);
  if (!Number.isFinite(numeric) || numeric < 0 || numeric > 100000) throw new Error(`NATIVE_DRAWING_CROP_INVALID: ${label}`);
  return numeric / 100000;
}

function fractionToXml(value: number): string {
  if (!Number.isFinite(value) || value < 0 || value > 1) throw new Error('UNSUPPORTED_FEATURE: image crop must be between 0 and 1');
  return String(Math.round(value * 100000));
}

function drawingColor(node: XmlNode | undefined): string | undefined {
  if (!node) return undefined;
  const srgb = child(node, 'srgbClr');
  const value = srgb?.attrs.val;
  if (value && /^[0-9a-f]{6}$/i.test(value)) return `#${value.toLowerCase()}`;
  return undefined;
}

function serializeColor(value: string): string {
  return serializeColorNode(value);
}

function serializeColorNode(value: string): string {
  const normalized = /^#[0-9a-f]{6}$/i.test(value) ? value.slice(1) : 'FFFFFF';
  return `<a:srgbClr val="${normalized.toUpperCase()}"/>`;
}

function serializeArrow(kind: 'headEnd' | 'tailEnd', arrow: ConnectorDrawingPayload['startArrowhead']): string {
  if (arrow === 'none') return '';
  const type = arrow === 'triangle' ? 'triangle' : arrow === 'stealth' ? 'stealth' : arrow === 'diamond' ? 'diamond' : 'oval';
  return `<a:${kind} type="${type}" w="med" len="med"/>`;
}

function connectionPoint(value: string | undefined): DrawingConnectionPoint {
  switch (Number(value)) {
    case 1: return 'top';
    case 2: return 'right';
    case 3: return 'bottom';
    case 4: return 'left';
    default: return 'center';
  }
}

function connectionIndex(value: DrawingConnectionPoint): number {
  return value === 'top' ? 1 : value === 'right' ? 2 : value === 'bottom' ? 3 : value === 'left' ? 4 : 0;
}

function arrowhead(value: string | undefined): ConnectorDrawingPayload['startArrowhead'] {
  if (value === 'triangle') return 'triangle';
  if (value === 'stealth') return 'stealth';
  if (value === 'diamond') return 'diamond';
  if (value === 'oval') return 'oval';
  return 'none';
}

function mediaPartFor(payload: ImageDrawingPayload): string {
  const extension = payload.asset.mimeType === 'image/jpeg' ? 'jpg' : payload.asset.mimeType.split('/')[1]?.toLowerCase();
  if (!extension || !['png', 'jpg', 'gif', 'webp', 'bmp', 'svg'].includes(extension)) throw new Error(`UNSUPPORTED_FEATURE: XLSX image MIME type ${payload.asset.mimeType}`);
  return `xl/media/${payload.asset.assetId}.${extension}`;
}

function addRelationship(existing: NativeRelationship[], request: Pick<NativeRelationship, 'type' | 'target' | 'targetMode'>): NativeRelationship[] {
  if (existing.some((entry) => entry.type === request.type && entry.target === request.target && entry.targetMode === request.targetMode)) return existing;
  const used = new Set(existing.map((entry) => entry.id));
  let index = 1;
  while (used.has(`rId${index}`)) index += 1;
  return [...existing, { ...request, id: `rId${index}` }];
}

function payloadFrom(payloads: DrawingWriterInput['drawingPayloads'], payloadId: string): DrawingPayload | undefined {
  return typeof (payloads as ReadonlyMap<string, DrawingPayload>).get === 'function'
    ? (payloads as ReadonlyMap<string, DrawingPayload>).get(payloadId)
    : (payloads as Readonly<Record<string, DrawingPayload>>)[payloadId];
}

function drawingAnchors(root: XmlNode): XmlNode[] {
  return root.children.filter(isDrawingAnchor);
}

function isDrawingAnchor(node: XmlNode): boolean {
  return ['absoluteAnchor', 'oneCellAnchor', 'twoCellAnchor'].includes(localName(node.name));
}

function drawingObjectNode(anchor: XmlNode): XmlNode | undefined {
  return anchor.children.find((node) => ['pic', 'sp', 'cxnSp', 'graphicFrame'].includes(localName(node.name)));
}

function drawingNodeKind(object: XmlNode): NativeDrawingNodeOwnership['kind'] {
  const name = localName(object.name);
  if (name === 'pic') return 'image';
  if (name === 'cxnSp') return 'connector';
  if (name === 'graphicFrame') return descendants(object, 'chart').length ? 'chart' : 'unknown';
  if (name === 'sp') return child(child(object, 'nvSpPr'), 'cNvSpPr')?.attrs.txBox === '1' ? 'textbox' : 'shape';
  return 'unknown';
}

function parseNativeId(value: string | undefined, label: string): number {
  const parsed = parseOptionalNativeId(value);
  if (parsed === undefined) throw new Error(`NATIVE_DRAWING_ID_INVALID: ${label}`);
  return parsed;
}

function parseOptionalNativeId(value: string | undefined): number | undefined {
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) && parsed > 0 ? parsed : undefined;
}

function integer(node: XmlNode, name: string, label: string): number {
  const value = Number(node.attrs[name] ?? textContent(child(node, name)));
  if (!Number.isSafeInteger(value) || value < 0) throw new Error(`NATIVE_DRAWING_COORDINATE_INVALID: ${label} ${name}`);
  return value;
}

function relationshipKind(type: string): string {
  const normalized = type.replace(/\/+$/, '');
  return normalized.slice(normalized.lastIndexOf('/') + 1).toLowerCase();
}

function resolveTarget(source: string, target: string): string {
  if (target.startsWith('/')) return normalizePart(target.slice(1));
  const base = source.includes('/') ? source.slice(0, source.lastIndexOf('/') + 1) : '';
  return normalizePart(`${base}${target}`);
}

function relativeTarget(source: string, target: string): string {
  const sourceParts = (source.includes('/') ? source.slice(0, source.lastIndexOf('/') + 1) : '').split('/').filter(Boolean);
  const targetParts = target.split('/').filter(Boolean);
  while (sourceParts.length && targetParts.length && sourceParts[0] === targetParts[0]) { sourceParts.shift(); targetParts.shift(); }
  return `${'../'.repeat(sourceParts.length)}${targetParts.join('/')}`;
}

function normalizePart(value: string): string {
  const parts: string[] = [];
  for (const part of value.replaceAll('\\', '/').split('/')) {
    if (!part || part === '.') continue;
    if (part === '..') { if (!parts.length) throw new Error(`Unsafe drawing relationship target: ${value}`); parts.pop(); }
    else parts.push(part);
  }
  return parts.join('/');
}

function firstDrawingRoot(document: XmlNode, part: string): XmlNode {
  const root = document.name === '#document' ? document.children[0] : document;
  if (!root || localName(root.name) !== 'wsDr') throw new Error(`NATIVE_DRAWING_ROOT_INVALID: ${part}`);
  return root;
}

function emptyDrawingXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="${DRAWING_NS}" xmlns:a="${DRAWING_MAIN_NS}" xmlns:r="${OFFICE_REL_NS}"/>`;
}

function mimeTypeForPart(part: string): string {
  const extension = part.slice(part.lastIndexOf('.') + 1).toLowerCase();
  const types: Record<string, string> = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', webp: 'image/webp', bmp: 'image/bmp', svg: 'image/svg+xml' };
  const mime = types[extension];
  if (!mime) throw new Error(`UNSUPPORTED_FEATURE: native drawing image type .${extension}`);
  return mime;
}

function decodeUtf8(bytes: Uint8Array): string {
  return new TextDecoder().decode(bytes);
}
