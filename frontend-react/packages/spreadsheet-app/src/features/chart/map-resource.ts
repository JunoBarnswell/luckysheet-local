import type { ChartMapCoordinate, ChartMapFeature, ChartMapResource, ChartMapRing } from '@react-sheets/core-model';

type JsonValue = null | boolean | number | string | JsonValue[] | { [key: string]: JsonValue };
type JsonObject = { [key: string]: JsonValue };

const MAX_FEATURES = 10_000;
const MAX_COORDINATES = 500_000;

function objectValue(value: JsonValue): JsonObject | undefined {
  return value !== null && typeof value === 'object' && !Array.isArray(value) ? value : undefined;
}

function textValue(value: JsonValue | undefined): string | undefined {
  if (typeof value === 'string' && value.trim()) return value.trim();
  if (typeof value === 'number' && Number.isFinite(value)) return String(value);
  return undefined;
}

function coordinate(value: JsonValue, label: string): ChartMapCoordinate {
  if (!Array.isArray(value) || value.length < 2 || typeof value[0] !== 'number' || typeof value[1] !== 'number'
    || !Number.isFinite(value[0]) || !Number.isFinite(value[1]) || value[0] < -180 || value[0] > 180 || value[1] < -90 || value[1] > 90) {
    throw new Error(`INVALID_MAP_RESOURCE: ${label} must be a finite longitude/latitude pair`);
  }
  return [value[0], value[1]];
}

function ring(value: JsonValue, label: string, coordinateCount: { value: number }): ChartMapRing {
  if (!Array.isArray(value) || value.length < 3) throw new Error(`INVALID_MAP_RESOURCE: ${label} must contain at least three points`);
  const points = value.map((entry, index) => {
    coordinateCount.value += 1;
    if (coordinateCount.value > MAX_COORDINATES) throw new Error(`INVALID_MAP_RESOURCE: geometry exceeds ${MAX_COORDINATES} coordinates`);
    return coordinate(entry, `${label}[${index}]`);
  });
  const first = points[0]!;
  const last = points[points.length - 1]!;
  if (first[0] !== last[0] || first[1] !== last[1]) points.push([first[0], first[1]]);
  return points;
}

function polygons(geometry: JsonObject, label: string, coordinateCount: { value: number }): readonly ChartMapRing[] {
  const type = geometry.type;
  if (type === 'Polygon') {
    if (!Array.isArray(geometry.coordinates)) throw new Error(`INVALID_MAP_RESOURCE: ${label} Polygon coordinates are missing`);
    return geometry.coordinates.map((entry, index) => ring(entry, `${label}.coordinates[${index}]`, coordinateCount));
  }
  if (type === 'MultiPolygon') {
    if (!Array.isArray(geometry.coordinates)) throw new Error(`INVALID_MAP_RESOURCE: ${label} MultiPolygon coordinates are missing`);
    return geometry.coordinates.flatMap((polygonValue, polygonIndex) => {
      if (!Array.isArray(polygonValue)) throw new Error(`INVALID_MAP_RESOURCE: ${label}.coordinates[${polygonIndex}] is invalid`);
      return polygonValue.map((entry, ringIndex) => ring(entry, `${label}.coordinates[${polygonIndex}][${ringIndex}]`, coordinateCount));
    });
  }
  throw new Error(`UNSUPPORTED_FEATURE: GeoJSON geometry ${String(type ?? 'missing')} is not Polygon or MultiPolygon`);
}

function feature(value: JsonValue, index: number, coordinateCount: { value: number }): ChartMapFeature {
  const record = objectValue(value);
  if (!record || record.type !== 'Feature') throw new Error(`INVALID_MAP_RESOURCE: features[${index}] is not a GeoJSON Feature`);
  const geometry = objectValue(record.geometry ?? null);
  if (!geometry) throw new Error(`INVALID_MAP_RESOURCE: features[${index}] has no geometry`);
  const properties = objectValue(record.properties ?? null);
  const id = textValue(record.id) ?? textValue(properties?.id) ?? textValue(properties?.ID) ?? textValue(properties?.code) ?? textValue(properties?.CODE);
  const label = textValue(properties?.name) ?? textValue(properties?.NAME) ?? textValue(properties?.label) ?? id;
  if (!id || !label) throw new Error(`INVALID_MAP_RESOURCE: features[${index}] requires a stable id or label`);
  return { id, label, polygons: polygons(geometry, `features[${index}]`, coordinateCount) };
}

async function checksum(text: string): Promise<string> {
  if (typeof crypto === 'undefined' || !crypto.subtle) throw new Error('MAP_RESOURCE_CHECKSUM_UNAVAILABLE: Web Crypto is required');
  const bytes = new TextEncoder().encode(text);
  const digest = await crypto.subtle.digest('SHA-256', bytes);
  return Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, '0')).join('');
}

/** Parse and normalize a GeoJSON FeatureCollection for offline map rendering. */
export async function parseGeoJsonMapResource(text: string, fileName = 'map.geojson'): Promise<ChartMapResource> {
  if (!text.trim()) throw new Error('INVALID_MAP_RESOURCE: GeoJSON file is empty');
  let parsed: JsonValue;
  try {
    parsed = JSON.parse(text) as JsonValue;
  } catch {
    throw new Error(`INVALID_MAP_RESOURCE: ${fileName} is not valid JSON`);
  }
  const root = objectValue(parsed);
  if (!root || root.type !== 'FeatureCollection' || !Array.isArray(root.features)) {
    throw new Error('INVALID_MAP_RESOURCE: expected a GeoJSON FeatureCollection');
  }
  if (root.features.length === 0 || root.features.length > MAX_FEATURES) {
    throw new Error(`INVALID_MAP_RESOURCE: feature count must be between 1 and ${MAX_FEATURES}`);
  }
  const coordinateCount = { value: 0 };
  const features = root.features.map((entry, index) => feature(entry, index, coordinateCount));
  const ids = new Set<string>();
  for (const entry of features) {
    if (ids.has(entry.id)) throw new Error(`INVALID_MAP_RESOURCE: duplicate feature id ${entry.id}`);
    ids.add(entry.id);
  }
  const contentHash = await checksum(text);
  return {
    schema: 'ChartMapResource',
    resourceId: `geojson-${contentHash.slice(0, 16)}`,
    source: 'geojson',
    checksum: contentHash,
    features,
  };
}
