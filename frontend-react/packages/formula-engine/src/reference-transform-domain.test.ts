import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import test from 'node:test';
import { MAX_COLUMN_INDEX, MAX_ROW_INDEX, ReferenceTransformDomain } from './reference-transform-domain';

interface SharedReferenceTransformVectors {
  readonly schema: string;
  readonly version: number;
  readonly points: readonly {
    readonly id: string;
    readonly axis: 'row' | 'column';
    readonly position: number;
    readonly at: number;
    readonly count: number;
    readonly operation: 'insert' | 'delete';
    readonly expected: { readonly kind: string; readonly position?: number };
  }[];
  readonly intervals: readonly {
    readonly id: string;
    readonly axis: 'row' | 'column';
    readonly start: number;
    readonly end: number;
    readonly at: number;
    readonly count: number;
    readonly operation: 'insert' | 'delete';
    readonly expected: { readonly kind: string; readonly start?: number; readonly end?: number };
  }[];
}

const sharedVectors = JSON.parse(readFileSync(
  new URL('../../../../contracts/reference-transform-vectors.json', import.meta.url),
  'utf8',
)) as SharedReferenceTransformVectors;

test('maps point references through insert and delete with distinct outcomes', () => {
  assert.deepEqual(ReferenceTransformDomain.mapPoint(2, 3, 2, 1, MAX_ROW_INDEX), { kind: 'mapped', position: 2 });
  assert.deepEqual(ReferenceTransformDomain.mapPoint(3, 3, 2, 1, MAX_ROW_INDEX), { kind: 'mapped', position: 5 });
  assert.deepEqual(ReferenceTransformDomain.mapPoint(2, 3, 2, -1, MAX_ROW_INDEX), { kind: 'mapped', position: 2 });
  assert.deepEqual(ReferenceTransformDomain.mapPoint(3, 3, 2, -1, MAX_ROW_INDEX), { kind: 'deleted' });
  assert.deepEqual(ReferenceTransformDomain.mapPoint(5, 3, 2, -1, MAX_ROW_INDEX), { kind: 'mapped', position: 3 });
  assert.deepEqual(ReferenceTransformDomain.mapPoint(MAX_ROW_INDEX, MAX_ROW_INDEX, 1, 1, MAX_ROW_INDEX), {
    kind: 'out-of-bounds',
    position: MAX_ROW_INDEX + 1,
  });
});

test('maps inclusive reference intervals without mapping endpoints independently', () => {
  assert.deepEqual(ReferenceTransformDomain.mapInterval(2, 4, { axis: 'row', at: 3, count: 2, op: 'insert' }), { kind: 'mapped', start: 2, end: 6 });
  assert.deepEqual(ReferenceTransformDomain.mapInterval(3, 4, { axis: 'column', at: 3, count: 2, op: 'insert' }), { kind: 'mapped', start: 5, end: 6 });
  assert.deepEqual(ReferenceTransformDomain.mapInterval(2, 5, { axis: 'row', at: 3, count: 2, op: 'delete' }), { kind: 'mapped', start: 2, end: 3 });
  assert.deepEqual(ReferenceTransformDomain.mapInterval(3, 4, { axis: 'row', at: 3, count: 2, op: 'delete' }), { kind: 'deleted' });
});

test('rejects mapped intervals outside Excel address limits and invalid domain inputs', () => {
  assert.deepEqual(ReferenceTransformDomain.mapInterval(MAX_ROW_INDEX, MAX_ROW_INDEX, { axis: 'row', at: MAX_ROW_INDEX, count: 1, op: 'insert' }), {
    kind: 'out-of-bounds', start: MAX_ROW_INDEX + 1, end: MAX_ROW_INDEX + 1,
  });
  assert.deepEqual(ReferenceTransformDomain.mapInterval(MAX_COLUMN_INDEX, MAX_COLUMN_INDEX, { axis: 'column', at: MAX_COLUMN_INDEX, count: 1, op: 'insert' }), {
    kind: 'out-of-bounds', start: MAX_COLUMN_INDEX + 1, end: MAX_COLUMN_INDEX + 1,
  });
  assert.throws(() => ReferenceTransformDomain.mapPoint(0, -1, 1, 1, MAX_ROW_INDEX), /invalid/);
  assert.throws(() => ReferenceTransformDomain.mapInterval(0, 1, { axis: 'row', at: 0, count: 0, op: 'delete' }), /invalid/);
  assert.throws(() => ReferenceTransformDomain.mapInterval(0, 0, { axis: 'row', at: MAX_ROW_INDEX + 1, count: 1, op: 'insert' }), /bounds/);
  assert.throws(() => ReferenceTransformDomain.mapInterval(0, 0, { axis: 'row', at: MAX_ROW_INDEX, count: 2, op: 'delete' }), /bounds/);
});

test('matches the shared TypeScript and Java structural-mapping vectors', () => {
  assert.equal(sharedVectors.schema, 'ReferenceTransformVectors');
  assert.equal(sharedVectors.version, 1);
  for (const vector of sharedVectors.points) {
    const maximum = vector.axis === 'row' ? MAX_ROW_INDEX : MAX_COLUMN_INDEX;
    assert.deepEqual(
      ReferenceTransformDomain.mapPoint(vector.position, vector.at, vector.count, vector.operation === 'insert' ? 1 : -1, maximum),
      vector.expected,
      vector.id,
    );
  }
  for (const vector of sharedVectors.intervals) {
    assert.deepEqual(
      ReferenceTransformDomain.mapInterval(vector.start, vector.end, {
        axis: vector.axis,
        at: vector.at,
        count: vector.count,
        op: vector.operation,
      }),
      vector.expected,
      vector.id,
    );
  }
});
