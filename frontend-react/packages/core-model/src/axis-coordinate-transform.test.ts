import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { mapAxisCoordinate } from './axis-coordinate-transform';

describe('mapAxisCoordinate', () => {
  it('maps coordinates at and after an insertion point', () => {
    assert.equal(mapAxisCoordinate(2, 3, 2, 1), 2);
    assert.equal(mapAxisCoordinate(3, 3, 2, 1), 5);
    assert.equal(mapAxisCoordinate(8, 3, 2, 1), 10);
  });

  it('removes coordinates in a deleted interval and closes the gap after it', () => {
    assert.equal(mapAxisCoordinate(2, 3, 2, -1), 2);
    assert.equal(mapAxisCoordinate(3, 3, 2, -1), null);
    assert.equal(mapAxisCoordinate(4, 3, 2, -1), null);
    assert.equal(mapAxisCoordinate(5, 3, 2, -1), 3);
  });
});
