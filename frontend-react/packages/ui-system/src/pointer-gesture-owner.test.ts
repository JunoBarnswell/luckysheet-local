import test from 'node:test';
import assert from 'node:assert/strict';
import {
  claimPointerGesture,
  ownsPointerGesture,
  releasePointerGesture,
  releasePointerGesturesForSurface,
} from './pointer-gesture-owner';

function documentIdentity(): Document {
  return {} as Document;
}

test('a pointerId has exactly one gesture owner until the owner releases it', () => {
  const ownerDocument = documentIdentity();
  const worksheet = new EventTarget();
  const scrollbar = new EventTarget();

  assert.equal(claimPointerGesture(ownerDocument, 7, 'worksheet', worksheet), true);
  assert.equal(claimPointerGesture(ownerDocument, 7, 'scrollbar-vertical', scrollbar), false);
  assert.equal(ownsPointerGesture(ownerDocument, 7, 'worksheet', worksheet), true);
  assert.equal(releasePointerGesture(ownerDocument, 7, 'scrollbar-vertical', scrollbar), false);
  assert.equal(ownsPointerGesture(ownerDocument, 7, 'worksheet', worksheet), true);
  assert.equal(releasePointerGesture(ownerDocument, 7, 'worksheet', worksheet), true);
  assert.equal(claimPointerGesture(ownerDocument, 7, 'scrollbar-vertical', scrollbar), true);
});

test('pointer ownership is isolated by document and cleanup is scoped to one surface', () => {
  const leftDocument = documentIdentity();
  const rightDocument = documentIdentity();
  const leftSurface = new EventTarget();
  const otherLeftSurface = new EventTarget();
  const rightSurface = new EventTarget();

  assert.equal(claimPointerGesture(leftDocument, 1, 'worksheet', leftSurface), true);
  assert.equal(claimPointerGesture(leftDocument, 2, 'worksheet', otherLeftSurface), true);
  assert.equal(claimPointerGesture(rightDocument, 1, 'scrollbar-horizontal', rightSurface), true);

  releasePointerGesturesForSurface(leftDocument, leftSurface);
  assert.equal(ownsPointerGesture(leftDocument, 1, 'worksheet', leftSurface), false);
  assert.equal(ownsPointerGesture(leftDocument, 2, 'worksheet', otherLeftSurface), true);
  assert.equal(ownsPointerGesture(rightDocument, 1, 'scrollbar-horizontal', rightSurface), true);
});
