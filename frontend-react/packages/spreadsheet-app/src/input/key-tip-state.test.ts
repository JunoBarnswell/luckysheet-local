import assert from 'node:assert/strict';
import test from 'node:test';
import { INITIAL_KEY_TIP_STATE, keyTipTransition } from './key-tip-state';

test('KeyTips enter a tab prefix and execute a leaf command', () => {
  const started = { active: true, prefix: '' };
  const home = keyTipTransition(started, 'h');
  assert.deepEqual(home.state, { active: true, prefix: 'H' });
  assert.deepEqual(home.action, { kind: 'tab', id: 'home' });
  const paste = keyTipTransition(home.state, 'v');
  assert.deepEqual(paste.state, INITIAL_KEY_TIP_STATE);
  assert.deepEqual(paste.action, { kind: 'command', id: 'paste' });
});

test('KeyTips close on Escape and reject an unknown sequence without executing', () => {
  const active = { active: true, prefix: 'H' };
  assert.deepEqual(keyTipTransition(active, 'Escape'), { state: INITIAL_KEY_TIP_STATE });
  assert.deepEqual(keyTipTransition(active, 'z'), { state: INITIAL_KEY_TIP_STATE });
});

test('KeyTip leaf sequences are not shadowed by longer menu sequences', () => {
  const active = { active: true, prefix: 'H' };
  assert.deepEqual(keyTipTransition(active, 'c').state, { active: true, prefix: 'HC' });
  assert.deepEqual(keyTipTransition({ active: true, prefix: 'HC' }, 'p').action, { kind: 'command', id: 'copy' });
  assert.deepEqual(keyTipTransition(active, 'f').state, { active: true, prefix: 'HF' });
  assert.deepEqual(keyTipTransition({ active: true, prefix: 'HF' }, 'p').action, { kind: 'command', id: 'format-painter' });
  assert.deepEqual(keyTipTransition({ active: true, prefix: 'N' }, 's').state, { active: true, prefix: 'NS' });
  assert.deepEqual(keyTipTransition({ active: true, prefix: 'NS' }, 'c').action, { kind: 'command', id: 'camera' });
});
