import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { ribbonLayoutModeForWidth } from './RibbonShell';

describe('RibbonShell responsive layout', () => {
  it('uses Office-style density breakpoints without a horizontal overflow mode', () => {
    assert.equal(ribbonLayoutModeForWidth(1400), 'wide');
    assert.equal(ribbonLayoutModeForWidth(900), 'compact');
    assert.equal(ribbonLayoutModeForWidth(600), 'narrow');
  });
});

