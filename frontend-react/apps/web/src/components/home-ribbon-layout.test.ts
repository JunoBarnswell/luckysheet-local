import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { ribbonGroupWidthClass } from './RibbonLayoutRenderer';

describe('Home Ribbon compact group geometry', () => {
  it('keeps the late command groups dense and horizontally reachable', () => {
    assert.equal(ribbonGroupWidthClass('alignment', 'compact', 1581, 'home'), 'w-[180px]');
    assert.equal(ribbonGroupWidthClass('styles', 'compact', 1581, 'home'), 'w-[166px]');
    assert.equal(ribbonGroupWidthClass('cells', 'compact', 1581, 'home'), 'w-[118px]');
    assert.equal(ribbonGroupWidthClass('editing', 'compact', 1581, 'home'), 'w-[204px]');
    assert.equal(ribbonGroupWidthClass('editing', 'narrow', 900, 'home'), 'w-[196px]');
  });
});
