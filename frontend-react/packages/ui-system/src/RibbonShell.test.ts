import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { ribbonLayoutModeForWidth } from './RibbonShell';
import { DESIGNER_GEOMETRY } from './shell-types';

describe('RibbonShell responsive layout', () => {
  it('uses Office-style density breakpoints without a horizontal overflow mode', () => {
    assert.equal(ribbonLayoutModeForWidth(1400), 'wide');
    assert.equal(ribbonLayoutModeForWidth(900), 'compact');
    assert.equal(ribbonLayoutModeForWidth(600), 'narrow');
  });

  it('keeps the 1280x720 Designer vertical contract exact', () => {
    assert.equal(DESIGNER_GEOMETRY.ribbonHeight + DESIGNER_GEOMETRY.formulaBarHeight + DESIGNER_GEOMETRY.workspaceHeight + DESIGNER_GEOMETRY.statusBarHeight, 720);
    assert.equal(DESIGNER_GEOMETRY.ribbonTabHeight + DESIGNER_GEOMETRY.ribbonContentHeight, 142);
  });
});
