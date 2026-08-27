import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { ribbonLayoutModeForWidth } from './RibbonShell';
import { DESIGNER_GEOMETRY, RIBBON_DENSITY } from './shell-types';

describe('RibbonShell responsive layout', () => {
  it('uses Office-style density breakpoints with one reachable command viewport', () => {
    assert.equal(ribbonLayoutModeForWidth(1920), 'wide');
    assert.equal(ribbonLayoutModeForWidth(1905), 'wide');
    assert.equal(ribbonLayoutModeForWidth(1783), 'wide');
    assert.equal(ribbonLayoutModeForWidth(1782), 'compact');
    assert.equal(ribbonLayoutModeForWidth(1400), 'compact');
    assert.equal(ribbonLayoutModeForWidth(1024), 'compact');
    assert.equal(ribbonLayoutModeForWidth(900), 'narrow');
    assert.equal(ribbonLayoutModeForWidth(600), 'narrow');
  });

  it('keeps the 1920x1080 Designer vertical contract exact', () => {
    assert.equal(DESIGNER_GEOMETRY.ribbonHeight + DESIGNER_GEOMETRY.formulaBarHeight + DESIGNER_GEOMETRY.workspaceHeight + DESIGNER_GEOMETRY.statusBarHeight, 1080);
    assert.equal(DESIGNER_GEOMETRY.ribbonTabHeight + DESIGNER_GEOMETRY.ribbonContentHeight, 167);
    assert.deepEqual(RIBBON_DENSITY, {
      shellHeight: 133,
      tabStripHeight: 32,
      commandAreaHeight: 101,
      groupContentHeight: 93,
      largeCommandHeight: 72,
      groupCaptionHeight: 12,
    });
  });
});
