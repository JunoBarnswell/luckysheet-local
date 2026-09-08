import assert from 'node:assert/strict';
import { existsSync, statSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { describe, it } from 'node:test';
import { HOME_RIBBON_ICON_NAMES } from './home/HomeRibbonIcon';
import { ribbonGroupWidthClass } from './RibbonLayoutRenderer';
import { ribbonLayoutModeForWidth } from '@react-sheets/ui-system';

describe('Home Ribbon responsive group geometry', () => {
  it('uses the canonical breakpoints and keeps each Home group explicit', () => {
    assert.equal(ribbonLayoutModeForWidth(1600), 'wide');
    assert.equal(ribbonLayoutModeForWidth(1581), 'compact');
    assert.equal(ribbonLayoutModeForWidth(1183), 'narrow');
    assert.equal(ribbonGroupWidthClass('clipboard', 'compact', 1581, 'home'), 'w-[141px]');
    assert.equal(ribbonGroupWidthClass('font', 'compact', 1581, 'home'), 'w-[347px]');
    assert.equal(ribbonGroupWidthClass('alignment', 'compact', 1581, 'home'), 'w-[342px]');
    assert.equal(ribbonGroupWidthClass('styles', 'narrow', 1183, 'home'), 'w-[220px]');
    assert.equal(ribbonGroupWidthClass('cells', 'narrow', 1183, 'home'), 'w-[144px]');
    assert.equal(ribbonGroupWidthClass('editing', 'narrow', 1183, 'home'), 'w-[300px]');
  });

  it('ships every Figma SVG and the exact Noto Sans SC font used by the Home design', () => {
    const publicRoot = fileURLToPath(new URL('../../public/', import.meta.url));
    for (const name of HOME_RIBBON_ICON_NAMES) {
      const path = `${publicRoot}figma/home-ribbon/${name}.svg`;
      assert.equal(existsSync(path), true, `${name} is missing`);
      assert.ok(statSync(path).size > 0, `${name} is empty`);
    }
    assert.equal(existsSync(`${publicRoot}figma/home-ribbon/divider.svg`), true);
    assert.ok(statSync(`${publicRoot}fonts/NotoSansSC-VF.ttf`).size > 0);
  });
});
