import assert from 'node:assert/strict';
import { existsSync, statSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { describe, it } from 'node:test';
import { HOME_RIBBON_ICON_NAMES } from './home/HomeRibbonIcon';
import { ribbonGroupWidthClass } from './RibbonLayoutRenderer';
import { ribbonLayoutModeForWidth } from '@react-sheets/ui-system';

describe('Home Ribbon responsive group geometry', () => {
  it('keeps the same wide group widths at every viewport', () => {
    assert.equal(ribbonLayoutModeForWidth(1581), 'wide');
    assert.equal(ribbonLayoutModeForWidth(1183), 'wide');
    for (const width of [1581, 1183]) {
      assert.equal(ribbonGroupWidthClass('clipboard', 'wide', width, 'home'), 'w-[141px]');
      assert.equal(ribbonGroupWidthClass('font', 'wide', width, 'home'), 'w-[347px]');
      assert.equal(ribbonGroupWidthClass('alignment', 'wide', width, 'home'), 'w-[342px]');
      assert.equal(ribbonGroupWidthClass('styles', 'wide', width, 'home'), 'w-[251px]');
      assert.equal(ribbonGroupWidthClass('cells', 'wide', width, 'home'), 'w-[191px]');
      assert.equal(ribbonGroupWidthClass('editing', 'wide', width, 'home'), 'w-[310px]');
    }
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
