import assert from 'node:assert/strict';
import { existsSync, statSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { describe, it } from 'node:test';
import { HOME_RIBBON_ICON_NAMES } from './home/HomeRibbonIcon';
import { ribbonGroupWidthClass } from './RibbonLayoutRenderer';

describe('Home Ribbon responsive group geometry', () => {
  it('keeps the wide and compact command groups reachable', () => {
    assert.equal(ribbonGroupWidthClass('clipboard', 'compact', 1581, 'home'), 'w-[141px]');
    assert.equal(ribbonGroupWidthClass('font', 'compact', 1581, 'home'), 'w-[347px]');
    assert.equal(ribbonGroupWidthClass('alignment', 'compact', 1581, 'home'), 'w-[342px]');
    assert.equal(ribbonGroupWidthClass('styles', 'compact', 1581, 'home'), 'w-[251px]');
    assert.equal(ribbonGroupWidthClass('cells', 'compact', 1581, 'home'), 'w-[191px]');
    assert.equal(ribbonGroupWidthClass('editing', 'compact', 1581, 'home'), 'w-[310px]');
    assert.equal(ribbonGroupWidthClass('clipboard', 'compact', 1183, 'home'), 'w-[124px]');
    assert.equal(ribbonGroupWidthClass('font', 'compact', 1183, 'home'), 'w-[224px]');
    assert.equal(ribbonGroupWidthClass('alignment', 'compact', 1183, 'home'), 'w-[248px]');
    assert.equal(ribbonGroupWidthClass('styles', 'compact', 1183, 'home'), 'w-[220px]');
    assert.equal(ribbonGroupWidthClass('cells', 'compact', 1183, 'home'), 'w-[144px]');
    assert.equal(ribbonGroupWidthClass('editing', 'compact', 1183, 'home'), 'w-[300px]');
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
