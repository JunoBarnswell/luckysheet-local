import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import enUS from '../locales/en-US.json';
import zhCN from '../locales/zh-CN.json';
import { INSERT_CHART_VARIANTS, INSERT_FORM_CONTROL_VARIANTS, INSERT_SHAPE_GALLERY, INSERT_SPARKLINE_VARIANTS } from './insert-ribbon-catalog';
import { getRibbonSurfaces } from '@react-sheets/spreadsheet-app';

const variants = [
  ...INSERT_CHART_VARIANTS,
  ...INSERT_SPARKLINE_VARIANTS,
  ...INSERT_SHAPE_GALLERY.flatMap((category) => category.variants),
  ...INSERT_FORM_CONTROL_VARIANTS,
];

describe('INSERT variant catalog localization', () => {
  it('has unique stable IDs and complete locale entries for every payload', () => {
    const ids = variants.map((variant) => variant.id);
    assert.equal(new Set(ids).size, ids.length);
    for (const variant of variants) {
      assert.notEqual(enUS.insertUi[variant.labelKey], undefined, `${variant.id} is missing from en-US`);
      assert.notEqual(zhCN.insertUi[variant.labelKey], undefined, `${variant.id} is missing from zh-CN`);
      assert.notEqual(enUS.insertUi[variant.ariaLabelKey], undefined, `${variant.id} aria label is missing from en-US`);
      assert.notEqual(zhCN.insertUi[variant.ariaLabelKey], undefined, `${variant.id} aria label is missing from zh-CN`);
      assert.notEqual(enUS.insertUi[variant.tooltipKey], undefined, `${variant.id} tooltip is missing from en-US`);
      assert.notEqual(zhCN.insertUi[variant.tooltipKey], undefined, `${variant.id} tooltip is missing from zh-CN`);
    }
  });

  it('does not expose internal enum identifiers as product copy', () => {
    const internalIdentifiers = /^(rounded-rectangle|spin-button|list-box|combo-box|option-button|group-box|win-loss)$/;
    for (const variant of variants) {
      assert.doesNotMatch(enUS.insertUi[variant.labelKey], internalIdentifiers, variant.id);
      assert.doesNotMatch(zhCN.insertUi[variant.labelKey], internalIdentifiers, variant.id);
    }
  });

  it('keeps English INSERT copy free of Chinese characters', () => {
    for (const variant of variants) assert.doesNotMatch(enUS.insertUi[variant.labelKey], /[\u3400-\u9fff]/, variant.id);
  });

  it('keeps the shape gallery categorized and limited to renderer-backed identities', () => {
    assert.deepEqual(INSERT_SHAPE_GALLERY.map((category) => category.id), ['basic-shapes', 'lines', 'callouts-and-stars']);
    assert.deepEqual(INSERT_SHAPE_GALLERY.flatMap((category) => category.variants.map((variant) => variant.value)), [
      'rectangle', 'rounded-rectangle', 'ellipse', 'line', 'arrow', 'callout', 'star',
    ]);
  });

  it('keeps every semantic INSERT group and surface reachable at every responsive breakpoint', () => {
    const groups = ['tables', 'illustrations', 'controls', 'charts', 'sparklines', 'filters', 'links', 'insertComments', 'text', 'symbols'] as const;
    const breakpoints = ['wide', 'compact', 'narrow'] as const;
    for (const group of groups) {
      const byBreakpoint = breakpoints.map((breakpoint) => getRibbonSurfaces('insert', group, breakpoint).map((surface) => surface.id));
      assert.ok(byBreakpoint.every((ids) => ids.length > 0), `${group} has no surface at one breakpoint`);
      assert.deepEqual(new Set(byBreakpoint[0]), new Set(byBreakpoint[1]), `${group} compact surface drift`);
      assert.deepEqual(new Set(byBreakpoint[0]), new Set(byBreakpoint[2]), `${group} narrow surface drift`);
    }
  });
});
