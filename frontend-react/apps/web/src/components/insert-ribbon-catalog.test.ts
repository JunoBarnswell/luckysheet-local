import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import enUS from '../locales/en-US.json';
import zhCN from '../locales/zh-CN.json';
import { INSERT_CHART_VARIANTS, INSERT_FORM_CONTROL_VARIANTS, INSERT_SHAPE_VARIANTS, INSERT_SPARKLINE_VARIANTS } from './insert-ribbon-catalog';

const variants = [
  ...INSERT_CHART_VARIANTS,
  ...INSERT_SPARKLINE_VARIANTS,
  ...INSERT_SHAPE_VARIANTS,
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
});
