import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import enUS from '../../locales/en-US.json';
import zhCN from '../../locales/zh-CN.json';
import { HOME_CELLS_ACTIONS, HOME_NUMBER_FORMAT_OPTIONS } from './home-localization';

describe('HOME responsive localization catalogs', () => {
  it('keeps canonical number-format values stable while localizing every label', () => {
    assert.deepEqual(HOME_NUMBER_FORMAT_OPTIONS.map(({ value }) => value), ['general', '$#,##0', '0%', '#,##0', '0.00']);
    for (const { value, labelKey } of HOME_NUMBER_FORMAT_OPTIONS) {
      assert.notEqual(enUS.homeUi[labelKey], undefined, `${value} is missing from en-US`);
      assert.notEqual(zhCN.homeUi[labelKey], undefined, `${value} is missing from zh-CN`);
    }
  });

  it('uses one complete localized descriptor for compact and wide Cells actions', () => {
    assert.deepEqual(HOME_CELLS_ACTIONS.map(({ id }) => id), ['columnWidth', 'autoFitColumnWidth', 'hideColumns', 'unhideColumns', 'defaultColumnWidth']);
    for (const { id, labelKey } of HOME_CELLS_ACTIONS) {
      assert.notEqual(enUS.homeUi[labelKey], undefined, `${id} is missing from en-US`);
      assert.notEqual(zhCN.homeUi[labelKey], undefined, `${id} is missing from zh-CN`);
    }
    assert.equal(enUS.homeUi.cells, 'Cells');
    assert.equal(zhCN.homeUi.cells, '单元格');
  });

  it('does not leak mixed-language HOME copy from the English responsive path', () => {
    for (const { labelKey } of HOME_CELLS_ACTIONS) assert.doesNotMatch(enUS.homeUi[labelKey], /[\u3400-\u9fff]/, labelKey);
    for (const { labelKey } of HOME_NUMBER_FORMAT_OPTIONS) assert.doesNotMatch(enUS.homeUi[labelKey], /[\u3400-\u9fff]/, labelKey);
  });
});
