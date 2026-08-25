import assert from 'node:assert/strict';
import test from 'node:test';
import {
  clampFilterPopoverPosition,
  clampFilterPopoverSize,
  filterModeOptions,
  filterOperatorsFor,
  filterScalarKey,
} from './FilterPopover';
import { filterText } from '../i18n';
import type { FilterDomainDescriptor } from '@react-sheets/sheet-features';

function descriptor(overrides: Partial<FilterDomainDescriptor> = {}): FilterDomainDescriptor {
  return {
    column: 0,
    values: ['Alpha', 'Beta'],
    scalarTypes: ['text'],
    dominantType: 'text',
    hasBlank: false,
    dateDomain: [],
    dateHierarchy: [],
    colorDomain: [],
    iconDomain: [],
    supportedFamilies: ['values', 'text'],
    ...overrides,
  };
}

test('FilterPopover exposes only the families in the resolved domain', () => {
  assert.deepEqual(filterModeOptions(descriptor()), ['values', 'text']);
  assert.deepEqual(filterModeOptions(descriptor({ dominantType: 'number', scalarTypes: ['number'], supportedFamilies: ['values', 'number'] })), ['values', 'number']);
  assert.deepEqual(filterModeOptions(descriptor({ dominantType: 'mixed', scalarTypes: ['text', 'number'], supportedFamilies: ['values'] })), ['values']);
  assert.deepEqual(filterModeOptions(descriptor({ dominantType: 'empty', scalarTypes: ['blank'], supportedFamilies: ['values'] })), ['values']);
});

test('FilterPopover operators are data-type-driven and reject text operators for numbers/dates', () => {
  assert.deepEqual(filterOperatorsFor('text'), ['equals', 'notEquals', 'contains', 'notContains', 'beginsWith', 'endsWith']);
  assert.deepEqual(filterOperatorsFor('number'), ['equals', 'notEquals', 'lessThan', 'lessThanOrEqual', 'greaterThan', 'greaterThanOrEqual']);
  assert.deepEqual(filterOperatorsFor('date'), ['equals', 'notEquals', 'lessThan', 'lessThanOrEqual', 'greaterThan', 'greaterThanOrEqual']);
  assert.equal(filterOperatorsFor('number').includes('contains'), false);
  assert.equal(filterOperatorsFor('date').includes('beginsWith'), false);
});

test('FilterPopover preserves typed value identity when labels collide', () => {
  assert.notEqual(filterScalarKey(1), filterScalarKey('1'));
  assert.notEqual(filterScalarKey(false), filterScalarKey('false'));
  assert.equal(filterScalarKey(null), 'null');
});

test('FilterPopover resize is bounded transient state and stays inside the viewport', () => {
  assert.deepEqual(clampFilterPopoverSize({ width: 100, height: 100 }, { width: 640, height: 480 }), { width: 280, height: 360 });
  assert.deepEqual(clampFilterPopoverSize({ width: 900, height: 900 }, { width: 640, height: 480 }), { width: 624, height: 464 });
  assert.deepEqual(clampFilterPopoverPosition(900, 900, { width: 320, height: 360 }, { width: 640, height: 480 }), { left: 316, top: 116 });
});

test('FilterPopover dynamic quarter labels are localized', () => {
  assert.equal(filterText('en-US', 'nextQuarter'), 'Next quarter');
  assert.equal(filterText('zh-CN', 'nextQuarter'), '下季度');
  assert.notEqual(filterText('zh-CN', 'apply'), 'Apply');
});
