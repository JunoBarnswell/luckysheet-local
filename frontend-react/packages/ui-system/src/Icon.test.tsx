import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { renderToStaticMarkup } from 'react-dom/server';
import { Icon, type IconName } from './Icon';

const ribbonIconNames: readonly IconName[] = [
  'scissors', 'copy', 'clipboard', 'palette', 'align-left', 'align-center', 'align-right', 'align-top', 'align-middle', 'align-bottom',
  'borders', 'dollar-sign', 'percent', 'comma', 'decimal-increase', 'decimal-decrease', 'filter', 'search', 'trash',
  'table-sheet', 'gantt-sheet', 'report-sheet', 'table', 'table-pivot', 'chart-column', 'barcode', 'sparkline',
  'picture', 'shape-square', 'camera', 'form-control', 'link', 'checkbox', 'textbox',
];

describe('Ribbon SVG icon coverage', () => {
  it('renders an independent SVG path for every visible Home/Insert icon', () => {
    for (const name of ribbonIconNames) {
      const markup = renderToStaticMarkup(<Icon name={name} size="md" />);
      assert.match(markup, /^<svg\b/, name);
      assert.match(markup, /<(?:path|rect|circle|line|ellipse)\b/, name);
      assert.doesNotMatch(markup, /undefined|null/, name);
    }
  });
});
