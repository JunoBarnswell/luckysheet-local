import assert from 'node:assert/strict';
import test from 'node:test';
import { buildExcelParityReport, createExcelFeatureRegistry, validateExcelFeatureRegistry } from './excel-parity';

test('Excel feature registry has one canonical ribbon surface graph', () => {
  const registry = createExcelFeatureRegistry();
  assert.deepEqual(validateExcelFeatureRegistry(registry), []);
  assert.ok(registry.ribbonCommands.length > 0);
  assert.ok(registry.ribbonSurfaces.length > 0);
  assert.ok(registry.shortcutBindings.some((binding) => binding.id === 'print.preview'));
  assert.deepEqual(registry.shortcutSequenceBindings.find((binding) => binding.id === 'clipboard.pasteSpecial.legacy')?.chords, [{ key: 'e', alt: true }, { key: 's' }]);
  assert.equal(registry.shortcutBindings.some((binding) => binding.id === 'commandPalette.open' && binding.key === 'p' && binding.primary && !binding.shift), false);
});

test('Excel parity report keeps failures and host preservation observable', () => {
  const report = buildExcelParityReport();
  assert.equal(report.schema, 'ExcelParityReport');
  assert.ok(report.total > 0);
  assert.equal(report.failures.some((item) => item.id === 'shortcut.print'), false);
  assert.ok(report.failures.some((item) => item.scope === 'visual'));
  assert.ok(report.byScope.object.preserveOnly > 0);
});

test('Excel parity report enforces the Issue #317 delivery thresholds', () => {
  const report = buildExcelParityReport();
  assert.ok(report.coreParity >= 0.95);
  assert.equal(report.homeVisibleCoverage, 1);
  assert.equal(report.insertVisibleCoverage, 1);
  assert.equal(report.officialShortcutCatalogCoverage, 1);
  assert.equal(report.nativeSilentLoss, 0);
});
