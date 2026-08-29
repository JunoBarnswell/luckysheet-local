import assert from 'node:assert/strict';
import test from 'node:test';
import { SpreadsheetFeatureRuntime, compileFeatureSurfaceSchema } from './feature-registry';

test('feature runtime is instance scoped and advances lifecycle in dependency order', () => {
  const events: string[] = [];
  const runtime = new SpreadsheetFeatureRuntime();
  runtime.load([
    { id: 'core', version: '1', commandIds: ['cell.set'], lifecycle: { ready: () => events.push('core.ready') } },
    { id: 'feature', version: '1', dependencies: ['core'], commandIds: ['feature.run'], ribbon: [{ id: 'run', tab: 'home', group: 'editing', label: 'Run', surfaceCommandId: 'fillDown', requiredCommandId: 'feature.run', icon: 'run' }], lifecycle: { ready: () => events.push('feature.ready'), steady: () => events.push('feature.steady') } },
  ]);
  runtime.activate({ documentType: 'spreadsheet', environment: 'browser' });
  runtime.advance('ready');
  runtime.advance('rendered');
  runtime.advance('steady');
  assert.equal(runtime.getPhase(), 'steady');
  assert.deepEqual(events, ['core.ready', 'feature.ready', 'feature.steady']);
  assert.deepEqual(compileFeatureSurfaceSchema(runtime.getActiveManifests()).ribbon.map((entry) => entry.featureId), ['feature']);
  const second = new SpreadsheetFeatureRuntime();
  assert.throws(() => second.load([{ id: 'other', version: '1', dependencies: ['missing'], commandIds: [] }]), /FEATURE_DEPENDENCY_UNAVAILABLE/);
});
