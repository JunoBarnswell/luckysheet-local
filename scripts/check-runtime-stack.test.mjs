import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { inspectRuntimeStack } from './check-runtime-stack.mjs';

function fixture(t) {
  const root = mkdtempSync(join(tmpdir(), 'react-sheets-stack-'));
  t.after(() => rmSync(root, { recursive: true, force: true }));
  const put = (name, text = '') => { const path = join(root, name); mkdirSync(dirname(path), { recursive: true }); writeFileSync(path, text); };
  put('frontend-react/package.json', JSON.stringify({ engines: { node: '24.x' } }));
  put('backend/pom.xml', '<project><properties><java.version>21</java.version></properties></project>');
  put('backend/src/main/java/com/xc/luckysheet/server/service/WorkbookOperationService.java');
  put('frontend-react/packages/core-model/src/domain.ts');
  put('frontend-react/packages/exchange-excel-ooxml/src/ooxml.ts');
  return { root, put };
}

test('accepts the configured Java and browser owners without native runtime', t => {
  const { root } = fixture(t);
  assert.deepEqual(inspectRuntimeStack(root), []);
});

test('rejects source, generated WASM and workflow dependencies together', t => {
  const { root, put } = fixture(t);
  put('Cargo.toml', '[workspace]');
  put('frontend-react/apps/web/public/kernel/kernel_host.wasm', 'binary');
  put('frontend-react/packages/kernel-client/src/index.ts', 'export {};');
  put('.github/workflows/verification.yml', 'run: cargo build --release');
  const errors = inspectRuntimeStack(root);
  assert.ok(errors.some(error => error.includes('Cargo.toml')));
  assert.ok(errors.some(error => error.includes('kernel_host.wasm')));
  assert.ok(errors.some(error => error.includes('kernel-client')));
  assert.ok(errors.some(error => error.includes('verification.yml')));
});

test('rejects hidden native host invocation even without a Cargo workspace', t => {
  const { root, put } = fixture(t);
  put('backend/src/main/java/Host.java', 'new ProcessBuilder("workbook-kernel-host");');
  assert.ok(inspectRuntimeStack(root).some(error => error.includes('Host.java')));
});

test('does not allow deleting the canonical owners or changing toolchain to pass', t => {
  const { root, put } = fixture(t);
  rmSync(join(root, 'frontend-react/packages/core-model/src/domain.ts'));
  put('backend/pom.xml', '<java.version>17</java.version>');
  put('frontend-react/package.json', JSON.stringify({ engines: { node: '22.x' } }));
  const errors = inspectRuntimeStack(root);
  assert.equal(errors.length, 3);
});
