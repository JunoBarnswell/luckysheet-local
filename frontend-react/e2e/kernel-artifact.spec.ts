import { expect, test } from '@playwright/test';

test('serves the canonical kernel manifest and raw wasm artifact', async ({ page }) => {
  const manifestResponse = await page.request.get('/kernel/kernel-manifest.json');
  expect(manifestResponse.ok()).toBeTruthy();
  const manifest = await manifestResponse.json() as { artifact: string; bytes: number; sha256: string };
  expect(manifest.artifact).toBe('kernel_host.wasm');
  expect(manifest.bytes).toBeGreaterThan(0);
  expect(manifest.sha256).toMatch(/^[a-f0-9]{64}$/);
  const wasmResponse = await page.request.get(`/kernel/${manifest.artifact}`);
  expect(wasmResponse.ok()).toBeTruthy();
  expect((await wasmResponse.body()).byteLength).toBe(manifest.bytes);
});
