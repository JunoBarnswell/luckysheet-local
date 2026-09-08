/** Explicit Node/test host; executes the same WASM artifact as the browser. */
import { readFile } from 'node:fs/promises';
import { initializeKernel } from './index';
export async function initializeNodeKernel(wasmPath: string | URL = new URL('../../../apps/web/public/kernel/kernel_host.wasm', import.meta.url)): Promise<void> {
  const bytes = await readFile(wasmPath);
  await initializeKernel({wasmBytes: Uint8Array.from(bytes)});
}

