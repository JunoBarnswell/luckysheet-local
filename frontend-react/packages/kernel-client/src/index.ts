/** The only browser/WASM host binding. Domain algorithms live in Rust. */
export const KERNEL_PROTOCOL_VERSION = 1 as const;
export const KERNEL_MANIFEST_VERSION = 11 as const;
const MAX_FRAME_BYTES = 16 * 1024 * 1024;
export interface KernelErrorDetail { code: string; message: string; object?: string | null; recovery: string; }
export class KernelInvocationError extends Error {
  readonly code: string; readonly object?: string | null; readonly recovery: string;
  constructor(detail: KernelErrorDetail) { super(detail.message); this.name = 'KernelInvocationError'; this.code = detail.code; this.object = detail.object; this.recovery = detail.recovery; }
}
export interface KernelBuildManifest {
  readonly schema: 'react-sheets.kernel-build.v1';
  readonly artifact: 'kernel_host.wasm';
  readonly bytes: number;
  readonly expectedSha256: string;
}
interface KernelExports extends WebAssembly.Exports {
  memory: WebAssembly.Memory;
  kernel_alloc(length: number): number;
  kernel_free(pointer: number, length: number): void;
  kernel_invoke(pointer: number, length: number): number;
  kernel_result_ptr(): number;
  kernel_result_len(): number;
}
let exports: KernelExports | null = null;
let initialization: Promise<void> | null = null;
let sequence = 0;
const encoder = new TextEncoder();
const decoder = new TextDecoder('utf-8', { fatal: true });
export interface KernelInitializationOptions { wasmBytes?: BufferSource; wasmUrl?: string | URL; expectedSha256?: string; }
function failure(code: string, message: string, recovery = 'reload-kernel'): KernelInvocationError { return new KernelInvocationError({code,message,recovery}); }
const SHA256_PATTERN = /^[a-f0-9]{64}$/;

export function parseKernelBuildManifest(value: unknown): KernelBuildManifest {
  if (!value || typeof value !== 'object') throw failure('KERNEL_MANIFEST_INVALID', 'Kernel build manifest is not an object.');
  const manifest = value as Record<string, unknown>;
  if (manifest.schema !== 'react-sheets.kernel-build.v1'
    || manifest.artifact !== 'kernel_host.wasm'
    || !Number.isSafeInteger(manifest.bytes) || (manifest.bytes as number) <= 0
    || typeof manifest.expectedSha256 !== 'string'
    || !SHA256_PATTERN.test(manifest.expectedSha256)) {
    throw failure('KERNEL_MANIFEST_INVALID', 'Kernel build manifest does not match the supported build-bound contract.');
  }
  return {
    schema: 'react-sheets.kernel-build.v1',
    artifact: 'kernel_host.wasm',
    bytes: manifest.bytes as number,
    expectedSha256: manifest.expectedSha256,
  };
}

export function versionedKernelWasmUrl(manifestUrl: string | URL, manifest: KernelBuildManifest): URL {
  const wasmUrl = new URL(manifest.artifact, manifestUrl);
  wasmUrl.searchParams.set('sha256', manifest.expectedSha256);
  return wasmUrl;
}

export async function initializeKernelFromManifest(manifestResource: string | URL): Promise<void> {
  const manifestUrl = new URL(manifestResource, typeof window === 'undefined' ? undefined : window.location.href);
  const buildId = manifestUrl.searchParams.get('build');
  if (!buildId || !SHA256_PATTERN.test(buildId)) throw failure('KERNEL_BUILD_CONTRACT_REQUIRED', 'Kernel manifest URL is missing its build binding.');
  const response = await fetch(manifestUrl, { cache: 'no-store' });
  if (!response.ok) throw failure('KERNEL_MANIFEST_UNAVAILABLE', `Unable to load kernel build manifest (${response.status}).`);
  let value: unknown;
  try { value = await response.json(); } catch { throw failure('KERNEL_MANIFEST_INVALID', 'Kernel build manifest is not valid JSON.'); }
  const manifest = parseKernelBuildManifest(value);
  if (manifest.expectedSha256 !== buildId) throw failure('KERNEL_BUILD_CONTRACT_MISMATCH', 'Kernel build manifest does not match the running frontend build.');
  await initializeKernel({ wasmUrl: versionedKernelWasmUrl(manifestUrl, manifest), expectedSha256: manifest.expectedSha256 });
}
export function isKernelReady(): boolean { return exports !== null; }
export async function initializeKernel(options: KernelInitializationOptions = {}): Promise<void> {
  if (exports) return;
  if (initialization) return initialization;
  initialization = (async () => {
    let bytes = options.wasmBytes;
    if (!bytes) {
      if (typeof window === 'undefined' && !options.wasmUrl) throw failure('KERNEL_INITIALIZATION_REQUIRED','Non-browser hosts must explicitly supply compiled WASM bytes.');
      if (!options.wasmUrl || !options.expectedSha256 || !SHA256_PATTERN.test(options.expectedSha256.toLowerCase())) {
        throw failure('KERNEL_BUILD_CONTRACT_REQUIRED', 'Browser kernel initialization requires a manifest-bound WASM URL and expectedSha256.');
      }
      const wasmUrl = new URL(options.wasmUrl, typeof window === 'undefined' ? undefined : window.location.href);
      if (wasmUrl.searchParams.get('sha256')?.toLowerCase() !== options.expectedSha256.toLowerCase()) {
        throw failure('KERNEL_BUILD_CONTRACT_MISMATCH', 'Kernel WASM URL is not bound to the expectedSha256 from the build manifest.');
      }
      const response = await fetch(wasmUrl, {cache:'no-store'});
      if (!response.ok) throw failure('KERNEL_UNAVAILABLE', 'Unable to load spreadsheet kernel (' + response.status + ').');
      bytes = await response.arrayBuffer();
    }
    if (options.expectedSha256) {
      const digest = await crypto.subtle.digest('SHA-256', bytes);
      const actual = Array.from(new Uint8Array(digest), value => value.toString(16).padStart(2,'0')).join('');
      if (actual !== options.expectedSha256.toLowerCase()) throw failure('KERNEL_CHECKSUM_MISMATCH','Compiled kernel checksum does not match the deployment manifest.');
    }
    const compiled = await WebAssembly.instantiate(bytes, { env: { analytics_now_ms: () => performance.now() } });
    const candidate = compiled.instance.exports as KernelExports;
    for (const name of ['kernel_alloc','kernel_free','kernel_invoke','kernel_result_ptr','kernel_result_len']) {
      if (typeof candidate[name] !== 'function') throw failure('KERNEL_ABI_MISMATCH','Missing kernel export: ' + name);
    }
    if (!(candidate.memory instanceof WebAssembly.Memory)) throw failure('KERNEL_ABI_MISMATCH','Kernel memory export is unavailable.');
    exports = candidate;
    try {
      const result = kernelInvoke<{protocolVersion:number;manifestVersion:number}>('init',{});
      if (result.protocolVersion !== KERNEL_PROTOCOL_VERSION || result.manifestVersion !== KERNEL_MANIFEST_VERSION) throw failure('KERNEL_VERSION_MISMATCH','Kernel protocol or manifest version is unsupported.');
    } catch (cause) { exports = null; throw cause; }
  })();
  try { await initialization; } catch (cause) { initialization = null; throw cause; }
}
/** A synchronous call only after host startup has awaited initializeKernel(). */
export function kernelInvoke<T = unknown>(operation: string, params: unknown): T {
  const runtime = exports;
  if (!runtime) throw failure('KERNEL_INITIALIZATION_REQUIRED','Spreadsheet kernel has not been initialized.','initialize-kernel');
  const requestId = String(++sequence);
  let bytes: Uint8Array;
  try {
    bytes = encoder.encode(JSON.stringify({protocolVersion:KERNEL_PROTOCOL_VERSION,requestId,operation,params}, (_key,value:unknown) => {
      if (typeof value === 'number' && !Number.isFinite(value)) throw failure('KERNEL_PAYLOAD_INVALID','Non-finite numbers cannot cross the kernel boundary.','correct-input');
      return value;
    }));
  } catch (cause) { if (cause instanceof KernelInvocationError) throw cause; throw failure('KERNEL_PAYLOAD_INVALID','Kernel request is not serializable.','correct-input'); }
  if (bytes.byteLength > MAX_FRAME_BYTES) throw failure('KERNEL_PAYLOAD_TOO_LARGE','Kernel control request exceeds 16 MiB.','use-data-pages');
  const pointer = runtime.kernel_alloc(bytes.byteLength);
  if (!pointer) throw failure('KERNEL_MEMORY_LIMIT','Kernel could not allocate request memory.','reduce-request');
  let text: string;
  try {
    new Uint8Array(runtime.memory.buffer,pointer,bytes.byteLength).set(bytes);
    const status = runtime.kernel_invoke(pointer,bytes.byteLength);
    const length = runtime.kernel_result_len(), resultPointer = runtime.kernel_result_ptr();
    if (status !== 0 || length === 0 || length > MAX_FRAME_BYTES || resultPointer + length > runtime.memory.buffer.byteLength) throw failure('KERNEL_PROTOCOL_ERROR','Invalid kernel response frame.');
    text = decoder.decode(new Uint8Array(runtime.memory.buffer,resultPointer,length));
  } catch (cause) { if (cause instanceof KernelInvocationError) throw cause; throw failure('KERNEL_EXECUTION_FAILED', cause instanceof Error ? cause.message : 'Kernel execution failed.'); }
  finally { runtime.kernel_free(pointer,bytes.byteLength); }
  let response: {protocolVersion:number;requestId:string;ok:boolean;result?:T;error?:KernelErrorDetail};
  try { response = JSON.parse(text); } catch { throw failure('KERNEL_PROTOCOL_ERROR','Kernel response is not valid JSON.'); }
  if (response.protocolVersion !== KERNEL_PROTOCOL_VERSION || response.requestId !== requestId || typeof response.ok !== 'boolean') throw failure('KERNEL_PROTOCOL_ERROR','Kernel response identity does not match the request.');
  if (!response.ok) {
    const error = response.error;
    if (!error || typeof error.code !== 'string' || typeof error.message !== 'string' || typeof error.recovery !== 'string') throw failure('KERNEL_PROTOCOL_ERROR','Kernel error payload is incomplete.');
    throw new KernelInvocationError(error);
  }
  if (!Object.hasOwn(response,'result')) throw failure('KERNEL_PROTOCOL_ERROR','Successful kernel response has no result.');
  return response.result as T;
}
