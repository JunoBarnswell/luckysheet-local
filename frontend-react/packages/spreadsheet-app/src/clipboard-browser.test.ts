import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import {
  CLIPBOARD_HTML_MIME,
  CLIPBOARD_INTERNAL_MIME,
  CLIPBOARD_TEXT_MIME,
  type ClipboardPayload,
} from '@react-sheets/sheet-features';
import { writeSystemClipboard } from './clipboard-browser';

function payload(): ClipboardPayload {
  return {
    range: { sheetId: 'sheet-1', startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 },
    values: [[{ value: 'Ada' }, { value: 42 }]],
    transfer: 'copy',
    rangeMetadata: { columnWidths: [], validations: [], conditionalFormats: [], notes: [], comments: [], hyperlinks: [] },
    representations: [
      { mime: CLIPBOARD_INTERNAL_MIME, data: '{"values":[[{"value":"Ada"},{"value":42}]]}' },
      { mime: CLIPBOARD_HTML_MIME, data: '<table><tr><td>Ada</td><td>42</td></tr></table>' },
      { mime: CLIPBOARD_TEXT_MIME, data: 'Ada\t42' },
    ],
  };
}

describe('browser clipboard publication boundary', () => {
  it('publishes text and HTML from the canonical payload and omits unsupported internal MIME', async () => {
    const original = (globalThis as typeof globalThis & { ClipboardItem?: unknown }).ClipboardItem;
    const written: unknown[] = [];
    class FakeClipboardItem {
      static supports(mime: string): boolean { return mime !== CLIPBOARD_INTERNAL_MIME; }
      constructor(readonly data: Record<string, Blob>) { written.push(data); }
    }
    Object.defineProperty(globalThis, 'ClipboardItem', { configurable: true, value: FakeClipboardItem });
    try {
      const outcome = await writeSystemClipboard(payload(), {
        write: async (items) => { written.push(items); },
      });
      assert.equal(outcome.status, 'published');
      if (outcome.status === 'published') assert.deepEqual(outcome.formats, [CLIPBOARD_TEXT_MIME, CLIPBOARD_HTML_MIME]);
      assert.equal(written.length, 2);
      assert.deepEqual(Object.keys(written[0] as Record<string, Blob>), [CLIPBOARD_TEXT_MIME, CLIPBOARD_HTML_MIME]);
    } finally {
      Object.defineProperty(globalThis, 'ClipboardItem', { configurable: true, value: original });
    }
  });

  it('uses an explicit text-only result when ClipboardItem is unavailable', async () => {
    const original = (globalThis as typeof globalThis & { ClipboardItem?: unknown }).ClipboardItem;
    Object.defineProperty(globalThis, 'ClipboardItem', { configurable: true, value: undefined });
    let text = '';
    try {
      const outcome = await writeSystemClipboard(payload(), {
        write: async () => { throw new Error('rich write must not be called'); },
        writeText: async (value) => { text = value; },
      });
      assert.equal(outcome.status, 'reduced');
      if (outcome.status === 'reduced') assert.deepEqual(outcome.formats, [CLIPBOARD_TEXT_MIME]);
      assert.equal(text, 'Ada\t42');
    } finally {
      Object.defineProperty(globalThis, 'ClipboardItem', { configurable: true, value: original });
    }
  });

  it('returns a failure when the system write is rejected without falling back after a capability exists', async () => {
    const original = (globalThis as typeof globalThis & { ClipboardItem?: unknown }).ClipboardItem;
    class FakeClipboardItem {
      static supports(): boolean { return true; }
      constructor(readonly _data: Record<string, Blob>) {}
    }
    Object.defineProperty(globalThis, 'ClipboardItem', { configurable: true, value: FakeClipboardItem });
    let fallbackCalled = false;
    try {
      const outcome = await writeSystemClipboard(payload(), {
        write: async () => { throw new Error('permission denied'); },
        writeText: async () => { fallbackCalled = true; },
      });
      assert.equal(outcome.status, 'failed');
      if (outcome.status === 'failed') assert.match(outcome.error.message, /permission denied/i);
      assert.equal(fallbackCalled, false);
    } finally {
      Object.defineProperty(globalThis, 'ClipboardItem', { configurable: true, value: original });
    }
  });

  it('fails explicitly when no browser clipboard capability exists', async () => {
    const outcome = await writeSystemClipboard(payload(), undefined);
    assert.equal(outcome.status, 'failed');
  });
});
