import {
  CLIPBOARD_HTML_MIME,
  CLIPBOARD_INTERNAL_MIME,
  CLIPBOARD_TEXT_MIME,
  clipboardRepresentations,
  type ClipboardPayload,
} from '@react-sheets/sheet-features';

export interface BrowserClipboardPort {
  write?: (items: ClipboardItems) => Promise<void>;
  writeText?: (text: string) => Promise<void>;
}

export type SystemClipboardWriteOutcome =
  | { status: 'published'; formats: readonly string[] }
  | { status: 'reduced'; formats: readonly string[]; missingFormats: readonly string[] }
  | { status: 'failed'; formats: readonly []; error: Error };

type ClipboardItemConstructor = new (items: Record<string, Blob>) => ClipboardItem;

function normalizeError(error: unknown, fallback: string): Error {
  return error instanceof Error ? error : new Error(fallback);
}

function representations(payload: ClipboardPayload): Map<string, string> {
  const values = new Map<string, string>();
  for (const representation of payload.representations ?? clipboardRepresentations(payload)) values.set(representation.mime, representation.data);
  if (!values.has(CLIPBOARD_TEXT_MIME) && payload.text !== undefined) values.set(CLIPBOARD_TEXT_MIME, payload.text);
  if (!values.has(CLIPBOARD_HTML_MIME) && payload.html !== undefined) values.set(CLIPBOARD_HTML_MIME, payload.html);
  const text = values.get(CLIPBOARD_TEXT_MIME);
  if (text === undefined) throw new Error('Clipboard payload has no text/plain representation');
  return values;
}

function supportsFormat(itemConstructor: ClipboardItemConstructor, mime: string, optional: boolean): boolean {
  const supports = (itemConstructor as unknown as { supports?: (type: string) => boolean }).supports;
  if (!supports) return !optional;
  try {
    return supports(mime);
  } catch {
    return false;
  }
}

/**
 * Publishes the same canonical payload used by in-app Paste to the browser
 * clipboard. This is the only boundary that knows about navigator.clipboard.
 */
export async function writeSystemClipboard(
  payload: ClipboardPayload,
  clipboard: BrowserClipboardPort | undefined = typeof navigator === 'undefined' ? undefined : navigator.clipboard,
): Promise<SystemClipboardWriteOutcome> {
  if (!clipboard) {
    return { status: 'failed', formats: [], error: new Error('System clipboard is unavailable') };
  }

  let serialized: Map<string, string>;
  try {
    serialized = representations(payload);
  } catch (error) {
    return { status: 'failed', formats: [], error: normalizeError(error, 'Clipboard payload could not be serialized') };
  }

  const itemConstructor = (globalThis as typeof globalThis & { ClipboardItem?: ClipboardItemConstructor }).ClipboardItem;
  if (clipboard.write && itemConstructor) {
    const optionalInternal = supportsFormat(itemConstructor, CLIPBOARD_INTERNAL_MIME, true);
    const supported = [CLIPBOARD_TEXT_MIME, CLIPBOARD_HTML_MIME, ...(optionalInternal ? [CLIPBOARD_INTERNAL_MIME] : [])]
      .filter((mime) => serialized.has(mime) && supportsFormat(itemConstructor, mime, false));
    if (supported.includes(CLIPBOARD_TEXT_MIME)) {
      const itemData: Record<string, Blob> = {};
      for (const mime of supported) itemData[mime] = new Blob([serialized.get(mime)!], { type: mime });
      try {
        const item = new itemConstructor(itemData);
        await clipboard.write([item]);
        const missingFormats = [CLIPBOARD_HTML_MIME]
          .filter((mime) => !supported.includes(mime));
        return missingFormats.length === 0
          ? { status: 'published', formats: supported }
          : { status: 'reduced', formats: supported, missingFormats };
      } catch (error) {
        return { status: 'failed', formats: [], error: normalizeError(error, 'System clipboard write was rejected') };
      }
    }
  }

  if (clipboard.writeText) {
    try {
      await clipboard.writeText(serialized.get(CLIPBOARD_TEXT_MIME)!);
      return {
        status: 'reduced',
        formats: [CLIPBOARD_TEXT_MIME],
        missingFormats: [CLIPBOARD_HTML_MIME, CLIPBOARD_INTERNAL_MIME].filter((mime) => serialized.has(mime)),
      };
    } catch (error) {
      return { status: 'failed', formats: [], error: normalizeError(error, 'System clipboard write was rejected') };
    }
  }

  return { status: 'failed', formats: [], error: new Error('System clipboard write is unsupported') };
}
