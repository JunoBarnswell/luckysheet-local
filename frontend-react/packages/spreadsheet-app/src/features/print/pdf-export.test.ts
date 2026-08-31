import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { zlibSync } from 'fflate';
import {
  browserPrintHook,
  createDefaultPrintLayout,
  createPdfBytes,
  nodePrintHook,
  PdfResourceError,
  type PrintLayoutModel,
  type PrintPageInfo,
} from './index';
import { decodePngToRgb, PngDecodeError } from './png-decoder';

const PNG_SIGNATURE = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

function concatBytes(parts: readonly Uint8Array[]): Uint8Array {
  const result = new Uint8Array(parts.reduce((total, part) => total + part.length, 0));
  let offset = 0;
  for (const part of parts) {
    result.set(part, offset);
    offset += part.length;
  }
  return result;
}

function crc32(bytes: Uint8Array): number {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit += 1) crc = (crc & 1) === 0 ? crc >>> 1 : 0xedb88320 ^ (crc >>> 1);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function pngChunk(type: string, data: Uint8Array): Uint8Array {
  const typeBytes = new TextEncoder().encode(type);
  const crcInput = concatBytes([typeBytes, data]);
  const header = new Uint8Array(4);
  new DataView(header.buffer).setUint32(0, data.length, false);
  const footer = new Uint8Array(4);
  new DataView(footer.buffer).setUint32(0, crc32(crcInput), false);
  return concatBytes([header, typeBytes, data, footer]);
}

function createPng(options: {
  width: number;
  height: number;
  colorType: 0 | 2 | 3 | 4 | 6;
  scanlines: Uint8Array;
  palette?: Uint8Array;
  transparency?: Uint8Array;
  interlace?: number;
}): Uint8Array {
  const header = new Uint8Array(13);
  const view = new DataView(header.buffer);
  view.setUint32(0, options.width, false);
  view.setUint32(4, options.height, false);
  header[8] = 8;
  header[9] = options.colorType;
  header[12] = options.interlace ?? 0;
  return concatBytes([
    PNG_SIGNATURE,
    pngChunk('IHDR', header),
    ...(options.palette ? [pngChunk('PLTE', options.palette)] : []),
    ...(options.transparency ? [pngChunk('tRNS', options.transparency)] : []),
    pngChunk('IDAT', zlibSync(options.scanlines)),
    pngChunk('IEND', new Uint8Array()),
  ]);
}

function filterRows(rows: readonly Uint8Array[], bytesPerPixel: number, filters: readonly number[]): Uint8Array {
  const rowBytes = rows[0]?.length ?? 0;
  const previous = new Uint8Array(rowBytes);
  const encoded = new Uint8Array(rows.length * (rowBytes + 1));
  let offset = 0;
  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex]!;
    const filter = filters[rowIndex]!;
    encoded[offset++] = filter;
    for (let index = 0; index < rowBytes; index += 1) {
      const left = index >= bytesPerPixel ? row[index - bytesPerPixel]! : 0;
      const above = previous[index]!;
      const upperLeft = index >= bytesPerPixel ? previous[index - bytesPerPixel]! : 0;
      const predictor = filter === 0 ? 0 : filter === 1 ? left : filter === 2 ? above : filter === 3 ? Math.floor((left + above) / 2) : (() => {
        const estimate = left + above - upperLeft;
        const leftDistance = Math.abs(estimate - left);
        const aboveDistance = Math.abs(estimate - above);
        const upperLeftDistance = Math.abs(estimate - upperLeft);
        return leftDistance <= aboveDistance && leftDistance <= upperLeftDistance ? left : aboveDistance <= upperLeftDistance ? above : upperLeft;
      })();
      encoded[offset++] = (row[index]! - predictor + 256) & 0xff;
    }
    previous.set(row);
  }
  return encoded;
}

function imageProjection(page: PrintPageInfo, mimeType: string, bytes: Uint8Array, width: number, height: number) {
  return {
    schema: 'PrintProjection' as const,
    page,
    cells: [{
      row: 0,
      column: 0,
      value: null,
      displayValue: '',
      rowOffsetPx: 0,
      columnOffsetPx: 0,
      widthPx: 40,
      heightPx: 20,
      image: {
        asset: { schema: 'AssetRef' as const, assetId: 'print-image', contentHash: 'a'.repeat(64), mimeType, byteLength: bytes.length, width, height },
        bytes,
      },
    }],
    drawings: [],
    visibleRows: [0],
    visibleColumns: [0],
    scaleX: 1,
    scaleY: 1,
    contentWidthPx: 40,
    contentHeightPx: 20,
  };
}

const layout: PrintLayoutModel = {
  ...createDefaultPrintLayout('wb-pdf', 'sheet-1'),
  pageSetup: {
    ...createDefaultPrintLayout('wb-pdf', 'sheet-1').pageSetup,
    paperSize: 'a4',
    orientation: 'portrait',
  },
  printAreas: [{ sheetId: 'sheet-1', range: { sheetId: 'sheet-1', startRow: 0, endRow: 4, startColumn: 0, endColumn: 2 } }],
};

const pages: PrintPageInfo[] = [{
  pageIndex: 0,
  sheetId: 'sheet-1',
  range: { sheetId: 'sheet-1', startRow: 0, endRow: 4, startColumn: 0, endColumn: 2 },
  widthPx: 240,
  heightPx: 140,
}];

describe('print PDF output', () => {
  it('creates a real PDF byte stream in Node', async () => {
    const bytes = createPdfBytes(pages, layout, { title: 'UAT 工作簿', author: '测试' });
    assert.equal(new TextDecoder().decode(bytes.subarray(0, 8)), '%PDF-1.7');
    const text = new TextDecoder().decode(bytes);
    assert.match(text, /xref/);
    assert.match(text, /%%EOF/);
    assert.match(text, /FEFF00550041005400205DE54F5C7C3F/i);
    assert.match(text, /6D4B8BD5/i);

    const result = await nodePrintHook.renderPages(pages, layout, { title: 'UAT 工作簿' });
    assert.ok(result instanceof Uint8Array);
    assert.equal(new TextDecoder().decode(result.subarray(0, 8)), '%PDF-1.7');
  });

  it('returns application/pdf bytes from the browser host port', async () => {
    const result = await browserPrintHook.renderPages(pages, layout, { title: 'Browser workbook' });
    if (result instanceof Blob) {
      assert.equal(result.type, 'application/pdf');
      const bytes = new Uint8Array(await result.arrayBuffer());
      assert.equal(new TextDecoder().decode(bytes.subarray(0, 8)), '%PDF-1.7');
    } else {
      assert.equal(new TextDecoder().decode(result.subarray(0, 8)), '%PDF-1.7');
    }
  });

  it('decodes PNG into RGB and embeds it with FlateDecode', () => {
    const png = createPng({
      width: 2,
      height: 5,
      colorType: 6,
      scanlines: filterRows([
        new Uint8Array([255, 0, 0, 255, 0, 255, 0, 255]),
        new Uint8Array([0, 0, 255, 255, 255, 255, 0, 255]),
        new Uint8Array([12, 34, 56, 255, 90, 80, 70, 255]),
        new Uint8Array([200, 100, 20, 255, 10, 30, 50, 255]),
        new Uint8Array([1, 2, 3, 255, 4, 5, 6, 255]),
      ], 4, [0, 1, 2, 3, 4]),
    });
    const decoded = decodePngToRgb(png);
    assert.equal(decoded.width, 2);
    assert.equal(decoded.height, 5);
    assert.deepEqual(Array.from(decoded.rgb.subarray(0, 6)), [255, 0, 0, 0, 255, 0]);
    const bytes = createPdfBytes(pages, layout, { pageProjection: (page) => imageProjection(page, 'image/png', png, 2, 5) });
    const pdf = new TextDecoder().decode(bytes);
    assert.match(pdf, /\/Filter \/FlateDecode/);
    assert.match(pdf, /\/DecodeParms << \/Colors 3 \/BitsPerComponent 8 \/Columns 2 >>/);
  });

  it('supports palette and tRNS alpha compositing onto white', () => {
    const palettePng = createPng({
      width: 2,
      height: 1,
      colorType: 3,
      palette: new Uint8Array([255, 0, 0, 0, 0, 255]),
      transparency: new Uint8Array([0, 255]),
      scanlines: new Uint8Array([0, 0, 1]),
    });
    assert.deepEqual(Array.from(decodePngToRgb(palettePng).rgb), [255, 255, 255, 0, 0, 255]);

    const grayAlphaPng = createPng({ width: 1, height: 1, colorType: 4, scanlines: new Uint8Array([0, 64, 128]) });
    assert.deepEqual(Array.from(decodePngToRgb(grayAlphaPng).rgb), [159, 159, 159]);

    const rgbTransparentPng = createPng({
      width: 1,
      height: 1,
      colorType: 2,
      transparency: new Uint8Array([0, 10, 0, 20, 0, 30]),
      scanlines: new Uint8Array([0, 10, 20, 30]),
    });
    assert.deepEqual(Array.from(decodePngToRgb(rgbTransparentPng).rgb), [255, 255, 255]);
  });

  it('fails closed for interlaced PNG and unsupported image formats', () => {
    const interlaced = createPng({ width: 1, height: 1, colorType: 6, interlace: 1, scanlines: new Uint8Array([0, 1, 2, 3, 255]) });
    assert.throws(() => decodePngToRgb(interlaced), (error: unknown) => error instanceof PngDecodeError && error.code === 'PNG_UNSUPPORTED_FORMAT');
    const webp = new Uint8Array([0x52, 0x49, 0x46, 0x46]);
    assert.throws(
      () => createPdfBytes(pages, layout, { pageProjection: (page) => imageProjection(page, 'image/webp', webp, 1, 1) }),
      (error: unknown) => error instanceof PdfResourceError && error.code === 'NATIVE_PRINT_RESOURCE_UNSUPPORTED_FORMAT',
    );
  });
});
