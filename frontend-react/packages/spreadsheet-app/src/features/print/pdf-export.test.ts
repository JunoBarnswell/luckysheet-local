import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import {
  browserPrintHook,
  createDefaultPrintLayout,
  createPdfBytes,
  nodePrintHook,
  type PrintLayoutModel,
  type PrintPageInfo,
} from './index';

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
});
