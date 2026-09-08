import assert from 'node:assert/strict';
import test from 'node:test';
import { getPrintDocument, type PageSetup } from './index';
import { registerPrintCommands } from './commands';
import { openCanonicalTestRuntime } from '../../../../core-model/src/canonical-test-runtime.test';

const pageSetup: PageSetup = {
  paperSize: 'letter',
  orientation: 'landscape',
  margins: { top: 36, right: 36, bottom: 36, left: 36, header: 18, footer: 18 },
  scale: 90,
  fitToWidth: 1,
  printGridlines: true,
  printHeadings: false,
  centerHorizontally: true,
  centerVertically: false,
};

async function setup(unitId: string) {
  const fixture = await openCanonicalTestRuntime(unitId, 'Print commands');
  registerPrintCommands(fixture.runtime.registry);
  fixture.runtime.registry.assertComplete();
  return fixture;
}

test('print commands publish only canonical typed mutations', async () => {
  const { workbook, runtime, close } = await setup('print-command-canonical');
  try {
    const planned: Array<{ id: string; params: unknown }> = [];
    runtime.onMutation((mutation) => planned.push({ id: mutation.id, params: mutation.params }));
    const sheetId = workbook.primarySheetId;
    const range = { sheetId, startRow: 2, endRow: 20, startColumn: 1, endColumn: 7 };

    await runtime.execute('pageLayout.pageSetup.set', { sheetId, pageSetup });
    await runtime.execute('pageLayout.printArea.set', { sheetId, range });
    await runtime.execute('pageLayout.pageBreak.insert', { sheetId, pageBreak: { sheetId, row: 10 } });

    assert.deepEqual(planned.map((entry) => entry.id), [
      'pageLayout.pageSetupDetail.set',
      'pageLayout.printArea.set',
      'pageLayout.pageBreak.insert',
    ]);
    assert.equal(planned.every((entry) => !Object.prototype.hasOwnProperty.call(entry, 'inverse')), true);
    assert.equal(runtime.registry.listMutationRegistrations().every((entry) => !Object.prototype.hasOwnProperty.call(entry.metadata, 'inversePolicy')), true);

    const document = getPrintDocument(workbook, sheetId);
    assert.deepEqual(document.pageSetup, pageSetup);
    assert.deepEqual(document.printAreas, [{ sheetId, range }]);
    assert.deepEqual(document.pageBreaks, [{ sheetId, row: 10 }]);
  } finally {
    close();
  }
});

test('print commands compose title updates and persist all page layout fields', async () => {
  const { workbook, runtime, close } = await setup('print-command-fields');
  try {
    const sheetId = workbook.primarySheetId;
    const planned: string[] = [];
    runtime.onMutation((mutation) => planned.push(mutation.id));

    const setupResult = await runtime.execute('pageLayout.pageSetup.set', {
      sheetId,
      layout: {
        paper: 'A4', orientation: 'portrait', margin: { top: 20, right: 20, bottom: 20, left: 20 },
        repeatRows: { sheetId, startRow: 0, endRow: 1, startColumn: 0, endColumn: 4 },
      },
    });
    await runtime.execute('pageLayout.printTitles.set', { sheetId, repeatColumns: { start: 0, end: 0 } });
    await runtime.execute('pageLayout.scaleToFit.set', { sheetId, scale: 70, fitToHeight: 1 });
    await runtime.execute('pageLayout.scaleToFit.set', { sheetId, scale: 80, fitToWidth: 1 });
    await runtime.execute('pageLayout.printGridlines.set', { sheetId, enabled: true });
    await runtime.execute('pageLayout.printHeadings.set', { sheetId, enabled: true });
    await runtime.execute('pageLayout.viewGridlines.set', { sheetId, enabled: false });
    await runtime.execute('pageLayout.viewHeadings.set', { sheetId, enabled: false });

    assert.equal(setupResult.mutationCount, 2);
    assert.deepEqual(planned, [
      'pageLayout.pageSetupDetail.set',
      'pageLayout.printTitles.set',
      'pageLayout.printTitles.set',
      'pageLayout.scaleToFit.set',
      'pageLayout.scaleToFit.set',
      'pageLayout.printGridlines.set',
      'pageLayout.printHeadings.set',
      'pageLayout.viewGridlines.set',
      'pageLayout.viewHeadings.set',
    ]);
    const document = getPrintDocument(workbook, sheetId);
    assert.deepEqual(document.repeatRows, { start: 0, end: 1 });
    assert.deepEqual(document.repeatColumns, { start: 0, end: 0 });
    assert.equal(document.pageSetup.scale, 80);
    assert.equal(document.pageSetup.fitToWidth, 1);
    assert.equal(document.pageSetup.fitToHeight, 1);
    assert.equal(document.pageSetup.printGridlines, true);
    assert.equal(document.pageSetup.printHeadings, true);
    assert.equal(workbook.getSheet(sheetId).showGridlines, false);
    assert.equal(workbook.getSheet(sheetId).showHeaders, false);
  } finally {
    close();
  }
});

test('print clear mutations use canonical empty-state commands', async () => {
  const { workbook, runtime, close } = await setup('print-command-clear');
  try {
    const sheetId = workbook.primarySheetId;
    const range = { sheetId, startRow: 1, endRow: 5, startColumn: 0, endColumn: 3 };
    await runtime.execute('pageLayout.printArea.set', { sheetId, range });
    await runtime.execute('pageLayout.pageBreak.insert', { sheetId, pageBreak: { sheetId, row: 3 } });
    await runtime.execute('pageLayout.printTitles.set', { sheetId, repeatRows: { start: 0, end: 1 } });
    await runtime.execute('pageLayout.printArea.clear', { sheetId });
    await runtime.execute('pageLayout.pageBreak.clear', { sheetId });
    await runtime.execute('pageLayout.printTitles.clear', { sheetId });

    const document = getPrintDocument(workbook, sheetId);
    assert.deepEqual(document.printAreas, []);
    assert.deepEqual(document.pageBreaks, []);
    assert.equal(document.repeatRows, undefined);
    assert.equal(document.repeatColumns, undefined);
  } finally {
    close();
  }
});
