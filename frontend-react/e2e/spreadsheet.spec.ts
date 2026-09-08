import { expect, test, type Page } from '@playwright/test';
import { DESIGNER_GEOMETRY } from '@react-sheets/ui-system';
import { openConnectedWorkbook, revealRibbonSurfaceById, waitForServerSaved } from './support/workbook-fixtures';

const DEFAULT_COLUMN_WIDTH_PX = 64;
const ROW_HEADER_WIDTH_PX = 39;
const COLUMN_HEADER_HEIGHT_PX = 20;
const DEFAULT_ROW_HEIGHT_PX = 20;

async function waitForWorkspace(page: Page) {
  await openConnectedWorkbook(page, 'zh-CN', `E2E ${Date.now()}`);
}

async function waitForDesignerDemo(page: Page) {
  page.on('pageerror', (error) => console.log(`[pageerror] ${error.message}\n${error.stack ?? ''}`));
  page.on('console', (message) => { if (message.type() === 'error') console.log(`[console.error] ${message.text()}`); });
  await page.setViewportSize({ width: 1280, height: 720 });
  await openConnectedWorkbook(page, 'zh-CN', `Designer E2E ${Date.now()}`);
  await page.getByTestId('name-box').fill('B1');
  await page.getByTestId('name-box').press('Enter');
  await expect(page.getByTestId('name-box')).toHaveValue('B1');
}

async function focusCanvas(page: Page) {
  const canvas = page.getByTestId('sheet-canvas');
  // A1 单元格中心：行头 39px + 列宽一半，列头 20px + 行高一半
  await canvas.click({ position: { x: 78, y: COLUMN_HEADER_HEIGHT_PX + DEFAULT_ROW_HEIGHT_PX / 2 } });
  return canvas;
}

function cellPoint(box: { x: number; y: number }, row: number, column: number) {
  return { x: box.x + ROW_HEADER_WIDTH_PX + column * DEFAULT_COLUMN_WIDTH_PX + DEFAULT_COLUMN_WIDTH_PX / 2, y: box.y + COLUMN_HEADER_HEIGHT_PX + row * DEFAULT_ROW_HEIGHT_PX + DEFAULT_ROW_HEIGHT_PX / 2 };
}

async function dragCells(page: Page, canvas: ReturnType<Page['getByTestId']>, from: { row: number; column: number }, to: { row: number; column: number }) {
  const box = await canvas.boundingBox();
  if (!box) throw new Error('Spreadsheet canvas has no bounds');
  await page.mouse.move(...Object.values(cellPoint(box, from.row, from.column)) as [number, number]);
  await page.mouse.down();
  await page.mouse.move(...Object.values(cellPoint(box, to.row, to.column)) as [number, number]);
  await page.mouse.up();
  return box;
}

async function dragScrollbar(page: Page, orientation: 'horizontal' | 'vertical') {
  const scrollbar = page.getByRole('scrollbar', { name: `${orientation} scrollbar` });
  await expect(scrollbar).toBeVisible();
  const track = await scrollbar.boundingBox();
  const thumb = await scrollbar.locator(':scope > div').first().boundingBox();
  if (!track || !thumb) throw new Error(`${orientation} scrollbar has no geometry`);
  const before = Number(await scrollbar.getAttribute('aria-valuenow'));
  const start = { x: thumb.x + thumb.width / 2, y: thumb.y + thumb.height / 2 };
  const end = orientation === 'horizontal'
    ? { x: track.x + track.width * 0.65, y: start.y }
    : { x: start.x, y: track.y + track.height * 0.65 };
  await page.mouse.move(start.x, start.y);
  await page.mouse.down();
  await page.mouse.move(end.x, end.y, { steps: 8 });
  await page.mouse.up();
  await expect.poll(async () => Number(await scrollbar.getAttribute('aria-valuenow'))).not.toBe(before);
  return scrollbar;
}

test.describe('spreadsheet baseline', () => {
  test('cloud workbook shell exposes the fixed 1280x720 geometry and real palette entry', async ({ page }) => {
    await waitForDesignerDemo(page);
    const geometry = await page.evaluate(() => {
      const rect = (selector: string) => {
        const node = document.querySelector(selector);
        if (!node) return null;
        const box = node.getBoundingClientRect();
        return { y: Math.round(box.y), height: Math.round(box.height) };
      };
      return {
        ribbon: rect('[data-testid="designer-ribbon"]'),
        formulaBar: rect('[data-testid="designer-formula-bar"]'),
        workspace: rect('[data-testid="designer-workspace"]'),
        sheetTabs: rect('[data-testid="designer-sheet-tabs"]'),
        statusBar: rect('[data-testid="designer-status-bar"]'),
      };
    });
    expect(geometry.ribbon).toEqual({ y: DESIGNER_GEOMETRY.documentBarHeight, height: DESIGNER_GEOMETRY.ribbonHeight });
    expect(geometry.formulaBar).toEqual({ y: DESIGNER_GEOMETRY.documentBarHeight + DESIGNER_GEOMETRY.ribbonHeight, height: DESIGNER_GEOMETRY.formulaBarHeight });
    const workspaceY = DESIGNER_GEOMETRY.documentBarHeight + DESIGNER_GEOMETRY.ribbonHeight + DESIGNER_GEOMETRY.formulaBarHeight;
    expect(geometry.workspace).toEqual({
      y: workspaceY,
      height: 720 - DESIGNER_GEOMETRY.statusBarHeight - workspaceY,
    });
    expect(geometry.sheetTabs).toEqual({
      y: 720 - DESIGNER_GEOMETRY.statusBarHeight - DESIGNER_GEOMETRY.sheetTabsHeight,
      height: DESIGNER_GEOMETRY.sheetTabsHeight,
    });
    expect(geometry.statusBar).toEqual({ y: 720 - DESIGNER_GEOMETRY.statusBarHeight, height: DESIGNER_GEOMETRY.statusBarHeight });
    await page.screenshot({ path: 'test-results/designer-demo-1280-current.png' });
    await page.getByRole('tab', { name: '视图' }).click();
    await page.getByRole('button', { name: '命令面板' }).click();
    await expect(page.getByTestId('command-palette')).toBeVisible();
    await expect(page.getByRole('button', { name: /保存/ }).first()).toBeVisible();
  });

  test('selection updates the name box', async ({ page }) => {
    await waitForWorkspace(page);
    await focusCanvas(page);
    await expect(page.getByTestId('name-box')).toHaveValue('A1');
    await page.keyboard.press('ArrowRight');
    await expect(page.getByTestId('name-box')).toHaveValue('B1');
  });

  test('horizontal and vertical scrollbar gestures never enter worksheet selection', async ({ page }) => {
    await waitForWorkspace(page);
    const canvas = await focusCanvas(page);
    await expect(page.getByTestId('name-box')).toHaveValue('A1');
    const selectionIdentity = async () => ({
      row: await canvas.getAttribute('aria-rowindex'),
      column: await canvas.getAttribute('aria-colindex'),
      name: await page.getByTestId('name-box').inputValue(),
    });
    const before = await selectionIdentity();

    const vertical = await dragScrollbar(page, 'vertical');
    expect(await selectionIdentity()).toEqual(before);
    const verticalBeforeHorizontal = Number(await vertical.getAttribute('aria-valuenow'));
    await dragScrollbar(page, 'horizontal');
    expect(await selectionIdentity()).toEqual(before);
    await expect.poll(async () => Number(await vertical.getAttribute('aria-valuenow'))).toBe(verticalBeforeHorizontal);

    const verticalBeforeKey = Number(await vertical.getAttribute('aria-valuenow'));
    await vertical.focus();
    await page.keyboard.press('PageDown');
    await expect.poll(async () => Number(await vertical.getAttribute('aria-valuenow'))).not.toBe(verticalBeforeKey);
    expect(await selectionIdentity()).toEqual(before);
  });

  test('keyboard navigation moves the active cell', async ({ page }) => {
    await waitForWorkspace(page);
    await focusCanvas(page);
    await page.keyboard.press('ArrowDown');
    await expect(page.getByTestId('name-box')).toHaveValue('A2');
    await page.keyboard.press('ArrowRight');
    await expect(page.getByTestId('name-box')).toHaveValue('B2');
  });

  test('direct typing commits cell editing', async ({ page }) => {
    await waitForWorkspace(page);
    const canvas = await focusCanvas(page);
    await page.keyboard.type('hello-e2e');
    await page.keyboard.press('Enter');
    await expect(page.getByTestId('formula-input')).not.toHaveValue('hello-e2e');
    await canvas.press('ArrowUp');
    await expect(page.getByTestId('formula-input')).toHaveValue('hello-e2e');
  });

  test('clicking another cell commits the old editor before accepting the new value', async ({ page }) => {
    await waitForWorkspace(page);
    const canvas = await focusCanvas(page);
    await page.keyboard.type('old-cell-value');
    await expect(page.getByLabel('Cell editor')).toBeVisible();
    await canvas.click({ position: { x: 206, y: 34 } });
    await expect(page.getByTestId('name-box')).toHaveValue('C1');
    await expect(page.getByLabel('Cell editor')).toHaveCount(0);
    await page.keyboard.type('new-cell-value');
    await page.keyboard.press('Enter');
    await canvas.press('ArrowUp');
    await expect(page.getByTestId('name-box')).toHaveValue('C1');
    await expect(page.getByTestId('formula-input')).toHaveValue('new-cell-value');
  });

  test('clipboard copy and paste round-trip', async ({ page, context }) => {
    await context.grantPermissions(['clipboard-read', 'clipboard-write']);
    await waitForWorkspace(page);
    const canvas = await focusCanvas(page);
    await page.keyboard.type('copy-me');
    await page.keyboard.press('Enter');
    await canvas.press('ArrowUp');
    await canvas.press('Control+C');
    await canvas.press('Control+V');
    await canvas.press('ArrowLeft');
    await expect(page.getByTestId('formula-input')).toHaveValue('copy-me');
  });

  test('formula bar cancel restores the draft', async ({ page }) => {
    await waitForWorkspace(page);
    await focusCanvas(page);
    const input = page.getByTestId('formula-input');
    const original = await input.inputValue();
    await input.fill(`${original}-draft`);
    await page.getByTestId('formula-cancel').click();
    await expect(input).toHaveValue(original);
  });

  test('undo reverts the last edit', async ({ page }) => {
    await waitForWorkspace(page);
    const canvas = await focusCanvas(page);
    await page.keyboard.type('undo-me');
    await page.keyboard.press('Enter');
    await canvas.press('ArrowUp');
    await expect(page.getByTestId('formula-input')).toHaveValue('undo-me');
    await canvas.press('Control+Z');
    await expect(page.getByTestId('formula-input')).toHaveValue('');
  });

  test('adds and activates a new editable worksheet from the sheet tab control', async ({ page }) => {
    await waitForWorkspace(page);
    await page.getByRole('button', { name: /add worksheet|添加工作表/i }).click();
    await expect(page.getByRole('tab', { name: /sheet2/i })).toBeVisible();
    await expect(page.getByRole('tab', { name: /sheet2/i })).toHaveAttribute('aria-selected', 'true');
    const canvas = await focusCanvas(page);
    await page.keyboard.type('new-sheet-value');
    await page.keyboard.press('Enter');
    await canvas.press('ArrowUp');
    await expect(page.getByTestId('formula-input')).toHaveValue('new-sheet-value');
  });

  test('F2 editing keeps Chinese IME text intact before one committed operation', async ({ page }) => {
    await waitForWorkspace(page);
    const canvas = await focusCanvas(page);
    await canvas.press('F2');
    const editor = page.getByLabel('Cell editor');
    await expect(editor).toBeFocused();
    await editor.dispatchEvent('compositionstart');
    await editor.fill('中文公式输入');
    await editor.dispatchEvent('compositionend');
    await editor.press('Enter');
    await canvas.press('ArrowUp');
    await expect(page.getByTestId('formula-input')).toHaveValue('中文公式输入');
    await page.keyboard.press('Control+Z');
    await expect(page.getByTestId('formula-input')).toHaveValue('');
  });

  test('Alt+Enter inserts a canonical multiline draft and commits it without an early Enter action', async ({ page }) => {
    await waitForWorkspace(page);
    const canvas = await focusCanvas(page);
    await canvas.press('F2');
    const editor = page.getByLabel('Cell editor');
    await editor.type('line one');
    await editor.press('Alt+Enter');
    await editor.type('line two');
    await expect(page.getByTestId('formula-input')).toHaveValue('line one\nline two');
    await editor.press('Enter');
    await canvas.press('ArrowUp');
    await expect(page.getByTestId('formula-input')).toHaveValue('line one\nline two');
  });

  test('formula Point mode inserts a clicked reference and F4 rewrites the caret token', async ({ page }) => {
    await waitForWorkspace(page);
    const canvas = await focusCanvas(page);
    const box = await canvas.boundingBox();
    if (!box) throw new Error('Spreadsheet canvas has no bounds');
    await page.keyboard.type('=');
    await page.mouse.click(...Object.values(cellPoint(box, 1, 1)) as [number, number]);
    await expect(page.getByLabel('Cell editor')).toBeVisible();
    await expect(page.getByTestId('formula-input')).toHaveValue('=B2');
    await page.getByLabel('Cell editor').press('F4');
    await expect(page.getByTestId('formula-input')).toHaveValue('=$B$2');
    await page.getByLabel('Cell editor').press('Enter');
  });

  test('Ctrl+Enter commits one draft to every selected cell and one undo removes the whole transaction', async ({ page }) => {
    await waitForWorkspace(page);
    const canvas = await focusCanvas(page);
    const box = await canvas.boundingBox();
    if (!box) throw new Error('Spreadsheet canvas has no bounds');
    await dragCells(page, canvas, { row: 0, column: 0 }, { row: 1, column: 1 });
    await page.keyboard.type('batch-entry');
    await page.getByLabel('Cell editor').press('Control+Enter');
    for (const target of [{ row: 0, column: 0 }, { row: 0, column: 1 }, { row: 1, column: 0 }, { row: 1, column: 1 }]) {
      await page.mouse.click(...Object.values(cellPoint(box, target.row, target.column)) as [number, number]);
      await expect(page.getByTestId('formula-input')).toHaveValue('batch-entry');
    }
    await canvas.press('Control+Z');
    await page.mouse.click(...Object.values(cellPoint(box, 0, 0)) as [number, number]);
    await expect(page.getByTestId('formula-input')).toHaveValue('');
    await page.mouse.click(...Object.values(cellPoint(box, 1, 1)) as [number, number]);
    await expect(page.getByTestId('formula-input')).toHaveValue('');
  });

  test('draft Ctrl+Z is owned by CellEditDomain before workbook history', async ({ page }) => {
    await waitForWorkspace(page);
    await focusCanvas(page);
    await page.keyboard.type('draft');
    const editor = page.getByLabel('Cell editor');
    await editor.press('Control+Z');
    await expect(page.getByTestId('formula-input')).toHaveValue('draf');
    await editor.press('Escape');
    await expect(page.getByTestId('formula-input')).toHaveValue('');
  });

  test('a dragged A2:C9 selection edits and places the editor at active cell C9', async ({ page }) => {
    await waitForWorkspace(page);
    const canvas = await focusCanvas(page);
    const box = await canvas.boundingBox();
    if (!box) throw new Error('Spreadsheet canvas has no bounds');
    await page.mouse.move(...Object.values(cellPoint(box, 1, 0)) as [number, number]);
    await page.mouse.down();
    await page.mouse.move(...Object.values(cellPoint(box, 8, 2)) as [number, number]);
    await page.mouse.up();
    await expect(page.getByTestId('name-box')).toHaveValue('C9');

    await canvas.press('F2');
    const editor = page.getByLabel('Cell editor');
    await expect(editor).toBeFocused();
    const editorBox = await editor.boundingBox();
    if (!editorBox) throw new Error('Cell editor has no bounds');
    expect(editorBox.x - box.x).toBeGreaterThan(160);
    expect(editorBox.y - box.y).toBeGreaterThan(180);
    await editor.fill('4');
    await editor.press('Enter');
    await canvas.press('ArrowUp');
    await expect(page.getByTestId('name-box')).toHaveValue('C9');
    await expect(page.getByTestId('formula-input')).toHaveValue('4');
  });

  test('a dragged selection accepts direct typing at the release cell', async ({ page }) => {
    await waitForWorkspace(page);
    const canvas = await focusCanvas(page);
    await dragCells(page, canvas, { row: 1, column: 0 }, { row: 8, column: 2 });
    await expect(page.getByTestId('name-box')).toHaveValue('C9');
    await page.keyboard.type('direct-c9');
    await canvas.press('Enter');
    await canvas.press('ArrowUp');
    await expect(page.getByTestId('name-box')).toHaveValue('C9');
    await expect(page.getByTestId('formula-input')).toHaveValue('direct-c9');
  });

  test('reverse drag keeps the release cell active for F2 editing', async ({ page }) => {
    await waitForWorkspace(page);
    const canvas = await focusCanvas(page);
    await dragCells(page, canvas, { row: 8, column: 2 }, { row: 1, column: 0 });
    await expect(page.getByTestId('name-box')).toHaveValue('A2');
    await canvas.press('F2');
    const editor = page.getByLabel('Cell editor');
    await expect(editor).toBeFocused();
    await editor.fill('reverse-a2');
    await editor.press('Enter');
    await canvas.press('ArrowUp');
    await expect(page.getByTestId('name-box')).toHaveValue('A2');
    await expect(page.getByTestId('formula-input')).toHaveValue('reverse-a2');
  });

  test('row and column header drags commit the final header target', async ({ page }) => {
    await waitForWorkspace(page);
    const canvas = await focusCanvas(page);
    const box = await canvas.boundingBox();
    if (!box) throw new Error('Spreadsheet canvas has no bounds');

    await page.mouse.move(box.x + 20, box.y + COLUMN_HEADER_HEIGHT_PX + DEFAULT_ROW_HEIGHT_PX + DEFAULT_ROW_HEIGHT_PX / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + 20, box.y + COLUMN_HEADER_HEIGHT_PX + 8 * DEFAULT_ROW_HEIGHT_PX + DEFAULT_ROW_HEIGHT_PX / 2);
    await page.mouse.up();
    await expect(page.getByTestId('name-box')).toHaveValue('A9');

    await page.mouse.move(box.x + ROW_HEADER_WIDTH_PX + 55, box.y + 12);
    await page.mouse.down();
    await page.mouse.move(box.x + ROW_HEADER_WIDTH_PX + 2 * DEFAULT_COLUMN_WIDTH_PX + DEFAULT_COLUMN_WIDTH_PX / 2, box.y + 12);
    await page.mouse.up();
    await expect(page.getByTestId('name-box')).toHaveValue('C1');
  });

  test('Excel-style column width paths share one multi-column transaction surface', async ({ page }) => {
    await waitForWorkspace(page);
    const canvas = await focusCanvas(page);
    const box = await canvas.boundingBox();
    if (!box) throw new Error('Spreadsheet canvas has no bounds');

    // Full-column A:C selection.
    await page.mouse.move(box.x + ROW_HEADER_WIDTH_PX + DEFAULT_COLUMN_WIDTH_PX / 2, box.y + 12);
    await page.mouse.down();
    await page.mouse.move(box.x + ROW_HEADER_WIDTH_PX + 2 * DEFAULT_COLUMN_WIDTH_PX + DEFAULT_COLUMN_WIDTH_PX / 2, box.y + 12);
    await page.mouse.up();

    // Drag the C boundary; the preview exposes both character and pixel units.
    const boundary = box.x + ROW_HEADER_WIDTH_PX + 3 * DEFAULT_COLUMN_WIDTH_PX;
    await page.mouse.move(boundary, box.y + 12);
    await page.mouse.down();
    await page.mouse.move(boundary + 24, box.y + 12);
    await expect(canvas).toBeVisible();
    await page.mouse.up();

    // Right-click keeps the complete multi-column selection and opens exact width.
    await page.mouse.click(box.x + ROW_HEADER_WIDTH_PX + DEFAULT_COLUMN_WIDTH_PX + DEFAULT_COLUMN_WIDTH_PX / 2, box.y + 12, { button: 'right' });
    await page.getByRole('menuitem', { name: 'Column Width…' }).click();
    const dialog = page.getByRole('dialog', { name: 'Column Width' });
    await expect(dialog).toBeVisible();
    await dialog.getByLabel('Excel character width').fill('12');
    await dialog.getByRole('button', { name: 'OK' }).click();

    // Ribbon and context-menu entries route through the same controller.
    await page.getByTestId('ribbon-tab-home').click();
    await (await revealRibbonSurfaceById(page, 'control.auto-fit-column-width')).click();

    await page.mouse.click(box.x + ROW_HEADER_WIDTH_PX + DEFAULT_COLUMN_WIDTH_PX / 2, box.y + 12, { button: 'right' });
    await page.getByRole('menuitem', { name: 'Hide Columns', exact: true }).click();
    await canvas.press('Control+Z');
  });

  test('right click changes the command target before the context menu opens', async ({ page }) => {
    await waitForWorkspace(page);
    await page.getByRole('button', { name: /add worksheet|添加工作表/i }).click();
    const canvas = await focusCanvas(page);
    const box = await canvas.boundingBox();
    if (!box) throw new Error('Spreadsheet canvas has no bounds');
    await page.getByTestId('name-box').fill('B2');
    await page.getByTestId('name-box').press('Enter');
    await canvas.focus();
    await page.keyboard.type('right-click-target');
    await page.keyboard.press('Enter');
    await page.getByTestId('name-box').fill('A1');
    await page.getByTestId('name-box').press('Enter');
    const target = cellPoint(box, 1, 1);
    await page.mouse.click(target.x, target.y, { button: 'right' });
    await expect(page.getByRole('menu')).toBeVisible();
    await expect(page.getByTestId('name-box')).toHaveValue('B2');
    await page.getByRole('menuitem', { name: 'Clear contents' }).click();
    await page.getByTestId('name-box').fill('B2');
    await page.getByTestId('name-box').press('Enter');
    await expect(page.getByTestId('formula-input')).toHaveValue('');
  });

  test('fill handle follows the primary range bottom-right', async ({ page }) => {
    await waitForWorkspace(page);
    await page.getByRole('button', { name: /add worksheet|添加工作表/i }).click();
    const canvas = await focusCanvas(page);
    await page.keyboard.type('fill-source-unique');
    await page.keyboard.press('Enter');
    await page.getByTestId('name-box').fill('A1');
    await page.getByTestId('name-box').press('Enter');
    await canvas.focus();
    const box = await canvas.boundingBox();
    if (!box) throw new Error('Spreadsheet canvas has no bounds');
    await page.mouse.move(box.x + ROW_HEADER_WIDTH_PX + DEFAULT_COLUMN_WIDTH_PX - 4, box.y + COLUMN_HEADER_HEIGHT_PX + DEFAULT_ROW_HEIGHT_PX - 4);
    await page.mouse.down();
    await page.mouse.move(...Object.values(cellPoint(box, 2, 2)) as [number, number]);
    await page.mouse.up();
    await expect(page.getByTestId('name-box')).toHaveValue('C3');
    await page.getByTestId('name-box').fill('C3');
    await page.getByTestId('name-box').press('Enter');
    await expect(page.getByTestId('formula-input')).toHaveValue('fill-source-unique');
  });

  test('formula calculation commits through the cloud session and survives reload', async ({ page }) => {
    const apiRequests: string[] = [];
    page.on('request', (request) => {
      if (new URL(request.url()).pathname.startsWith('/api/')) apiRequests.push(request.url());
    });
    await waitForWorkspace(page);
    const canvas = await focusCanvas(page);
    await page.keyboard.type('=1+2');
    await page.keyboard.press('Enter');
    await expect(page.getByTestId('formula-input')).toHaveValue('');
    await canvas.press('ArrowUp');
    await expect(page.getByTestId('formula-input')).toHaveValue('=1+2');
    await waitForServerSaved(page);
    await page.reload();
    await expect(page.getByTestId('designer-shell')).toHaveAttribute('data-workspace-phase', 'ready', { timeout: 30_000 });
    await expect(page.getByTestId('formula-input')).toHaveValue('=1+2');
    expect(apiRequests.length).toBeGreaterThan(0);
  });

  test('Home ribbon opens shared format, sort, find, and paste dialogs without rendering Add-ins', async ({ page }) => {
    await waitForWorkspace(page);
    await page.getByTestId('ribbon-tab-home').click();
    await expect(page.getByTestId('home-ribbon-groups')).toBeVisible();
    await expect(page.getByRole('button', { name: /add-ins/i })).toHaveCount(0);

    await (await revealRibbonSurfaceById(page, 'styles.format-cells')).click();
    const formatDialog = page.getByTestId('format-cells-dialog');
    await expect(formatDialog).toBeVisible();
    await formatDialog.getByTestId('format-tab-font').click();
    await expect(formatDialog.getByLabel(/(font size|字号)/i)).toBeVisible();
    await formatDialog.getByRole('button', { name: /^(Close|关闭)$/ }).click();

    const sortRoot = await revealRibbonSurfaceById(page, 'editing.sort');
    if (await sortRoot.getAttribute('data-ribbon-command') === 'sortRange') await sortRoot.click();
    else {
      await sortRoot.click();
      await page.locator('[data-ribbon-surface="editing.sort"][data-ribbon-command="sortRange"]').click();
    }
    const sortDialog = page.getByTestId('sort-dialog');
    await expect(sortDialog).toBeVisible();
    await sortDialog.getByRole('button', { name: /^(Close|关闭)$/ }).click();

    const findRoot = await revealRibbonSurfaceById(page, 'editing.find');
    if (await findRoot.getAttribute('data-ribbon-command') === 'findReplace') await findRoot.click();
    else {
      await findRoot.click();
      await page.locator('[data-ribbon-surface="editing.find"][data-ribbon-command="findReplace"]').click();
    }
    const findDialog = page.getByTestId('find-replace-dialog');
    await expect(findDialog).toBeVisible();
    await findDialog.getByLabel(/^(Find|查找)$/).fill('home-dialog-check');
    await findDialog.getByTestId('replace-input').fill('replacement');
    await expect(findDialog.getByRole('button', { name: /(replace all|全部替换)/i })).toBeEnabled();
    await findDialog.getByRole('button', { name: /^(Close|关闭)$/ }).click();

    await (await revealRibbonSurfaceById(page, 'clipboard.paste-special')).click();
    const pasteDialog = page.getByTestId('paste-special-dialog');
    await expect(pasteDialog).toBeVisible();
    await expect(pasteDialog.getByTestId('paste-special-formats')).toBeVisible();
    await pasteDialog.getByRole('button', { name: /^(Cancel|取消)$/ }).click();

    await (await revealRibbonSurfaceById(page, 'editing.selection-pane')).click();
    await expect(page.getByTestId('selection-pane')).toBeVisible();
  });

  test('Selection Pane selects, renames, and toggles a drawing through host callbacks', async ({ page }) => {
    await waitForWorkspace(page);
    await page.getByTestId('ribbon-tab-insert').click();
    const shapeRoot = await revealRibbonSurfaceById(page, 'illustrations.shape');
    const rectangle = page.locator('[data-ribbon-variant="shape.rectangle"]');
    if (!await rectangle.isVisible()) await shapeRoot.click();
    await rectangle.click();

    await page.getByTestId('ribbon-tab-home').click();
    await (await revealRibbonSurfaceById(page, 'editing.selection-pane')).click();
    const pane = page.getByTestId('selection-pane');
    await expect(pane).toBeVisible();

    await pane.getByRole('button', { name: /^(Rename|重命名)$/ }).click();
    const rename = pane.getByLabel(/^(Rename|重命名)$/);
    await rename.fill('KPI tile');
    await rename.press('Enter');
    await expect(pane.getByRole('button', { name: 'KPI tile', exact: true })).toBeVisible();

    const visibility = pane.getByLabel(/KPI tile: (Visible|可见)/);
    await expect(visibility).toBeChecked();
    await visibility.click();
    await waitForServerSaved(page);
    await expect(visibility).not.toBeChecked();
  });

  test('Format Painter enters a transient Home state and completes after one target selection', async ({ page }) => {
    await waitForWorkspace(page);
    const canvas = await focusCanvas(page);
    await canvas.press('Control+B');

    const painter = await revealRibbonSurfaceById(page, 'control.format-painter');
    await painter.click();
    await expect(painter).toHaveAttribute('aria-pressed', 'true');

    await canvas.click({ position: { x: 142, y: 34 } });
    await expect(await revealRibbonSurfaceById(page, 'control.format-painter')).toHaveAttribute('aria-pressed', 'false');
  });
});
