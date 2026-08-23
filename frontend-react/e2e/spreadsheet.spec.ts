import { expect, test, type Page } from '@playwright/test';

async function waitForWorkspace(page: Page) {
  await page.goto('/');
  await expect(page.getByTestId('app-shell')).toBeVisible();
  await expect(page.getByTestId('app-shell')).toHaveAttribute('data-workspace-phase', 'ready');
}

async function focusCanvas(page: Page) {
  const canvas = page.getByTestId('sheet-canvas');
  // A1 单元格中心：行头 46px + 列宽一半，列头 24px + 行高一半
  await canvas.click({ position: { x: 100, y: 38 } });
  return canvas;
}

function cellPoint(box: { x: number; y: number }, row: number, column: number) {
  return { x: box.x + 46 + column * 110 + 55, y: box.y + 24 + row * 28 + 14 };
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

test.describe('spreadsheet baseline', () => {
  test('selection updates the name box', async ({ page }) => {
    await waitForWorkspace(page);
    await focusCanvas(page);
    await expect(page.getByTestId('name-box')).toHaveValue('A1');
    await page.keyboard.press('ArrowRight');
    await expect(page.getByTestId('name-box')).toHaveValue('B1');
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

  test('a dragged A2:C9 selection edits and places the editor at active cell C9', async ({ page }) => {
    await waitForWorkspace(page);
    const canvas = await focusCanvas(page);
    const box = await canvas.boundingBox();
    if (!box) throw new Error('Spreadsheet canvas has no bounds');
    await page.mouse.move(box.x + 101, box.y + 66);
    await page.mouse.down();
    await page.mouse.move(box.x + 321, box.y + 262);
    await page.mouse.up();
    await expect(page.getByTestId('name-box')).toHaveValue('C9');

    await canvas.press('F2');
    const editor = page.getByLabel('Cell editor');
    await expect(editor).toBeFocused();
    const editorBox = await editor.boundingBox();
    if (!editorBox) throw new Error('Cell editor has no bounds');
    expect(editorBox.x - box.x).toBeGreaterThan(250);
    expect(editorBox.y - box.y).toBeGreaterThan(220);
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

    await page.mouse.move(box.x + 20, box.y + 24 + 28 + 14);
    await page.mouse.down();
    await page.mouse.move(box.x + 20, box.y + 24 + 8 * 28 + 14);
    await page.mouse.up();
    await expect(page.getByTestId('name-box')).toHaveValue('A9');

    await page.mouse.move(box.x + 46 + 55, box.y + 12);
    await page.mouse.down();
    await page.mouse.move(box.x + 46 + 2 * 110 + 55, box.y + 12);
    await page.mouse.up();
    await expect(page.getByTestId('name-box')).toHaveValue('C1');
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
    await page.mouse.move(box.x + 46 + 110 - 4, box.y + 24 + 28 - 4);
    await page.mouse.down();
    await page.mouse.move(box.x + 46 + 110, box.y + 24 + 3 * 28);
    await page.mouse.up();
    await page.getByTestId('name-box').fill('A3');
    await page.getByTestId('name-box').press('Enter');
    await expect(page.getByTestId('formula-input')).toHaveValue('fill-source-unique');
  });

  test('local formula calculation and IndexedDB restore work without an API request', async ({ page }) => {
    const apiRequests: string[] = [];
    let socketCount = 0;
    page.on('request', (request) => {
      if (new URL(request.url()).pathname.startsWith('/api/')) apiRequests.push(request.url());
    });
    page.on('websocket', (socket) => {
      if (new URL(socket.url()).pathname === '/ws') socketCount += 1;
    });
    await waitForWorkspace(page);
    const canvas = await focusCanvas(page);
    await page.keyboard.type('=1+2');
    await page.keyboard.press('Enter');
    await expect(page.getByTestId('formula-input')).toHaveValue('');
    await canvas.press('ArrowUp');
    await expect(page.getByTestId('formula-input')).toHaveValue('=1+2');

    await page.reload();
    await expect(page.getByTestId('app-shell')).toHaveAttribute('data-workspace-phase', 'ready');
    await focusCanvas(page);
    await expect(page.getByTestId('formula-input')).toHaveValue('=1+2');
    expect(apiRequests).toEqual([]);
    expect(socketCount).toBe(0);
  });
});
