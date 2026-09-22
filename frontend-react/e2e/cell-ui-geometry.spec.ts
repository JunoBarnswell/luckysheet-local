import { expect, test, type Page } from '@playwright/test';

const ROW_HEADER_WIDTH = 39;
const COLUMN_HEADER_HEIGHT = 20;
const COLUMN_WIDTH = 64;
const ROW_HEIGHT = 20;

async function openWorkbook(page: Page): Promise<{ canvas: ReturnType<Page['getByTestId']>; box: NonNullable<Awaited<ReturnType<ReturnType<Page['getByTestId']>['boundingBox']>>> }> {
  await page.goto('/');
  await expect(page.getByTestId('workbook-hub')).toBeVisible();
  await page.getByRole('button', { name: '新建工作簿' }).click();
  const dialog = page.getByTestId('create-workbook-dialog');
  await expect(dialog).toBeVisible();
  await dialog.getByLabel('工作簿名称').fill(`Cell UI ${Date.now()}`);
  await dialog.getByLabel('保存位置').selectOption('local');
  await dialog.getByRole('button', { name: '创建工作簿' }).click();
  await expect(page).toHaveURL(/\/workbooks\/[^/]+(?:\?.*)?$/);
  await expect(page.getByTestId('designer-shell')).toHaveAttribute('data-workspace-phase', 'ready', { timeout: 30_000 });
  const canvas = page.getByTestId('sheet-canvas');
  const box = await canvas.boundingBox();
  if (!box) throw new Error('Spreadsheet canvas has no bounds');
  return { canvas, box };
}

function cellCenter(box: { x: number; y: number }, row: number, column: number): { x: number; y: number } {
  return {
    x: box.x + ROW_HEADER_WIDTH + column * COLUMN_WIDTH + COLUMN_WIDTH / 2,
    y: box.y + COLUMN_HEADER_HEIGHT + row * ROW_HEIGHT + ROW_HEIGHT / 2,
  };
}

function columnHeaderPoint(box: { x: number; y: number }, column: number): { x: number; y: number } {
  return { x: box.x + ROW_HEADER_WIDTH + column * COLUMN_WIDTH + COLUMN_WIDTH / 2, y: box.y + COLUMN_HEADER_HEIGHT / 2 };
}

function rowHeaderPoint(box: { x: number; y: number }, row: number): { x: number; y: number } {
  return { x: box.x + ROW_HEADER_WIDTH / 2, y: box.y + COLUMN_HEADER_HEIGHT + row * ROW_HEIGHT + ROW_HEIGHT / 2 };
}

test.describe('Issue 316 cell layout and header semantics', () => {
  test('edit surface uses measured width, crosses occupied neighbours, and removes clipping on cancel', async ({ page }) => {
    const { canvas, box } = await openWorkbook(page);
    const nameBox = page.getByTestId('name-box');

    await nameBox.fill('C1');
    await nameBox.press('Enter');
    await canvas.focus();
    await page.keyboard.type('occupied neighbour');
    await page.keyboard.press('Enter');

    await nameBox.fill('B1');
    await nameBox.press('Enter');
    await canvas.press('F2');
    const editor = page.getByLabel('Cell editor');
    await expect(editor).toBeFocused();
    await editor.fill('A measured draft that is wider than a cell');
    const editorBox = await editor.boundingBox();
    if (!editorBox) throw new Error('Cell editor has no bounds');
    expect(editorBox.width).toBeGreaterThan(COLUMN_WIDTH);
    expect(await editor.evaluate((node) => getComputedStyle(node).overflow)).not.toBe('hidden');
    expect(editorBox.x).toBeGreaterThanOrEqual(box.x);

    await editor.press('Escape');
    await expect(editor).toHaveCount(0);
    await nameBox.fill('C1');
    await nameBox.press('Enter');
    await expect(page.getByTestId('formula-input')).toHaveValue('occupied neighbour');
  });

  test('multiline editor grows from real line layout and keeps the canonical newline after commit', async ({ page }) => {
    const { canvas } = await openWorkbook(page);
    await canvas.click({ position: { x: ROW_HEADER_WIDTH + COLUMN_WIDTH / 2, y: COLUMN_HEADER_HEIGHT + ROW_HEIGHT / 2 } });
    await canvas.press('F2');
    const editor = page.getByLabel('Cell editor');
    await editor.fill('line one');
    await editor.press('Alt+Enter');
    await editor.type('line two');
    const editorBox = await editor.boundingBox();
    if (!editorBox) throw new Error('Multiline cell editor has no bounds');
    expect(editorBox.height).toBeGreaterThan(ROW_HEIGHT);
    expect(await editor.evaluate((node) => getComputedStyle(node).overflow)).not.toBe('hidden');
    await expect(page.getByTestId('formula-input')).toHaveValue('line one\nline two');
    await editor.press('Enter');
    await canvas.press('ArrowUp');
    await expect(page.getByTestId('formula-input')).toHaveValue('line one\nline two');
  });

  test('column and row header selections expose symmetric structural context actions', async ({ page }) => {
    const { canvas, box } = await openWorkbook(page);
    const columnA = columnHeaderPoint(box, 0);
    const columnC = columnHeaderPoint(box, 2);
    await page.mouse.click(columnA.x, columnA.y);
    await page.keyboard.down('Control');
    await page.mouse.click(columnC.x, columnC.y);
    await page.keyboard.up('Control');
    await page.mouse.click(columnC.x, columnC.y, { button: 'right' });
    const menu = page.getByRole('menu');
    await expect(menu).toBeVisible();
    for (const label of ['Cut', 'Copy', 'Paste', 'Paste Special…', 'Insert Columns', 'Delete Columns', 'Clear Contents', 'Format Cells…', 'Column Width…', 'Hide Columns', 'Unhide Columns']) {
      await expect(menu.getByRole('menuitem', { name: label, exact: true })).toBeVisible();
    }
    await page.keyboard.press('Escape');

    const rowA = rowHeaderPoint(box, 0);
    const rowC = rowHeaderPoint(box, 2);
    await page.mouse.click(rowA.x, rowA.y);
    await page.keyboard.down('Shift');
    await page.mouse.click(rowC.x, rowC.y);
    await page.keyboard.up('Shift');
    await page.mouse.click(rowC.x, rowC.y, { button: 'right' });
    await expect(menu).toBeVisible();
    for (const label of ['Cut', 'Copy', 'Paste', 'Paste Special…', 'Insert Rows', 'Delete Rows', 'Clear Contents', 'Format Cells…', 'Row Height…', 'Hide Rows', 'Unhide Rows']) {
      await expect(menu.getByRole('menuitem', { name: label, exact: true })).toBeVisible();
    }
    await canvas.press('Escape');
  });
});
