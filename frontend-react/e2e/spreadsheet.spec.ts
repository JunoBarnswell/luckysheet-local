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
    await page.keyboard.press('ArrowUp');
    await expect(page.getByTestId('formula-input')).toHaveValue('hello-e2e');
  });

  test('clipboard copy and paste round-trip', async ({ page, context }) => {
    await context.grantPermissions(['clipboard-read', 'clipboard-write']);
    await waitForWorkspace(page);
    const canvas = await focusCanvas(page);
    await page.keyboard.type('copy-me');
    await page.keyboard.press('Enter');
    await page.keyboard.press('ArrowUp');
    await page.keyboard.press('Control+C');
    await page.keyboard.press('Control+V');
    await page.keyboard.press('ArrowLeft');
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
    await page.keyboard.press('ArrowUp');
    await expect(page.getByTestId('formula-input')).toHaveValue('undo-me');
    await page.keyboard.press('Control+Z');
    await expect(page.getByTestId('formula-input')).toHaveValue('');
  });
});
