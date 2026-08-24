# Instructions

- Following Playwright test failed.
- Explain why, be concise, respect Playwright best practices.
- Provide a snippet of code with the fix, if possible.

# Test info

- Name: spreadsheet.spec.ts >> spreadsheet baseline >> Format Painter enters a transient Home state and completes after one target selection
- Location: e2e\spreadsheet.spec.ts:431:3

# Error details

```
Test timeout of 30000ms exceeded.
```

```
Error: locator.click: Test timeout of 30000ms exceeded.
Call log:
  - waiting for getByTestId('home-format-painter')

```

# Page snapshot

```yaml
- application "Spreadsheet Designer" [ref=e3]:
  - generic [ref=e5]:
    - generic [ref=e6]:
      - button "Open workbook menu" [ref=e9] [cursor=pointer]: 文件
      - tablist "Workbook ribbon tabs" [ref=e10]:
        - tab "Home" [selected] [ref=e11] [cursor=pointer]
        - tab "Insert" [ref=e12] [cursor=pointer]
        - tab "Page Layout" [ref=e13] [cursor=pointer]
        - tab "Formulas" [ref=e14] [cursor=pointer]
        - tab "Data" [ref=e15] [cursor=pointer]
        - tab "Review" [ref=e16] [cursor=pointer]
        - tab "View" [ref=e17] [cursor=pointer]
        - tab "Settings" [ref=e18] [cursor=pointer]
        - tab "Automate" [ref=e19] [cursor=pointer]
      - generic [ref=e20]: Engine Connected
    - generic [ref=e26]:
      - generic [ref=e27]:
        - generic [ref=e29]:
          - button "Undo (Ctrl+Z)" [ref=e30] [cursor=pointer]
          - button "Redo (Ctrl+Y)" [ref=e34] [cursor=pointer]
        - generic: History
      - separator
      - generic [ref=e38]:
        - generic [ref=e40]:
          - button "Paste" [ref=e41] [cursor=pointer]
          - generic [ref=e45]:
            - button "Cut (Ctrl+X)" [ref=e46] [cursor=pointer]
            - button "Copy (Ctrl+C)" [ref=e51] [cursor=pointer]
            - button "Format Painter" [ref=e55] [cursor=pointer]
        - generic: Clipboard
      - separator
      - generic [ref=e58]:
        - generic [ref=e60]:
          - generic [ref=e61]:
            - combobox "Font family" [ref=e64]:
              - option "微软雅黑" [selected]
              - option "Arial"
              - option "Calibri"
              - option "Segoe UI"
              - option "Times New Roman"
            - textbox "Font size" [ref=e66]: "11"
            - button "Increase font size" [ref=e67] [cursor=pointer]: A
            - button "Decrease font size" [ref=e68] [cursor=pointer]: A
          - generic [ref=e69]:
            - button "Bold (Ctrl+B)" [pressed] [ref=e70] [cursor=pointer]
            - button "Italic (Ctrl+I)" [ref=e73] [cursor=pointer]
            - button "Underline (Ctrl+U)" [ref=e76] [cursor=pointer]
            - button "Strikethrough" [ref=e79] [cursor=pointer]
            - button "All Borders" [ref=e82] [cursor=pointer]
            - button "Text color" [ref=e88] [cursor=pointer]
            - button "Fill background" [ref=e93] [cursor=pointer]
        - generic: Font
      - separator
      - generic [ref=e98]:
        - generic [ref=e100]:
          - generic [ref=e101]:
            - generic [ref=e102]:
              - button "Align Left" [ref=e103] [cursor=pointer]
              - button "Align Center" [ref=e105] [cursor=pointer]
              - button "Align Right" [ref=e107] [cursor=pointer]
            - generic [ref=e109]:
              - button "Align Top" [ref=e110] [cursor=pointer]
              - button "Align Middle" [ref=e113] [cursor=pointer]
              - button "Align Bottom" [ref=e116] [cursor=pointer]
          - generic [ref=e119]:
            - button "Wrap text" [ref=e120] [cursor=pointer]
            - button "Merge and center" [ref=e124] [cursor=pointer]: 合并后居中
            - generic [ref=e129]:
              - button "Decrease Indent" [ref=e130] [cursor=pointer]
              - button "Increase Indent" [ref=e134] [cursor=pointer]
              - button "Text Orientation" [ref=e138] [cursor=pointer]
        - generic: Alignment
      - separator
      - generic [ref=e141]:
        - generic [ref=e143]:
          - combobox "Number" [ref=e145]:
            - option "常规" [selected]
            - option "货币"
            - option "百分比"
            - option "千位分隔"
            - option "小数"
          - generic [ref=e146]:
            - button "Currency Format" [ref=e147] [cursor=pointer]
            - button "Percent Format" [ref=e150] [cursor=pointer]
            - button "Comma Format" [ref=e155] [cursor=pointer]
            - button "Decimal formats" [ref=e161] [cursor=pointer]
        - generic: Number
      - separator
      - generic [ref=e166]:
        - generic [ref=e168]:
          - button "Conditional Format" [ref=e169] [cursor=pointer]
          - button "Format as Table" [ref=e173] [cursor=pointer]
          - button "Format Cells" [ref=e177] [cursor=pointer]
          - button "Cell Style Template" [ref=e181] [cursor=pointer]
          - button "Cell Editor" [ref=e184] [cursor=pointer]
        - generic: Styles
      - separator
      - generic [ref=e189]:
        - generic [ref=e191]:
          - button "插入" [ref=e194] [cursor=pointer]
          - button "删除" [ref=e199] [cursor=pointer]
          - button "Format cells" [ref=e206] [cursor=pointer]: 格式
        - generic: Cells
      - separator
      - generic [ref=e210]:
        - generic [ref=e212]:
          - button "AutoSum =SUM()" [ref=e213] [cursor=pointer]
          - button "Fill Down" [ref=e217] [cursor=pointer]
          - button "More editing commands" [ref=e223] [cursor=pointer]: 编辑
        - generic: Editing
  - form "Formula bar" [ref=e229]:
    - textbox "Selected cell" [ref=e231]: A1
    - button "Open Name Manager" [ref=e232] [cursor=pointer]
    - button "Cancel formula edit" [ref=e237] [cursor=pointer]
    - button "Apply formula" [ref=e240] [cursor=pointer]
    - button "Insert Function Wizard" [ref=e243] [cursor=pointer]: fx
    - textbox "Formula input" [ref=e248]:
      - /placeholder: ""
  - generic [ref=e249]:
    - grid "Spreadsheet canvas" [active] [ref=e256]:
      - scrollbar "horizontal scrollbar" [ref=e258]
      - scrollbar "vertical scrollbar" [ref=e260]
    - navigation "Worksheets" [ref=e262]:
      - navigation "Worksheets" [ref=e263]:
        - generic [ref=e264]:
          - button "Scroll worksheets left" [ref=e265] [cursor=pointer]
          - button "Scroll worksheets right" [ref=e268] [cursor=pointer]
          - button "Worksheet list" [ref=e273] [cursor=pointer]
          - tab "Sheet1" [selected] [ref=e280] [cursor=pointer]
          - button "Add worksheet" [ref=e282] [cursor=pointer]
          - button "More worksheet actions" [ref=e287] [cursor=pointer]
  - contentinfo [ref=e295]:
    - generic "Workbook status bar" [ref=e296]:
      - generic [ref=e297]: 就绪
      - button "Open keyboard shortcuts" [ref=e298] [cursor=pointer]: 快捷键
      - generic [ref=e299]:
        - button "Zoom out" [ref=e300] [cursor=pointer]: −
        - button "Zoom in" [ref=e303] [cursor=pointer]: +
        - generic [ref=e304]: 100%
```

# Test source

```ts
  337 |     const box = await canvas.boundingBox();
  338 |     if (!box) throw new Error('Spreadsheet canvas has no bounds');
  339 |     await page.mouse.move(box.x + ROW_HEADER_WIDTH_PX + DEFAULT_COLUMN_WIDTH_PX - 4, box.y + COLUMN_HEADER_HEIGHT_PX + DEFAULT_ROW_HEIGHT_PX - 4);
  340 |     await page.mouse.down();
  341 |     await page.mouse.move(...Object.values(cellPoint(box, 2, 2)) as [number, number]);
  342 |     await page.mouse.up();
  343 |     await expect(page.getByTestId('name-box')).toHaveValue('C3');
  344 |     await page.getByTestId('name-box').fill('C3');
  345 |     await page.getByTestId('name-box').press('Enter');
  346 |     await expect(page.getByTestId('formula-input')).toHaveValue('fill-source-unique');
  347 |   });
  348 | 
  349 |   test('local formula calculation and IndexedDB restore work without an API request', async ({ page }) => {
  350 |     const apiRequests: string[] = [];
  351 |     let socketCount = 0;
  352 |     page.on('request', (request) => {
  353 |       if (new URL(request.url()).pathname.startsWith('/api/')) apiRequests.push(request.url());
  354 |     });
  355 |     page.on('websocket', (socket) => {
  356 |       if (new URL(socket.url()).pathname === '/ws') socketCount += 1;
  357 |     });
  358 |     await waitForWorkspace(page);
  359 |     const canvas = await focusCanvas(page);
  360 |     await page.keyboard.type('=1+2');
  361 |     await page.keyboard.press('Enter');
  362 |     await expect(page.getByTestId('formula-input')).toHaveValue('');
  363 |     await canvas.press('ArrowUp');
  364 |     await expect(page.getByTestId('formula-input')).toHaveValue('=1+2');
  365 | 
  366 |     await page.reload();
  367 |   await expect(page.getByTestId('designer-shell')).toHaveAttribute('data-workspace-phase', 'ready');
  368 |     await focusCanvas(page);
  369 |     await expect(page.getByTestId('formula-input')).toHaveValue('=1+2');
  370 |     expect(apiRequests).toEqual([]);
  371 |     expect(socketCount).toBe(0);
  372 |   });
  373 | 
  374 |   test('Home ribbon opens shared format, sort, find, and paste dialogs without rendering Add-ins', async ({ page }) => {
  375 |     await waitForWorkspace(page);
  376 |     await page.getByTestId('ribbon-tab-home').click();
  377 |     await expect(page.getByTestId('home-ribbon-groups')).toBeVisible();
  378 |     await expect(page.getByRole('button', { name: /add-ins/i })).toHaveCount(0);
  379 | 
  380 |     await page.getByTestId('ribbon-format-cells').click();
  381 |     const formatDialog = page.getByTestId('format-cells-dialog');
  382 |     await expect(formatDialog).toBeVisible();
  383 |     await formatDialog.getByTestId('format-tab-font').click();
  384 |     await expect(formatDialog.getByLabel(/(font size|字号)/i)).toBeVisible();
  385 |     await formatDialog.getByRole('button', { name: /^(Close|关闭)$/ }).click();
  386 | 
  387 |     await page.getByRole('button', { name: /(sort range|排序区域)/i }).click();
  388 |     const sortDialog = page.getByTestId('sort-dialog');
  389 |     await expect(sortDialog).toBeVisible();
  390 |     await sortDialog.getByRole('button', { name: /^(Close|关闭)$/ }).click();
  391 | 
  392 |     await page.getByRole('button', { name: /(find & replace|查找与替换)/i }).click();
  393 |     const findDialog = page.getByTestId('find-replace-dialog');
  394 |     await expect(findDialog).toBeVisible();
  395 |     await findDialog.getByLabel(/^(Find|查找)$/).fill('home-dialog-check');
  396 |     await expect(findDialog.getByRole('button', { name: /(replace all|全部替换)/i })).toBeEnabled();
  397 |     await findDialog.getByRole('button', { name: /^(Close|关闭)$/ }).click();
  398 | 
  399 |     await page.getByRole('button', { name: /(paste special|选择性粘贴)/i }).click();
  400 |     const pasteDialog = page.getByTestId('paste-special-dialog');
  401 |     await expect(pasteDialog).toBeVisible();
  402 |     await expect(pasteDialog.getByTestId('paste-special-formats')).toBeVisible();
  403 |     await pasteDialog.getByRole('button', { name: /^(Cancel|取消)$/ }).click();
  404 | 
  405 |     await page.getByTestId('home-selection-pane').click();
  406 |     await expect(page.getByTestId('selection-pane')).toBeVisible();
  407 |   });
  408 | 
  409 |   test('Selection Pane selects, renames, and toggles a drawing through host callbacks', async ({ page }) => {
  410 |     await waitForWorkspace(page);
  411 |     await page.getByTestId('ribbon-tab-insert').click();
  412 |     await page.getByRole('button', { name: /(rectangle|矩形)/i }).click();
  413 | 
  414 |     await page.getByTestId('ribbon-tab-home').click();
  415 |     await page.getByTestId('home-selection-pane').click();
  416 |     const pane = page.getByTestId('selection-pane');
  417 |     await expect(pane).toBeVisible();
  418 | 
  419 |     await pane.getByRole('button', { name: /^(Rename|重命名)$/ }).click();
  420 |     const rename = pane.getByLabel(/^(Rename|重命名)$/);
  421 |     await rename.fill('KPI tile');
  422 |     await rename.press('Enter');
  423 |     await expect(pane.getByRole('button', { name: 'KPI tile', exact: true })).toBeVisible();
  424 | 
  425 |     const visibility = pane.getByLabel(/KPI tile: (Visible|可见)/);
  426 |     await expect(visibility).toBeChecked();
  427 |     await visibility.uncheck();
  428 |     await expect(visibility).not.toBeChecked();
  429 |   });
  430 | 
  431 |   test('Format Painter enters a transient Home state and completes after one target selection', async ({ page }) => {
  432 |     await waitForWorkspace(page);
  433 |     const canvas = await focusCanvas(page);
  434 |     await canvas.press('Control+B');
  435 | 
  436 |     const painter = page.getByTestId('home-format-painter');
> 437 |     await painter.click();
      |                   ^ Error: locator.click: Test timeout of 30000ms exceeded.
  438 |     await expect(painter).toHaveAttribute('aria-pressed', 'true');
  439 | 
  440 |     await canvas.click({ position: { x: 142, y: 34 } });
  441 |     await expect(painter).toHaveAttribute('aria-pressed', 'false');
  442 |   });
  443 | });
  444 | 
```