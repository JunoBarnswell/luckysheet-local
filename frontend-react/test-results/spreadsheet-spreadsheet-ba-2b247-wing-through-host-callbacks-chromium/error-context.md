# Instructions

- Following Playwright test failed.
- Explain why, be concise, respect Playwright best practices.
- Provide a snippet of code with the fix, if possible.

# Test info

- Name: spreadsheet.spec.ts >> spreadsheet baseline >> Selection Pane selects, renames, and toggles a drawing through host callbacks
- Location: e2e\spreadsheet.spec.ts:409:3

# Error details

```
Test timeout of 30000ms exceeded.
```

```
Error: locator.click: Test timeout of 30000ms exceeded.
Call log:
  - waiting for getByRole('button', { name: /(rectangle|矩形)/i })

```

# Page snapshot

```yaml
- application "Spreadsheet Designer" [ref=e3]:
  - generic [ref=e5]:
    - generic [ref=e6]:
      - button "Open workbook menu" [ref=e9] [cursor=pointer]: 文件
      - tablist "Workbook ribbon tabs" [ref=e10]:
        - tab "Home" [ref=e11] [cursor=pointer]
        - tab "Insert" [active] [selected] [ref=e12] [cursor=pointer]
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
        - generic [ref=e28]:
          - button "TableSheet" [ref=e29] [cursor=pointer]
          - button "GanttSheet" [ref=e33] [cursor=pointer]
          - button "ReportSheet" [ref=e40] [cursor=pointer]
        - generic [ref=e45]: Sheets
      - separator [ref=e46]
      - generic [ref=e47]:
        - generic [ref=e48]:
          - button "Table" [ref=e49] [cursor=pointer]
          - button "Pivot Table" [ref=e53] [cursor=pointer]
        - generic [ref=e57]: Tables
      - separator [ref=e58]
      - generic [ref=e59]:
        - generic [ref=e60]:
          - button "图表" [ref=e63] [cursor=pointer]
          - generic [ref=e68]:
            - generic [ref=e69]:
              - button "柱形图" [ref=e70] [cursor=pointer]
              - button "条形图" [ref=e75] [cursor=pointer]
              - button "折线图" [ref=e80] [cursor=pointer]
            - generic [ref=e86]:
              - button "面积图" [ref=e87] [cursor=pointer]
              - button "饼图" [ref=e91] [cursor=pointer]
              - button "散点图" [ref=e95] [cursor=pointer]
          - button "Barcode" [ref=e101] [cursor=pointer]
          - button "迷你图" [ref=e107] [cursor=pointer]
        - generic [ref=e111]: Charts
      - separator [ref=e112]
      - generic [ref=e113]:
        - button "Data Chart" [ref=e115] [cursor=pointer]
        - generic [ref=e121]: Data Charts
      - separator [ref=e122]
      - generic [ref=e123]:
        - generic [ref=e124]:
          - button "Picture" [ref=e125] [cursor=pointer]
          - button "形状" [ref=e132] [cursor=pointer]
          - button "Camera" [ref=e135] [cursor=pointer]
          - button "控件" [ref=e141] [cursor=pointer]
        - generic [ref=e147]: Illustrations
      - separator [ref=e148]
      - generic [ref=e149]:
        - button "Hyperlink" [ref=e151] [cursor=pointer]
        - generic [ref=e155]: Links
      - separator [ref=e156]
      - generic [ref=e157]:
        - generic [ref=e158]:
          - button "Checkbox" [ref=e159] [cursor=pointer]
          - button "Text Box" [ref=e163] [cursor=pointer]
        - generic [ref=e167]: Controls
  - form "Formula bar" [ref=e169]:
    - textbox "Selected cell" [ref=e171]: A1
    - button "Open Name Manager" [ref=e172] [cursor=pointer]
    - button "Cancel formula edit" [ref=e177] [cursor=pointer]
    - button "Apply formula" [ref=e180] [cursor=pointer]
    - button "Insert Function Wizard" [ref=e183] [cursor=pointer]: fx
    - textbox "Formula input" [ref=e188]:
      - /placeholder: ""
  - generic [ref=e189]:
    - grid "Spreadsheet canvas" [ref=e196]:
      - scrollbar "horizontal scrollbar" [ref=e198]
      - scrollbar "vertical scrollbar" [ref=e200]
    - navigation "Worksheets" [ref=e202]:
      - navigation "Worksheets" [ref=e203]:
        - generic [ref=e204]:
          - button "Scroll worksheets left" [ref=e205] [cursor=pointer]
          - button "Scroll worksheets right" [ref=e208] [cursor=pointer]
          - button "Worksheet list" [ref=e213] [cursor=pointer]
          - tab "Sheet1" [selected] [ref=e220] [cursor=pointer]
          - button "Add worksheet" [ref=e222] [cursor=pointer]
          - button "More worksheet actions" [ref=e227] [cursor=pointer]
  - contentinfo [ref=e235]:
    - generic "Workbook status bar" [ref=e236]:
      - generic [ref=e237]: 就绪
      - button "Open keyboard shortcuts" [ref=e238] [cursor=pointer]: 快捷键
      - generic [ref=e239]:
        - button "Zoom out" [ref=e240] [cursor=pointer]: −
        - button "Zoom in" [ref=e243] [cursor=pointer]: +
        - generic [ref=e244]: 100%
```

# Test source

```ts
  312 |     await page.getByTestId('name-box').press('Enter');
  313 |     await canvas.focus();
  314 |     await page.keyboard.type('right-click-target');
  315 |     await page.keyboard.press('Enter');
  316 |     await page.getByTestId('name-box').fill('A1');
  317 |     await page.getByTestId('name-box').press('Enter');
  318 |     const target = cellPoint(box, 1, 1);
  319 |     await page.mouse.click(target.x, target.y, { button: 'right' });
  320 |     await expect(page.getByRole('menu')).toBeVisible();
  321 |     await expect(page.getByTestId('name-box')).toHaveValue('B2');
  322 |     await page.getByRole('menuitem', { name: 'Clear contents' }).click();
  323 |     await page.getByTestId('name-box').fill('B2');
  324 |     await page.getByTestId('name-box').press('Enter');
  325 |     await expect(page.getByTestId('formula-input')).toHaveValue('');
  326 |   });
  327 | 
  328 |   test('fill handle follows the primary range bottom-right', async ({ page }) => {
  329 |     await waitForWorkspace(page);
  330 |     await page.getByRole('button', { name: /add worksheet|添加工作表/i }).click();
  331 |     const canvas = await focusCanvas(page);
  332 |     await page.keyboard.type('fill-source-unique');
  333 |     await page.keyboard.press('Enter');
  334 |     await page.getByTestId('name-box').fill('A1');
  335 |     await page.getByTestId('name-box').press('Enter');
  336 |     await canvas.focus();
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
> 412 |     await page.getByRole('button', { name: /(rectangle|矩形)/i }).click();
      |                                                                 ^ Error: locator.click: Test timeout of 30000ms exceeded.
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
  437 |     await painter.click();
  438 |     await expect(painter).toHaveAttribute('aria-pressed', 'true');
  439 | 
  440 |     await canvas.click({ position: { x: 142, y: 34 } });
  441 |     await expect(painter).toHaveAttribute('aria-pressed', 'false');
  442 |   });
  443 | });
  444 | 
```