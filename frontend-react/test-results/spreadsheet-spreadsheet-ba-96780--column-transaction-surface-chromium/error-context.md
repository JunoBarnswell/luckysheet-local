# Instructions

- Following Playwright test failed.
- Explain why, be concise, respect Playwright best practices.
- Provide a snippet of code with the fix, if possible.

# Test info

- Name: spreadsheet.spec.ts >> spreadsheet baseline >> Excel-style column width paths share one multi-column transaction surface
- Location: e2e\spreadsheet.spec.ts:266:3

# Error details

```
Test timeout of 30000ms exceeded.
```

```
Error: locator.click: Test timeout of 30000ms exceeded.
Call log:
  - waiting for getByRole('button', { name: 'Format', exact: true })

```

# Page snapshot

```yaml
- application "Spreadsheet Designer" [ref=e3]:
  - generic [ref=e5]:
    - generic [ref=e6]:
      - button "Open workbook menu" [ref=e9] [cursor=pointer]: 文件
      - tablist "Workbook ribbon tabs" [ref=e10]:
        - tab "Home" [active] [selected] [ref=e11] [cursor=pointer]
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
            - button "Bold (Ctrl+B)" [ref=e70] [cursor=pointer]
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
    - textbox "Selected cell" [ref=e231]: B1
    - button "Open Name Manager" [ref=e232] [cursor=pointer]
    - button "Cancel formula edit" [ref=e237] [cursor=pointer]
    - button "Apply formula" [ref=e240] [cursor=pointer]
    - button "Insert Function Wizard" [ref=e243] [cursor=pointer]: fx
    - textbox "Formula input" [ref=e248]:
      - /placeholder: ""
  - generic [ref=e249]:
    - grid "Spreadsheet canvas" [ref=e256]:
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
  196 |     await waitForWorkspace(page);
  197 |     const canvas = await focusCanvas(page);
  198 |     const box = await canvas.boundingBox();
  199 |     if (!box) throw new Error('Spreadsheet canvas has no bounds');
  200 |     await page.mouse.move(...Object.values(cellPoint(box, 1, 0)) as [number, number]);
  201 |     await page.mouse.down();
  202 |     await page.mouse.move(...Object.values(cellPoint(box, 8, 2)) as [number, number]);
  203 |     await page.mouse.up();
  204 |     await expect(page.getByTestId('name-box')).toHaveValue('C9');
  205 | 
  206 |     await canvas.press('F2');
  207 |     const editor = page.getByLabel('Cell editor');
  208 |     await expect(editor).toBeFocused();
  209 |     const editorBox = await editor.boundingBox();
  210 |     if (!editorBox) throw new Error('Cell editor has no bounds');
  211 |     expect(editorBox.x - box.x).toBeGreaterThan(160);
  212 |     expect(editorBox.y - box.y).toBeGreaterThan(180);
  213 |     await editor.fill('4');
  214 |     await editor.press('Enter');
  215 |     await canvas.press('ArrowUp');
  216 |     await expect(page.getByTestId('name-box')).toHaveValue('C9');
  217 |     await expect(page.getByTestId('formula-input')).toHaveValue('4');
  218 |   });
  219 | 
  220 |   test('a dragged selection accepts direct typing at the release cell', async ({ page }) => {
  221 |     await waitForWorkspace(page);
  222 |     const canvas = await focusCanvas(page);
  223 |     await dragCells(page, canvas, { row: 1, column: 0 }, { row: 8, column: 2 });
  224 |     await expect(page.getByTestId('name-box')).toHaveValue('C9');
  225 |     await page.keyboard.type('direct-c9');
  226 |     await canvas.press('Enter');
  227 |     await canvas.press('ArrowUp');
  228 |     await expect(page.getByTestId('name-box')).toHaveValue('C9');
  229 |     await expect(page.getByTestId('formula-input')).toHaveValue('direct-c9');
  230 |   });
  231 | 
  232 |   test('reverse drag keeps the release cell active for F2 editing', async ({ page }) => {
  233 |     await waitForWorkspace(page);
  234 |     const canvas = await focusCanvas(page);
  235 |     await dragCells(page, canvas, { row: 8, column: 2 }, { row: 1, column: 0 });
  236 |     await expect(page.getByTestId('name-box')).toHaveValue('A2');
  237 |     await canvas.press('F2');
  238 |     const editor = page.getByLabel('Cell editor');
  239 |     await expect(editor).toBeFocused();
  240 |     await editor.fill('reverse-a2');
  241 |     await editor.press('Enter');
  242 |     await canvas.press('ArrowUp');
  243 |     await expect(page.getByTestId('name-box')).toHaveValue('A2');
  244 |     await expect(page.getByTestId('formula-input')).toHaveValue('reverse-a2');
  245 |   });
  246 | 
  247 |   test('row and column header drags commit the final header target', async ({ page }) => {
  248 |     await waitForWorkspace(page);
  249 |     const canvas = await focusCanvas(page);
  250 |     const box = await canvas.boundingBox();
  251 |     if (!box) throw new Error('Spreadsheet canvas has no bounds');
  252 | 
  253 |     await page.mouse.move(box.x + 20, box.y + COLUMN_HEADER_HEIGHT_PX + DEFAULT_ROW_HEIGHT_PX + DEFAULT_ROW_HEIGHT_PX / 2);
  254 |     await page.mouse.down();
  255 |     await page.mouse.move(box.x + 20, box.y + COLUMN_HEADER_HEIGHT_PX + 8 * DEFAULT_ROW_HEIGHT_PX + DEFAULT_ROW_HEIGHT_PX / 2);
  256 |     await page.mouse.up();
  257 |     await expect(page.getByTestId('name-box')).toHaveValue('A9');
  258 | 
  259 |     await page.mouse.move(box.x + ROW_HEADER_WIDTH_PX + 55, box.y + 12);
  260 |     await page.mouse.down();
  261 |     await page.mouse.move(box.x + ROW_HEADER_WIDTH_PX + 2 * DEFAULT_COLUMN_WIDTH_PX + DEFAULT_COLUMN_WIDTH_PX / 2, box.y + 12);
  262 |     await page.mouse.up();
  263 |     await expect(page.getByTestId('name-box')).toHaveValue('C1');
  264 |   });
  265 | 
  266 |   test('Excel-style column width paths share one multi-column transaction surface', async ({ page }) => {
  267 |     await waitForWorkspace(page);
  268 |     const canvas = await focusCanvas(page);
  269 |     const box = await canvas.boundingBox();
  270 |     if (!box) throw new Error('Spreadsheet canvas has no bounds');
  271 | 
  272 |     // Full-column A:C selection.
  273 |     await page.mouse.move(box.x + ROW_HEADER_WIDTH_PX + DEFAULT_COLUMN_WIDTH_PX / 2, box.y + 12);
  274 |     await page.mouse.down();
  275 |     await page.mouse.move(box.x + ROW_HEADER_WIDTH_PX + 2 * DEFAULT_COLUMN_WIDTH_PX + DEFAULT_COLUMN_WIDTH_PX / 2, box.y + 12);
  276 |     await page.mouse.up();
  277 | 
  278 |     // Drag the C boundary; the preview exposes both character and pixel units.
  279 |     const boundary = box.x + ROW_HEADER_WIDTH_PX + 3 * DEFAULT_COLUMN_WIDTH_PX;
  280 |     await page.mouse.move(boundary, box.y + 12);
  281 |     await page.mouse.down();
  282 |     await page.mouse.move(boundary + 24, box.y + 12);
  283 |     await expect(canvas).toBeVisible();
  284 |     await page.mouse.up();
  285 | 
  286 |     // Right-click keeps the complete multi-column selection and opens exact width.
  287 |     await page.mouse.click(box.x + ROW_HEADER_WIDTH_PX + DEFAULT_COLUMN_WIDTH_PX + DEFAULT_COLUMN_WIDTH_PX / 2, box.y + 12, { button: 'right' });
  288 |     await page.getByRole('menuitem', { name: 'Column Width…' }).click();
  289 |     const dialog = page.getByRole('dialog', { name: 'Column Width' });
  290 |     await expect(dialog).toBeVisible();
  291 |     await dialog.getByLabel('Excel character width').fill('12');
  292 |     await dialog.getByRole('button', { name: 'OK' }).click();
  293 | 
  294 |     // Ribbon and context-menu entries route through the same controller.
  295 |     await page.getByTestId('ribbon-tab-home').click();
> 296 |     await page.getByRole('button', { name: 'Format', exact: true }).click();
      |                                                                     ^ Error: locator.click: Test timeout of 30000ms exceeded.
  297 |     await expect(page.getByRole('button', { name: 'AutoFit Column Width', exact: true })).toBeVisible();
  298 |     await page.getByRole('button', { name: 'AutoFit Column Width', exact: true }).click();
  299 | 
  300 |     await page.mouse.click(box.x + ROW_HEADER_WIDTH_PX + DEFAULT_COLUMN_WIDTH_PX / 2, box.y + 12, { button: 'right' });
  301 |     await page.getByRole('menuitem', { name: 'Hide Columns', exact: true }).click();
  302 |     await canvas.press('Control+Z');
  303 |   });
  304 | 
  305 |   test('right click changes the command target before the context menu opens', async ({ page }) => {
  306 |     await waitForWorkspace(page);
  307 |     await page.getByRole('button', { name: /add worksheet|添加工作表/i }).click();
  308 |     const canvas = await focusCanvas(page);
  309 |     const box = await canvas.boundingBox();
  310 |     if (!box) throw new Error('Spreadsheet canvas has no bounds');
  311 |     await page.getByTestId('name-box').fill('B2');
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
```