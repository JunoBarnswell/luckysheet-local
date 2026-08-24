# Instructions

- Following Playwright test failed.
- Explain why, be concise, respect Playwright best practices.
- Provide a snippet of code with the fix, if possible.

# Test info

- Name: spreadsheet.spec.ts >> spreadsheet baseline >> clicking another cell commits the old editor before accepting the new value
- Location: e2e\spreadsheet.spec.ts:118:3

# Error details

```
Error: expect(locator).toBeVisible() failed

Locator: getByLabel('Cell editor')
Expected: visible
Error: strict mode violation: getByLabel('Cell editor') resolved to 2 elements:
    1) <button type="button" title="Cell Editor" aria-label="Cell Editor" class="inline-flex shrink-0 items-center justify-center gap-2 whitespace-nowrap font-medium transition-colors focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-accent/50 focus-visible:ring-offset-2 disabled:cursor-not-allowed disabled:opacity-45 min-h-8 rounded-md px-2.5 text-xs bg-transparent text-muted hover:bg-slate-100 hover:text-ink !h-[64px] !min-h-0 !w-[68px] flex-col gap-0 rounded-none px-1 text-[10px] leading-…>…</button> aka getByRole('button', { name: 'Cell Editor' })
    2) <textarea aria-label="Cell editor" class="w-full rounded-md border border-slate-200 bg-white px-2.5 py-1.5 text-xs text-slate-700 outline-none transition-colors focus:border-blue-500 focus:ring-2 focus:ring-blue-100 h-full min-h-0 w-full resize-none overflow-hidden rounded-none border-0 bg-transparent px-1 py-0 text-[13px] leading-[inherit] text-slate-800 outline-none focus:border-0 focus:ring-0">old-cell-value</textarea> aka getByRole('textbox', { name: 'Cell editor' })

Call log:
  - Expect "toBeVisible" with timeout 5000ms
  - waiting for getByLabel('Cell editor')

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
    - textbox "Selected cell" [ref=e231]: A1
    - button "Open Name Manager" [ref=e232] [cursor=pointer]
    - button "Cancel formula edit" [ref=e237] [cursor=pointer]
    - button "Apply formula" [ref=e240] [cursor=pointer]
    - button "Insert Function Wizard" [ref=e243] [cursor=pointer]: fx
    - textbox "Formula input" [ref=e248]:
      - /placeholder: ""
      - text: old-cell-value
  - generic [ref=e249]:
    - generic [ref=e255]:
      - grid "Spreadsheet canvas" [ref=e256]:
        - scrollbar "horizontal scrollbar" [ref=e258]
        - scrollbar "vertical scrollbar" [ref=e260]
      - textbox "Cell editor" [active] [ref=e264]: old-cell-value
    - navigation "Worksheets" [ref=e265]:
      - navigation "Worksheets" [ref=e266]:
        - generic [ref=e267]:
          - button "Scroll worksheets left" [ref=e268] [cursor=pointer]
          - button "Scroll worksheets right" [ref=e271] [cursor=pointer]
          - button "Worksheet list" [ref=e276] [cursor=pointer]
          - tab "Sheet1" [selected] [ref=e283] [cursor=pointer]
          - button "Add worksheet" [ref=e285] [cursor=pointer]
          - button "More worksheet actions" [ref=e290] [cursor=pointer]
  - contentinfo [ref=e298]:
    - generic "Workbook status bar" [ref=e299]:
      - generic [ref=e300]: 就绪
      - button "Open keyboard shortcuts" [ref=e301] [cursor=pointer]: 快捷键
      - generic [ref=e302]:
        - button "Zoom out" [ref=e303] [cursor=pointer]: −
        - button "Zoom in" [ref=e306] [cursor=pointer]: +
        - generic [ref=e307]: 100%
```

# Test source

```ts
  22  | async function waitForDesignerDemo(page: Page) {
  23  |   page.on('pageerror', (error) => console.log(`[pageerror] ${error.message}\n${error.stack ?? ''}`));
  24  |   page.on('console', (message) => { if (message.type() === 'error') console.log(`[console.error] ${message.text()}`); });
  25  |   await page.addInitScript(() => window.localStorage.setItem('react-sheets:locale', 'zh-CN'));
  26  |   await page.setViewportSize({ width: 1280, height: 720 });
  27  |   await page.goto('/');
  28  |   await expect(page.getByTestId('workbook-hub')).toBeVisible();
  29  |   await page.getByRole('button', { name: 'Designer Demo' }).click();
  30  |   const createDialog = page.getByTestId('create-workbook-dialog');
  31  |   await expect(createDialog).toBeVisible();
  32  |   await createDialog.getByLabel('工作簿名称').fill(`Designer Demo E2E ${Date.now()}`);
  33  |   await createDialog.getByLabel('保存位置').selectOption('local');
  34  |   await createDialog.getByRole('button', { name: '创建工作簿' }).click();
  35  |   await expect(page).toHaveURL(/\/workbooks\/[^/]+(?:\?.*)?$/);
  36  |   await expect(page.getByTestId('designer-shell')).toBeVisible();
  37  |   await expect(page.getByTestId('designer-shell')).toHaveAttribute('data-workspace-phase', 'ready', { timeout: 30_000 });
  38  |   await expect(page.getByTestId('name-box')).toHaveValue('B1');
  39  | }
  40  | 
  41  | async function focusCanvas(page: Page) {
  42  |   const canvas = page.getByTestId('sheet-canvas');
  43  |   // A1 单元格中心：行头 39px + 列宽一半，列头 20px + 行高一半
  44  |   await canvas.click({ position: { x: 78, y: COLUMN_HEADER_HEIGHT_PX + DEFAULT_ROW_HEIGHT_PX / 2 } });
  45  |   return canvas;
  46  | }
  47  | 
  48  | function cellPoint(box: { x: number; y: number }, row: number, column: number) {
  49  |   return { x: box.x + ROW_HEADER_WIDTH_PX + column * DEFAULT_COLUMN_WIDTH_PX + DEFAULT_COLUMN_WIDTH_PX / 2, y: box.y + COLUMN_HEADER_HEIGHT_PX + row * DEFAULT_ROW_HEIGHT_PX + DEFAULT_ROW_HEIGHT_PX / 2 };
  50  | }
  51  | 
  52  | async function dragCells(page: Page, canvas: ReturnType<Page['getByTestId']>, from: { row: number; column: number }, to: { row: number; column: number }) {
  53  |   const box = await canvas.boundingBox();
  54  |   if (!box) throw new Error('Spreadsheet canvas has no bounds');
  55  |   await page.mouse.move(...Object.values(cellPoint(box, from.row, from.column)) as [number, number]);
  56  |   await page.mouse.down();
  57  |   await page.mouse.move(...Object.values(cellPoint(box, to.row, to.column)) as [number, number]);
  58  |   await page.mouse.up();
  59  |   return box;
  60  | }
  61  | 
  62  | test.describe('spreadsheet baseline', () => {
  63  |   test('Designer Demo shell exposes the fixed 1280x720 geometry and real palette entry', async ({ page }) => {
  64  |     await waitForDesignerDemo(page);
  65  |     const geometry = await page.evaluate(() => {
  66  |       const rect = (selector: string) => {
  67  |         const node = document.querySelector(selector);
  68  |         if (!node) return null;
  69  |         const box = node.getBoundingClientRect();
  70  |         return { y: Math.round(box.y), height: Math.round(box.height) };
  71  |       };
  72  |       return {
  73  |         ribbon: rect('[data-testid="designer-ribbon"]'),
  74  |         formulaBar: rect('[data-testid="designer-formula-bar"]'),
  75  |         workspace: rect('[data-testid="designer-workspace"]'),
  76  |         sheetTabs: rect('[data-testid="designer-sheet-tabs"]'),
  77  |         statusBar: rect('[data-testid="designer-status-bar"]'),
  78  |       };
  79  |     });
  80  |     expect(geometry.ribbon).toEqual({ y: 0, height: 142 });
  81  |     expect(geometry.formulaBar).toEqual({ y: 142, height: 37 });
  82  |     expect(geometry.workspace).toEqual({ y: 179, height: 519 });
  83  |     expect(geometry.statusBar).toEqual({ y: 698, height: 22 });
  84  |     await page.screenshot({ path: 'test-results/designer-demo-1280-current.png' });
  85  |     await page.getByRole('tab', { name: '视图' }).click();
  86  |     await page.getByRole('button', { name: '命令面板' }).click();
  87  |     await expect(page.getByTestId('command-palette')).toBeVisible();
  88  |     await expect(page.getByRole('button', { name: /保存/ }).first()).toBeVisible();
  89  |   });
  90  | 
  91  |   test('selection updates the name box', async ({ page }) => {
  92  |     await waitForWorkspace(page);
  93  |     await focusCanvas(page);
  94  |     await expect(page.getByTestId('name-box')).toHaveValue('A1');
  95  |     await page.keyboard.press('ArrowRight');
  96  |     await expect(page.getByTestId('name-box')).toHaveValue('B1');
  97  |   });
  98  | 
  99  |   test('keyboard navigation moves the active cell', async ({ page }) => {
  100 |     await waitForWorkspace(page);
  101 |     await focusCanvas(page);
  102 |     await page.keyboard.press('ArrowDown');
  103 |     await expect(page.getByTestId('name-box')).toHaveValue('A2');
  104 |     await page.keyboard.press('ArrowRight');
  105 |     await expect(page.getByTestId('name-box')).toHaveValue('B2');
  106 |   });
  107 | 
  108 |   test('direct typing commits cell editing', async ({ page }) => {
  109 |     await waitForWorkspace(page);
  110 |     const canvas = await focusCanvas(page);
  111 |     await page.keyboard.type('hello-e2e');
  112 |     await page.keyboard.press('Enter');
  113 |     await expect(page.getByTestId('formula-input')).not.toHaveValue('hello-e2e');
  114 |     await canvas.press('ArrowUp');
  115 |     await expect(page.getByTestId('formula-input')).toHaveValue('hello-e2e');
  116 |   });
  117 | 
  118 |   test('clicking another cell commits the old editor before accepting the new value', async ({ page }) => {
  119 |     await waitForWorkspace(page);
  120 |     const canvas = await focusCanvas(page);
  121 |     await page.keyboard.type('old-cell-value');
> 122 |     await expect(page.getByLabel('Cell editor')).toBeVisible();
      |                                                  ^ Error: expect(locator).toBeVisible() failed
  123 |     await canvas.click({ position: { x: 206, y: 34 } });
  124 |     await expect(page.getByTestId('name-box')).toHaveValue('C1');
  125 |     await expect(page.getByLabel('Cell editor')).toHaveCount(0);
  126 |     await page.keyboard.type('new-cell-value');
  127 |     await page.keyboard.press('Enter');
  128 |     await canvas.press('ArrowUp');
  129 |     await expect(page.getByTestId('name-box')).toHaveValue('C1');
  130 |     await expect(page.getByTestId('formula-input')).toHaveValue('new-cell-value');
  131 |   });
  132 | 
  133 |   test('clipboard copy and paste round-trip', async ({ page, context }) => {
  134 |     await context.grantPermissions(['clipboard-read', 'clipboard-write']);
  135 |     await waitForWorkspace(page);
  136 |     const canvas = await focusCanvas(page);
  137 |     await page.keyboard.type('copy-me');
  138 |     await page.keyboard.press('Enter');
  139 |     await canvas.press('ArrowUp');
  140 |     await canvas.press('Control+C');
  141 |     await canvas.press('Control+V');
  142 |     await canvas.press('ArrowLeft');
  143 |     await expect(page.getByTestId('formula-input')).toHaveValue('copy-me');
  144 |   });
  145 | 
  146 |   test('formula bar cancel restores the draft', async ({ page }) => {
  147 |     await waitForWorkspace(page);
  148 |     await focusCanvas(page);
  149 |     const input = page.getByTestId('formula-input');
  150 |     const original = await input.inputValue();
  151 |     await input.fill(`${original}-draft`);
  152 |     await page.getByTestId('formula-cancel').click();
  153 |     await expect(input).toHaveValue(original);
  154 |   });
  155 | 
  156 |   test('undo reverts the last edit', async ({ page }) => {
  157 |     await waitForWorkspace(page);
  158 |     const canvas = await focusCanvas(page);
  159 |     await page.keyboard.type('undo-me');
  160 |     await page.keyboard.press('Enter');
  161 |     await canvas.press('ArrowUp');
  162 |     await expect(page.getByTestId('formula-input')).toHaveValue('undo-me');
  163 |     await canvas.press('Control+Z');
  164 |     await expect(page.getByTestId('formula-input')).toHaveValue('');
  165 |   });
  166 | 
  167 |   test('adds and activates a new editable worksheet from the sheet tab control', async ({ page }) => {
  168 |     await waitForWorkspace(page);
  169 |     await page.getByRole('button', { name: /add worksheet|添加工作表/i }).click();
  170 |     await expect(page.getByRole('tab', { name: /sheet2/i })).toBeVisible();
  171 |     await expect(page.getByRole('tab', { name: /sheet2/i })).toHaveAttribute('aria-selected', 'true');
  172 |     const canvas = await focusCanvas(page);
  173 |     await page.keyboard.type('new-sheet-value');
  174 |     await page.keyboard.press('Enter');
  175 |     await canvas.press('ArrowUp');
  176 |     await expect(page.getByTestId('formula-input')).toHaveValue('new-sheet-value');
  177 |   });
  178 | 
  179 |   test('F2 editing keeps Chinese IME text intact before one committed operation', async ({ page }) => {
  180 |     await waitForWorkspace(page);
  181 |     const canvas = await focusCanvas(page);
  182 |     await canvas.press('F2');
  183 |     const editor = page.getByLabel('Cell editor');
  184 |     await expect(editor).toBeFocused();
  185 |     await editor.dispatchEvent('compositionstart');
  186 |     await editor.fill('中文公式输入');
  187 |     await editor.dispatchEvent('compositionend');
  188 |     await editor.press('Enter');
  189 |     await canvas.press('ArrowUp');
  190 |     await expect(page.getByTestId('formula-input')).toHaveValue('中文公式输入');
  191 |     await page.keyboard.press('Control+Z');
  192 |     await expect(page.getByTestId('formula-input')).toHaveValue('');
  193 |   });
  194 | 
  195 |   test('a dragged A2:C9 selection edits and places the editor at active cell C9', async ({ page }) => {
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
```