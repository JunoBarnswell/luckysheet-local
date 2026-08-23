# Instructions

- Following Playwright test failed.
- Explain why, be concise, respect Playwright best practices.
- Provide a snippet of code with the fix, if possible.

# Test info

- Name: spreadsheet.spec.ts >> spreadsheet baseline >> right click changes the command target before the context menu opens
- Location: e2e\spreadsheet.spec.ts:191:3

# Error details

```
Error: expect(locator).toHaveValue(expected) failed

Locator:  getByTestId('formula-input')
Expected: ""
Received: "right-click-target"
Timeout:  5000ms

Call log:
  - Expect "toHaveValue" with timeout 5000ms
  - waiting for getByTestId('formula-input')
    14 × locator resolved to <input aria-label="Formula input" value="right-click-target" data-testid="formula-input" placeholder="Enter a value or formula (=SUM, =IF, ...)" class="min-h-9 w-full min-w-0 rounded-lg border bg-white px-3 text-sm text-ink outline-none transition placeholder:text-slate-400 focus:border-accent focus:ring-2 focus:ring-accent/15 disabled:cursor-not-allowed disabled:bg-slate-50 disabled:text-slate-400 read-only:bg-slate-50/70 border-line font-mono text-xs"/>
       - unexpected value "right-click-target"

```

```yaml
- textbox "Formula input":
  - /placeholder: Enter a value or formula (=SUM, =IF, ...)
  - text: right-click-target
```

# Test source

```ts
  111 |     await editor.fill('中文公式输入');
  112 |     await editor.dispatchEvent('compositionend');
  113 |     await editor.press('Enter');
  114 |     await canvas.press('ArrowUp');
  115 |     await expect(page.getByTestId('formula-input')).toHaveValue('中文公式输入');
  116 |     await page.keyboard.press('Control+Z');
  117 |     await expect(page.getByTestId('formula-input')).toHaveValue('');
  118 |   });
  119 | 
  120 |   test('a dragged A2:C9 selection edits and places the editor at active cell C9', async ({ page }) => {
  121 |     await waitForWorkspace(page);
  122 |     const canvas = await focusCanvas(page);
  123 |     const box = await canvas.boundingBox();
  124 |     if (!box) throw new Error('Spreadsheet canvas has no bounds');
  125 |     await page.mouse.move(box.x + 101, box.y + 66);
  126 |     await page.mouse.down();
  127 |     await page.mouse.move(box.x + 321, box.y + 262);
  128 |     await page.mouse.up();
  129 |     await expect(page.getByTestId('name-box')).toHaveValue('C9');
  130 | 
  131 |     await canvas.press('F2');
  132 |     const editor = page.getByLabel('Cell editor');
  133 |     await expect(editor).toBeFocused();
  134 |     const editorBox = await editor.boundingBox();
  135 |     if (!editorBox) throw new Error('Cell editor has no bounds');
  136 |     expect(editorBox.x - box.x).toBeGreaterThan(250);
  137 |     expect(editorBox.y - box.y).toBeGreaterThan(220);
  138 |     await editor.fill('4');
  139 |     await editor.press('Enter');
  140 |     await canvas.press('ArrowUp');
  141 |     await expect(page.getByTestId('name-box')).toHaveValue('C9');
  142 |     await expect(page.getByTestId('formula-input')).toHaveValue('4');
  143 |   });
  144 | 
  145 |   test('a dragged selection accepts direct typing at the release cell', async ({ page }) => {
  146 |     await waitForWorkspace(page);
  147 |     const canvas = await focusCanvas(page);
  148 |     await dragCells(page, canvas, { row: 1, column: 0 }, { row: 8, column: 2 });
  149 |     await expect(page.getByTestId('name-box')).toHaveValue('C9');
  150 |     await page.keyboard.type('direct-c9');
  151 |     await canvas.press('Enter');
  152 |     await canvas.press('ArrowUp');
  153 |     await expect(page.getByTestId('name-box')).toHaveValue('C9');
  154 |     await expect(page.getByTestId('formula-input')).toHaveValue('direct-c9');
  155 |   });
  156 | 
  157 |   test('reverse drag keeps the release cell active for F2 editing', async ({ page }) => {
  158 |     await waitForWorkspace(page);
  159 |     const canvas = await focusCanvas(page);
  160 |     await dragCells(page, canvas, { row: 8, column: 2 }, { row: 1, column: 0 });
  161 |     await expect(page.getByTestId('name-box')).toHaveValue('A2');
  162 |     await canvas.press('F2');
  163 |     const editor = page.getByLabel('Cell editor');
  164 |     await expect(editor).toBeFocused();
  165 |     await editor.fill('reverse-a2');
  166 |     await editor.press('Enter');
  167 |     await canvas.press('ArrowUp');
  168 |     await expect(page.getByTestId('name-box')).toHaveValue('A2');
  169 |     await expect(page.getByTestId('formula-input')).toHaveValue('reverse-a2');
  170 |   });
  171 | 
  172 |   test('row and column header drags commit the final header target', async ({ page }) => {
  173 |     await waitForWorkspace(page);
  174 |     const canvas = await focusCanvas(page);
  175 |     const box = await canvas.boundingBox();
  176 |     if (!box) throw new Error('Spreadsheet canvas has no bounds');
  177 | 
  178 |     await page.mouse.move(box.x + 20, box.y + 24 + 28 + 14);
  179 |     await page.mouse.down();
  180 |     await page.mouse.move(box.x + 20, box.y + 24 + 8 * 28 + 14);
  181 |     await page.mouse.up();
  182 |     await expect(page.getByTestId('name-box')).toHaveValue('A9');
  183 | 
  184 |     await page.mouse.move(box.x + 46 + 55, box.y + 12);
  185 |     await page.mouse.down();
  186 |     await page.mouse.move(box.x + 46 + 2 * 110 + 55, box.y + 12);
  187 |     await page.mouse.up();
  188 |     await expect(page.getByTestId('name-box')).toHaveValue('C1');
  189 |   });
  190 | 
  191 |   test('right click changes the command target before the context menu opens', async ({ page }) => {
  192 |     await waitForWorkspace(page);
  193 |     await page.getByRole('button', { name: /add worksheet|添加工作表/i }).click();
  194 |     const canvas = await focusCanvas(page);
  195 |     const box = await canvas.boundingBox();
  196 |     if (!box) throw new Error('Spreadsheet canvas has no bounds');
  197 |     await page.getByTestId('name-box').fill('B2');
  198 |     await page.getByTestId('name-box').press('Enter');
  199 |     await canvas.focus();
  200 |     await page.keyboard.type('right-click-target');
  201 |     await page.keyboard.press('Enter');
  202 |     await page.getByTestId('name-box').fill('A1');
  203 |     await page.getByTestId('name-box').press('Enter');
  204 |     const target = cellPoint(box, 1, 1);
  205 |     await page.mouse.click(target.x, target.y, { button: 'right' });
  206 |     await expect(page.getByRole('menu')).toBeVisible();
  207 |     await expect(page.getByTestId('name-box')).toHaveValue('B2');
  208 |     await page.getByRole('menuitem', { name: 'Clear contents' }).click();
  209 |     await page.getByTestId('name-box').fill('B2');
  210 |     await page.getByTestId('name-box').press('Enter');
> 211 |     await expect(page.getByTestId('formula-input')).toHaveValue('');
      |                                                     ^ Error: expect(locator).toHaveValue(expected) failed
  212 |   });
  213 | 
  214 |   test('fill handle follows the primary range bottom-right', async ({ page }) => {
  215 |     await waitForWorkspace(page);
  216 |     await page.getByRole('button', { name: /add worksheet|添加工作表/i }).click();
  217 |     const canvas = await focusCanvas(page);
  218 |     await page.keyboard.type('fill-source-unique');
  219 |     await page.keyboard.press('Enter');
  220 |     await page.getByTestId('name-box').fill('A1');
  221 |     await page.getByTestId('name-box').press('Enter');
  222 |     await canvas.focus();
  223 |     const box = await canvas.boundingBox();
  224 |     if (!box) throw new Error('Spreadsheet canvas has no bounds');
  225 |     await page.mouse.move(box.x + 46 + 110 - 4, box.y + 24 + 28 - 4);
  226 |     await page.mouse.down();
  227 |     await page.mouse.move(box.x + 46 + 110, box.y + 24 + 3 * 28);
  228 |     await page.mouse.up();
  229 |     await page.getByTestId('name-box').fill('A3');
  230 |     await page.getByTestId('name-box').press('Enter');
  231 |     await expect(page.getByTestId('formula-input')).toHaveValue('fill-source-unique');
  232 |   });
  233 | 
  234 |   test('local formula calculation and IndexedDB restore work without an API request', async ({ page }) => {
  235 |     const apiRequests: string[] = [];
  236 |     let socketCount = 0;
  237 |     page.on('request', (request) => {
  238 |       if (new URL(request.url()).pathname.startsWith('/api/')) apiRequests.push(request.url());
  239 |     });
  240 |     page.on('websocket', (socket) => {
  241 |       if (new URL(socket.url()).pathname === '/ws') socketCount += 1;
  242 |     });
  243 |     await waitForWorkspace(page);
  244 |     const canvas = await focusCanvas(page);
  245 |     await page.keyboard.type('=1+2');
  246 |     await page.keyboard.press('Enter');
  247 |     await expect(page.getByTestId('formula-input')).toHaveValue('');
  248 |     await canvas.press('ArrowUp');
  249 |     await expect(page.getByTestId('formula-input')).toHaveValue('=1+2');
  250 | 
  251 |     await page.reload();
  252 |     await expect(page.getByTestId('app-shell')).toHaveAttribute('data-workspace-phase', 'ready');
  253 |     await focusCanvas(page);
  254 |     await expect(page.getByTestId('formula-input')).toHaveValue('=1+2');
  255 |     expect(apiRequests).toEqual([]);
  256 |     expect(socketCount).toBe(0);
  257 |   });
  258 | });
  259 | 
```