# React Sheets 内置浏览器逐项功能点验收方案与操作手册

本方案为基于 Claude Code 内置浏览器工具链（`preview_start`, `preview_snapshot`, `preview_eval`, `preview_click`, `preview_inspect`, `preview_console_logs`, `preview_network`, `preview_resize`）进行端到端逐项功能点验收的标准规范与操作步骤。

---

## 1. 验收环境与启动准备

### 1.1 启动服务
通过 `preview_start(name: "react-sheets")` 启动服务：
- 前端 Web: `http://localhost:4180`
- 后端 API/WS: `http://localhost:4181`

### 1.2 控制台零报错门禁
使用 `preview_console_logs(level: "error")` 确认：
- 页面初始化阶段 0 个未捕获异常。

---

## 2. 逐项功能点验收清单与断言标准

### 验收项 1：AppShell 完整工作台结构
- **目标**：验证标题栏、Ribbon 工具栏、公式栏、表格画板、工作表标签栏、状态栏及右侧工具面板全部正常挂载。
- **浏览器指令**：
  ```js
  preview_snapshot()
  ```
- **断言标准**：
  1. Header 包含标题 `Q3 Growth Planning` 及保存状态徽章。
  2. Ribbon 包含 `Home`, `Insert`, `Data`, `Review`, `View` 标签。
  3. FormulaBar 包含当前激活单元格坐标（如 `A1`）及输入框。
  4. SheetTabs 包含 `Sheet1` 及 `+` 新建按钮。
  5. StatusBar 包含缩放比例 `100%` 及状态描述。

---

### 验收项 2：Ribbon 工具栏多 Tab 切换
- **目标**：验证 Ribbon 工具栏在各个功能大类之间切换自如，展示对应的功能分组。
- **浏览器指令**：
  ```js
  // 切换到 Insert Tab
  preview_eval(`(() => {
    const tab = Array.from(document.querySelectorAll('[role="tab"]')).find(t => t.textContent.trim() === 'Insert');
    tab?.click();
    return tab ? 'ok' : 'fail';
  })()`)
  ```
- **断言标准**：
  - `Insert` Tab 下显示 `Pivot Table`, `Chart Builder`, `Sparkline`, `Shapes & Lines`, `Insert Function (fx)`。
  - `Data` Tab 下显示 `Sort A to Z`, `Sort Z to A`, `Custom Sort...`, `Data Validation`。
  - `View` Tab 下显示 `Freeze Panes`, `Zoom In`, `Zoom Out`, `Print & PDF`。

---

### 验收项 3：单元格选区与数据联动
- **目标**：验证鼠标点击单元格后，选区框正确绘制，公式栏同步联动。
- **浏览器指令**：
  ```js
  preview_eval(`(() => {
    const grid = document.querySelector('[role="grid"]');
    grid?.dispatchEvent(new PointerEvent('pointerdown', { clientX: 200, clientY: 150, bubbles: true }));
    return 'selected';
  })()`)
  ```
- **断言标准**：
  1. 页面渲染蓝色边框的激活单元格选区矩形及右下角填充柄（Fill Handle）。
  2. FormulaBar 中的 `Selected cell` 文本框同步显示对应单元格坐标（如 `C4`）。
  3. 公式输入框显示该单元格的模型值或公式字符串。

---

### 验收项 4：行内单元格编辑器 (CellEditor)
- **目标**：验证按 `F2` 或双击单元格激活行内浮动编辑器，编辑并提交新内容。
- **浏览器指令**：
  ```js
  // 1. 触发编辑态
  preview_eval(`(() => {
    const grid = document.querySelector('[role="grid"]');
    grid?.dispatchEvent(new KeyboardEvent('keydown', { key: 'F2', bubbles: true }));
    return 'editing';
  })()`)
  ```
- **断言标准**：
  1. 悬浮在选区单元格正上方的 `CellEditor`（textarea）弹出并获得焦点。
  2. 输入新文本（如 `Revenue Target`）按 Enter 后，编辑器自动关闭，新值写入模型并在 Canvas 及 FormulaBar 实时更新。

---

### 验收项 5：全局键盘导航
- **目标**：验证表格内键盘导航的完整性。
- **浏览器指令**：
  ```js
  preview_eval(`(() => {
    const grid = document.querySelector('[role="grid"]');
    grid?.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowRight', bubbles: true }));
    grid?.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
    return 'navigated';
  })()`)
  ```
- **断言标准**：
  1. `ArrowUp`, `ArrowDown`, `ArrowLeft`, `ArrowRight` 平滑移动选区。
  2. `Tab` 向右移动，`Shift+Tab` 向左移动。
  3. `Enter` 向下移动，`Shift+Enter` 向上移动。
  4. `Delete` / `Backspace` 清空选中单元格内容。

---

### 验收项 6：单元格富样式格式化
- **目标**：验证 Ribbon 中的字体、颜色、对齐、换行和合并功能。
- **浏览器指令**：
  ```js
  // 触发加粗与对齐
  preview_eval(`(() => {
    const boldBtn = document.querySelector('button[aria-label="Bold (Ctrl+B)"]');
    boldBtn?.click();
    const alignCenterBtn = document.querySelector('button[aria-label="Align Center"]');
    alignCenterBtn?.click();
    return 'styled';
  })()`)
  ```
- **断言标准**：
  1. 选中单元格应用粗体，Canvas 实时以 `bold 13px Inter` 重绘。
  2. 文本对齐更新为居中对齐。
  3. 调色板可选择前景色与背景填充色并即时生效。
  4. 点击 `Merge & Center` 成功跨列合并单元格。

---

### 验收项 7：数字格式化 (Currency / Percent / Comma)
- **目标**：验证数字格式的切换与显示。
- **浏览器指令**：
  ```js
  preview_eval(`(() => {
    const select = document.querySelector('select');
    if (select) {
      select.value = '$#,##0';
      select.dispatchEvent(new Event('change', { bubbles: true }));
      return 'currency formatted';
    }
    return 'select not found';
  })()`)
  ```
- **断言标准**：
  1. 数值 `132000` 格式化为 `$132,000`。
  2. 数值 `0.42` 格式化为 `42%`。

---

### 验收项 8：公式计算引擎与自动重算
- **目标**：验证 60+ 内置函数计算及依赖图拓扑更新。
- **浏览器指令**：
  ```js
  preview_eval(`(() => {
    const input = document.querySelector('input[aria-label="Formula input"]');
    if (input) {
      input.value = '=SUM(D2:D5)';
      input.dispatchEvent(new Event('input', { bubbles: true }));
      input.closest('form')?.dispatchEvent(new Event('submit', { bubbles: true, cancelable: true }));
      return 'formula evaluated';
    }
    return 'input not found';
  })()`)
  ```
- **断言标准**：
  1. 公式 `=SUM(D2:D5)` 正确求和并在单元格输出计算结果。
  2. 公式 `=IF(D2>0.4, "High", "Low")` 正确输出逻辑判断分支。
  3. 修改基础单元格数值，依赖它的公式单元格自动联动刷新。

---

### 验收项 9：函数向导对话框 (Function Wizard - fx)
- **目标**：验证点击 FormulaBar 的 `fx` 按钮打开交互式函数向导。
- **浏览器指令**：
  ```js
  preview_eval(`(() => {
    const fxBtn = Array.from(document.querySelectorAll('button')).find(b => b.textContent?.trim() === 'fx');
    fxBtn?.click();
    return 'opened wizard';
  })()`)
  ```
- **断言标准**：
  1. 弹出 `FunctionWizardDialog`。
  2. 搜索框支持按函数名过滤（如 `VLOOKUP`, `INDEX`, `DATE`）。
  3. 右侧面板展示语法格式、参数说明和使用描述。
  4. 点击确定自动将公式草稿插入激活单元格。

---

### 验收项 10：Pro 功能 - 图表构建 (Chart Builder)
- **目标**：验证在侧边栏配置并生成 Canvas 矢量图表。
- **浏览器指令**：
  ```js
  preview_eval(`(() => {
    const tab = Array.from(document.querySelectorAll('[role="tab"]')).find(t => t.textContent?.trim() === 'Chart');
    tab?.click();
    const btn = Array.from(document.querySelectorAll('button')).find(b => b.textContent?.includes('Insert Chart to Canvas'));
    btn?.click();
    return 'chart inserted';
  })()`)
  ```
- **断言标准**：
  1. `ChartModel` 创建成功，侧边栏显示在工作表图表列表中。
  2. Canvas Overlay 层绘制出柱状图/折线图/饼图矢量图、坐标轴刻度、背景网格线及图例。

---

### 验收项 11：Pro 功能 - 透视表聚合 (Pivot Table)
- **目标**：验证透视表多维聚合引擎计算与展示。
- **浏览器指令**：
  ```js
  preview_eval(`(() => {
    const tab = Array.from(document.querySelectorAll('[role="tab"]')).find(t => t.textContent?.trim() === 'Pivot');
    tab?.click();
    const btn = Array.from(document.querySelectorAll('button')).find(b => b.textContent?.includes('Generate Pivot Table'));
    btn?.click();
    return 'pivot generated';
  })()`)
  ```
- **断言标准**：
  1. Pivot 引擎按指定 Row Field 进行分组。
  2. 按 Value Field 执行 SUM / COUNT / AVERAGE / MIN / MAX 聚合。
  3. 生成包含分组数据行与 Grand Total 汇总行的透视表。

---

### 验收项 12：Pro 功能 - 浮动几何图形 (Shapes)
- **目标**：验证几何图形与标注绘制。
- **浏览器指令**：
  ```js
  preview_eval(`(() => {
    const tab = Array.from(document.querySelectorAll('[role="tab"]')).find(t => t.textContent?.trim() === 'Shape');
    tab?.click();
    const btn = Array.from(document.querySelectorAll('button')).find(b => b.textContent?.includes('Place Shape on Canvas'));
    btn?.click();
    return 'shape placed';
  })()`)
  ```
- **断言标准**：
  1. 支持圆角矩形、椭圆、箭头、星形、标注气泡。
  2. Canvas 准确渲染图形填充色、边框描边，并居中排版图形内部文本。

---

### 验收项 13：Pro 功能 - 单元格迷你图 (Sparkline)
- **目标**：验证在表格单元格内绘制微型趋势图。
- **浏览器指令**：
  ```js
  preview_eval(`(() => {
    const tab = Array.from(document.querySelectorAll('[role="tab"]')).find(t => t.textContent?.trim() === 'Spark');
    tab?.click();
    const btn = Array.from(document.querySelectorAll('button')).find(b => b.textContent?.includes('Insert In-Cell Sparkline'));
    btn?.click();
    return 'sparkline inserted';
  })()`)
  ```
- **断言标准**：
  1. 在目标单元格内绘制微型折线图（带最大/最小值高亮标记点）、柱状图或胜负图。

---

### 验收项 14：Pro 功能 - 条件格式与数据验证
- **目标**：验证规则添加与单元格校验。
- **浏览器指令**：
  - 条件格式：在 `Format` 面板添加大于阈值高亮规则。
  - 数据验证：在 `Validate` 面板配置下拉列表规则。
- **断言标准**：
  1. 规则成功注入 WorksheetModel，并在侧边栏规则列表中展示。

---

### 验收项 15：多工作表管理与数据隔离
- **目标**：验证新建、切换、重命名与删除工作表。
- **浏览器指令**：
  ```js
  preview_eval(`(() => {
    const addBtn = document.querySelector('button[aria-label="Add worksheet"]');
    addBtn?.click();
    return 'sheet added';
  })()`)
  ```
- **断言标准**：
  1. 成功新增 `Sheet 2` 并自动切换。
  2. 切换回 `Sheet1`，各工作表的单元格数据、选区与图表完全隔离。
  3. 支持右键 Sheet Tab 进行重命名和删除操作。

---

### 验收项 16：数据持久化与实时同步
- **目标**：验证与独立 SQLite WAL 服务端的接口通信。
- **浏览器指令**：
  ```js
  preview_network()
  ```
- **断言标准**：
  1. 初始化发送 `POST /api/v1/workbooks` 返回 `201 Created`。
  2. 每次单元格修改发送 `POST /api/v1/workbooks/:id/changesets` 返回 `200 OK` 及自增 revision。
  3. 服务端 `data/react-sheets.sqlite` 数据库正确落盘。

---

### 验收项 17：视口与缩放自适应
- **目标**：验证不同缩放比例与分辨率下的渲染对齐。
- **浏览器指令**：
  ```js
  preview_resize(width: 1280, height: 800)
  ```
- **断言标准**：
  1. 缩放 75%、100%、125% 比例下，Canvas 网格线与 React 选区框像素级对齐。
  2. 窄屏分辨率下右侧面板自适应折叠，主体表格区域保持可用。
