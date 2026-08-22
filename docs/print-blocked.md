# 打印能力 Blocked 评估

> **状态更新（2026-08-23）**：打印模块已在 `frontend/src/expendPlugins/print/` **全量自研实现**（对标 `@univerjs-pro/sheets-print` 公开能力面）。本文档保留历史评估，**Blocked 已解除**。

## 当前实现摘要

| 模块 | 路径 |
|---|---|
| Facade | `print.js` |
| 布局/分页 | `printLayout.js` |
| 状态/编排 | `printManager.js` |
| 高保真渲染 | `printRenderer.js` |
| 对话框 | `printDialog.js` |
| PDF | `printPdf.js`（jspdf） |
| 截图 | `printScreenshot.js` |
| 事件钩子 | `printEvents.js` |
| 参考映射 | [`print-pro-reference.md`](print-pro-reference.md) |

## 启用方式

```javascript
luckysheet.create({
  plugins: [{ name: "print", config: { enforceWatermark: false } }],
});
luckysheet.openPrintDialog();
luckysheet.print();
luckysheet.exportPrintPdf();
luckysheet.saveScreenshotToClipboard();
```

## 与 Excel / Pro 仍存在的差距

- 分页与 Excel 像素级对齐可能有偏差
- 复杂条件格式 / 部分图表类型保真度有限
- 协同编辑时打印配置广播未定义

## 历史 Blocked 原因（2026-08-22）

原 `print.js` 为空壳；现已替换为完整引擎，见 git 历史与本目录 [`print-ls-uv-map.md`](print-ls-uv-map.md)。
