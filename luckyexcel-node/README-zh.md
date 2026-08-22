# Luckyexcel-node

> **已废弃。** Excel 导入/导出已合并到 Java 后端，不再需要单独启动本 Node 服务。
>
> - 导出：`POST /luckysheet/luckyToXlsx`
> - 导入：`POST /luckysheet/luckyexcel/upload`
>
> 实现见 `backend/luckysheet` 的 `ExcelIoController` / `ExcelIoService`（luckysheet-lib）。本目录保留一个观察期，后续可删除。

简体中文 | [English](./README.md)

## 介绍
Luckyexcel-node，是一个excel导入导出库 [Luckyexcel](https://github.com/mengshukeji/Luckyexcel) 的koa2服务端解析案例。

## 开发

### 环境
[Node.js](https://nodejs.org/en/) Version >= 6 

### 安装
```
npm install
```
### 启动
```
node app.js
```
然后访问`http://localhost:3000/luckyexcel`。

即可看到后端返回的转换好的excel json

核心代码在[luckyexcel.js](./controllers/luckyexcel.js)

## 资源
- [Luckysheet](https://github.com/mengshukeji/Luckysheet)
- [Luckysheet-vue](https://github.com/mengshukeji/Luckysheet-vue)

## 版权信息
[MIT](http://opensource.org/licenses/MIT)

Copyright (c) 2020-present, mengshukeji
