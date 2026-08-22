# 同页多实例架构

## 运行时边界

- `createWorkbookContext()` 为每个 workbook 创建独立 Store、undo/redo、公式状态、计时器、portal 和模块运行时。
- legacy controller 的导入对象通过 `createContextualModule()` 保持稳定，但其可变属性按 focused `InstanceContext` 解析；不再快照或恢复共享 controller。
- `Store` Proxy 只服务旧同步 API。异步资源必须捕获 `instanceId`，并通过 `withInstance()` 回到所属 context。
- `server` 的 socket、重连和 online 回调均绑定创建它的实例；销毁只关闭目标 socket、计时器、DOM 和事件。

## DOM 与 API

- `InstanceDom` 为 LuckySheet 生成 ID 前缀，并用 scoped jQuery adapter 替代 jQuery/`Document` 原型 monkey patch；组合选择器、HTML 插入和 document 事件均按实例作用域处理。
- body portal 含 `data-ls-instance` 并登记到 context，`destroy(id)` 只回收其资源。
- `create()` 保持销毁后新建的历史行为；`create({ multi: true })` 返回 `{ instanceId }` 并保留其他实例。
- `use(id)` 聚焦实例；`getInstance(id)` 提供绑定 API facade；`listInstances()` 返回 `{ instanceId, container, focused, state }[]`；`order` 始终是 sheet 索引。

## 不支持的语义

- 不支持跨工作簿公式引用；它不属于 LuckySheet 原有模型。
- 普通旧 API 默认作用于 focused 实例；异步插件实现必须使用 bound facade 或显式 instanceId。
