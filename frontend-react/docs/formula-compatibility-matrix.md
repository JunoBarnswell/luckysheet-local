# Formula Compatibility Matrix

本矩阵记录 React Sheets 公式引擎相对 Excel 365 的兼容级别,随 M2+ 里程碑持续更新。

## 图例

| 级别 | 含义 |
|------|------|
| **Full** | 与 Excel 365 同输入同语义(含 spill / 动态数组) |
| **Partial** | 常见路径可用;边界/ locale / 1904 日期等可能差异 |
| **Preserve** | 导入保留;重算可能降级或仅显示缓存值 |
| **N/A** | 未实现;求值返回 `#NAME?` 或文档标注不支持 |

## 错误值

| 错误 | 状态 | 说明 |
|------|------|------|
| `#NULL!` | Full | 交集为空 |
| `#DIV/0!` | Full | 除零 |
| `#VALUE!` | Full | 类型错误 |
| `#REF!` | Full | 无效引用 |
| `#NAME?` | Full | 未知名称/函数 |
| `#NUM!` | Partial | 数值域错误 |
| `#N/A` | Partial | 查找未命中 |
| `#CALC!` | Partial | 计算引擎限制 |
| `#BLOCKED!` | Partial | 被保护/策略阻止 |
| `#SPILL!` | Partial | 溢出区被阻挡 |
| `#PARSE!` | N/A (引擎扩展) | 解析失败;Excel 显示为公式栏错误 |
| `#CYCLE!` | Full (引擎扩展) | 循环引用检测 |

## 引用语义

| 能力 | 状态 |
|------|------|
| A1 / `$A$1` / `A$1` / `$A1` | Full |
| 同表 / 跨表 `'Sheet 1'!A1` | Full |
| 定义名称 | Full |
| 结构化表引用 `Table1[Col]` / `[@Col]` / `#All` / `#Headers` / `#Data` / `#Totals` | Partial |
| 3D 引用 / 外部簿 | Preserve |
| 联合 `@` / `#`  spilled range | Partial |

## 依赖图与重算

| 能力 | 状态 |
|------|------|
| 单元格依赖 | Full |
| 矩形区域依赖 (SUM(A1:B2)) | Full |
| 增量失效 (RangeIndex) | Full |
| Volatile (NOW/RAND) | Partial |
| 手动 / 部分重算模式 | Partial |
| 结构变更 remapStructure | Full |

## Spill 模型

| 能力 | 状态 |
|------|------|
| 动态数组锚点 + 溢出区 | Partial |
| `#SPILL!` blocker 检测 | Partial |
| 子格只读(不存公式) | Partial |

## 函数优先集 (M2)

### 查找

| 函数 | 状态 |
|------|------|
| XLOOKUP | Partial |
| XMATCH | Partial |
| INDEX | Partial |
| MATCH | Partial |
| VLOOKUP | Partial |
| HLOOKUP | Partial |
| OFFSET | Partial |
| INDIRECT | Partial |

### 动态数组

| 函数 | 状态 |
|------|------|
| FILTER | Partial |
| SORT / SORTBY | Partial |
| UNIQUE | Partial |
| SEQUENCE / RANDARRAY | Partial |
| TAKE / DROP | Partial |
| HSTACK / VSTACK | Partial |

### 文本

| 函数 | 状态 |
|------|------|
| TEXTJOIN | Partial |
| CONCAT / CONCATENATE | Partial |
| LEFT / RIGHT / MID | Partial |

### 逻辑 / 数学

| 函数 | 状态 |
|------|------|
| IF / AND / OR | Partial |
| SUM / AVERAGE / COUNT | Partial |
| LET / LAMBDA | N/A |

### M18 矩阵

| 函数 | 状态 |
|------|------|
| GROUPBY | Partial |
| PIVOTBY | Partial |

## 维护说明

- 新增函数必须更新本表与 `packages/formula-engine/src/index.test.ts`
- L2 门禁:优先集标记 **Full** 或 **Partial** 且具备单测
- 导入 xlsx 时 **Preserve** 项写入 Compatibility Report (M14)
