# 数据模型

## 概览

当前原型使用前端 mock 数据，但数据模型已经按未来接真实服务的方向组织。核心思想是：表格可见值和隐藏取数配置分离。

用户看到的是 Excel-like 表格；系统内部维护指标定义、指标绑定、输出配置、刷新状态和模板配置。

## 核心对象

### MetricDefinition

指标定义，相当于 DM 数据库里的指标字典。

关键字段：

| 字段 | 含义 |
|---|---|
| `code` | 指标代码，例如 `CN.HOG.PRICE.NATIONAL` |
| `label` | 用户看到的指标名称 |
| `category` | 指标分类 |
| `unit` | 单位 |
| `source` | 数据源 |
| `description` | 口径说明 |
| `aliases` | 搜索别名 |

它解决的是“用户输入一个自然语言词，系统应该联想到哪些指标”。

### MetricBinding

指标绑定，表示某一行输入区已经确认了指标口径。

关键字段：

| 字段 | 含义 |
|---|---|
| `row` | 指标所在行 |
| `cell` | 指标单元格，例如 `A13` |
| `displayName` | 表格中展示的指标名 |
| `metricCode` | 已确认的指标代码 |
| `source` | 数据源 |
| `disambiguationStatus` | 是否已经完成口径确认 |

它解决的是“这一行到底代表哪个指标”。

### OutputCellConfig

输出单元格的隐藏取数配置。

关键字段：

| 字段 | 含义 |
|---|---|
| `targetCell` | 输出单元格，例如 `H13` |
| `metricCell` | 指标来源单元格，例如 `A13` |
| `yearCell` | 年份来源单元格，例如 `H1` |
| `metricCode` | 指标代码 |
| `year` | 年份参数 |
| `status` | 输出状态 |
| `lastRefreshAt` | 最近刷新时间 |
| `value` | 最近刷新值 |

它解决的是“这个结果单元格应该按什么规则刷新”。

### OutputStatus

输出状态枚举：

| 状态 | 含义 |
|---|---|
| `empty` | 没有配置或参数 |
| `pendingRefresh` | 参数变化后等待刷新 |
| `loading` | 正在刷新 |
| `ready` | 已有可用结果 |
| `failed` | 刷新失败 |

### TemplateConfig

模板配置，表示一个可复用的表格框架。

关键字段：

| 字段 | 含义 |
|---|---|
| `templateId` | 模板 ID |
| `templateName` | 模板名称 |
| `description` | 模板说明 |
| `inputRanges` | 输入区范围 |
| `outputRange` | 输出区范围 |
| `metricBindings` | 指标绑定 |
| `refreshRules` | 刷新规则 |
| `formatPreset` | 展示格式 |
| `type` | 默认模板或自定义模板 |

## 数据流

### 新增指标

1. 用户选中 `A13`。
2. 用户输入“生猪价格”。
3. 系统从 `MetricDefinition` 中找到候选指标。
4. DM AI 弹出确认卡。
5. 用户选择口径。
6. 系统生成 `MetricBinding`。
7. 系统为 `B13:H13` 生成 `OutputCellConfig`。
8. 输出区显示 `待刷新`。

### 修改年份

1. 用户编辑 `H1` 为 `2026`。
2. 系统找到所有引用 `H1` 的 `OutputCellConfig`。
3. 将这些输出配置状态改为 `pendingRefresh`。
4. 表格显示 `待刷新`。

### 刷新取数

1. 用户点击 Ribbon 的“刷新”。
2. 系统找到所有 `pendingRefresh` 单元格。
3. 状态改为 `loading`。
4. 调用取数函数。
5. 成功后写入 `value`，状态改为 `ready`。
6. 失败后状态改为 `failed`。

当前原型中，取数函数是 `mockFetchMetricValue`。未来接真实服务时，应保持输入输出形态稳定，只替换内部实现。

## 当前代码位置

| 内容 | 文件 |
|---|---|
| 类型定义 | `src/types.ts` |
| mock 指标和模板 | `src/mockData.ts` |
| 主要交互状态 | `src/App.tsx` |
| 表格和面板样式 | `src/styles.css` |

## GitHub 协作建议

这些数据模型文档应该提交到 GitHub，因为它们解释了产品和技术之间的共同语言。

不建议提交临时工作日志，例如：

- `task_plan.md`
- `findings.md`
- `progress.md`

如果其中有长期价值，应提炼到 `docs/product/` 或 `docs/architecture/`。

