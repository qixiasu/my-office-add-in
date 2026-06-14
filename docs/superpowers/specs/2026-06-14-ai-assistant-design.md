# AI 助手功能设计文档

> Excel 加载项中集成 DeepSeek AI 的聊天面板功能

## 概述

在已有的 Excel 数据处理加载项中新增"AI 助手"功能。用户在 Excel 中选中数据后，通过 Ribbon 栏的"AI 助手"按钮打开聊天面板，以自然语言与 AI 交互，完成公式生成、数据分析和文本分类/提取等任务。

## 核心需求

| 需求 | 说明 |
|------|------|
| AI 提供商 | DeepSeek API（兼容 OpenAI Function Calling 格式） |
| 交互方式 | 统一聊天面板（类似 ChatGPT 的任务面板） |
| 输出模式 | 混合模式：分析结果在面板显示，操作类自动写回 Excel |
| 输入模式 | 只发送数据摘要 + 样本行，不发送全部原始数据到 API |

### 三种能力

1. **公式生成（A）** — 用户用自然语言描述需求，AI 生成 Excel 公式写入指定单元格
2. **数据问答分析（C）** — AI 读取选中数据的摘要，在聊天面板中返回分析结论
3. **文本分类/提取（E）** — AI 根据规则对数据进行分类、提取或标注，结果写回新列

## 架构

### 系统架构

```
Excel 选中数据 → 聊天面板（Taskpane） → DeepSeek API（Function Calling）
     ↑                                      ↓
     └────────── 前端执行 Tool ←───────────┘
```

- 用户选中 Excel 数据 → 打开 AI 助手聊天面板
- 面板自动读取当前选区摘要 + 样本数据
- 用户输入指令，发送到 DeepSeek API
- DeepSeek 通过 Function Calling 返回要调用的 Tool 及参数
- 前端解析 Tool 调用，通过 Office JS API 执行对应操作
- 操作结果回传给 AI / 在面板中显示

### 数据流

**三个阶段：**

1. **数据采集** — 用户选中单元格后，前端通过 Office JS API 读取：
   - 选区地址、行数、列数
   - 列标题（首行）
   - 前 N 行样本数据（可配置，默认 10 行）
   - 数值列统计信息（最大/最小/平均/中位数）

2. **AI 处理** — 数据摘要 + 用户指令发送到 DeepSeek API：
   - System Prompt 定义 AI 角色和可用工具
   - 数据摘要注入对话上下文
   - AI 选择调用哪个 Tool，返回结构化参数

3. **结果执行** — 前端收到 Tool 调用后：
   - 分析类 → 在聊天面板显示结果文本
   - 操作类（写公式、分类结果） → 通过 Office JS API 写回 Excel，操作前显示确认框

### 上下文管理

| 组件 | 大小 | 说明 |
|------|------|------|
| System Prompt | ~500 tokens | AI 角色定义 + Tool Schema |
| 数据上下文 | ~500-1500 tokens | 当前选区摘要 + 样本行 |
| 对话历史 | ~2000-4000 tokens | 最近 3 轮对话（滑动窗口） |
| **总计** | **~4K-6K tokens** | 在 DeepSeek 上下文限制内安全 |

裁剪规则：
- 保留 System Prompt（始终不变）
- 保留当前数据上下文（选区变化时更新）
- 保留最近 3 轮对话（用户-AI-工具结果为一轮）
- 超出阈值时丢弃最早的对话轮次

## Tool 定义（Function Calling Schema）

DeepSeek API 通过以下 5 个工具与 Excel 交互：

### 1. generate_formula

根据用户描述生成 Excel 公式并写入单元格。

```json
{
  "name": "generate_formula",
  "description": "根据用户描述生成 Excel 公式并写入单元格",
  "parameters": {
    "formula": "=IF(A2>100,\"高\",\"低\")",
    "targetCell": "B2",
    "fillRange": "B2:B50",
    "explanation": "如果 A 列大于 100 则显示'高'，否则显示'低'"
  }
}
```

### 2. analyze_data

分析选中的 Excel 数据并返回分析结论（结果在面板显示）。

```json
{
  "name": "analyze_data",
  "description": "分析选中的 Excel 数据并返回分析结论",
  "parameters": {
    "dataSummary": {
      "rowCount": 100,
      "columns": ["日期", "销售额", "地区"],
      "sampleRows": 5,
      "stats": { "销售额": { "max": 50000, "min": 100, "avg": 8500 } }
    },
    "question": "哪个地区的销售额最高？"
  }
}
```

### 3. classify_data

对选中数据进行分类、提取或标注，结果写回新列。

```json
{
  "name": "classify_data",
  "description": "对选中数据进行分类、提取或标注",
  "parameters": {
    "range": "A2:A100",
    "task": "分类",
    "instruction": "判断每行的评论是正面、负面还是中性",
    "outputColumn": "C",
    "labelMapping": { "正面": 1, "负面": -1, "中性": 0 }
  }
}
```

### 4. write_to_cells

将数据数组写入 Excel 指定区域（通用写入工具）。

```json
{
  "name": "write_to_cells",
  "description": "将数据数组写入 Excel 指定区域",
  "parameters": {
    "range": "C2:C100",
    "values": [["正面"], ["负面"], ["正面"]],
    "insertBefore": false
  }
}
```

### 5. get_selection

获取用户当前选中的 Excel 数据摘要（自动执行，不依赖 AI 调用）。

返回值：
```json
{
  "address": "A1:D100",
  "columnCount": 4,
  "rowCount": 100,
  "headers": ["姓名", "年龄", "城市", "备注"],
  "sampleData": [["张三", 28, "北京", ""]],
  "stats": { "年龄": { "max": 65, "min": 18, "avg": 35 } }
}
```

## UI / 交互设计

### 布局

```
┌─────────────────────────────────┐
│  🤖 AI 助手           ⚙️ 设置   │  ← 标题栏
├─────────────────────────────────┤
│                                 │
│  AI: 已获取选区 A1:D100         │
│  请问你想做什么？               │  ← 对话气泡
│                                 │
│  你: 把备注列按情感分类         │
│                                 │
│  AI: 好的，正在处理...          │
│                                 │
├─────────────────────────────────┤
│  [输入框...            ] [发送] │  ← 输入区
└─────────────────────────────────┘
```

### 交互流程

1. 用户点击 Ribbon 栏"AI 助手"按钮 → 打开聊天面板
2. 面板自动获取当前选区摘要 → 显示"已获取选区"提示
3. 用户输入自然语言指令 → 点击发送
4. AI 通过 Function Calling 返回 Tool 调用
5. 分析类操作 → 在面板直接显示结果
6. 写入类操作 → 弹出确认框 → 用户确认后执行 → 反馈完成
7. 对话持续进行，历史保留最近 3 轮

### 设置弹窗

点击右上角齿轮图标弹出设置：
- API Key 输入框（存储在 localStorage）
- 模型选择（默认 deepseek-chat）
- 温度调节（默认 0.3）
- 最大 Token 数（默认 4096）

## API Key 管理

**初始方案**：存储在浏览器 localStorage 中
- 用户首次使用 → 弹出 API Key 输入对话框
- 后续自动读取 localStorage
- 不随代码上传，不存储在 git 中

**可选增强**：后期可添加后端代理转发，API Key 存储在服务器端

## 确认机制

所有会修改 Excel 数据的操作在执行前都显示确认框：
- 复用现有代码中的 `showConfirmBox` 模式
- 确认框显示即将执行的操作描述
- 用户确认后才真正执行 Excel 写入
- 用户取消则只返回文本提示

## System Prompt 设计

```
你是 Excel 数据处理助手，运行在 Excel 加载项中。
你拥有以下工具可以使用：
- generate_formula：根据描述生成 Excel 公式
- analyze_data：分析选中的数据，返回分析结论
- classify_data：对数据进行分类、提取或标注
- write_to_cells：将数据写入 Excel 指定区域

当前用户选中的数据：
[选区地址，列名，样本数据，数值统计]

注意事项：
1. 分析类操作 → 在聊天中显示结果
2. 修改类操作 → 通过 write_to_cells 写回 Excel
3. 数据批量处理 → 分类任务请返回分类规则，由前端逐行执行
4. 解释你的每一步操作，让用户知道发生了什么
```

## 文件结构

### 新增文件

| 文件 | 用途 |
|------|------|
| `src/taskpane/ai-assistant-taskpane.html` | AI 聊天面板 HTML（Fabric UI 风格） |
| `src/taskpane/ai-assistant-taskpane.js` | 面板主逻辑：Office.onReady → Excel.run |
| `src/taskpane/ai-assistant-taskpane.css` | 对话气泡、输入框样式 |
| `src/utils/ai-utils.js` | 核心工具模块 |
| `src/utils/ai-utils.test.js` | 单元测试 |

### 修改文件

| 文件 | 修改内容 |
|------|----------|
| `manifest.xml` | 新增 "AI 助手" ribbon 按钮，ShowTaskpane 指向 ai-assistant-taskpane.html |
| `webpack.config.js` | 新增 ai-assistant-taskpane 构建入口 |
| `assets/` | 新增 AI 按钮图标（16/32/80 px） |

### ai-utils.js 核心模块

**DeepSeekClient**
- DeepSeek API 调用封装（fetch）
- Function Calling 参数组装
- 错误重试机制

**ToolRegistry**
- 注册 5 个 Tool 定义
- 生成 OpenAI-compatible Schema
- 分派 Tool 调用到对应处理函数

**ContextManager**
- 对话历史存储
- 滑动窗口裁剪
- 数据上下文注入
- System Prompt 管理

**ExcelDataService**
- 读取当前选区数据
- 生成数据摘要 + 样本
- 执行 Excel 写入操作

## 非功能需求

- **安全性**：API Key 存储在 localStorage，不做硬编码
- **性能**：不发送全部原始数据到 API，只发摘要 + 样本
- **可用性**：写入操作前显示确认框，防止误操作
- **兼容性**：保持与现有代码一致的 Fabric UI 风格
- **可测试性**：ai-utils.js 的四个核心模块均可独立单元测试
