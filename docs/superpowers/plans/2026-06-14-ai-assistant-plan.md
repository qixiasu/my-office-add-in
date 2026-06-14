# AI 助手功能 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 在 Excel 数据处理加载项中新增 AI 助手聊天面板，通过 DeepSeek API + Function Calling 实现公式生成、数据分析和文本分类/提取三种能力。

**Architecture:** 聊天面板（Taskpane）→ DeepSeek API（Function Calling）→ 前端解析 Tool 调用 → Office JS API 执行。数据只发摘要+样本，操作类写回 Excel 前弹确认框。

**Tech Stack:** Office JS API, DeepSeek API (OpenAI-compatible), localStorage, Jest

---

## 文件结构

### 新增文件

| # | 文件 | 职责 |
|---|------|------|
| 1 | `assets/ai-16.png` | AI 助手 ribbon 按钮图标 16px |
| 2 | `assets/ai-32.png` | AI 助手 ribbon 按钮图标 32px |
| 3 | `assets/ai-80.png` | AI 助手 ribbon 按钮图标 80px |
| 4 | `src/utils/ai-utils.js` | 核心工具：DeepSeekClient + ToolRegistry + ContextManager + ExcelDataService |
| 5 | `src/utils/ai-utils.test.js` | 单元测试（覆盖全部 4 个模块） |
| 6 | `src/taskpane/ai-assistant-taskpane.html` | AI 聊天面板 HTML |
| 7 | `src/taskpane/ai-assistant-taskpane.css` | AI 聊天面板样式 |
| 8 | `src/taskpane/ai-assistant-taskpane.js` | AI 聊天面板主逻辑 |

### 修改文件

| # | 文件 | 修改内容 |
|---|------|----------|
| 9 | `webpack.config.js` | 新增 `ai-assistant-taskpane` 入口 + HtmlWebpackPlugin |
| 10 | `manifest.xml` | 新增 "AI 助手" ribbon 按钮（ShowTaskpane） |

---

## Task 1: 创建 AI 按钮图标 + webpack 入口 + manifest 注册

**依赖：** 无（独立的基础设置）

### Sub-task 1.1: 创建 AI 按钮图标

**文件：** 创建 `assets/ai-16.png`, `assets/ai-32.png`, `assets/ai-80.png`

使用以下 Node.js 脚本生成三个尺寸的图标（用简单的蓝色机器人头 SVG 渲染）：

```bash
node -e "
const { createCanvas } = (() => { try { return require('canvas'); } catch(e) { return null; } })();
// Fallback: copy an existing icon as placeholder
require('fs').copyFileSync('assets/tools-16.png', 'assets/ai-16.png');
require('fs').copyFileSync('assets/tools-32.png', 'assets/ai-32.png');
require('fs').copyFileSync('assets/tools-80.png', 'assets/ai-80.png');
console.log('AI icons created (placeholder from tools icons)');
"
```

运行以上命令。如无 canvas 模块可用，先用 tools 图标作为占位符。

- [ ] 运行命令生成三个 AI 图标文件
- [ ] 确认文件存在：`ls assets/ai-*`

### Sub-task 1.2: 添加 webpack 入口

**文件：** 修改 `webpack.config.js`

在 `entry` 对象中追加 `ai-assistant-taskpane` 入口（按字母顺序排在其他 taskpane 之后）：

```js
      "split-sheet-taskpane": ["./src/taskpane/split-sheet-taskpane.js", "./src/taskpane/split-sheet-taskpane.html"],
      "sql-query-taskpane": [
        "./src/taskpane/sql-query-taskpane.js",
        "./src/taskpane/sql-query-taskpane.html",
      ],
      // ↓ 在 sql-query-taskpane 后面追加：
      "ai-assistant-taskpane": [
        "./src/taskpane/ai-assistant-taskpane.js",
        "./src/taskpane/ai-assistant-taskpane.html",
      ],
```

在 `plugins` 数组中追加 HtmlWebpackPlugin（放在最后一个 HtmlWebpackPlugin 之后）：

```js
      new HtmlWebpackPlugin({
        filename: "ai-assistant-taskpane.html",
        template: "./src/taskpane/ai-assistant-taskpane.html",
        chunks: ["polyfill", "ai-assistant-taskpane"],
      }),
```

- [ ] 在 entry 对象中追加 ai-assistant-taskpane 入口
- [ ] 在 plugins 数组中追加 ai-assistant-taskpane 的 HtmlWebpackPlugin

### Sub-task 1.3: manifest 注册 AI 助手按钮

**文件：** 修改 `manifest.xml`

在 `<Group id="DatabaseQueryGroup">` 后面、`<Label resid="CustomTab.Label"/>` 之前插入新的 Group：

```xml
              <Group id="AiAssistantGroup">
                <Label resid="AiAssistantGroup.Label"/>
                <Icon>
                    <bt:Image size="16" resid="AiIcon.16x16"/>
                    <bt:Image size="32" resid="AiIcon.32x32"/>
                    <bt:Image size="80" resid="AiIcon.80x80"/>
                </Icon>
                <Control xsi:type="Button" id="AiAssistantButton">
                    <Label resid="AiAssistantButton.Label"/>
                    <Supertip>
                        <Title resid="AiAssistantButton.Label"/>
                        <Description resid="AiAssistantButton.Tooltip"/>
                    </Supertip>
                    <Icon>
                        <bt:Image size="16" resid="AiIcon.16x16"/>
                        <bt:Image size="32" resid="AiIcon.32x32"/>
                        <bt:Image size="80" resid="AiIcon.80x80"/>
                    </Icon>
                    <Action xsi:type="ShowTaskpane">
                        <TaskpaneId>AiAssistantTaskpaneId</TaskpaneId>
                        <SourceLocation resid="AiAssistantTaskpane.Url"/>
                    </Action>
                </Control>
              </Group>
```

在 `<bt:Images>` 中追加图标资源：

```xml
        <bt:Image id="AiIcon.16x16" DefaultValue="https://localhost:3000/assets/ai-16.png"/>
        <bt:Image id="AiIcon.32x32" DefaultValue="https://localhost:3000/assets/ai-32.png"/>
        <bt:Image id="AiIcon.80x80" DefaultValue="https://localhost:3000/assets/ai-80.png"/>
```

在 `<bt:Urls>` 中追加面板 URL：

```xml
        <bt:Url id="AiAssistantTaskpane.Url" DefaultValue="https://localhost:3000/ai-assistant-taskpane.html"/>
```

在 `<bt:ShortStrings>` 中追加字符串资源：

```xml
        <bt:String id="AiAssistantGroup.Label" DefaultValue="AI 助手"/>
        <bt:String id="AiAssistantButton.Label" DefaultValue="AI 助手"/>
```

在 `<bt:LongStrings>` 中追加工具提示：

```xml
        <bt:String id="AiAssistantButton.Tooltip" DefaultValue="打开 AI 助手聊天面板，用自然语言处理 Excel 数据"/>
```

- [ ] 插入 AiAssistantGroup 到 manifest
- [ ] 追加 AiIcon 资源到 bt:Images
- [ ] 追加 AiAssistantTaskpane.Url 到 bt:Urls
- [ ] 追加 AiAssistantGroup.Label / AiAssistantButton.Label 到 bt:ShortStrings
- [ ] 追加 AiAssistantButton.Tooltip 到 bt:LongStrings

### Sub-task 1.4: 验证构建

- [ ] 运行 `npm run build:dev` 确认构建成功
- [ ] 确认 dist 目录包含 `ai-assistant-taskpane.html`
- [ ] 提交：`git add -A && git commit -m "feat: add AI assistant button, webpack entry, and manifest registration"`

---

## Task 2: 编写并测试 ai-utils.js

**依赖：** 无（纯工具模块，不依赖 UI）

**文件：** 创建 `src/utils/ai-utils.js`

核心模块，包含四个子模块：DeepSeekClient, ToolRegistry, ContextManager, ExcelDataService。

### Sub-task 2.1: 实现 DeepSeekClient

```js
/**
 * DeepSeek API 客户端
 * 封装与 DeepSeek Chat API 的通信，支持 Function Calling
 */

var DEEPSEEK_API_URL = "https://api.deepseek.com/chat/completions";
var DEEPSEEK_MODEL = "deepseek-chat";

/**
 * 发送聊天请求到 DeepSeek API
 * @param {string} apiKey - DeepSeek API Key
 * @param {Array} messages - 消息数组 [{role, content, ...}]
 * @param {Array} tools - Function Calling 工具定义
 * @param {Object} [options] - 可选参数
 * @param {number} [options.temperature=0.3] - 生成温度
 * @param {number} [options.maxTokens=4096] - 最大 Token 数
 * @returns {Promise<Object>} API 响应
 */
function sendChatRequest(apiKey, messages, tools, options) {
  options = options || {};
  var temperature = options.temperature !== undefined ? options.temperature : 0.3;
  var maxTokens = options.maxTokens || 4096;

  return fetch(DEEPSEEK_API_URL, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: "Bearer " + apiKey,
    },
    body: JSON.stringify({
      model: DEEPSEEK_MODEL,
      messages: messages,
      tools: tools,
      temperature: temperature,
      max_tokens: maxTokens,
    }),
  })
    .then(function (response) {
      if (!response.ok) {
        return response.json().then(function (err) {
          throw new Error("DeepSeek API 错误: " + (err.error?.message || response.statusText));
        });
      }
      return response.json();
    })
    .then(function (data) {
      return data;
    });
}

/**
 * 解析 API 响应的 choice 内容
 * @param {Object} response - API 完整响应
 * @returns {Object} {message, toolCalls}
 *   - message: AI 回复的消息对象
 *   - toolCalls: 如果有 tool_calls 则返回数组，否则 null
 */
function parseResponse(response) {
  var choice = response.choices && response.choices[0];
  if (!choice) {
    throw new Error("API 返回为空");
  }

  var message = choice.message;
  var toolCalls = message.tool_calls || null;

  return {
    message: message,
    toolCalls: toolCalls,
  };
}
```

- [ ] 实现 `sendChatRequest(apiKey, messages, tools, options)`
- [ ] 实现 `parseResponse(response)`

### Sub-task 2.2: 实现 ToolRegistry

```js
/**
 * 工具注册表
 * 管理 5 个 Function Calling 工具的 Schema 定义和执行函数
 */

var TOOLS = [
  {
    type: "function",
    function: {
      name: "generate_formula",
      description: "根据用户描述生成 Excel 公式并写入单元格",
      parameters: {
        type: "object",
        properties: {
          formula: {
            type: "string",
            description: "Excel 公式字符串，如 =IF(A2>100,\"高\",\"低\")",
          },
          targetCell: {
            type: "string",
            description: "公式写入的目标单元格，如 B2",
          },
          fillRange: {
            type: "string",
            description: "可选，自动填充范围，如 B2:B50",
          },
          explanation: {
            type: "string",
            description: "对公式功能的自然语言解释，展示给用户",
          },
        },
        required: ["formula", "targetCell", "explanation"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "analyze_data",
      description: "分析选中的 Excel 数据，返回分析结论（在聊天面板中显示结果）",
      parameters: {
        type: "object",
        properties: {
          conclusion: {
            type: "string",
            description: "分析结论的文本内容，使用 Markdown 格式",
          },
          insights: {
            type: "array",
            items: { type: "string" },
            description: "关键洞察点列表",
          },
        },
        required: ["conclusion"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "classify_data",
      description: "对选中数据进行分类、提取或标注，结果写回新列",
      parameters: {
        type: "object",
        properties: {
          range: {
            type: "string",
            description: "要处理的数据范围，如 A2:A100",
          },
          task: {
            type: "string",
            enum: ["分类", "提取", "标注"],
            description: "任务类型：分类、提取或标注",
          },
          instruction: {
            type: "string",
            description: "处理规则说明，如"判断每行的评论是正面、负面还是中性"",
          },
          outputColumn: {
            type: "string",
            description: "结果写入的目标列字母，如 C",
          },
          labelMapping: {
            type: "object",
            description: "可选，分类标签到数值的映射，如 {\"正面\": 1, \"负面\": -1}",
          },
        },
        required: ["range", "task", "instruction", "outputColumn"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "write_to_cells",
      description: "将数据数组写入 Excel 指定区域",
      parameters: {
        type: "object",
        properties: {
          range: {
            type: "string",
            description: "写入的目标范围，如 C2:C100",
          },
          values: {
            type: "array",
            items: {
              type: "array",
              items: {},
            },
            description: "要写入的二维数组，每行一个子数组",
          },
          insertBefore: {
            type: "boolean",
            description: "是否在写入前先插入新列",
          },
        },
        required: ["range", "values"],
      },
    },
  },
];

/**
 * 获取 Function Calling 工具定义数组
 * @returns {Array} tools 数组
 */
function getToolDefinitions() {
  return TOOLS;
}

/**
 * 根据 tool name 查找对应的工具定义
 * @param {string} name - 工具名称
 * @returns {Object|null} 工具定义或 null
 */
function findToolDefinition(name) {
  for (var i = 0; i < TOOLS.length; i++) {
    if (TOOLS[i].function.name === name) {
      return TOOLS[i];
    }
  }
  return null;
}
```

- [ ] 实现 `getToolDefinitions()`
- [ ] 实现 `findToolDefinition(name)`

### Sub-task 2.3: 实现 ContextManager

```js
/**
 * 对话上下文管理器
 * 管理消息历史、滑动窗口裁剪、数据上下文注入
 */

var MAX_HISTORY_ROUNDS = 3; // 保留最近 3 轮对话

/**
 * 构建 System Prompt（含数据上下文）
 * @param {Object|null} selectionSummary - Excel 选区摘要（可选）
 * @returns {string} system prompt 字符串
 */
function buildSystemPrompt(selectionSummary) {
  var prompt = [
    "你是 Excel 数据处理助手，运行在 Excel 加载项中。",
    "你拥有以下工具可以使用：",
    "- generate_formula：根据用户描述生成 Excel 公式并写入单元格",
    "- analyze_data：分析选中的数据，返回分析结论",
    "- classify_data：对数据进行分类、提取或标注",
    "- write_to_cells：将数据写入 Excel 指定区域",
    "",
    "注意事项：",
    "1. 分析类操作 → 在聊天中显示结果（使用 analyze_data 工具）",
    "2. 修改类操作 → 使用 write_to_cells 工具写回 Excel",
    "3. 公式生成 → 使用 generate_formula 工具",
    "4. 分类/提取 → 使用 classify_data 工具",
    "5. 解释你的每一步操作，让用户知道发生了什么",
  ];

  if (selectionSummary) {
    prompt.push("");
    prompt.push("当前用户选中的数据：");
    prompt.push("选区地址：" + (selectionSummary.address || "未知"));
    prompt.push("列数：" + (selectionSummary.columnCount || 0));
    prompt.push("行数：" + (selectionSummary.rowCount || 0));
    if (selectionSummary.headers && selectionSummary.headers.length) {
      prompt.push("列标题：" + selectionSummary.headers.join(" | "));
    }
    if (selectionSummary.sampleData && selectionSummary.sampleData.length) {
      prompt.push("样本数据（前 " + selectionSummary.sampleData.length + " 行）：");
      selectionSummary.sampleData.forEach(function (row, idx) {
        prompt.push("  第" + (idx + 1) + "行: " + JSON.stringify(row));
      });
    }
    if (selectionSummary.stats) {
      var statsText = [];
      for (var col in selectionSummary.stats) {
        if (selectionSummary.stats.hasOwnProperty(col)) {
          var s = selectionSummary.stats[col];
          statsText.push(col + ": 最大=" + s.max + " 最小=" + s.min + " 平均=" + s.avg);
        }
      }
      if (statsText.length) {
        prompt.push("数值统计：" + statsText.join("；"));
      }
    }
  }

  return prompt.join("\n");
}

/**
 * 创建初始对话上下文
 * @param {Object|null} selectionSummary - Excel 选区摘要
 * @returns {Array} messages 数组
 */
function createContext(selectionSummary) {
  return [
    { role: "system", content: buildSystemPrompt(selectionSummary) },
  ];
}

/**
 * 向上下文中追加一条消息
 * @param {Array} context - 当前上下文数组
 * @param {Object} message - 要追加的消息 {role, content, ...}
 * @returns {Array} 新的上下文数组（不可变更新）
 */
function addMessage(context, message) {
  return context.concat([message]);
}

/**
 * 裁剪上下文：保留 system prompt + 最近 N 轮对话
 * @param {Array} context - 当前上下文数组
 * @returns {Array} 裁剪后的上下文数组
 */
function trimContext(context) {
  var systemMsg = null;
  var otherMessages = [];

  for (var i = 0; i < context.length; i++) {
    if (context[i].role === "system") {
      systemMsg = context[i];
    } else {
      otherMessages.push(context[i]);
    }
  }

  // 计算对话轮次：每轮 = user message + assistant message (含可能的 tool call)
  // 从后往前保留 MAX_HISTORY_ROUNDS 轮
  var rounds = [];
  var currentRound = [];
  for (var j = otherMessages.length - 1; j >= 0; j--) {
    currentRound.unshift(otherMessages[j]);
    if (otherMessages[j].role === "user") {
      rounds.unshift(currentRound);
      currentRound = [];
      if (rounds.length >= MAX_HISTORY_ROUNDS) {
        break;
      }
    }
  }

  var kept = [];
  if (currentRound.length > 0) {
    kept = currentRound;
  }
  rounds.forEach(function (round) {
    kept = kept.concat(round);
  });

  var result = systemMsg ? [systemMsg] : [];
  return result.concat(kept);
}

/**
 * 更新上下文中的数据摘要（替换 system prompt + 保留历史）
 * @param {Array} context - 当前上下文
 * @param {Object} selectionSummary - 新的选区摘要
 * @returns {Array} 更新后的上下文
 */
function updateSelection(context, selectionSummary) {
  var trimmed = trimContext(context);
  var systemMsg = { role: "system", content: buildSystemPrompt(selectionSummary) };
  var others = [];
  for (var i = 0; i < trimmed.length; i++) {
    if (trimmed[i].role !== "system") {
      others.push(trimmed[i]);
    }
  }
  return [systemMsg].concat(others);
}
```

- [ ] 实现 `buildSystemPrompt(selectionSummary)`
- [ ] 实现 `createContext(selectionSummary)`
- [ ] 实现 `addMessage(context, message)`
- [ ] 实现 `trimContext(context)`
- [ ] 实现 `updateSelection(context, selectionSummary)`

### Sub-task 2.4: 实现 ExcelDataService

```js
/**
 * Excel 数据服务
 * 读取选区数据、生成摘要、执行写入操作（通过 Office JS API）
 */

var SAMPLE_SIZE = 10; // 默认样本行数

/**
 * 从当前选区提取摘要信息（不发送全部数据到 API）
 * 必须在 Excel.run 回调中调用
 * @param {Excel.RequestContext} context - Excel 请求上下文
 * @returns {Promise} 解析为 {address, columnCount, rowCount, headers, sampleData, stats}
 */
function getSelectionSummary(context) {
  var range = context.workbook.getSelectedRange();
  range.load(["address", "columnCount", "rowCount", "values"]);

  return context.sync().then(function () {
    var values = range.values; // 二维数组
    var headers = values.length > 0 ? values[0].map(String) : [];
    var dataRows = values.slice(1); // 剔除标题行

    // 提取样本
    var sampleData = dataRows.slice(0, SAMPLE_SIZE);

    // 计算数值列统计
    var stats = {};
    for (var colIdx = 0; colIdx < headers.length; colIdx++) {
      var nums = [];
      for (var r = 0; r < dataRows.length; r++) {
        var val = dataRows[r][colIdx];
        if (typeof val === "number" && !isNaN(val)) {
          nums.push(val);
        }
      }
      if (nums.length > 1) {
        var max = nums[0], min = nums[0], sum = 0;
        for (var k = 0; k < nums.length; k++) {
          if (nums[k] > max) max = nums[k];
          if (nums[k] < min) min = nums[k];
          sum += nums[k];
        }
        stats[headers[colIdx]] = {
          max: max,
          min: min,
          avg: Math.round((sum / nums.length) * 100) / 100,
          count: nums.length,
        };
      }
    }

    return {
      address: range.address,
      columnCount: range.columnCount,
      rowCount: Math.max(0, dataRows.length),
      headers: headers,
      sampleData: sampleData,
      stats: stats,
    };
  });
}

/**
 * 将公式写入指定单元格
 * @param {Excel.RequestContext} context - 请求上下文
 * @param {string} targetCell - 目标单元格，如 B2
 * @param {string} formula - 公式字符串
 * @param {string|null} fillRange - 可选自动填充范围
 * @returns {Promise}
 */
function writeFormula(context, targetCell, formula, fillRange) {
  var worksheet = context.workbook.worksheets.getActiveWorksheet();
  var cell = worksheet.getRange(targetCell);
  cell.formulas = [[formula]];
  return context.sync().then(function () {
    if (fillRange) {
      var fillTarget = worksheet.getRange(fillRange);
      cell.autoFill(fillTarget);
      return context.sync();
    }
  });
}

/**
 * 将数据写入指定范围
 * @param {Excel.RequestContext} context - 请求上下文
 * @param {string} rangeAddress - 范围地址，如 C2:C100
 * @param {Array} values - 二维数组
 * @param {boolean} insertBefore - 是否先插入列
 * @returns {Promise}
 */
function writeValues(context, rangeAddress, values, insertBefore) {
  var worksheet = context.workbook.worksheets.getActiveWorksheet();

  if (insertBefore) {
    // 插入新列
    var colLetter = rangeAddress.replace(/[0-9]/g, "").replace(":", "");
    worksheet.getRange(colLetter + ":" + colLetter).insert(Excel.InsertShiftDirection.right);
    return context.sync().then(function () {
      var targetRange = worksheet.getRange(rangeAddress);
      targetRange.values = values;
      return context.sync();
    });
  }

  var targetRange = worksheet.getRange(rangeAddress);
  targetRange.values = values;
  return context.sync();
}

/**
 * 对数据进行分类/提取/标注处理
 * 读取指定范围的数据，逐行应用分类规则，写入输出列
 * @param {Excel.RequestContext} context - 请求上下文
 * @param {string} sourceRange - 数据源范围，如 A2:A100
 * @param {function} classifyFn - 分类函数 (cellValue) => resultValue
 * @param {string} outputColumn - 输出列字母，如 C
 * @returns {Promise}
 */
function applyClassification(context, sourceRange, classifyFn, outputColumn) {
  var worksheet = context.workbook.worksheets.getActiveWorksheet();
  var source = worksheet.getRange(sourceRange);
  source.load("values");
  return context.sync().then(function () {
    var values = source.values;
    var results = [];
    for (var i = 0; i < values.length; i++) {
      var result = classifyFn(values[i][0]);
      results.push([result]);
    }

    // 写入结果列
    var startRow = parseInt(sourceRange.match(/[0-9]+/)[0], 10);
    var endRow = startRow + values.length - 1;
    var targetAddress = outputColumn + startRow + ":" + outputColumn + endRow;
    var targetRange = worksheet.getRange(targetAddress);
    targetRange.values = results;
    return context.sync();
  });
}
```

- [ ] 实现 `getSelectionSummary(context)`
- [ ] 实现 `writeFormula(context, targetCell, formula, fillRange)`
- [ ] 实现 `writeValues(context, rangeAddress, values, insertBefore)`
- [ ] 实现 `applyClassification(context, sourceRange, classifyFn, outputColumn)`

### Sub-task 2.5: 模块导出

在 `ai-utils.js` 文件末尾添加导出：

```js
module.exports = {
  // DeepSeekClient
  sendChatRequest: sendChatRequest,
  parseResponse: parseResponse,
  // ToolRegistry
  getToolDefinitions: getToolDefinitions,
  findToolDefinition: findToolDefinition,
  // ContextManager
  buildSystemPrompt: buildSystemPrompt,
  createContext: createContext,
  addMessage: addMessage,
  trimContext: trimContext,
  updateSelection: updateSelection,
  // ExcelDataService
  getSelectionSummary: getSelectionSummary,
  writeFormula: writeFormula,
  writeValues: writeValues,
  applyClassification: applyClassification,
};
```

- [ ] 添加 `module.exports`
- [ ] 运行 `npm run lint` 确认无 lint 错误

### Sub-task 2.6: 提交

- [ ] `git add -A && git commit -m "feat: add ai-utils.js with DeepSeekClient, ToolRegistry, ContextManager, ExcelDataService"`

---

## Task 3: 编写 ai-utils.test.js 单元测试

**依赖：** Task 2 完成后

**文件：** 创建 `src/utils/ai-utils.test.js`

### Sub-task 3.1: 测试 DeepSeekClient

```js
var {
  sendChatRequest,
  parseResponse,
  getToolDefinitions,
  findToolDefinition,
  buildSystemPrompt,
  createContext,
  addMessage,
  trimContext,
  updateSelection,
  getSelectionSummary,
  writeFormula,
  writeValues,
  applyClassification,
} = require("./ai-utils");

describe("DeepSeekClient", function () {
  describe("parseResponse", function () {
    it("extracts message from valid response", function () {
      var response = {
        choices: [
          {
            message: { role: "assistant", content: "Hello" },
          },
        ],
      };
      var result = parseResponse(response);
      expect(result.message.content).toBe("Hello");
      expect(result.toolCalls).toBeNull();
    });

    it("extracts tool_calls when present", function () {
      var response = {
        choices: [
          {
            message: {
              role: "assistant",
              content: null,
              tool_calls: [{ id: "call_1", function: { name: "test", arguments: "{}" } }],
            },
          },
        ],
      };
      var result = parseResponse(response);
      expect(result.toolCalls).toHaveLength(1);
      expect(result.toolCalls[0].function.name).toBe("test");
    });

    it("throws for empty choices", function () {
      expect(function () {
        parseResponse({ choices: [] });
      }).toThrow();
    });
  });

  describe("sendChatRequest", function () {
    // 不测试真实网络请求，仅验证函数存在
    it("is a function", function () {
      expect(typeof sendChatRequest).toBe("function");
    });
  });
});
```

- [ ] 编写 parseResponse 测试（有效响应、含 tool_calls、空响应三种情况）

### Sub-task 3.2: 测试 ToolRegistry

```js
describe("ToolRegistry", function () {
  describe("getToolDefinitions", function () {
    it("returns 4 tools", function () {
      var tools = getToolDefinitions();
      expect(tools).toHaveLength(4);
    });

    it("includes generate_formula", function () {
      var names = getToolDefinitions().map(function (t) { return t.function.name; });
      expect(names).toContain("generate_formula");
      expect(names).toContain("analyze_data");
      expect(names).toContain("classify_data");
      expect(names).toContain("write_to_cells");
    });

    it("each tool has required fields", function () {
      var tools = getToolDefinitions();
      for (var i = 0; i < tools.length; i++) {
        expect(tools[i].type).toBe("function");
        expect(tools[i].function.name).toBeDefined();
        expect(tools[i].function.description).toBeDefined();
        expect(tools[i].function.parameters).toBeDefined();
      }
    });
  });

  describe("findToolDefinition", function () {
    it("finds existing tool by name", function () {
      var tool = findToolDefinition("generate_formula");
      expect(tool).not.toBeNull();
      expect(tool.function.name).toBe("generate_formula");
    });

    it("returns null for unknown tool", function () {
      expect(findToolDefinition("nonexistent")).toBeNull();
    });
  });
});
```

- [ ] 编写 getToolDefinitions 测试（数量、名称、结构）
- [ ] 编写 findToolDefinition 测试（存在/不存在）

### Sub-task 3.3: 测试 ContextManager

```js
describe("ContextManager", function () {
  describe("buildSystemPrompt", function () {
    it("returns prompt without selection summary", function () {
      var prompt = buildSystemPrompt(null);
      expect(prompt).toContain("Excel 数据处理助手");
      expect(prompt).toContain("generate_formula");
    });

    it("includes selection summary when provided", function () {
      var summary = {
        address: "A1:D100",
        columnCount: 4,
        rowCount: 100,
        headers: ["姓名", "年龄"],
        sampleData: [["张三", 28]],
        stats: { "年龄": { max: 65, min: 18, avg: 35 } },
      };
      var prompt = buildSystemPrompt(summary);
      expect(prompt).toContain("A1:D100");
      expect(prompt).toContain("姓名");
      expect(prompt).toContain("年龄");
      expect(prompt).toContain("最大=65");
    });
  });

  describe("createContext", function () {
    it("creates context with system message", function () {
      var ctx = createContext(null);
      expect(ctx).toHaveLength(1);
      expect(ctx[0].role).toBe("system");
    });
  });

  describe("addMessage", function () {
    it("appends message without mutating original", function () {
      var ctx = createContext(null);
      var msg = { role: "user", content: "hello" };
      var newCtx = addMessage(ctx, msg);
      expect(ctx).toHaveLength(1); // 原数组不变
      expect(newCtx).toHaveLength(2);
      expect(newCtx[1].content).toBe("hello");
    });
  });

  describe("trimContext", function () {
    it("preserves system message", function () {
      var ctx = createContext(null);
      ctx = addMessage(ctx, { role: "user", content: "hi" });
      ctx = addMessage(ctx, { role: "assistant", content: "hello" });
      var trimmed = trimContext(ctx);
      expect(trimmed[0].role).toBe("system");
    });

    it("keeps only recent rounds when limit exceeded", function () {
      var ctx = createContext(null);
      // 添加 5 轮对话
      for (var i = 0; i < 5; i++) {
        ctx = addMessage(ctx, { role: "user", content: "q" + i });
        ctx = addMessage(ctx, { role: "assistant", content: "a" + i });
      }
      var trimmed = trimContext(ctx);
      // 应保留 system message + 最近 3 轮 = 1 + 6 = 7 条
      expect(trimmed.length).toBeLessThanOrEqual(7);
      // 最早的 q0 应该被裁剪掉
      var allContent = trimmed.map(function (m) { return m.content; }).join(" ");
      expect(allContent).not.toContain("q0");
      expect(allContent).toContain("q4");
    });
  });

  describe("updateSelection", function () {
    it("replaces system prompt with new selection info", function () {
      var ctx = createContext(null);
      var newSummary = { address: "NEW_RANGE", columnCount: 2, rowCount: 5 };
      var updated = updateSelection(ctx, newSummary);
      expect(updated[0].content).toContain("NEW_RANGE");
    });
  });
});
```

- [ ] 编写 buildSystemPrompt 测试（有/无摘要两种情况）
- [ ] 编写 createContext 测试
- [ ] 编写 addMessage 测试（不可变性验证）
- [ ] 编写 trimContext 测试（保留 system、裁剪旧轮次）
- [ ] 编写 updateSelection 测试

### Sub-task 3.4: 运行测试确认全部通过

- [ ] 运行 `npx jest src/utils/ai-utils.test.js --verbose`
- [ ] 确认全部测试 PASS

### Sub-task 3.5: 提交

- [ ] `git add -A && git commit -m "test: add ai-utils unit tests"`

---

## Task 4: 编写 AI 聊天面板 HTML

**依赖：** Task 1 完成后

**文件：** 创建 `src/taskpane/ai-assistant-taskpane.html`

```html
<!-- Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT License. -->

<!DOCTYPE html>
<html>

<head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>AI 助手</title>

    <!-- Office JavaScript API -->
    <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>

    <!-- Fluent UI Core -->
    <link rel="stylesheet" href="https://res-1.cdn.office.net/files/fabric-cdn-prod_20230815.002/office-ui-fabric-core/11.1.0/css/fabric.min.css"/>

    <link href="taskpane.css" rel="stylesheet" type="text/css" />
    <link href="ai-assistant-taskpane.css" rel="stylesheet" type="text/css" />
</head>

<body class="ms-font-m ms-welcome ms-Fabric">
    <div class="ai-container">
        <!-- Header -->
        <div class="ai-header">
            <div class="ai-header-title">
                <span class="ai-header-icon">🤖</span>
                <span>AI 助手</span>
            </div>
            <button id="settingsBtn" class="ai-icon-btn" title="设置">⚙️</button>
        </div>

        <!-- Chat Messages -->
        <div id="chatMessages" class="ai-chat-messages">
            <div class="ai-message ai-message-assistant">
                <div class="ai-message-content">
                    👋 你好！我是 AI 助手。选中 Excel 中的数据后，告诉我你想做什么。
                </div>
            </div>
        </div>

        <!-- Status Bar -->
        <div id="statusBar" class="ai-status-bar ai-status-idle">
            <span id="selectionInfo">未选中数据</span>
        </div>

        <!-- Input Area -->
        <div class="ai-input-area">
            <textarea id="userInput" class="ai-input" rows="2" placeholder="输入指令...（Enter 发送，Shift+Enter 换行）"></textarea>
            <button id="sendBtn" class="ai-send-btn" disabled>发送</button>
        </div>

        <!-- Settings Modal -->
        <div id="settingsModal" class="ai-modal" style="display:none;">
            <div class="ai-modal-content">
                <div class="ai-modal-header">
                    <span>⚙️ 设置</span>
                    <button id="closeSettingsBtn" class="ai-icon-btn">✕</button>
                </div>
                <div class="ai-modal-body">
                    <div class="ai-form-group">
                        <label for="apiKeyInput" class="ai-label">DeepSeek API Key</label>
                        <input type="password" id="apiKeyInput" class="ai-form-input" placeholder="sk-..." />
                    </div>
                    <div class="ai-form-group">
                        <label for="modelSelect" class="ai-label">模型</label>
                        <select id="modelSelect" class="ai-form-input">
                            <option value="deepseek-chat">deepseek-chat</option>
                            <option value="deepseek-reasoner">deepseek-reasoner</option>
                        </select>
                    </div>
                    <div class="ai-form-group">
                        <label for="temperatureInput" class="ai-label">温度 (0-1)</label>
                        <input type="number" id="temperatureInput" class="ai-form-input" value="0.3" min="0" max="1" step="0.1" />
                    </div>
                </div>
                <div class="ai-modal-footer">
                    <button id="saveSettingsBtn" class="concat-button concat-button--small">保存</button>
                </div>
            </div>
        </div>

        <!-- Confirm Box -->
        <div id="confirmBox" class="confirm-box" style="display:none;">
            <div id="confirmMsg" class="confirm-msg"></div>
            <div class="confirm-buttons">
                <button id="confirmYes" class="concat-button concat-button--small">确定</button>
                <button id="confirmNo" class="concat-button concat-button--small concat-button--secondary">取消</button>
            </div>
        </div>
    </div>
</body>

</html>
```

- [ ] 编写 HTML 文件（聊天区域、输入框、设置弹窗、确认框）

---

## Task 5: 编写 AI 聊天面板 CSS

**文件：** 创建 `src/taskpane/ai-assistant-taskpane.css`

```css
/* AI 助手聊天面板样式 */

.ai-container {
  display: flex;
  flex-direction: column;
  height: 100vh;
  padding: 0;
  background: #fafafa;
}

/* Header */
.ai-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 12px 16px;
  background: #0078d4;
  color: white;
  flex-shrink: 0;
}

.ai-header-title {
  display: flex;
  align-items: center;
  gap: 8px;
  font-size: 15px;
  font-weight: 600;
}

.ai-header-icon {
  font-size: 18px;
}

.ai-icon-btn {
  background: none;
  border: none;
  color: white;
  font-size: 16px;
  cursor: pointer;
  padding: 4px 8px;
  border-radius: 4px;
  opacity: 0.8;
}

.ai-icon-btn:hover {
  opacity: 1;
  background: rgba(255,255,255,0.15);
}

/* Chat Messages */
.ai-chat-messages {
  flex: 1;
  overflow-y: auto;
  padding: 12px 16px;
  display: flex;
  flex-direction: column;
  gap: 10px;
}

.ai-message {
  max-width: 90%;
  animation: fadeIn 0.2s ease;
}

@keyframes fadeIn {
  from { opacity: 0; transform: translateY(4px); }
  to { opacity: 1; transform: translateY(0); }
}

.ai-message-user {
  align-self: flex-end;
}

.ai-message-assistant {
  align-self: flex-start;
}

.ai-message-content {
  padding: 10px 14px;
  border-radius: 12px;
  font-size: 13px;
  line-height: 1.5;
  word-break: break-word;
  white-space: pre-wrap;
}

.ai-message-user .ai-message-content {
  background: #0078d4;
  color: white;
  border-bottom-right-radius: 4px;
}

.ai-message-assistant .ai-message-content {
  background: white;
  color: #333;
  border: 1px solid #e0e0e0;
  border-bottom-left-radius: 4px;
}

.ai-message-assistant .ai-message-content code {
  background: #f0f0f0;
  padding: 1px 4px;
  border-radius: 3px;
  font-size: 12px;
}

/* Loading indicator */
.ai-loading {
  display: flex;
  align-items: center;
  gap: 4px;
  padding: 10px 14px;
}

.ai-loading-dot {
  width: 6px;
  height: 6px;
  background: #999;
  border-radius: 50%;
  animation: bounce 1.4s ease-in-out infinite;
}

.ai-loading-dot:nth-child(2) { animation-delay: 0.2s; }
.ai-loading-dot:nth-child(3) { animation-delay: 0.4s; }

@keyframes bounce {
  0%, 80%, 100% { transform: scale(0.6); opacity: 0.4; }
  40% { transform: scale(1); opacity: 1; }
}

/* Status Bar */
.ai-status-bar {
  padding: 6px 16px;
  font-size: 11px;
  color: #888;
  border-top: 1px solid #e0e0e0;
  flex-shrink: 0;
}

.ai-status-idle { background: #fafafa; }
.ai-status-loading { background: #fff3cd; color: #856404; }
.ai-status-error { background: #f8d7da; color: #721c24; }
.ai-status-success { background: #d4edda; color: #155724; }

/* Input Area */
.ai-input-area {
  display: flex;
  gap: 8px;
  padding: 10px 16px;
  border-top: 1px solid #e0e0e0;
  background: white;
  flex-shrink: 0;
}

.ai-input {
  flex: 1;
  padding: 8px 12px;
  border: 2px solid #ddd;
  border-radius: 8px;
  font-size: 13px;
  outline: none;
  resize: none;
  font-family: inherit;
  transition: border-color 0.2s;
  box-sizing: border-box;
}

.ai-input:focus {
  border-color: #0078d4;
}

.ai-send-btn {
  padding: 8px 16px;
  background: #0078d4;
  color: white;
  border: none;
  border-radius: 8px;
  font-size: 13px;
  font-weight: 600;
  cursor: pointer;
  transition: background 0.2s;
  align-self: flex-end;
}

.ai-send-btn:hover {
  background: #005a9e;
}

.ai-send-btn:disabled {
  background: #ccc;
  cursor: not-allowed;
}

/* Settings Modal */
.ai-modal {
  position: fixed;
  top: 0;
  left: 0;
  right: 0;
  bottom: 0;
  background: rgba(0,0,0,0.4);
  display: flex;
  align-items: center;
  justify-content: center;
  z-index: 1000;
}

.ai-modal-content {
  background: white;
  border-radius: 8px;
  width: 90%;
  max-width: 320px;
  box-shadow: 0 8px 32px rgba(0,0,0,0.2);
}

.ai-modal-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 14px 16px;
  border-bottom: 1px solid #e0e0e0;
  font-weight: 600;
  font-size: 14px;
}

.ai-modal-body {
  padding: 16px;
}

.ai-modal-footer {
  padding: 12px 16px;
  border-top: 1px solid #e0e0e0;
  text-align: right;
}

.ai-form-group {
  margin-bottom: 12px;
}

.ai-label {
  display: block;
  font-size: 12px;
  color: #666;
  margin-bottom: 4px;
}

.ai-form-input {
  width: 100%;
  padding: 8px 10px;
  border: 1px solid #ddd;
  border-radius: 4px;
  font-size: 13px;
  outline: none;
  box-sizing: border-box;
}

.ai-form-input:focus {
  border-color: #0078d4;
}

/* Reuse concat-button styles for consistency */
.concat-button {
  width: 100%;
  padding: 10px 16px;
  background: transparent;
  color: #0078d4;
  border: 2px solid #0078d4;
  border-radius: 6px;
  font-size: 14px;
  font-weight: 600;
  cursor: pointer;
  transition: all 0.2s;
}

.concat-button:hover {
  background: #0078d4;
  color: white;
}

.concat-button:disabled {
  background: #f5f5f5;
  color: #999;
  border-color: #999;
  cursor: not-allowed;
}

.concat-button--small {
  padding: 6px 16px;
  font-size: 13px;
  width: auto;
}

.concat-button--secondary {
  background: transparent;
  color: #666;
  border-color: #999;
}

.concat-button--secondary:hover {
  background: #f0f0f0;
  color: #333;
}

/* Confirm box */
.confirm-box {
  padding: 16px;
  background: #fff3cd;
  border: 2px solid #ffc107;
  border-radius: 6px;
  margin: 8px 16px;
}

.confirm-msg {
  font-size: 13px;
  color: #856404;
  margin-bottom: 12px;
  line-height: 1.5;
}

.confirm-buttons {
  text-align: right;
}

.confirm-buttons .concat-button + .concat-button {
  margin-left: 8px;
}
```

- [ ] 编写 CSS 文件

---

## Task 6: 编写 AI 聊天面板主逻辑 JS

**依赖：** Task 2, 4, 5 完成后

**文件：** 创建 `src/taskpane/ai-assistant-taskpane.js`

完整代码：

```js
/* global console, document, Excel, Office, localStorage */

var aiUtils = require("../utils/ai-utils");

// ---- 状态 ----
var apiKey = localStorage.getItem("deepseek_api_key") || "";
var settings = {
  model: localStorage.getItem("deepseek_model") || "deepseek-chat",
  temperature: parseFloat(localStorage.getItem("deepseek_temperature") || "0.3"),
  maxTokens: 4096,
};
var context = aiUtils.createContext(null); // 对话上下文
var currentSelection = null; // 当前选区摘要
var isProcessing = false;

// ---- DOM 引用 ----
var chatMessages, userInput, sendBtn, statusBar, selectionInfo;
var settingsModal, apiKeyInput, modelSelect, temperatureInput;
var confirmBox, confirmMsg, confirmYes, confirmNo;

// ---- 初始化 ----
Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    cacheDom();
    bindEvents();
    refreshSelection();
    restoreApiKeyState();
  }
});

function cacheDom() {
  chatMessages = document.getElementById("chatMessages");
  userInput = document.getElementById("userInput");
  sendBtn = document.getElementById("sendBtn");
  statusBar = document.getElementById("statusBar");
  selectionInfo = document.getElementById("selectionInfo");
  settingsModal = document.getElementById("settingsModal");
  apiKeyInput = document.getElementById("apiKeyInput");
  modelSelect = document.getElementById("modelSelect");
  temperatureInput = document.getElementById("temperatureInput");
  confirmBox = document.getElementById("confirmBox");
  confirmMsg = document.getElementById("confirmMsg");
  confirmYes = document.getElementById("confirmYes");
  confirmNo = document.getElementById("confirmNo");
}

function bindEvents() {
  document.getElementById("settingsBtn").addEventListener("click", openSettings);
  document.getElementById("closeSettingsBtn").addEventListener("click", closeSettings);
  document.getElementById("saveSettingsBtn").addEventListener("click", saveSettings);
  sendBtn.addEventListener("click", handleSend);
  userInput.addEventListener("keydown", function (e) {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      handleSend();
    }
  });
  userInput.addEventListener("input", function () {
    sendBtn.disabled = !userInput.value.trim() || !apiKey || isProcessing;
  });
  confirmYes.addEventListener("click", confirmAction);
  confirmNo.addEventListener("click", cancelConfirm);
}

function restoreApiKeyState() {
  if (apiKey) {
    apiKeyInput.value = apiKey;
    modelSelect.value = settings.model;
    temperatureInput.value = settings.temperature;
    sendBtn.disabled = !userInput.value.trim();
  } else {
    openSettings();
  }
}

// ---- 选区刷新 ----
function refreshSelection() {
  setStatus("正在获取选区...", "loading");
  Excel.run(function (ctx) {
    return aiUtils.getSelectionSummary(ctx).then(function (summary) {
      currentSelection = summary;
      context = aiUtils.updateSelection(context, summary);
      selectionInfo.textContent = "已选中: " + summary.address + " (" + summary.rowCount + "行 × " + summary.columnCount + "列)";
      setStatus("已就绪", "idle");
      // 添加选区变更提示
      addAssistantMessage("已获取选区 " + summary.address + "（" + summary.rowCount + "行，" + summary.columnCount + "列），请问你想做什么？");
    });
  }).catch(function (err) {
    currentSelection = null;
    selectionInfo.textContent = "未选中数据";
    setStatus("无法获取选区: " + err.message, "error");
  });
}

// ---- 发送消息 ----
function handleSend() {
  if (isProcessing || !apiKey) return;

  var text = userInput.value.trim();
  if (!text) return;

  userInput.value = "";
  sendBtn.disabled = true;
  isProcessing = true;

  addUserMessage(text);
  context = aiUtils.addMessage(context, { role: "user", content: text });
  showLoading();

  aiUtils.sendChatRequest(apiKey, aiUtils.trimContext(context), aiUtils.getToolDefinitions(), {
    temperature: settings.temperature,
    maxTokens: settings.maxTokens,
  })
    .then(function (response) {
      return aiUtils.parseResponse(response);
    })
    .then(function (parsed) {
      hideLoading();
      context = aiUtils.addMessage(context, parsed.message);

      if (parsed.toolCalls && parsed.toolCalls.length > 0) {
        return handleToolCalls(parsed.toolCalls);
      }

      // 纯文本回复
      var content = parsed.message.content || "";
      addAssistantMessage(content);
    })
    .catch(function (err) {
      hideLoading();
      addAssistantMessage("❌ " + err.message);
      setStatus("错误", "error");
    })
    .finally(function () {
      isProcessing = false;
      sendBtn.disabled = !userInput.value.trim() || !apiKey;
    });
}

// ---- 处理 Tool 调用 ----
function handleToolCalls(toolCalls) {
  var results = [];

  // 每个 tool call 依次执行（串行，避免 Excel 操作冲突）
  var sequence = Promise.resolve();

  toolCalls.forEach(function (toolCall) {
    sequence = sequence.then(function () {
      var name = toolCall.function.name;
      var args = JSON.parse(toolCall.function.arguments);

      setStatus("正在执行: " + name + "...", "loading");

      if (name === "generate_formula") {
        return handleGenerateFormula(args);
      } else if (name === "analyze_data") {
        return handleAnalyzeData(args);
      } else if (name === "classify_data") {
        return handleClassifyData(args);
      } else if (name === "write_to_cells") {
        return handleWriteToCells(args);
      } else {
        return { error: "未知工具: " + name };
      }
    }).then(function (result) {
      results.push(result);
      // 将工具结果加入上下文
      context = aiUtils.addMessage(context, {
        role: "tool",
        tool_call_id: toolCall.id,
        content: JSON.stringify(result),
      });
    });
  });

  // 所有 tool 执行完后，将结果发给 AI 继续对话
  return sequence.then(function () {
    return aiUtils.sendChatRequest(apiKey, aiUtils.trimContext(context), aiUtils.getToolDefinitions(), {
      temperature: settings.temperature,
      maxTokens: settings.maxTokens,
    });
  }).then(function (response) {
    return aiUtils.parseResponse(response);
  }).then(function (parsed) {
    context = aiUtils.addMessage(context, parsed.message);
    addAssistantMessage(parsed.message.content || "");
    setStatus("完成", "success");
  });
}

// ---- Tool 处理函数 ----

// pendingConfirmation 用于确认框
var pendingConfirmAction = null;

function handleGenerateFormula(args) {
  return new Promise(function (resolve, reject) {
    var msg = "即将在 " + args.targetCell + " 写入公式：\n" + args.formula;
    if (args.fillRange) {
      msg += "\n并自动填充至 " + args.fillRange;
    }
    msg += "\n\n" + args.explanation;

    showConfirm(msg, function (confirmed) {
      if (!confirmed) {
        resolve({ status: "cancelled", message: "用户取消了操作" });
        return;
      }
      Excel.run(function (ctx) {
        return aiUtils.writeFormula(ctx, args.targetCell, args.formula, args.fillRange || null);
      }).then(function () {
        resolve({ status: "success", message: "公式已写入 " + args.targetCell });
      }).catch(function (err) {
        reject(err);
      });
    });
  });
}

function handleAnalyzeData(args) {
  // analyze_data 的结果直接在聊天中显示
  addAssistantMessage("📊 **分析结果**\n\n" + args.conclusion);
  if (args.insights && args.insights.length) {
    var insightsText = "**关键洞察：**\n";
    args.insights.forEach(function (item, i) {
      insightsText += (i + 1) + ". " + item + "\n";
    });
    addAssistantMessage(insightsText);
  }
  setStatus("分析完成", "success");
  return Promise.resolve({ status: "success" });
}

function handleClassifyData(args) {
  var classifyFn = createClassifyFunction(args.instruction, args.labelMapping);

  return new Promise(function (resolve, reject) {
    var msg = "即将对 " + args.range + " 执行" + args.task + "操作：\n" + args.instruction;
    msg += "\n结果写入 " + args.outputColumn + " 列";

    showConfirm(msg, function (confirmed) {
      if (!confirmed) {
        resolve({ status: "cancelled", message: "用户取消了操作" });
        return;
      }
      Excel.run(function (ctx) {
        return aiUtils.applyClassification(ctx, args.range, classifyFn, args.outputColumn);
      }).then(function () {
        resolve({ status: "success", message: args.task + "完成，结果已写入 " + args.outputColumn + " 列" });
      }).catch(function (err) {
        reject(err);
      });
    });
  });
}

function createClassifyFunction(instruction, labelMapping) {
  // 简单的关键词匹配分类器
  // 对于复杂分类，可以在这里调用 DeepSeek API 逐行处理
  return function (cellValue) {
    if (cellValue === null || cellValue === undefined || cellValue === "") {
      return "";
    }
    var text = String(cellValue).toLowerCase();

    if (instruction.indexOf("情感") >= 0 || instruction.indexOf("正面") >= 0) {
      var positive = ["好", "棒", "优秀", "满意", "推荐", "喜欢", "赞", "good", "great", "excellent", "positive"];
      var negative = ["差", "烂", "失望", "糟糕", "不满", "投诉", "bad", "poor", "terrible", "negative"];
      var posScore = positive.filter(function (w) { return text.indexOf(w) >= 0; }).length;
      var negScore = negative.filter(function (w) { return text.indexOf(w) >= 0; }).length;
      var label = posScore > negScore ? "正面" : (negScore > posScore ? "负面" : "中性");
      return labelMapping && labelMapping[label] !== undefined ? labelMapping[label] : label;
    }

    if (instruction.indexOf("提取") >= 0) {
      // 提取邮箱
      if (instruction.indexOf("邮箱") >= 0 || instruction.indexOf("邮件") >= 0) {
        var emailMatch = text.match(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/);
        return emailMatch ? emailMatch[0] : "";
      }
      // 提取手机号
      if (instruction.indexOf("手机") >= 0 || instruction.indexOf("电话") >= 0) {
        var phoneMatch = text.match(/1[3-9]\d{9}/);
        return phoneMatch ? phoneMatch[0] : "";
      }
    }

    // 默认返回原值
    return cellValue;
  };
}

function handleWriteToCells(args) {
  return new Promise(function (resolve, reject) {
    var msg = "即将在 " + args.range + " 写入 " + args.values.length + " 行数据";
    if (args.insertBefore) {
      msg += "（将先插入新列）";
    }

    showConfirm(msg, function (confirmed) {
      if (!confirmed) {
        resolve({ status: "cancelled", message: "用户取消了操作" });
        return;
      }
      Excel.run(function (ctx) {
        return aiUtils.writeValues(ctx, args.range, args.values, args.insertBefore || false);
      }).then(function () {
        resolve({ status: "success", message: "数据已写入 " + args.range });
      }).catch(function (err) {
        reject(err);
      });
    });
  });
}

// ---- 确认框 ----
var pendingConfirm = null;

function showConfirm(message, callback) {
  confirmMsg.textContent = message;
  confirmBox.style.display = "block";
  pendingConfirm = callback;
}

function confirmAction() {
  confirmBox.style.display = "none";
  if (pendingConfirm) {
    pendingConfirm(true);
    pendingConfirm = null;
  }
}

function cancelConfirm() {
  confirmBox.style.display = "none";
  if (pendingConfirm) {
    pendingConfirm(false);
    pendingConfirm = null;
  }
}

// ---- UI 辅助 ----
function addUserMessage(text) {
  var div = document.createElement("div");
  div.className = "ai-message ai-message-user";
  div.innerHTML = '<div class="ai-message-content">' + escapeHtml(text) + '</div>';
  chatMessages.appendChild(div);
  scrollToBottom();
}

function addAssistantMessage(text) {
  // 移除 loading 指示器
  removeLoading();
  var div = document.createElement("div");
  div.className = "ai-message ai-message-assistant";
  div.innerHTML = '<div class="ai-message-content">' + renderMarkdown(escapeHtml(text)) + '</div>';
  chatMessages.appendChild(div);
  scrollToBottom();
}

function showLoading() {
  removeLoading();
  var div = document.createElement("div");
  div.className = "ai-message ai-message-assistant";
  div.id = "loadingIndicator";
  div.innerHTML = '<div class="ai-message-content ai-loading"><span class="ai-loading-dot"></span><span class="ai-loading-dot"></span><span class="ai-loading-dot"></span></div>';
  chatMessages.appendChild(div);
  scrollToBottom();
}

function hideLoading() {
  removeLoading();
}

function removeLoading() {
  var el = document.getElementById("loadingIndicator");
  if (el) el.remove();
}

function scrollToBottom() {
  chatMessages.scrollTop = chatMessages.scrollHeight;
}

function setStatus(message, type) {
  statusBar.className = "ai-status-bar ai-status-" + type;
  // 保留 selectionInfo 不变，只更新状态文字
  var currentSel = selectionInfo.textContent;
  statusBar.innerHTML = '<span id="selectionInfo">' + currentSel + '</span> | ' + escapeHtml(message);
  selectionInfo = document.getElementById("selectionInfo");
}

// ---- 设置 ----
function openSettings() {
  apiKeyInput.value = apiKey;
  modelSelect.value = settings.model;
  temperatureInput.value = settings.temperature;
  settingsModal.style.display = "flex";
}

function closeSettings() {
  settingsModal.style.display = "none";
}

function saveSettings() {
  apiKey = apiKeyInput.value.trim();
  settings.model = modelSelect.value;
  settings.temperature = parseFloat(temperatureInput.value) || 0.3;

  localStorage.setItem("deepseek_api_key", apiKey);
  localStorage.setItem("deepseek_model", settings.model);
  localStorage.setItem("deepseek_temperature", String(settings.temperature));

  sendBtn.disabled = !userInput.value.trim() || !apiKey || isProcessing;
  closeSettings();
}

// ---- 工具函数 ----
function escapeHtml(str) {
  if (!str) return "";
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function renderMarkdown(escapedText) {
  // 简单渲染：**粗体**、`行内代码`、换行
  return escapedText
    .replace(/\*\*(.+?)\*\*/g, "<strong>$1</strong>")
    .replace(/`([^`]+)`/g, "<code>$1</code>")
    .replace(/\n/g, "<br>");
}
```

- [ ] 编写 ai-assistant-taskpane.js 完整代码
- [ ] 运行 `npm run build:dev` 确认构建成功

### Sub-task 6.1: 提交

- [ ] `git add -A && git commit -m "feat: add AI assistant taskpane UI (HTML, CSS, JS)"`

---

## Task 7: 集成验证

**依赖：** Task 1-6 全部完成

### Sub-task 7.1: 完整构建

- [ ] 运行 `npm run build`（生产模式构建）
- [ ] 确认构建无错误

### Sub-task 7.2: 运行全部测试

- [ ] 运行 `npx jest --verbose`
- [ ] 确认全部测试 PASS（包括已有的测试和新增的 ai-utils 测试）

### Sub-task 7.3: 代码质量检查

- [ ] 运行 `npm run lint` 确认无 lint 问题

### Sub-task 7.4: 最终提交

- [ ] `git add -A && git commit -m "feat: integrate AI assistant feature - full build and tests passing"`

---

## 自审清单

- [ ] 所有 spec 需求都有对应 task 覆盖
- [ ] 无 "TBD"、"TODO" 等占位符
- [ ] 所有文件路径和函数签名在各 task 间一致
- [ ] 每步都是可执行的 2-5 分钟操作
- [ ] Task 依赖顺序合理（无循环依赖）
