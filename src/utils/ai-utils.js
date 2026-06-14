/**
 * ai-utils.js
 *
 * AI 助手核心工具模块，包含 4 个子模块：
 *   1. DeepSeekClient - DeepSeek Chat API 客户端
 *   2. ToolRegistry - Function Calling 工具注册表
 *   3. ContextManager - 对话上下文管理器
 *   4. ExcelDataService - Excel 数据服务（通过 Office JS API）
 */

// =============================================================================
// Module 1: DeepSeekClient
// =============================================================================

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
  }).then(function (response) {
    if (!response.ok) {
      return response.json().then(function (err) {
        throw new Error("DeepSeek API 错误: " + (err.error?.message || response.statusText));
      });
    }
    return response.json();
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

// =============================================================================
// Module 2: ToolRegistry
// =============================================================================

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
            description: 'Excel 公式字符串，如 =IF(A2>100,"高","低")',
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
            description: '处理规则说明，如"判断每行的评论是正面、负面还是中性"',
          },
          outputColumn: {
            type: "string",
            description: "结果写入的目标列字母，如 C",
          },
          labelMapping: {
            type: "object",
            description: '可选，分类标签到数值的映射，如 {"正面": 1, "负面": -1}',
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

// =============================================================================
// Module 3: ContextManager
// =============================================================================

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
        if (Object.prototype.hasOwnProperty.call(selectionSummary.stats, col)) {
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
  return [{ role: "system", content: buildSystemPrompt(selectionSummary) }];
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
  var systemMsg = { role: "system", content: buildSystemPrompt(selectionSummary) };
  var others = [];
  for (var i = 0; i < context.length; i++) {
    if (context[i].role !== "system") {
      others.push(context[i]);
    }
  }
  return [systemMsg].concat(others);
}

// =============================================================================
// Module 4: ExcelDataService
// =============================================================================

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
        var max = nums[0],
          min = nums[0],
          sum = 0;
        for (var k = 0; k < nums.length; k++) {
          if (nums[k] > max) {
            max = nums[k];
          }
          if (nums[k] < min) {
            min = nums[k];
          }
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
    // 插入新列（提取列字母前缀，如 "C2:C100" → "C"）
    var colMatch = rangeAddress.match(/^[A-Za-z]+/);
    if (!colMatch) {
      return Promise.reject(new Error("无效的范围地址: " + rangeAddress));
    }
    var colLetter = colMatch[0];
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

// =============================================================================
// Module Exports
// =============================================================================

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
