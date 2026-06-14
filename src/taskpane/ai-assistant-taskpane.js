/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Excel, Office, localStorage */

var aiUtils = require("../utils/ai-utils");

// ---- 状态 ----
var apiKey = localStorage.getItem("deepseek_api_key") || "";
var settings = {
  model: localStorage.getItem("deepseek_model") || "deepseek-v4-flash",
  temperature: parseFloat(localStorage.getItem("deepseek_temperature") || "0.3"),
  maxTokens: 4096,
};
var context = aiUtils.createContext(null);
var isProcessing = false;
var selectionTimer = null;

// ---- DOM 引用 ----
var chatMessages, userInput, sendBtn, statusBar, selectionInfo;
var settingsModal, apiKeyInput, modelSelect, temperatureInput;
var confirmBox, confirmMsg, confirmYes, confirmNo;

// ---- 初始化 ----
Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    cacheDom();
    bindEvents();
    restoreApiKeyState();
    refreshSelection(false);

    // 自动检测选区变化（防抖 500ms）
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, function () {
      if (selectionTimer) clearTimeout(selectionTimer);
      selectionTimer = setTimeout(function () {
        refreshSelection(true);
      }, 500);
    });
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
  // 关闭弹窗（点击背景）
  settingsModal.addEventListener("click", function (e) {
    if (e.target === settingsModal) closeSettings();
  });
  // 手动刷新选区
  document.getElementById("refreshSelectionBtn").addEventListener("click", function () {
    refreshSelection(false);
  });
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
function refreshSelection(silent) {
  setStatus("正在获取选区...", "loading");
  Excel.run(function (ctx) {
    return aiUtils.getSelectionSummary(ctx).then(function (summary) {
      context = aiUtils.updateSelection(context, summary);
      if (summary.rowCount === 0) {
        selectionInfo.textContent = "选区：" + summary.address + "（无数据）";
        setStatus("所选区域无数据，请选择有内容的单元格区域", "warning");
        if (!silent) {
          addAssistantMessage(
            "所选区域 " + summary.address + " 为空，请选择包含数据的区域后再提问。"
          );
        }
      } else {
        selectionInfo.textContent =
          "已选中: " +
          summary.address +
          " (" +
          summary.rowCount +
          "行 × " +
          summary.columnCount +
          "列)";
        setStatus("已就绪", "idle");
        if (!silent) {
          addAssistantMessage(
            "已获取选区 " +
              summary.address +
              "（" +
              summary.rowCount +
              "行，" +
              summary.columnCount +
              "列），请问你想做什么？"
          );
        }
      }
    });
  }).catch(function (err) {
    selectionInfo.textContent = "选区获取失败";
    setStatus("获取选区异常: " + err.message, "error");
  });
}

// ---- 发送消息 ----
function handleSend() {
  if (isProcessing || !apiKey) {
    if (!apiKey) {
      addAssistantMessage("请先在设置中配置 DeepSeek API Key ⚙️");
    }
    return;
  }

  var text = userInput.value.trim();
  if (!text) return;

  userInput.value = "";
  sendBtn.disabled = true;
  isProcessing = true;

  addUserMessage(text);
  context = aiUtils.addMessage(context, { role: "user", content: text });
  toolRoundCount = 0; // 重置工具调用轮次计数
  showLoading();

  aiUtils
    .sendChatRequest(
      apiKey,
      aiUtils.ensureValidContext(aiUtils.trimContext(context)),
      aiUtils.getToolDefinitions(),
      {
        temperature: settings.temperature,
        maxTokens: settings.maxTokens,
      }
    )
    .then(function (response) {
      return aiUtils.parseResponse(response);
    })
    .then(function (parsed) {
      hideLoading();
      context = aiUtils.addMessage(context, parsed.message);

      if (parsed.toolCalls && parsed.toolCalls.length > 0) {
        return handleToolCalls(parsed.toolCalls);
      }

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
// 递归处理 AI 返回的 tool_calls，支持多轮工具调用（如先分析再写入）
var MAX_TOOL_ROUNDS = 5;
var toolRoundCount = 0; // 跨递归调用的总轮次计数

function handleToolCalls(toolCalls) {
  var sequence = Promise.resolve();

  toolCalls.forEach(function (toolCall) {
    sequence = sequence
      .then(function () {
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
      })
      .then(function (result) {
        context = aiUtils.addMessage(context, {
          role: "tool",
          tool_call_id: toolCall.id,
          content: JSON.stringify(result),
        });
      })
      .catch(function (err) {
        // 即使执行失败也要添加 tool 响应，否则上下文出现孤立 tool_calls
        // 后续调用 DeepSeek API 会报错:
        // "assistant message with tool_calls must be followed by tool messages"
        context = aiUtils.addMessage(context, {
          role: "tool",
          tool_call_id: toolCall.id,
          content: JSON.stringify({ error: err.message }),
        });
      });
  });

  // 递归跟进：AI 可能在后续响应中继续调用工具
  function followUp() {
    toolRoundCount++;
    if (toolRoundCount > MAX_TOOL_ROUNDS) {
      setStatus("达到最大工具调用轮次", "warning");
      addAssistantMessage(
        "⚠️ 已达到最大工具调用轮次（" + MAX_TOOL_ROUNDS + "），操作可能未完全完成。请继续输入指令。"
      );
      return Promise.resolve();
    }

    return sequence
      .then(function () {
        showLoading();
        return aiUtils.sendChatRequest(
          apiKey,
          aiUtils.ensureValidContext(aiUtils.trimContext(context)),
          aiUtils.getToolDefinitions(),
          { temperature: settings.temperature, maxTokens: settings.maxTokens }
        );
      })
      .then(function (response) {
        hideLoading();
        return aiUtils.parseResponse(response);
      })
      .then(function (parsed) {
        context = aiUtils.addMessage(context, parsed.message);

        if (parsed.toolCalls && parsed.toolCalls.length > 0) {
          // 先显示本轮文本内容（保持时间顺序），再递归处理新的 tool_calls
          if (parsed.message.content) {
            addAssistantMessage(parsed.message.content);
          }
          return handleToolCalls(parsed.toolCalls);
        }

        // 没有更多 tool_calls，显示文本并完成
        addAssistantMessage(parsed.message.content || "");
        setStatus("完成", "success");
      });
  }

  return followUp();
}

// ---- Tool 处理函数 ----
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
      })
        .then(function () {
          resolve({ status: "success", message: "公式已写入 " + args.targetCell });
        })
        .catch(function (err) {
          reject(err);
        });
    });
  });
}

function handleAnalyzeData(args) {
  addAssistantMessage("📊 **分析结果**\n\n" + args.conclusion);
  if (args.insights && args.insights.length) {
    var insightsText = "**关键洞察：**\n";
    args.insights.forEach(function (item, i) {
      insightsText += i + 1 + ". " + item + "\n";
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
      })
        .then(function () {
          resolve({
            status: "success",
            message: args.task + "完成，结果已写入 " + args.outputColumn + " 列",
          });
        })
        .catch(function (err) {
          reject(err);
        });
    });
  });
}

function createClassifyFunction(instruction, labelMapping) {
  return function (cellValue) {
    if (cellValue === null || cellValue === undefined || cellValue === "") {
      return "";
    }
    var text = String(cellValue).toLowerCase();

    if (instruction.indexOf("情感") >= 0 || instruction.indexOf("正面") >= 0) {
      var positive = [
        "好",
        "棒",
        "优秀",
        "满意",
        "推荐",
        "喜欢",
        "赞",
        "good",
        "great",
        "excellent",
        "positive",
      ];
      var negative = [
        "差",
        "烂",
        "失望",
        "糟糕",
        "不满",
        "投诉",
        "bad",
        "poor",
        "terrible",
        "negative",
      ];
      var posScore = 0;
      var negScore = 0;
      positive.forEach(function (w) {
        if (text.indexOf(w) >= 0) posScore++;
      });
      negative.forEach(function (w) {
        if (text.indexOf(w) >= 0) negScore++;
      });
      var label = posScore > negScore ? "正面" : negScore > posScore ? "负面" : "中性";
      return labelMapping && labelMapping[label] !== undefined ? labelMapping[label] : label;
    }

    if (instruction.indexOf("提取") >= 0) {
      if (instruction.indexOf("邮箱") >= 0 || instruction.indexOf("邮件") >= 0) {
        var emailMatch = text.match(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/);
        return emailMatch ? emailMatch[0] : "";
      }
      if (instruction.indexOf("手机") >= 0 || instruction.indexOf("电话") >= 0) {
        var phoneMatch = text.match(/1[3-9]\d{9}/);
        return phoneMatch ? phoneMatch[0] : "";
      }
    }

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
      })
        .then(function () {
          resolve({ status: "success", message: "数据已写入 " + args.range });
        })
        .catch(function (err) {
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
  div.innerHTML = '<div class="ai-message-content">' + escapeHtml(text) + "</div>";
  chatMessages.appendChild(div);
  scrollToBottom();
}

function addAssistantMessage(text) {
  removeLoading();
  var div = document.createElement("div");
  div.className = "ai-message ai-message-assistant";
  div.innerHTML = '<div class="ai-message-content">' + renderMarkdown(escapeHtml(text)) + "</div>";
  chatMessages.appendChild(div);
  scrollToBottom();
}

function showLoading() {
  removeLoading();
  var div = document.createElement("div");
  div.className = "ai-message ai-message-assistant";
  div.id = "loadingIndicator";
  div.innerHTML =
    '<div class="ai-message-content ai-loading"><span class="ai-loading-dot"></span><span class="ai-loading-dot"></span><span class="ai-loading-dot"></span></div>';
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
  document.getElementById("statusText").textContent = message ? " | " + message : "";
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
  return escapedText
    .replace(/\*\*(.+?)\*\*/g, "<strong>$1</strong>")
    .replace(/`([^`]+)`/g, "<code>$1</code>")
    .replace(/\n/g, "<br>");
}
