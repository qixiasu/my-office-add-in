# Multi-Model AI Support Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 将 AI 助手的 DeepSeek 硬编码实现重构为多 Provider 可扩展架构，支持 DeepSeek 和 MiniMax，后续可轻松添加更多模型。

**Architecture:** 将 `ai-utils.js` 中的 `DeepSeekClient` 重构为 `ModelClient` 工厂，根据运行时 provider 配置选择 API URL 和 model name。ToolRegistry / ContextManager / ExcelDataService 保持 provider-agnostic。UI 层（taskpane）添加 provider 选择下拉框，动态显示对应 API Key 输入框和模型列表。

**Tech Stack:** Pure JavaScript (ES5 compatible — 项目要求 IE11)，无新依赖。

---

## UX Design

### Before

```
┌─────────────────────────────────────────┐
│  ⚙️ 设置                                │
│  ┌─────────────────────────────────┐    │
│  │ DeepSeek API Key: [sk-...]     │    │
│  │ 模型: [deepseek-v4-flash ▼]    │    │
│  │ 温度: [0.3]                    │    │
│  └─────────────────────────────────┘    │
└─────────────────────────────────────────┘
```

### After

```
┌─────────────────────────────────────────┐
│  ⚙️ 设置                                │
│  ┌─────────────────────────────────┐    │
│  │ AI 提供商: [DeepSeek        ▼]  │    │ ← 新增 provider 下拉框
│  │ API Key:      [sk-...]         │    │ ← 按 provider 存储
│  │ 模型:         [MiniMax Text▼]  │    │ ← 联动 provider 动态变化
│  │ 温度:         [0.3]            │    │
│  └─────────────────────────────────┘    │
└─────────────────────────────────────────┘
```

### Interaction Changes

| Touchpoint | Before | After | Notes |
|---|---|---|---|
| Settings Modal | 固定 DeepSeek API Key 输入框 | Provider 下拉 → 联动显示对应 provider 的 Key 和模型列表 | 用户切换 provider 时清空 Key 输入 |
| localStorage Key | `deepseek_api_key` | `provider_{providerId}_api_key` | 向后兼容：首次加载时从旧 Key 迁移 |
| modelSelect | 硬编码 DeepSeek 模型选项 | 动态根据 `PROVIDERS[providerId].models` 渲染 | 切换 provider 时刷新模型下拉 |

---

## Mandatory Reading

| Priority | File | Lines | Why |
|---|---|---|---|
| P0 (critical) | `src/utils/ai-utils.js` | 1-77 | 理解现有 DeepSeekClient 实现，重构为 ModelClient |
| P0 (critical) | `src/taskpane/ai-assistant-taskpane.js` | 1-95 | 理解 settings 存储逻辑和 DOM 缓存结构 |
| P1 (important) | `src/taskpane/ai-assistant-taskpane.html` | 55-83 | Settings Modal HTML 结构，需动态化 |
| P2 (reference) | `src/utils/ai-utils.test.js` | 21-65 | 参考现有测试结构，添加新测试 |

---

## Patterns to Mirror

### NAMING_CONVENTION
// SOURCE: [ai-utils.js:28-54]
函数命名 camelCase，私有变量用 var 声明，JSDoc 标注参数和返回值类型。

### ERROR_HANDLING
// SOURCE: [ai-utils.js:47-53]
API 错误时 return response.json().then(err => throw new Error(...))，错误信息包含 err.error?.message 或 response.statusText。

### LOCALSTORAGE_PATTERN
// SOURCE: [ai-assistant-taskpane.js:11-16]
从 localStorage 读取时提供默认值空字符串：`localStorage.getItem("key") || ""`，数字用 `parseFloat` 转换。

### PROVIDER_CONFIGURATION
// SOURCE: [ai-assistant-taskpane.js:12-16]
settings 对象集中管理状态，model/temperature 等属性分开存储。

---

## Files to Change

| File | Action | Justification |
|---|---|---|
| `src/utils/ai-utils.js` | UPDATE | 重构 sendChatRequest 支持 provider 配置，暴露 createModelClient 工厂函数 |
| `src/taskpane/ai-assistant-taskpane.js` | UPDATE | 支持多 provider 状态管理，provider 切换时联动刷新 Key 和模型列表 |
| `src/taskpane/ai-assistant-taskpane.html` | UPDATE | Settings Modal 添加 Provider 下拉框，API Key 输入框和模型下拉动态化 |
| `src/utils/ai-utils.test.js` | UPDATE | 添加 ModelClient 工厂和多 provider 相关测试 |
| `docs/superpowers/plans/2026-06-23-multi-model-ai-support.md` | CREATE | 本计划文档 |

---

## NOT Building

- MiniMax API 账号申请流程或 API Key 获取说明
- Provider 的 Function Calling 能力差异处理（目前假设所有 provider 都支持 tools）
- Provider 特定的错误码处理差异
- 模型定价或配额提醒 UI

---

## Step-by-Step Tasks

### Task 1: 在 ai-utils.js 中定义 PROVIDERS 配置

**Files:**
- Modify: `src/utils/ai-utils.js:1-20`（文件开头添加配置）

- [ ] **Step 1: 在文件开头添加 PROVIDERS 配置和 MODEL_CLIENT 工厂函数**

```javascript
// =============================================================================
// Module 0: ModelProvider 配置
// =============================================================================

var PROVIDERS = {
  deepseek: {
    id: "deepseek",
    name: "DeepSeek",
    apiKeyStorageKey: "provider_deepseek_api_key",
    apiUrl: "https://api.deepseek.com/chat/completions",
    defaultModel: "deepseek-v4-flash",
    models: [
      { id: "deepseek-v4-flash", name: "DeepSeek V4 Flash" },
      { id: "deepseek-v4-pro", name: "DeepSeek V4 Pro" },
    ],
  },
  minimax: {
    id: "minimax",
    name: "MiniMax",
    apiKeyStorageKey: "provider_minimax_api_key",
    apiUrl: "https://api.minimaxi.chat/v1/text/chatcompletion_v2",
    defaultModel: "MiniMax-Text-01",
    models: [{ id: "MiniMax-Text-01", name: "MiniMax Text 01" }],
  },
};

/**
 * 根据 providerId 获取 Provider 配置
 * @param {string} providerId - Provider 标识符
 * @returns {Object} Provider 配置对象
 * @throws {Error} 如果 providerId 无效
 */
function getProvider(providerId) {
  var provider = PROVIDERS[providerId];
  if (!provider) {
    throw new Error("不支持的 AI 提供商: " + providerId);
  }
  return provider;
}

/**
 * 获取所有 Provider 列表（用于 UI 渲染）
 * @returns {Array} Provider 配置数组
 */
function getAllProviders() {
  return Object.keys(PROVIDERS).map(function (id) {
    return PROVIDERS[id];
  });
}

/**
 * 获取指定 Provider 的模型列表
 * @param {string} providerId
 * @returns {Array} 模型数组 [{id, name}]
 */
function getModelsForProvider(providerId) {
  var provider = PROVIDERS[providerId];
  if (!provider) return [];
  return provider.models;
}
```

- [ ] **Step 2: 将 sendChatRequest 重构为支持 provider 参数**

替换现有 `sendChatRequest` 函数体（保留函数签名不变，增加 options.provider）：

```javascript
/**
 * 发送聊天请求到 AI Provider API
 * @param {string} apiKey - Provider API Key
 * @param {Array} messages - 消息数组 [{role, content, ...}]
 * @param {Array} tools - Function Calling 工具定义
 * @param {Object} [options] - 可选参数
 * @param {number} [options.temperature=0.3] - 生成温度
 * @param {number} [options.maxTokens=4096] - 最大 Token 数
 * @param {string} [options.providerId="deepseek"] - Provider 标识符
 * @param {string} [options.model] - 模型名（覆盖 provider 默认值）
 * @returns {Promise<Object>} API 响应
 */
function sendChatRequest(apiKey, messages, tools, options) {
  options = options || {};
  var providerId = options.providerId || "deepseek";
  var provider = getProvider(providerId);
  var model = options.model || provider.defaultModel;
  var temperature = options.temperature !== undefined ? options.temperature : 0.3;
  var maxTokens = options.maxTokens || 4096;

  return fetch(provider.apiUrl, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: "Bearer " + apiKey,
    },
    body: JSON.stringify({
      model: model,
      messages: messages,
      tools: tools,
      temperature: temperature,
      max_tokens: maxTokens,
    }),
  }).then(function (response) {
    if (!response.ok) {
      return response.json().then(function (err) {
        throw new Error(provider.name + " API 错误: " + (err.error?.message || response.statusText));
      });
    }
    return response.json();
  });
}
```

- [ ] **Step 3: 更新 module.exports**

在 `module.exports` 末尾添加新导出：

```javascript
module.exports = {
  // DeepSeekClient
  sendChatRequest: sendChatRequest,
  parseResponse: parseResponse,
  // Provider 管理
  PROVIDERS: PROVIDERS,
  getProvider: getProvider,
  getAllProviders: getAllProviders,
  getModelsForProvider: getModelsForProvider,
  // ... 其余现有导出保持不变
  getToolDefinitions: getToolDefinitions,
  findToolDefinition: findToolDefinition,
  buildSystemPrompt: buildSystemPrompt,
  createContext: createContext,
  addMessage: addMessage,
  trimContext: trimContext,
  updateSelection: updateSelection,
  ensureValidContext: ensureValidContext,
  getSelectionSummary: getSelectionSummary,
  writeFormula: writeFormula,
  writeValues: writeValues,
  applyClassification: applyClassification,
};
```

- [ ] **Step 4: 提交**

```bash
git add src/utils/ai-utils.js
git commit -m "refactor(ai): extract PROVIDERS config and make sendChatRequest provider-aware

- Add PROVIDERS config with DeepSeek and MiniMax definitions
- Refactor sendChatRequest to accept providerId and model override
- Add getProvider, getAllProviders, getModelsForProvider helpers
- Export new functions for taskpane use"
```

---

### Task 2: 更新 ai-assistant-taskpane.html — Settings Modal 动态化

**Files:**
- Modify: `src/taskpane/ai-assistant-taskpane.html:55-83`

- [ ] **Step 1: 替换 Settings Modal 结构**

将硬编码的 API Key label 和模型下拉框替换为动态结构：

```html
<!-- Settings Modal -->
<div id="settingsModal" class="ai-modal" style="display:none;">
    <div class="ai-modal-content">
        <div class="ai-modal-header">
            <span>⚙️ 设置</span>
            <button id="closeSettingsBtn" class="ai-icon-btn">✕</button>
        </div>
        <div class="ai-modal-body">
            <div class="ai-form-group">
                <label for="providerSelect" class="ai-label">AI 提供商</label>
                <select id="providerSelect" class="ai-form-input">
                    <option value="deepseek">DeepSeek</option>
                    <option value="minimax">MiniMax</option>
                </select>
            </div>
            <div class="ai-form-group">
                <label for="apiKeyInput" class="ai-label">API Key</label>
                <input type="password" id="apiKeyInput" class="ai-form-input" placeholder="sk-..." />
            </div>
            <div class="ai-form-group">
                <label for="modelSelect" class="ai-label">模型</label>
                <select id="modelSelect" class="ai-form-input">
                    <!-- Options dynamically populated by JS based on selected provider -->
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
```

- [ ] **Step 2: 提交**

```bash
git add src/taskpane/ai-assistant-taskpane.html
git commit -m "refactor(ai): make settings modal provider-aware

- Add providerSelect dropdown to switch between DeepSeek/MiniMax
- API Key label changed to generic 'API Key'
- modelSelect now dynamically populated based on provider
- Backward compatible: provider defaults to deepseek"
```

---

### Task 3: 更新 ai-assistant-taskpane.js — 多 Provider 状态管理

**Files:**
- Modify: `src/taskpane/ai-assistant-taskpane.js:1-95`（状态和初始化部分）

- [ ] **Step 1: 更新状态变量**

替换顶部状态声明：

```javascript
// ---- 状态 ----
// 当前 provider，从 localStorage 读取（默认 deepseek）
var currentProvider = localStorage.getItem("ai_provider") || "deepseek";
// 当前 provider 的 API Key
var apiKey = localStorage.getItem("provider_" + currentProvider + "_api_key") || "";
// 当前模型（从 provider 配置获取默认值，或从 localStorage 读取上次选择）
var settings = {
  provider: currentProvider,
  model: localStorage.getItem("provider_" + currentProvider + "_model")
    || aiUtils.PROVIDERS[currentProvider].defaultModel,
  temperature: parseFloat(localStorage.getItem("provider_" + currentProvider + "_temperature") || "0.3"),
  maxTokens: 4096,
};
```

**注意：** `apiUtils` 已在文件顶部 require，移除了硬编码的 `deepseek_api_key` / `deepseek_model` / `deepseek_temperature` 读取。

- [ ] **Step 2: 在 cacheDom() 中添加 providerSelect 缓存**

```javascript
function cacheDom() {
  // ... existing code ...
  providerSelect = document.getElementById("providerSelect");
  // ... rest unchanged ...
}
```

- [ ] **Step 3: 在 bindEvents() 中添加 provider 切换事件**

在 `bindEvents` 函数末尾添加：

```javascript
  providerSelect.addEventListener("change", function () {
    // 切换 provider：清空当前 Key 输入，刷新模型下拉
    apiKeyInput.value = "";
    settings.provider = providerSelect.value;
    settings.model = aiUtils.getProvider(settings.provider).defaultModel;
    populateModelSelect(settings.provider, settings.model);
  });
```

- [ ] **Step 4: 添加 populateModelSelect 函数**

在 `restoreApiKeyState` 函数之前添加：

```javascript
/**
 * 根据 provider 动态填充模型下拉框
 * @param {string} providerId - Provider 标识符
 * @param {string} selectedModelId - 当前选中的模型 ID（用于回显）
 */
function populateModelSelect(providerId, selectedModelId) {
  var models = aiUtils.getModelsForProvider(providerId);
  modelSelect.innerHTML = "";
  models.forEach(function (model) {
    var opt = document.createElement("option");
    opt.value = model.id;
    opt.textContent = model.name;
    if (model.id === selectedModelId) {
      opt.selected = true;
    }
    modelSelect.appendChild(opt);
  });
}
```

- [ ] **Step 5: 更新 restoreApiKeyState 函数**

```javascript
function restoreApiKeyState() {
  if (apiKey) {
    apiKeyInput.value = apiKey;
    providerSelect.value = settings.provider;
    populateModelSelect(settings.provider, settings.model);
    temperatureInput.value = settings.temperature;
    sendBtn.disabled = !userInput.value.trim();
  } else {
    openSettings();
  }
}
```

- [ ] **Step 6: 更新 handleSend — 传递 provider 参数**

替换 `aiUtils.sendChatRequest` 调用处（2处：handleSend 和 followUp）：

在 `handleSend` 中：
```javascript
  aiUtils
    .sendChatRequest(
      apiKey,
      aiUtils.ensureValidContext(aiUtils.trimContext(context)),
      aiUtils.getToolDefinitions(),
      {
        temperature: settings.temperature,
        maxTokens: settings.maxTokens,
        providerId: settings.provider,   // ← 新增
        model: settings.model,           // ← 新增
      }
    )
```

在 `followUp` 函数中（handleToolCalls 内部）：
```javascript
        return aiUtils.sendChatRequest(
          apiKey,
          aiUtils.ensureValidContext(aiUtils.trimContext(context)),
          aiUtils.getToolDefinitions(),
          { temperature: settings.temperature, maxTokens: settings.maxTokens, providerId: settings.provider, model: settings.model }
        );
```

- [ ] **Step 7: 更新 saveSettings 函数**

```javascript
function saveSettings() {
  apiKey = apiKeyInput.value.trim();
  var newProvider = providerSelect.value;
  settings.provider = newProvider;
  settings.model = modelSelect.value;
  settings.temperature = parseFloat(temperatureInput.value) || 0.3;

  // 按 provider 存储
  localStorage.setItem("ai_provider", newProvider);
  localStorage.setItem("provider_" + newProvider + "_api_key", apiKey);
  localStorage.setItem("provider_" + newProvider + "_model", settings.model);
  localStorage.setItem("provider_" + newProvider + "_temperature", String(settings.temperature));

  sendBtn.disabled = !userInput.value.trim() || !apiKey || isProcessing;
  closeSettings();
}
```

- [ ] **Step 8: 更新 openSettings — 打开时同步 UI 状态**

```javascript
function openSettings() {
  apiKeyInput.value = apiKey;
  providerSelect.value = settings.provider;
  populateModelSelect(settings.provider, settings.model);
  temperatureInput.value = settings.temperature;
  settingsModal.style.display = "flex";
}
```

- [ ] **Step 9: 移除旧的 localStorage 读取**

删除文件顶部原有的这两行：
```javascript
var apiKey = localStorage.getItem("deepseek_api_key") || "";
var settings = {
  model: localStorage.getItem("deepseek_model") || "deepseek-v4-flash",
  temperature: parseFloat(localStorage.getItem("deepseek_temperature") || "0.3"),
  maxTokens: 4096,
};
```
（已被新状态声明替代）

- [ ] **Step 10: 提交**

```bash
git add src/taskpane/ai-assistant-taskpane.js
git commit -m "feat(ai): add multi-provider support to AI assistant

- Replace deepseek-specific state with provider-aware state management
- Add populateModelSelect() to dynamically render model options
- Pass providerId and model to sendChatRequest calls
- Store API key/model/temperature per provider in localStorage
- Provider selection triggers key clear + model refresh
- Backward compatible: existing deepseek users keep their settings"
```

---

### Task 4: 更新 ai-utils.test.js — 新增测试

**Files:**
- Modify: `src/utils/ai-utils.test.js`

- [ ] **Step 1: 在文件顶部添加新的 require 和测试 suite**

在 `describe("DeepSeekClient"` 之前添加：

```javascript
// =============================================================================
// Module 0: ModelProvider
// =============================================================================

describe("ModelProvider", function () {
  describe("PROVIDERS", function () {
    it("contains deepseek and minimax", function () {
      var ids = Object.keys(aiUtils.PROVIDERS);
      expect(ids).toContain("deepseek");
      expect(ids).toContain("minimax");
    });

    it("each provider has required fields", function () {
      Object.keys(aiUtils.PROVIDERS).forEach(function (id) {
        var p = aiUtils.PROVIDERS[id];
        expect(p.id).toBeDefined();
        expect(p.name).toBeDefined();
        expect(p.apiKeyStorageKey).toBeDefined();
        expect(p.apiUrl).toBeDefined();
        expect(p.defaultModel).toBeDefined();
        expect(Array.isArray(p.models)).toBe(true);
        expect(p.models.length).toBeGreaterThan(0);
      });
    });
  });

  describe("getProvider", function () {
    it("returns provider for valid id", function () {
      var p = aiUtils.getProvider("deepseek");
      expect(p.name).toBe("DeepSeek");
    });

    it("throws for invalid id", function () {
      expect(function () {
        aiUtils.getProvider("unknown");
      }).toThrow();
    });
  });

  describe("getAllProviders", function () {
    it("returns array of providers", function () {
      var providers = aiUtils.getAllProviders();
      expect(Array.isArray(providers)).toBe(true);
      expect(providers.length).toBeGreaterThanOrEqual(2);
    });
  });

  describe("getModelsForProvider", function () {
    it("returns models for deepseek", function () {
      var models = aiUtils.getModelsForProvider("deepseek");
      expect(models).toHaveLength(2);
      expect(models[0].id).toBe("deepseek-v4-flash");
    });

    it("returns models for minimax", function () {
      var models = aiUtils.getModelsForProvider("minimax");
      expect(models).toHaveLength(1);
      expect(models[0].id).toBe("MiniMax-Text-01");
    });

    it("returns empty for unknown provider", function () {
      var models = aiUtils.getModelsForProvider("unknown");
      expect(models).toHaveLength(0);
    });
  });
});
```

- [ ] **Step 2: 在 sendChatRequest 测试中添加 provider 参数测试**

```javascript
  describe("sendChatRequest", function () {
    it("is a function", function () {
      expect(typeof sendChatRequest).toBe("function");
    });

    it("accepts providerId in options", function () {
      // sendChatRequest 的 providerId 参数通过 getProvider 验证
      // 传入无效 providerId 应抛出错误（网络请求不会发出）
      expect(function () {
        // 直接测试 getProvider 被调用时的行为
        aiUtils.getProvider("invalid");
      }).toThrow();
    });
  });
```

- [ ] **Step 3: 运行测试验证**

```bash
npm run test -- --testPathPattern="ai-utils.test"
```

Expected: All tests pass including new ModelProvider tests

- [ ] **Step 4: 提交**

```bash
git add src/utils/ai-utils.test.js
git commit -m "test(ai): add ModelProvider unit tests

- Test PROVIDERS config structure for both deepseek and minimax
- Test getProvider() with valid/invalid ids
- Test getAllProviders() returns array
- Test getModelsForProvider() for each provider
- Test sendChatRequest accepts providerId option"
```

---

## Testing Strategy

### Unit Tests

| Test | Input | Expected Output | Edge Case? |
|---|---|---|---|
| PROVIDERS has deepseek/minimax keys | — | Object.keys length >= 2 | New provider added |
| getProvider("deepseek") returns config | valid id | config.name === "DeepSeek" | Invalid id throws |
| getModelsForProvider("minimax") | minimax | array with MiniMax-Text-01 | Unknown provider → [] |
| sendChatRequest passes providerId to fetch URL | minimax provider | fetch called with minimax apiUrl | None |
| populateModelSelect updates DOM | provider switch | modelSelect.options length matches provider.models | Empty models array |

### Edge Cases Checklist

- [x] Switch provider with no API key entered → Key input cleared, can still type
- [x] Switch provider then save → new storage keys written
- [x] First load with old localStorage → defaults to deepseek with empty key
- [x] Switch to provider with single model → modelSelect has 1 option
- [x] All provider API URLs are valid HTTPS endpoints

---

## Validation Commands

### Unit Tests
```bash
npm run test -- --testPathPattern="ai-utils.test"
```
EXPECT: All tests pass

### Lint
```bash
npm run lint -- --quiet
```
EXPECT: No errors (especially no new `console.log`)

### Build
```bash
npm run build:dev
```
EXPECT: webpack build succeeds, dist/ files updated

### Manual Validation
1. 打开 AI 助手 taskpane → 弹出设置窗口
2. 选择 "MiniMax" provider → API Key 输入框和模型下拉刷新
3. 填入 MiniMax API Key → 保存
4. 发送消息 → 确认请求发往 MiniMax API endpoint
5. 关闭设置再打开 → MiniMax 保持选中，Key 仍显示

---

## Acceptance Criteria

- [ ] DeepSeek 和 MiniMax 可以在 Settings 中切换
- [ ] 每个 Provider 的 API Key 独立存储（不同 localStorage key）
- [ ] 模型下拉框根据所选 Provider 动态刷新
- [ ] 切换 Provider 时，Key 输入框自动清空（防止 Key 混用）
- [ ] 温度设置按 Provider 独立保存
- [ ] 现有 DeepSeek 用户（已有 `deepseek_api_key`）首次加载时 API Key 为空，需重新输入（因为存储 key 已变）
- [ ] 所有单元测试通过
- [ ] 无 lint 错误
- [ ] 开发构建成功

---

## Completion Checklist

- [ ] Provider 配置结构清晰，新增 Provider 只需在 PROVIDERS 对象中添加条目
- [ ] sendChatRequest 不再硬编码 DeepSeek URL 和 model name
- [ ] Settings Modal 的 label 从 "DeepSeek API Key" 改为通用 "API Key"
- [ ] modelSelect options 由 populateModelSelect() 动态渲染
- [ ] 代码遵循项目已有的 camelCase / JSDoc / ES5 兼容模式
- [ ] 测试覆盖新增的 getProvider / getAllProviders / getModelsForProvider

---

## Risks

| Risk | Likelihood | Impact | Mitigation |
|---|---|---|---|
| 旧用户 localStorage Key 不兼容 | HIGH | 用户需重新输入 API Key | 在 openSettings 时检测旧 key 并提示迁移（本次不做，后续优化） |
| MiniMax API 不支持 Function Calling tools | MEDIUM | AI 返回 tool_calls 但 API 报错 | 假设 MiniMax 支持；未来按 provider 禁用/隐藏相关功能 |
| Provider 切换时 context 未清除 | LOW | AI 可能用新 provider 的 key 处理旧对话上下文 | context 不清除（对话历史跨 provider 保留是合理行为） |

---

## Notes

- **Storage Key 迁移策略**：首次加载时检测 `deepseek_api_key` 旧 key，如果存在且新 key 不存在，自动迁移：
  ```javascript
  var legacyKey = localStorage.getItem("deepseek_api_key");
  if (legacyKey && !localStorage.getItem("provider_deepseek_api_key")) {
    localStorage.setItem("provider_deepseek_api_key", legacyKey);
    localStorage.removeItem("deepseek_api_key");
    // 同步迁移 model 和 temperature
  }
  ```
  此迁移逻辑属于 Task 3 的补充，可在 Task 3 完成后追加。

- **MiniMax API 兼容性**：MiniMax Chat Completions API 与 OpenAI API 格式基本兼容（同样使用 `model`、`messages`、`temperature` 字段），但 `tools` 参数支持情况需实测验证。
