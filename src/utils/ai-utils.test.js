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

var aiUtils = require("./ai-utils");

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

    it("deepseek has correct apiUrl", function () {
      expect(aiUtils.PROVIDERS.deepseek.apiUrl).toBe("https://api.deepseek.com/chat/completions");
    });

    it("minimax has correct apiUrl", function () {
      expect(aiUtils.PROVIDERS.minimax.apiUrl).toBe("https://api.minimaxi.chat/v1/text/chatcompletion_v2");
    });
  });

  describe("getProvider", function () {
    it("returns provider config for deepseek", function () {
      var p = aiUtils.getProvider("deepseek");
      expect(p.name).toBe("DeepSeek");
      expect(p.id).toBe("deepseek");
    });

    it("returns provider config for minimax", function () {
      var p = aiUtils.getProvider("minimax");
      expect(p.name).toBe("MiniMax");
      expect(p.id).toBe("minimax");
    });

    it("throws for invalid providerId", function () {
      expect(function () {
        aiUtils.getProvider("unknown");
      }).toThrow();
    });

    it("error message includes providerId", function () {
      try {
        aiUtils.getProvider("unknown");
        fail("should have thrown");
      } catch (e) {
        expect(e.message).toContain("unknown");
      }
    });
  });

  describe("getAllProviders", function () {
    it("returns array", function () {
      expect(Array.isArray(aiUtils.getAllProviders())).toBe(true);
    });

    it("returns at least 2 providers", function () {
      expect(aiUtils.getAllProviders().length).toBeGreaterThanOrEqual(2);
    });

    it("each returned provider has id and name", function () {
      aiUtils.getAllProviders().forEach(function (p) {
        expect(p.id).toBeDefined();
        expect(p.name).toBeDefined();
      });
    });
  });

  describe("getModelsForProvider", function () {
    it("returns models array for deepseek", function () {
      var models = aiUtils.getModelsForProvider("deepseek");
      expect(Array.isArray(models)).toBe(true);
      expect(models.length).toBeGreaterThan(0);
    });

    it("deepseek models have id and name", function () {
      var models = aiUtils.getModelsForProvider("deepseek");
      models.forEach(function (m) {
        expect(m.id).toBeDefined();
        expect(m.name).toBeDefined();
      });
    });

    it("returns models array for minimax", function () {
      var models = aiUtils.getModelsForProvider("minimax");
      expect(Array.isArray(models)).toBe(true);
      expect(models.length).toBeGreaterThan(0);
    });

    it("throws for unknown provider", function () {
      // After fix in ai-utils.js, getModelsForProvider now throws like getProvider
      expect(function () {
        aiUtils.getModelsForProvider("unknown");
      }).toThrow();
    });
  });
});

// =============================================================================
// Module 1: DeepSeekClient
// =============================================================================

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
    it("is a function", function () {
      expect(typeof sendChatRequest).toBe("function");
    });
  });
});

// =============================================================================
// Module 2: ToolRegistry
// =============================================================================

describe("ToolRegistry", function () {
  describe("getToolDefinitions", function () {
    it("returns 4 tools", function () {
      var tools = getToolDefinitions();
      expect(tools).toHaveLength(4);
    });

    it("includes generate_formula", function () {
      var names = getToolDefinitions().map(function (t) {
        return t.function.name;
      });
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

// =============================================================================
// Module 3: ContextManager
// =============================================================================

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
        stats: { 年龄: { max: 65, min: 18, avg: 35 } },
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
      var allContent = trimmed
        .map(function (m) {
          return m.content;
        })
        .join(" ");
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

// =============================================================================
// Module 4: ExcelDataService (stub - requires Excel.run context)
// =============================================================================

describe("ExcelDataService", function () {
  describe("function existence", function () {
    it("getSelectionSummary is a function", function () {
      expect(typeof getSelectionSummary).toBe("function");
    });

    it("writeFormula is a function", function () {
      expect(typeof writeFormula).toBe("function");
    });

    it("writeValues is a function", function () {
      expect(typeof writeValues).toBe("function");
    });

    it("applyClassification is a function", function () {
      expect(typeof applyClassification).toBe("function");
    });
  });
});
