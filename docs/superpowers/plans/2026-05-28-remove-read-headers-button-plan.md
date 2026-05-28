# 移除读取表头按钮实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 移除"读取表头"按钮，改为用户在表头行号输入框中输入值时自动触发读取，同时修复预览表格的 DOM 结构缺失问题。

**Architecture:** 保持现有文件结构不变，只修改 vlookup-taskpane.html 和 vlookup-taskpane.js 两处：HTML 修复表格结构 + 移除按钮，JS 移除事件绑定 + 添加 input 监听 + 确保初始化时自动加载。

**Tech Stack:** Vanilla JS / Office.js / HTML

---

## 文件映射

| 文件 | 变更 |
|------|------|
| `src/taskpane/vlookup-taskpane.html` | 移除按钮，修复 `<table>` 结构添加 thead/tbody |
| `src/taskpane/vlookup-taskpane.js` | 移除 readHeaders 事件绑定，添加 headerRow input 监听，初始化时自动调用 |

---

## Task 1: 修复 HTML 表格结构

**Files:**
- Modify: `src/taskpane/vlookup-taskpane.html:119-123`

当前代码:
```html
<div id="vlookupPreview" class="vlookup-preview">
    <table id="previewTable">
        <!-- Preview content will be populated by JavaScript -->
    </table>
</div>
```

修改为:
```html
<div id="vlookupPreview" class="vlookup-preview">
    <table id="previewTable">
        <thead id="previewHead"></thead>
        <tbody id="previewBody"></tbody>
    </table>
</div>
```

同时移除第 62-63 行的读取表头按钮:
```html
<input type="number" id="headerRow" ... />
<button type="button" id="readHeaders" class="vlookup-btn">读取表头</button>
```
改为:
```html
<input type="number" id="headerRow" class="vlookup-input vlookup-input--number" value="1" min="1" />
```

---

## Task 2: 修改 JS 事件绑定

**Files:**
- Modify: `src/taskpane/vlookup-taskpane.js:36` (移除 readHeaders onclick)
- Modify: `src/taskpane/vlookup-taskpane.js:54-57` (添加 headerRow input 监听)
- Modify: `src/taskpane/vlookup-taskpane.js:59-72` (loadInitialSelection 结束时自动调用 readHeaders)

### Step 1: 移除 readHeaders onclick 绑定

第 36 行:
```javascript
document.getElementById("readHeaders").onclick = readHeaders;
```
删除此行。

### Step 2: 在 headerRow 输入框添加 input 事件监听

在 `initEventListeners()` 函数的 `form.addEventListener("input", validateForm);` 之后添加:

```javascript
document.getElementById("headerRow").addEventListener("input", function () {
  var tableInput = document.getElementById("lookupTable").value;
  if (tableInput) {
    loadTableHeaders();
  }
});
```

### Step 3: 将 readHeaders 重命名为 loadTableHeaders

将函数名 `readHeaders` 改为 `loadTableHeaders`，对应调用处也要更新:
- 第 36 行移除的事件绑定（已删除）
- Task 2 Step 2 中的调用
- `refreshColumns()` 之后不再被自动调用，需要确保初始化时调用

### Step 4: loadInitialSelection 结束时自动调用 loadTableHeaders

在 `loadInitialSelection()` 的 `.catch()` 之后，添加自动调用:

```javascript
}).catch(function (error) {
  setStatus("加载选择失败: " + error.message, "error");
});

// Auto-load headers if lookupTable has value
if (document.getElementById("lookupTable").value) {
  loadTableHeaders();
}
```

---

## Task 3: 验证

1. 打开 Excel 侧边栏，增强查找功能
2. 在"表头行号"输入框输入 `1`，验证匹配列下拉选项自动更新为实际列名
3. 验证预览表格正确渲染（thead 和 tbody 有内容）
4. 刷新选择后，验证表头自动重新加载

---

## 预期结果

- 移除"读取表头"按钮
- 用户修改表头行号 → 下拉选项自动更新
- 初始化时（如果有查找表区域）自动加载表头
- 预览表格正常显示，无 null 错误