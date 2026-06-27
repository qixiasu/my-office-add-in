# Markdown 导出界面美化实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 美化 markdown-export 界面，参考 vlookup/concat 风格，添加刷新按钮、生成预览按钮和复制按钮

**Architecture:** 保持现有功能不变，仅重构 UI 布局和样式。HTML 结构重新设计以匹配项目其他 taskpane 的风格，CSS 使用统一的蓝色主题。

**Tech Stack:** 原生 HTML/CSS/JS，Office UI Fabric 图标，无新依赖

---

## 用户故事

As a Excel 用户，我想要一个美观的 Markdown 导出界面，当我选择不同区域时可以刷新预览，并一键生成和复制 Markdown 表格。

## 问题 → 解决方案

当前界面简陋（只有简单文本标题+原始预览区） → 参考 vlookup/concat 的专业 UI 风格（SVG图标标题+引导卡片+双色调按钮）

## 元数据

- **复杂度**: Small
- **源 PRD**: N/A
- **PRD Phase**: N/A
- **预估文件数**: 3（html/css/js）

---

## UX 设计

### Before

```
┌─────────────────────────────┐
│  导出 Markdown 表格          │  (纯文本标题)
├─────────────────────────────┤
│  ⚠️ 警告区域                 │  (简单警告框)
├─────────────────────────────┤
│  预览                    [复制]│  (简陋预览区)
│  | A | B | C |              │
│  |---|---|---|              │
└─────────────────────────────┘
```

### After

```
┌─────────────────────────────┐
│  📋 导出 Markdown 表格       │  (SVG图标标题)
├─────────────────────────────┤
│  📌 选择区域后点击"生成预览"  │  (引导卡片)
├─────────────────────────────┤
│  [刷新选择] [生成 Markdown]  │  (操作按钮行)
├─────────────────────────────┤
│  预览                       │
│  | A | B | C |              │
│  |---|---|---|              │
│  [复制到剪贴板]              │  (独立复制按钮)
└─────────────────────────────┘
```

### 交互变化

| 交互点 | 之前 | 之后 | 备注 |
|---|---|---|---|
| 加载时机 | 页面打开自动加载 | 需点击"生成预览"按钮 | 允许用户先看界面说明 |
| 刷新选区 | 无 | 点击"刷新选择"按钮 | 支持切换选区后刷新 |
| 复制功能 | 在预览标题栏 | 独立按钮在预览下方 | 更明显 |

---

## 强制阅读

| 优先级 | 文件 | 行数 | 为什么 |
|---|---|---|---|
| P0 (关键) | `src/taskpane/vlookup-taskpane.html` | 1-50 | 标题+按钮的 HTML 结构参考 |
| P0 (关键) | `src/taskpane/vlookup-taskpane.css` | 1-100 | 蓝色主题样式参考 |
| P1 (重要) | `src/taskpane/markdown-export-taskpane.js` | 全部 | 现有逻辑保持不变 |
| P2 (参考) | `src/taskpane/concat-taskpane.css` | 1-80 | 另一个 UI 风格参考 |

---

## 需镜像的模式

### VLOOKUP_TITLE_STYLE
// SOURCE: `src/taskpane/vlookup-taskpane.html:26-31`
```html
<h1 class="vlookup-title">
    <svg width="20" height="20" viewBox="0 0 20 20" fill="#0078d4">
        <path d="..."/>
    </svg>
    增强查找
</h1>
```

### GUIDE_CARD_STYLE
// SOURCE: `src/taskpane/vlookup-taskpane.css:21-34`
```css
.guide-card {
  background: #f0f7ff;
  border-left: 3px solid #0078d4;
  padding: 12px;
  border-radius: 4px;
  margin-bottom: 16px;
}
```

### BUTTON_STYLE
// SOURCE: `src/taskpane/vlookup-taskpane.css:121-136`
```css
.vlookup-button {
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

.vlookup-button:hover {
  background: #0078d4;
  color: white;
}
```

### PRIMARY_BUTTON_STYLE
// SOURCE: `src/taskpane/vlookup-taskpane.css:150-158`
```css
.vlookup-button--primary {
  background: #0078d4;
  color: white;
}

.vlookup-button--primary:hover {
  background: #005a9e;
  border-color: #005a9e;
}
```

---

## 文件变更

| 文件 | 操作 | 理由 |
|---|---|---|
| `src/taskpane/markdown-export-taskpane.html` | 修改 | 重新设计 HTML 结构，添加按钮和引导卡片 |
| `src/taskpane/markdown-export-taskpane.css` | 修改 | 应用蓝色主题样式，删除冗余样式 |
| `src/taskpane/markdown-export-taskpane.js` | 修改 | 添加刷新和生成按钮的事件处理 |

## 不构建

- 不修改 Markdown 生成逻辑（保持现有 generateMarkdownTable 功能）
- 不添加单元测试（纯 UI 重构）
- 不引入新依赖

---

## 步骤任务

### Task 1: 重构 HTML 结构

**文件:**
- 修改: `src/taskpane/markdown-export-taskpane.html`

- [ ] **Step 1: 重写 HTML 结构**

```html
<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <meta http-equiv="Content-Security-Policy" content="default-src 'self' 'unsafe-inline' https://localhost:* https://appsforoffice.microsoft.com;">
  <title>导出 Markdown</title>

  <!-- Office JavaScript API -->
  <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>

  <!-- Template styles -->
  <link href="taskpane.css" rel="stylesheet" type="text/css" />
  <link href="markdown-export-taskpane.css" rel="stylesheet" type="text/css" />
</head>
<body class="ms-font-m ms-welcome ms-Fabric">
  <div class="markdown-export-container">
    <!-- Title with SVG icon -->
    <h1 class="md-title">
      <svg width="20" height="20" viewBox="0 0 20 20" fill="#0078d4">
        <path d="M3 3h14v2H3V3zm0 4h14v2H3V7zm0 4h8v2H3v-2zm10 4h4v2h-4v-2z"/>
      </svg>
      导出 Markdown 表格
    </h1>

    <!-- Guide card -->
    <div class="guide-card">
      <p>选择 Excel 中的数据区域，点击"生成预览"查看 Markdown 格式效果，然后复制到剪贴板</p>
    </div>

    <!-- Warning box (hidden by default) -->
    <div id="warningBox" class="warning-box" style="display: none;">
      <span class="warning-icon">⚠️</span>
      <span id="warningText"></span>
    </div>

    <!-- Action buttons row -->
    <div class="md-btn-row">
      <button type="button" id="refreshBtn" class="md-button">
        刷新选择
      </button>
      <button type="button" id="generateBtn" class="md-button md-button--primary">
        生成 Markdown
      </button>
    </div>

    <!-- Preview section -->
    <div class="preview-section">
      <div class="preview-header">
        <span>预览</span>
        <span id="previewInfo" class="preview-info"></span>
      </div>
      <pre id="markdownPreview" class="md-preview"></pre>
    </div>

    <!-- Copy button -->
    <button type="button" id="copyBtn" class="md-button md-button--copy" disabled>
      复制到剪贴板
    </button>

    <!-- Toast notification -->
    <div id="toast" class="md-toast">已复制到剪贴板</div>
  </div>
</body>
</html>
```

- [ ] **Step 2: 验证文件写入**

确认文件内容已更新。

---

### Task 2: 重构 CSS 样式

**文件:**
- 修改: `src/taskpane/markdown-export-taskpane.css`

- [ ] **Step 1: 写入新的 CSS 样式**

```css
/* Markdown Export Taskpane Styles - 蓝色主题 */

.markdown-export-container {
  padding: 20px;
}

.md-title {
  font-size: 16px;
  font-weight: 700;
  color: #0078d4;
  margin-bottom: 16px;
  display: flex;
  align-items: center;
  gap: 8px;
}

.md-title svg {
  flex-shrink: 0;
}

.guide-card {
  background: #f0f7ff;
  border-left: 3px solid #0078d4;
  padding: 12px;
  border-radius: 4px;
  margin-bottom: 16px;
}

.guide-card p {
  font-size: 13px;
  color: #005a9e;
  line-height: 1.6;
  margin: 0;
}

.warning-box {
  background: #f8d7da;
  border: 1px solid #f5c6cb;
  border-radius: 4px;
  padding: 10px 12px;
  margin-bottom: 16px;
  display: flex;
  align-items: center;
  gap: 8px;
  font-size: 13px;
  color: #721c24;
}

.warning-icon {
  font-size: 16px;
}

.md-btn-row {
  display: flex;
  gap: 8px;
  margin-bottom: 16px;
}

.md-button {
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

.md-button:hover {
  background: #0078d4;
  color: white;
}

.md-button:active {
  background: #005a9e;
  border-color: #005a9e;
}

.md-button:disabled {
  background: #f5f5f5;
  color: #999;
  border-color: #999;
  cursor: not-allowed;
}

.md-button--primary {
  background: #0078d4;
  color: white;
}

.md-button--primary:hover {
  background: #005a9e;
  border-color: #005a9e;
}

.md-button--copy {
  width: 100%;
  margin-top: 12px;
}

.preview-section {
  flex: 1;
  display: flex;
  flex-direction: column;
  border: 1px solid #ddd;
  border-radius: 6px;
  overflow: hidden;
  min-height: 200px;
}

.preview-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 10px 12px;
  background: #f6f8fa;
  border-bottom: 1px solid #ddd;
  font-size: 13px;
  font-weight: 500;
  color: #333;
}

.preview-info {
  font-size: 12px;
  color: #666;
  font-weight: 400;
}

.md-preview {
  flex: 1;
  margin: 0;
  padding: 12px;
  overflow: auto;
  font-family: 'Consolas', 'Monaco', monospace;
  font-size: 13px;
  line-height: 1.5;
  white-space: pre-wrap;
  word-break: break-word;
  background: #fff;
  min-height: 100px;
}

.md-preview:empty::before {
  content: '点击"生成 Markdown"预览效果';
  color: #999;
}

.md-toast {
  position: fixed;
  bottom: 20px;
  left: 50%;
  transform: translateX(-50%) translateY(100px);
  background: #333;
  color: #fff;
  padding: 10px 20px;
  border-radius: 6px;
  font-size: 14px;
  opacity: 0;
  transition: transform 0.3s, opacity 0.3s;
  z-index: 1000;
}

.md-toast.show {
  transform: translateX(-50%) translateY(0);
  opacity: 1;
}
```

- [ ] **Step 2: 验证文件写入**

确认文件内容已更新。

---

### Task 3: 重构 JS 逻辑

**文件:**
- 修改: `src/taskpane/markdown-export-taskpane.js`

- [ ] **Step 1: 重写 JS 逻辑，添加刷新和生成按钮事件处理**

```javascript
/* global Office, Excel */

import { generateMarkdownTable } from "../utils/markdown-table-utils";

let toastTimer = null;
let currentMarkdown = ""; // 缓存当前生成的 markdown

Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    initialize();
  }
});

function initialize() {
  document.getElementById("refreshBtn").addEventListener("click", refreshSelection);
  document.getElementById("generateBtn").addEventListener("click", generateMarkdown);
  document.getElementById("copyBtn").addEventListener("click", copyToClipboard);

  // 初始化按钮状态
  updateButtonStates(false);
}

function updateButtonStates(hasPreview) {
  document.getElementById("copyBtn").disabled = !hasPreview;
}

/**
 * 刷新选区信息（从 Excel 获取当前选区范围地址）
 */
async function refreshSelection() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(["address"]);
      await context.sync();

      // 更新预览信息显示选区地址
      const address = range.address;
      document.getElementById("previewInfo").textContent = `已选: ${address}`;
    });
  } catch (error) {
    showError("读取选区失败：" + error.message);
  }
}

/**
 * 生成 Markdown 预览
 */
async function generateMarkdown() {
  try {
    // 清除之前的警告
    hideWarning();

    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(["values", "mergedRanges", "address", "rowCount", "columnCount"]);

      await context.sync();

      const values = range.values;
      const mergedRanges = range.mergedRanges;
      const address = range.address;

      if (!values || !values.length || !values[0].length) {
        showError("请先选择数据区域");
        updateButtonStates(false);
        return;
      }

      // 检查数据量
      const rowCount = values.length;
      const colCount = values[0].length;
      if (rowCount > 10000) {
        showError("数据量过大（超过 10000 行），请缩小选区");
        updateButtonStates(false);
        return;
      }

      // 检查合并单元格
      if (mergedRanges && mergedRanges.length > 0) {
        showWarning("此表格含合并单元格，生成的 Markdown 可能不完美");
      }

      // 生成 Markdown（只取前10行预览）
      const previewRows = Math.min(values.length, 10);
      const previewValues = values.slice(0, previewRows);
      const isTruncated = values.length > 10;

      currentMarkdown = generateMarkdownTable(values, {
        includeAlignment: true,
        preserveFormat: true,
      });

      // 显示预览（带截断提示）
      let previewMarkdown = generateMarkdownTable(previewValues, {
        includeAlignment: true,
        preserveFormat: true,
      });

      if (isTruncated) {
        previewMarkdown += "\n\n*... 共 " + values.length + " 行，仅显示前 10 行 ...*";
      }

      document.getElementById("markdownPreview").textContent = previewMarkdown;
      document.getElementById("previewInfo").textContent = `${address} | ${rowCount}行 × ${colCount}列`;

      updateButtonStates(true);
    });
  } catch (error) {
    showError("生成 Markdown 失败：" + error.message);
    updateButtonStates(false);
  }
}

/**
 * 复制到剪贴板
 */
async function copyToClipboard() {
  if (!currentMarkdown) {
    showError("请先生成 Markdown");
    return;
  }

  let success = false;

  try {
    await navigator.clipboard.writeText(currentMarkdown);
    success = true;
  } catch (error) {
    // Fallback for IE11
    try {
      const textarea = document.createElement("textarea");
      textarea.value = currentMarkdown;
      textarea.style.position = "fixed";
      textarea.style.opacity = "0";
      document.body.appendChild(textarea);
      textarea.select();
      document.execCommand("copy");
      document.body.removeChild(textarea);
      success = true;
    } catch (e) {
      try {
        document.body.removeChild(textarea);
      } catch (e2) {}
    }
  }

  if (success) {
    showToast();
  } else {
    showError("复制失败，请手动复制");
  }
}

function showToast() {
  const toast = document.getElementById("toast");
  toast.classList.add("show");

  if (toastTimer) clearTimeout(toastTimer);
  toastTimer = setTimeout(() => {
    toast.classList.remove("show");
  }, 2000);
}

function showError(message) {
  document.getElementById("warningBox").style.display = "flex";
  document.getElementById("warningText").textContent = message;
  document.getElementById("markdownPreview").textContent = "";
  currentMarkdown = "";
}

function showWarning(message) {
  document.getElementById("warningBox").style.display = "flex";
  document.getElementById("warningText").textContent = message;
}

function hideWarning() {
  document.getElementById("warningBox").style.display = "none";
}
```

- [ ] **Step 2: 验证文件写入**

确认文件内容已更新。

---

### Task 4: 验证构建

- [ ] **Step 1: 运行 lint 检查**

Run: `npm run lint`
EXPECT: 无 lint 错误

- [ ] **Step 2: 运行开发服务器验证**

Run: `npm run dev-server`
EXPECT: 服务器启动成功，https://localhost:3000/markdown-export-taskpane.html 可访问

- [ ] **Step 3: 手动验证清单**

检查清单：
- [ ] 标题显示 SVG 图标 + "导出 Markdown 表格"
- [ ] 引导卡片显示操作说明
- [ ] "刷新选择" 按钮可见
- [ ] "生成 Markdown" 按钮可见
- [ ] 点击 "生成 Markdown" 后预览区显示表格（最多10行）
- [ ] "复制到剪贴板" 按钮在有预览后启用
- [ ] 点击 "复制到剪贴板" 显示 toast 提示

---

## 验收标准

- [ ] HTML 结构与 vlookup/concat 风格一致
- [ ] CSS 使用蓝色主题 (#0078d4)
- [ ] 刷新按钮可获取当前选区地址
- [ ] 生成按钮生成 Markdown 预览（最多10行）
- [ ] 复制按钮在有内容时启用
- [ ] 无 lint 错误
- [ ] 界面美观，与项目其他 taskpane 风格统一

## 完成检查清单

- [ ] 代码遵循 discovered patterns
- [ ] 无 lint 错误
- [ ] 功能与设计一致
- [ ] 现有 Markdown 生成逻辑保持不变
- [ ] 无控制台错误
- [ ] 自包含 — 实现过程无需额外搜索