# 连接列侧边栏添加执行按钮实施计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 在连接列侧边栏添加执行按钮，使用轮廓按钮样式，移除 Enter 键触发

**Architecture:** 修改现有 concat-taskpane.html/js/css，添加按钮元素和点击事件，移除 Enter 监听

**Tech Stack:** Pure HTML/CSS/JavaScript，遵循 Fluent UI 风格

---

## 文件结构

- Modify: `src/taskpane/concat-taskpane.html` — 添加按钮，移除 autofocus
- Modify: `src/taskpane/concat-taskpane.css` — 添加按钮样式
- Modify: `src/taskpane/concat-taskpane.js` — 移除 Enter 监听，添加按钮点击事件

---

## 实施步骤

### Task 1: 添加按钮样式到 concat-taskpane.css

**Files:**
- Modify: `src/taskpane/concat-taskpane.css`

- [ ] **Step 1: 添加按钮样式**

在现有样式末尾添加：

```css
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
  margin-top: 12px;
}

.concat-button:hover {
  background: #0078d4;
  color: white;
}

.concat-button:active {
  background: #005a9e;
  border-color: #005a9e;
}

.concat-button:disabled {
  background: #f5f5f5;
  color: #999;
  border-color: #999;
  cursor: not-allowed;
}
```

- [ ] **Step 2: 提交**

```bash
git add src/taskpane/concat-taskpane.css
git commit -m "feat: add execute button styles to concat-taskpane"
```

---

### Task 2: 修改 concat-taskpane.html 添加按钮

**Files:**
- Modify: `src/taskpane/concat-taskpane.html`

- [ ] **Step 1: 更新 HTML 结构**

修改步骤引导：
```html
<span class="guide-step">2. 输入连接符，点击执行</span>
```

添加按钮元素（在 status div 之前）：
```html
<button id="executeBtn" class="concat-button">执行连接</button>
```

移除输入框的 autofocus 属性。

完整 HTML：

```html
<!-- Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT License. -->

<!DOCTYPE html>
<html>

<head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>连接列工具</title>

    <!-- Office JavaScript API -->
    <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>

    <!-- For more information on Fluent UI, visit https://developer.microsoft.com/fluentui#/. -->
    <link rel="stylesheet" href="https://res-1.cdn.office.net/files/fabric-cdn-prod_20230815.002/office-ui-fabric-core/11.1.0/css/fabric.min.css"/>

    <!-- Template styles -->
    <link href="taskpane.css" rel="stylesheet" type="text/css" />
    <link href="concat-taskpane.css" rel="stylesheet" type="text/css" />
</head>

<body class="ms-font-m ms-welcome ms-Fabric">
    <div class="concat-container">
        <h1 class="concat-title">
            <svg width="20" height="20" viewBox="0 0 20 20" fill="#0078d4">
                <path d="M14 3a1 1 0 0 1 1 1v5.586l2.707 2.707A1 1 0 0 1 17 14.414V15a1 1 0 0 1-2 0v-1a1 1 0 0 0-1-1H9a1 1 0 0 0-1 1v1a1 1 0 0 1-2 0v-.586a1 1 0 0 1 .293-.707L11 9.586V4a1 1 0 0 1 1-1h2z"/>
            </svg>
            连接列工具
        </h1>

        <div class="guide-card">
            <p>
                <span class="guide-step">1. 选中 Excel 中要连接的两列</span><br>
                <span class="guide-step">2. 输入连接符，点击执行</span>
            </p>
        </div>

        <div class="concat-form-group">
            <label for="connector" class="concat-label">连接符</label>
            <input type="text" id="connector" class="concat-input" value="_" />
        </div>

        <button id="executeBtn" class="concat-button">执行连接</button>

        <div id="status" class="status-message status-idle">状态：等待操作...</div>
    </div>
</body>

</html>
```

- [ ] **Step 2: 提交**

```bash
git add src/taskpane/concat-taskpane.html
git commit -m "feat: add execute button to concat-taskpane HTML"
```

---

### Task 3: 修改 concat-taskpane.js 移除 Enter 监听，添加按钮事件

**Files:**
- Modify: `src/taskpane/concat-taskpane.js`

- [ ] **Step 1: 更新 JavaScript 逻辑**

读取当前文件内容，找到并修改 Office.onReady 函数和 runConcat 函数。

修改 Office.onReady 函数（移除 Enter 监听）：
```javascript
Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    var connectorInput = document.getElementById("connector");
    var executeBtn = document.getElementById("executeBtn");

    // 自动聚焦输入框
    connectorInput.focus();

    // 按钮点击触发执行
    executeBtn.addEventListener("click", function() {
      runConcat();
    });
  }
});
```

修改 runConcat 函数（添加按钮禁用控制）：
```javascript
function runConcat() {
  var statusEl = document.getElementById("status");
  var connectorInput = document.getElementById("connector");
  var executeBtn = document.getElementById("executeBtn");
  var connector = connectorInput.value || "_";

  // 禁用按钮，显示加载状态
  executeBtn.disabled = true;
  setStatus("处理中...", "loading");

  Excel.run(function (context) {
    // ... 现有逻辑 ...

    // 成功后启用按钮
    executeBtn.disabled = false;
    setStatus("完成! 已在第 " + targetColLetter + " 列写入 " + rowCount + " 行公式", "success");
  }).catch(function (error) {
    // 错误时启用按钮
    executeBtn.disabled = false;
    setStatus("错误: " + error.message, "error");
  });
}
```

- [ ] **Step 2: 提交**

```bash
git add src/taskpane/concat-taskpane.js
git commit -m "feat: add button click handler, remove Enter key listener"
```

---

### Task 4: 测试完整流程

**Files:**
- Test: `src/taskpane/concat-taskpane.html`
- Test: `src/taskpane/concat-taskpane.js`

- [ ] **Step 1: 启动开发服务器**

```bash
npm run dev-server
```

- [ ] **Step 2: 启动 Excel 调试**

```bash
npm run start
```

- [ ] **Step 3: 测试交互流程**

1. 点击「连接列」按钮打开侧边栏
2. 验证输入框自动聚焦
3. 输入连接符，点击按钮执行（验证 Enter 键不再触发）
4. 验证按钮在处理中时禁用
5. 验证成功状态显示（绿色）
6. 验证错误状态显示（红色）

- [ ] **Step 4: 提交最终版本**

```bash
git add .
git commit -m "feat: complete concat-sidebar execute button"
```

