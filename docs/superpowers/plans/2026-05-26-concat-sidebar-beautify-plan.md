# 连接列侧边栏美化实施计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 对连接列侧边栏进行视觉美化，采用「友好引导」设计风格

**Architecture:** 修改现有 concat-taskpane.html 结构，添加新样式文件 concat-taskpane.css，实现蓝色强调色 + 信息卡片布局

**Tech Stack:** Pure HTML/CSS/JavaScript，遵循 Fluent UI 风格

---

## 文件结构

- 修改: `src/taskpane/concat-taskpane.html` — 调整 HTML 结构
- 创建: `src/taskpane/concat-taskpane.css` — 新样式文件
- 修改: `src/taskpane/concat-taskpane.js` — 更新状态显示逻辑

---

## 实施步骤

### Task 1: 创建样式文件 concat-taskpane.css

**Files:**
- Create: `src/taskpane/concat-taskpane.css`

- [ ] **Step 1: 创建样式文件**

```css
/* 连接列侧边栏样式 - 友好引导风格 */

.concat-container {
  padding: 20px;
}

.concat-title {
  font-size: 16px;
  font-weight: 700;
  color: #0078d4;
  margin-bottom: 16px;
  display: flex;
  align-items: center;
  gap: 8px;
}

.concat-title svg {
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

.guide-step {
  margin-bottom: 6px;
}

.guide-step:last-child {
  margin-bottom: 0;
}

.concat-form-group {
  margin-bottom: 16px;
}

.concat-label {
  display: block;
  font-size: 13px;
  color: #666;
  margin-bottom: 6px;
}

.concat-input {
  width: 100%;
  padding: 10px 12px;
  border: 2px solid #0078d4;
  border-radius: 6px;
  font-size: 15px;
  outline: none;
  transition: box-shadow 0.2s;
  box-sizing: border-box;
}

.concat-input:focus {
  box-shadow: 0 0 0 3px rgba(0,120,212,0.15);
}

.status-message {
  padding: 10px;
  border-radius: 6px;
  font-size: 13px;
}

.status-idle {
  background: #f5f5f5;
  color: #666;
}

.status-success {
  background: #d4edda;
  border: 1px solid #c3e6cb;
  color: #155724;
}

.status-error {
  background: #f8d7da;
  border: 1px solid #f5c6cb;
  color: #721c24;
}

.status-loading {
  background: #fff3cd;
  border: 1px solid #ffeeba;
  color: #856404;
}
```

- [ ] **Step 2: 提交**

```bash
git add src/taskpane/concat-taskpane.css
git commit -m "feat: add concat-taskpane styles with friendly design"
```

---

### Task 2: 修改 concat-taskpane.html 结构

**Files:**
- Modify: `src/taskpane/concat-taskpane.html:1-37`

- [ ] **Step 1: 更新 HTML 结构**

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
                <span class="guide-step">2. 输入连接符，按 Enter 执行</span>
            </p>
        </div>

        <div class="concat-form-group">
            <label for="connector" class="concat-label">连接符</label>
            <input type="text" id="connector" class="concat-input" value="_" autofocus />
        </div>

        <div id="status" class="status-message status-idle">状态：等待操作...</div>
    </div>
</body>

</html>
```

- [ ] **Step 2: 提交**

```bash
git add src/taskpane/concat-taskpane.html
git commit -m "feat: update concat-taskpane HTML with guide card layout"
```

---

### Task 3: 修改 concat-taskpane.js 状态显示逻辑

**Files:**
- Modify: `src/taskpane/concat-taskpane.js`

- [ ] **Step 1: 更新状态显示函数**

找到状态相关的代码，替换为使用新样式类：

在 `runConcat` 函数中，更新状态显示逻辑：

```javascript
function setStatus(message, type) {
  var statusEl = document.getElementById("status");
  statusEl.textContent = message;
  statusEl.className = "status-message status-" + type;
}
```

更新 `runConcat` 函数中的状态调用：

```javascript
// 原来的：
statusEl.textContent = "处理中...";
statusEl.style.color = "green";

// 改为：
setStatus("处理中...", "loading");

// 成功：
setStatus("完成! 已在第 " + targetColLetter + " 列写入 " + rowCount + " 行公式", "success");

// 错误：
setStatus("错误: " + error.message, "error");
```

- [ ] **Step 2: 提交**

```bash
git add src/taskpane/concat-taskpane.js
git commit -m "feat: update concat-taskpane status display with CSS classes"
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
3. 输入连接符，按 Enter 执行
4. 验证成功状态显示（绿色）
5. 验证错误状态显示（红色）

- [ ] **Step 4: 提交最终版本**

```bash
git add .
git commit -m "feat: complete concat-sidebar beautification"
```

