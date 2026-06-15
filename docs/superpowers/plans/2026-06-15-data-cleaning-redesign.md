# 数据清洗面板美化实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 将数据清洗面板重设计为 Fluent UI 蓝色主调风格，包含标签页切换、引导卡片、蓝边信息框等元素

**Architecture:** 参照 concat-taskpane 的 HTML/CSS 结构，重写 data-cleaning-taskpane 的 HTML 结构（添加 Fluent UI CSS、body class）、CSS（完整重写）、JS（新增标签切换逻辑和引导卡片动态内容）

**Tech Stack:** 纯 HTML/CSS/JS + Fluent UI fabric CSS + office.js

---

## 文件结构

| 文件 | 变更类型 | 职责 |
|------|---------|------|
| `src/taskpane/data-cleaning-taskpane.html` | 重写 | 结构：标题、引导卡片、标签栏、参数区、预览、选区、操作按钮 |
| `src/taskpane/data-cleaning-taskpane.css` | 重写 | 样式：蓝色主调、引导卡片、标签栏、参数控件、预览表格、按钮 |
| `src/taskpane/data-cleaning-taskpane.js` | 修改 | 新增：标签切换函数、引导卡片动态内容更新 |

---

## Task 1: 重写 HTML 结构

**Files:**
- Modify: `src/taskpane/data-cleaning-taskpane.html`

- [ ] **Step 1: 写入新的 HTML 结构**

完全重写 `data-cleaning-taskpane.html`，参照 concat-taskpane.html 的模板格式：

```html
<!-- Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT License. -->

<!DOCTYPE html>
<html>

<head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>数据清洗</title>

    <!-- Office JavaScript API -->
    <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>

    <!-- Fluent UI -->
    <link rel="stylesheet" href="https://res-1.cdn.office.net/files/fabric-cdn-prod_20230815.002/office-ui-fabric-core/11.1.0/css/fabric.min.css"/>

    <!-- Template styles -->
    <link href="taskpane.css" rel="stylesheet" type="text/css" />
    <link href="data-cleaning-taskpane.css" rel="stylesheet" type="text/css" />
</head>

<body class="ms-font-m ms-welcome ms-Fabric">
    <div class="cleaning-container">
        <!-- 标题 -->
        <h1 class="cleaning-title">
            <svg width="20" height="20" viewBox="0 0 20 20" fill="#0078d4">
                <path d="M14 3a1 1 0 0 1 1 1v5.586l2.707 2.707A1 1 0 0 1 17 14.414V15a1 1 0 0 1-2 0v-1a1 1 0 0 0-1-1H9a1 1 0 0 0-1 1v1a1 1 0 0 1-2 0v-.586a1 1 0 0 1 .293-.707L11 9.586V4a1 1 0 0 1 1-1h2z"/>
            </svg>
            数据清洗
        </h1>

        <!-- 引导卡片（动态内容，由 JS 控制） -->
        <div class="guide-card" id="guide-card">
            <p id="guide-text"></p>
        </div>

        <!-- 标签栏 -->
        <div class="tab-bar" id="tab-bar">
            <button class="tab-btn selected" data-op="trimSpaces">修剪空格</button>
            <button class="tab-btn" data-op="removeEmpty">删除空行</button>
            <button class="tab-btn" data-op="convertCase">大小写转换</button>
            <button class="tab-btn" data-op="removeInvisible">清除不可见</button>
            <button class="tab-btn" data-op="removeDuplicates">移除重复行</button>
        </div>

        <!-- 参数设置区域 -->
        <div class="params-section" id="params-section">
            <div id="params-container"></div>
        </div>

        <!-- 预览表格 -->
        <div class="preview-section">
            <h3 class="section-title">👁 预览（前 5 行）</h3>
            <div class="preview-container">
                <table id="preview-table">
                    <thead><tr><th>原值</th><th>清洗后</th></tr></thead>
                    <tbody id="preview-body"></tbody>
                </table>
            </div>
        </div>

        <!-- 选区信息 -->
        <div class="selection-bar">
            <span id="selection-address" class="label">选区：未选择</span>
            <button id="refresh-btn" class="refresh-btn" title="刷新选区">↻</button>
        </div>

        <!-- 操作按钮 -->
        <div class="action-bar">
            <button id="execute-btn" class="cleaning-btn-primary" disabled>🚀 执行清洗</button>
            <button id="undo-btn" class="cleaning-btn-secondary" disabled>↩ 撤销</button>
        </div>

        <!-- 状态 -->
        <div id="status" class="status-message status-idle"></div>
    </div>
</body>

</html>
```

- [ ] **Step 2: 提交 HTML 变更**

```bash
git add src/taskpane/data-cleaning-taskpane.html
git commit -m "refactor: rewrite data-cleaning HTML with Fluent UI structure"
```

---

## Task 2: 重写 CSS 样式

**Files:**
- Modify: `src/taskpane/data-cleaning-taskpane.css`

- [ ] **Step 1: 写入完整 CSS**

完全重写 `data-cleaning-taskpane.css`，参照 `concat-taskpane.css` 的风格：

```css
/* 数据清洗面板样式 - Fluent UI 蓝色主调 */

.cleaning-container {
    padding: 20px;
}

/* 标题 */
.cleaning-title {
    font-size: 16px;
    font-weight: 700;
    color: #0078d4;
    margin-bottom: 16px;
    display: flex;
    align-items: center;
    gap: 8px;
}

.cleaning-title svg {
    flex-shrink: 0;
}

/* 引导卡片 - 左侧蓝边 */
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

/* 标签栏 */
.tab-bar {
    display: flex;
    flex-wrap: wrap;
    gap: 6px;
    margin-bottom: 16px;
}

.tab-btn {
    padding: 8px 14px;
    border: none;
    border-radius: 4px;
    background: #f0f0f0;
    color: #666;
    font-size: 13px;
    cursor: pointer;
    transition: all 0.15s;
    position: relative;
}

.tab-btn:hover {
    background: #e8f4ff;
    color: #0078d4;
}

.tab-btn.selected {
    background: #0078d4;
    color: #fff;
}

.tab-btn.selected::after {
    content: "";
    position: absolute;
    bottom: 0;
    left: 0;
    right: 0;
    height: 3px;
    background: #005a9e;
    border-radius: 0 0 2px 2px;
}

/* 参数设置区域 */
.params-section {
    background: #fafafa;
    border: 1px solid #eee;
    border-radius: 4px;
    padding: 12px;
    margin-bottom: 16px;
    min-height: 50px;
}

.param-row {
    display: flex;
    align-items: center;
    gap: 10px;
    margin-bottom: 10px;
}

.param-row:last-child {
    margin-bottom: 0;
}

.param-row label {
    font-size: 13px;
    color: #666;
    min-width: 70px;
}

.param-row select,
.param-row input[type="number"] {
    flex: 1;
    padding: 8px 10px;
    border: 2px solid #0078d4;
    border-radius: 6px;
    font-size: 14px;
    outline: none;
    transition: box-shadow 0.2s;
    box-sizing: border-box;
}

.param-row select:focus,
.param-row input:focus {
    box-shadow: 0 0 0 3px rgba(0,120,212,0.15);
}

.checkbox-group {
    display: flex;
    flex-wrap: wrap;
    gap: 8px;
}

.checkbox-group label {
    font-size: 12px;
    display: flex;
    align-items: center;
    gap: 4px;
    cursor: pointer;
}

/* 预览表格 */
.preview-section {
    margin-bottom: 16px;
}

.section-title {
    font-size: 13px;
    font-weight: 600;
    color: #333;
    margin-bottom: 8px;
}

.preview-container {
    max-height: 200px;
    overflow-y: auto;
    border: 1px solid #eee;
    border-radius: 4px;
}

#preview-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 12px;
}

#preview-table th,
#preview-table td {
    padding: 6px 8px;
    border: 1px solid #eee;
    text-align: left;
    word-break: break-all;
}

#preview-table th {
    background: #f0f0f0;
    font-weight: 600;
}

#preview-table td:first-child {
    background: #fff3f3;
}

#preview-table td:last-child {
    background: #f3fff3;
}

/* 选区栏 */
.selection-bar {
    display: flex;
    align-items: center;
    gap: 8px;
    padding: 8px 12px;
    background: #f5f5f5;
    border-radius: 4px;
    margin-bottom: 12px;
}

.label {
    font-size: 12px;
    color: #666;
    flex: 1;
}

.refresh-btn {
    padding: 4px 10px;
    background: transparent;
    color: #0078d4;
    border: 1px solid #0078d4;
    border-radius: 4px;
    cursor: pointer;
    font-size: 13px;
    transition: all 0.15s;
}

.refresh-btn:hover {
    background: #0078d4;
    color: #fff;
}

/* 操作按钮 */
.action-bar {
    display: flex;
    gap: 10px;
    margin-bottom: 12px;
}

.cleaning-btn-primary {
    flex: 1;
    padding: 10px 16px;
    background: #0078d4;
    color: #fff;
    border: none;
    border-radius: 6px;
    font-size: 14px;
    font-weight: 600;
    cursor: pointer;
    transition: all 0.2s;
}

.cleaning-btn-primary:hover {
    background: #106ebe;
}

.cleaning-btn-primary:disabled {
    background: #ccc;
    cursor: not-allowed;
}

.cleaning-btn-secondary {
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

.cleaning-btn-secondary:hover {
    background: #f0f7ff;
}

.cleaning-btn-secondary:disabled {
    color: #999;
    border-color: #999;
    cursor: not-allowed;
}

/* 状态栏 */
.status-message {
    padding: 10px;
    border-radius: 6px;
    font-size: 13px;
    text-align: center;
}

.status-idle { display: none; }
.status-loading { background: #fff3cd; color: #856404; border: 1px solid #ffeeba; }
.status-success { background: #d4edda; color: #155724; border: 1px solid #c3e6cb; }
.status-error { background: #f8d7da; color: #721c24; border: 1px solid #f5c6cb; }

/* 空状态 */
.empty-state {
    padding: 16px;
    text-align: center;
    color: #999;
    font-size: 13px;
}
```

- [ ] **Step 2: 提交 CSS 变更**

```bash
git add src/taskpane/data-cleaning-taskpane.css
git commit -m "refactor: rewrite data-cleaning CSS with Fluent UI blue theme"
```

---

## Task 3: 修改 JS — 标签切换逻辑和引导卡片

**Files:**
- Modify: `src/taskpane/data-cleaning-taskpane.js`

需要修改的部分：

1. 新增 `GUIDE_TEXTS` 配置对象（每个操作对应的引导文字）
2. 新增 `switchTab()` 函数（标签切换 + 更新引导卡片）
3. 修改 `selectOperation()` 函数，调用 `switchTab()` 更新标签高亮和引导卡片
4. 初始化时默认选中第一个标签

- [ ] **Step 1: 在 JS 文件开头 state 对象后添加 GUIDE_TEXTS 配置**

在 `state` 对象声明之后、`OPERATION_PARAMS` 之前添加：

```javascript
// ===== 引导文案配置 =====
var GUIDE_TEXTS = {
  trimSpaces: "1. 选中要清洗的单元格区域 2. 选择修剪模式 3. 查看预览后执行",
  removeEmpty: "1. 选中数据区域 2. 选择判定方式 3. 查看预览后执行",
  convertCase: "1. 选中要转换的单元格 2. 选择转换模式 3. 查看预览后执行",
  removeInvisible: "1. 选中单元格区域 2. 选择清除类型 3. 查看预览后执行",
  removeDuplicates: "1. 选中数据区域 2. 选择依据列和保留规则 3. 查看预览后执行",
};
```

- [ ] **Step 2: 在 `Office.onReady` 中修改操作按钮点击逻辑**

将原来的操作按钮循环绑定改为标签点击逻辑。在 `Office.onReady` 函数中：

```javascript
// 标签按钮点击切换
var tabBtns = document.querySelectorAll(".tab-btn");
for (var i = 0; i < tabBtns.length; i++) {
    tabBtns[i].addEventListener("click", function () {
        var opKey = this.dataset.op;
        switchTab(opKey);
        selectOperation(opKey);
    });
}
```

- [ ] **Step 3: 添加 `switchTab()` 函数**

在 `GUIDE_TEXTS` 配置之后、`OPERATION_PARAMS` 之前添加：

```javascript
// ===== 标签切换 =====
function switchTab(opKey) {
    // 高亮标签
    var tabBtns = document.querySelectorAll(".tab-btn");
    for (var i = 0; i < tabBtns.length; i++) {
        tabBtns[i].classList.remove("selected");
    }
    var activeTab = document.querySelector('.tab-btn[data-op="' + opKey + '"]');
    if (activeTab) activeTab.classList.add("selected");

    // 更新引导卡片
    var guideText = document.getElementById("guide-text");
    if (guideText && GUIDE_TEXTS[opKey]) {
        guideText.textContent = GUIDE_TEXTS[opKey];
    }
}
```

- [ ] **Step 4: 初始化时默认选中第一个标签**

在 `Office.onReady` 函数末尾（`loadSelection()` 调用之前）添加：

```javascript
// 默认选中第一个标签
switchTab("trimSpaces");
```

- [ ] **Step 5: 提交 JS 变更**

```bash
git add src/taskpane/data-cleaning-taskpane.js
git commit -m "feat: add tab switching and dynamic guide card in data-cleaning"
```

---

## Task 4: 构建验证

- [ ] **Step 1: 运行开发构建**

```bash
cd f:/projects/My\ Office\ Add-in && npm run build:dev
```

- [ ] **Step 2: 验证输出文件**

检查 `dist/data-cleaning-taskpane.html` 包含：
- Fluent UI CSS 链接
- `office.js` 脚本
- body 有 `ms-font-m ms-welcome ms-Fabric` class
- `.cleaning-container` 容器
- `.tab-bar` 标签栏
- `.guide-card` 引导卡片

检查 `dist/data-cleaning-taskpane.css` 存在且包含样式定义。

- [ ] **Step 3: 提交构建产物**

```bash
git add dist/
git commit -m "build: include redesigned data-cleaning taskpane assets"
```

---

## 自检清单

- [ ] HTML 中引入了 Fluent UI CSS 和 office.js
- [ ] body 有正确的 class
- [ ] 5 个标签按钮，data-op 属性正确
- [ ] 引导卡片默认显示"修剪空格"的引导文案
- [ ] 标签切换时引导卡片内容同步更新
- [ ] 参数控件样式为蓝色边框
- [ ] 预览表格原值列红底、清洗后列绿底
- [ ] 主按钮蓝底白字，次按钮透明蓝边
- [ ] 状态栏样式完整（idle/loading/success/error）
- [ ] 构建无报错