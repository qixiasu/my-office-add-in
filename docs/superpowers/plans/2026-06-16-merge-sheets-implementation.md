# 多 Sheet 合并工具实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 实现 Excel 多 Sheet 合并工具，用户可选择多个 Sheet 并指定表头行号后合并为一个新 Sheet。

**Architecture:** 单面板工具，使用 Office JavaScript API 读取工作簿 Sheet 列表，验证列数一致性，合并数据到新建的"合并结果"Sheet。

**Tech Stack:** Vanilla JavaScript, Office.js, Webpack

---

## 文件结构

| 操作 | 文件路径 |
|------|----------|
| 创建 | `src/taskpane/merge-sheets-taskpane.html` |
| 创建 | `src/taskpane/merge-sheets-taskpane.js` |
| 创建 | `src/taskpane/merge-sheets-taskpane.css` |
| 创建 | `src/utils/merge-sheets-utils.js` |
| 修改 | `manifest.xml` — 添加按钮和资源 |
| 修改 | `webpack.config.js` — 添加入口和插件 |

---

## Task 1: 创建 merge-sheets-utils.js 工具函数

**Files:**
- Create: `src/utils/merge-sheets-utils.js`

- [ ] **Step 1: 编写工具函数**

```javascript
/**
 * 获取 Excel 列字母（A-Z, AA-ZZ, ...）
 * @param {number} colIndex - 0-based column index
 * @returns {string} Column letter(s)
 */
function getColumnLetter(colIndex) {
  var letter = "";
  var remaining = colIndex;
  do {
    letter = String.fromCharCode(65 + (remaining % 26)) + letter;
    remaining = Math.floor(remaining / 26) - 1;
  } while (remaining >= 0);
  return letter;
}

/**
 * 验证 Sheet 数据列数是否一致
 * @param {Array<{name: string, columnCount: number}>} sheets
 * @returns {{valid: boolean, expectedColumnCount: number|null, error: string|null}}
 */
function validateColumnConsistency(sheets) {
  if (!sheets || sheets.length < 2) {
    return {
      valid: false,
      expectedColumnCount: null,
      error: "请至少选择两个 Sheet"
    };
  }
  var expected = sheets[0].columnCount;
  for (var i = 1; i < sheets.length; i++) {
    if (sheets[i].columnCount !== expected) {
      return {
        valid: false,
        expectedColumnCount: expected,
        error: "所选 Sheet 列数不一致，请重新选择"
      };
    }
  }
  return { valid: true, expectedColumnCount: expected, error: null };
}

/**
 * 生成唯一的 Sheet 名称
 * @param {string} baseName - 基础名称，如"合并结果"
 * @param {string[]} existingNames - 已存在的 Sheet 名称数组
 * @returns {string} 唯一可用名称
 */
function generateUniqueSheetName(baseName, existingNames) {
  if (existingNames.indexOf(baseName) === -1) {
    return baseName;
  }
  for (var i = 1; i <= 100; i++) {
    var name = baseName + "_" + i;
    if (existingNames.indexOf(name) === -1) {
      return name;
    }
  }
  return baseName + "_" + Date.now();
}

module.exports = {
  getColumnLetter: getColumnLetter,
  validateColumnConsistency: validateColumnConsistency,
  generateUniqueSheetName: generateUniqueSheetName
};
```

- [ ] **Step 2: 提交**

```bash
git add src/utils/merge-sheets-utils.js
git commit -m "feat: 添加多Sheet合并工具辅助函数"
```

---

## Task 2: 创建 merge-sheets-taskpane.html 界面模板

**Files:**
- Create: `src/taskpane/merge-sheets-taskpane.html`

- [ ] **Step 1: 编写 HTML**

```html
<!-- Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT License. -->

<!DOCTYPE html>
<html>

<head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>合并 Sheet</title>

    <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    <link rel="stylesheet" href="https://res-1.cdn.office.net/files/fabric-cdn-prod_20230815.002/office-ui-fabric-core/11.1.0/css/fabric.min.css"/>
    <link href="taskpane.css" rel="stylesheet" type="text/css" />
    <link href="merge-sheets-taskpane.css" rel="stylesheet" type="text/css" />
</head>

<body class="ms-font-m ms-welcome ms-Fabric">
    <div class="merge-container">
        <h1 class="merge-title">
            <svg width="20" height="20" viewBox="0 0 20 20" fill="#0078d4">
                <path d="M16 10a1 1 0 0 1-1 1H5a1 1 0 0 1 0-2h10a1 1 0 0 1 1 1z"/>
                <path d="M8 5a1 1 0 0 1 1 1v3.586l2.707 2.707A1 1 0 0 1 12 14.414V15a1 1 0 0 1-2 0v-1a1 1 0 0 0-1-1H4a1 1 0 0 0-1 1v1a1 1 0 0 1-2 0V9a1 1 0 0 1 1-1h5z"/>
            </svg>
            合并 Sheet
        </h1>

        <div class="guide-card">
            <p>
                <span class="guide-step">1. 勾选要合并的 Sheet（至少2个）</span><br>
                <span class="guide-step">2. 设置每个 Sheet 的表头行号（0=无表头）</span><br>
                <span class="guide-step">3. 点击执行合并</span>
            </p>
        </div>

        <div class="merge-list-header">
            <span class="merge-list-title">工作表列表</span>
            <div class="merge-list-actions">
                <button id="selectAllBtn" class="merge-btn merge-btn--small">全选</button>
                <button id="deselectAllBtn" class="merge-btn merge-btn--small">取消全选</button>
            </div>
        </div>

        <div id="sheetList" class="merge-sheet-list">
            <div class="merge-loading">加载中...</div>
        </div>

        <button id="executeBtn" class="merge-button merge-button--primary">执行合并</button>

        <div id="status" class="merge-status merge-status--idle">状态：等待操作...</div>

        <div id="confirmBox" class="merge-confirm-box" style="display: none;">
            <div id="confirmMsg" class="merge-confirm-msg"></div>
            <div class="merge-confirm-buttons">
                <button id="confirmOverwrite" class="merge-btn merge-btn--small">覆盖</button>
                <button id="confirmRename" class="merge-btn merge-btn--small">重命名</button>
                <button id="confirmCancel" class="merge-btn merge-btn--small merge-btn--secondary">取消</button>
            </div>
        </div>
    </div>
</body>

</html>
```

- [ ] **Step 2: 提交**

```bash
git add src/taskpane/merge-sheets-taskpane.html
git commit -m "feat: 添加多Sheet合并工具HTML模板"
```

---

## Task 3: 创建 merge-sheets-taskpane.css 样式

**Files:**
- Create: `src/taskpane/merge-sheets-taskpane.css`

- [ ] **Step 1: 编写样式**

```css
/* Merge Sheets Taskpane Styles */

.merge-container {
  padding: 20px;
}

.merge-title {
  font-size: 16px;
  font-weight: 700;
  color: #0078d4;
  margin-bottom: 16px;
  display: flex;
  align-items: center;
  gap: 8px;
}

.merge-title svg {
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
  font-weight: 600;
}

.merge-list-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 12px;
}

.merge-list-title {
  font-size: 14px;
  font-weight: 600;
  color: #333;
}

.merge-list-actions {
  display: flex;
  gap: 8px;
}

.merge-sheet-list {
  max-height: 300px;
  overflow-y: auto;
  border: 1px solid #edebe9;
  border-radius: 6px;
  margin-bottom: 16px;
}

.merge-sheet-item {
  display: flex;
  align-items: center;
  padding: 10px 12px;
  border-bottom: 1px solid #edebe9;
}

.merge-sheet-item:last-child {
  border-bottom: none;
}

.merge-sheet-item:hover {
  background: #f5f5f5;
}

.merge-sheet-checkbox {
  width: 18px;
  height: 18px;
  margin-right: 12px;
  cursor: pointer;
}

.merge-sheet-name {
  flex: 1;
  font-size: 14px;
  color: #333;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.merge-sheet-header-input {
  width: 60px;
  padding: 6px 8px;
  border: 1px solid #edebe9;
  border-radius: 4px;
  font-size: 13px;
  text-align: center;
}

.merge-sheet-header-input:focus {
  border-color: #0078d4;
  outline: none;
}

.merge-loading {
  padding: 20px;
  text-align: center;
  color: #666;
  font-size: 13px;
}

.merge-empty {
  padding: 20px;
  text-align: center;
  color: #999;
  font-size: 13px;
}

.merge-btn {
  padding: 6px 12px;
  background: transparent;
  color: #0078d4;
  border: 1px solid #0078d4;
  border-radius: 4px;
  font-size: 12px;
  cursor: pointer;
  transition: all 0.2s;
  white-space: nowrap;
}

.merge-btn:hover {
  background: #0078d4;
  color: white;
}

.merge-btn--secondary {
  color: #666;
  border-color: #999;
}

.merge-btn--secondary:hover {
  background: #f5f5f5;
  color: #333;
}

.merge-button {
  width: 100%;
  padding: 12px 16px;
  background: transparent;
  color: #0078d4;
  border: 2px solid #0078d4;
  border-radius: 6px;
  font-size: 14px;
  font-weight: 600;
  cursor: pointer;
  transition: all 0.2s;
}

.merge-button:hover {
  background: #0078d4;
  color: white;
}

.merge-button:active {
  background: #005a9e;
  border-color: #005a9e;
}

.merge-button:disabled {
  background: #f5f5f5;
  color: #999;
  border-color: #999;
  cursor: not-allowed;
}

.merge-button--primary {
  background: #0078d4;
  color: white;
}

.merge-button--primary:hover {
  background: #005a9e;
  border-color: #005a9e;
}

.merge-status {
  padding: 10px;
  border-radius: 6px;
  font-size: 13px;
  margin-top: 12px;
}

.merge-status--idle {
  background: #f5f5f5;
  color: #666;
}

.merge-status--success {
  background: #d4edda;
  border: 1px solid #c3e6cb;
  color: #155724;
}

.merge-status--error {
  background: #f8d7da;
  border: 1px solid #f5c6cb;
  color: #721c24;
}

.merge-status--loading {
  background: #fff3cd;
  border: 1px solid #ffeeba;
  color: #856404;
}

.merge-confirm-box {
  margin-top: 12px;
  padding: 12px;
  background: #fff3cd;
  border: 1px solid #ffeeba;
  border-radius: 6px;
}

.merge-confirm-msg {
  font-size: 13px;
  color: #856404;
  margin-bottom: 12px;
}

.merge-confirm-buttons {
  display: flex;
  gap: 8px;
  justify-content: flex-end;
}
```

- [ ] **Step 2: 提交**

```bash
git add src/taskpane/merge-sheets-taskpane.css
git commit -m "feat: 添加多Sheet合并工具样式"
```

---

## Task 4: 创建 merge-sheets-taskpane.js 主逻辑

**Files:**
- Create: `src/taskpane/merge-sheets-taskpane.js`

- [ ] **Step 1: 编写主逻辑**

```javascript
/* global console, document, Excel, Office */

var mergeSheetsUtils = require("../utils/merge-sheets-utils");

var pendingConfirmation = null;
var allSheets = [];

Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    initEventListeners();
    loadSheetList();
  }
});

function initEventListeners() {
  document.getElementById("selectAllBtn").onclick = function () {
    setAllCheckboxes(true);
  };

  document.getElementById("deselectAllBtn").onclick = function () {
    setAllCheckboxes(false);
  };

  document.getElementById("executeBtn").onclick = executeMerge;

  document.getElementById("confirmOverwrite").onclick = function () {
    handleConfirm("overwrite");
  };

  document.getElementById("confirmRename").onclick = function () {
    handleConfirm("rename");
  };

  document.getElementById("confirmCancel").onclick = function () {
    handleConfirm("cancel");
  };
}

function setAllCheckboxes(checked) {
  var checkboxes = document.querySelectorAll(".merge-sheet-checkbox");
  checkboxes.forEach(function (cb) {
    cb.checked = checked;
  });
}

function loadSheetList() {
  Excel.run(async function (context) {
    var worksheets = context.workbook.worksheets;
    worksheets.load("items/name");
    await context.sync();

    allSheets = worksheets.items.map(function (ws) {
      return { name: ws.name, checked: false, headerRow: 1 };
    });

    renderSheetList();
  }).catch(function (error) {
    setStatus("加载工作表失败: " + error.message, "error");
  });
}

function renderSheetList() {
  var listEl = document.getElementById("sheetList");

  if (allSheets.length === 0) {
    listEl.innerHTML = '<div class="merge-empty">未找到工作表</div>';
    return;
  }

  var html = "";
  allSheets.forEach(function (sheet, index) {
    html +=
      '<div class="merge-sheet-item">' +
      '<input type="checkbox" class="merge-sheet-checkbox" data-index="' +
      index +
      '" />' +
      '<span class="merge-sheet-name">' +
      escapeHtml(sheet.name) +
      "</span>" +
      '<input type="number" class="merge-sheet-header-input" data-index="' +
      index +
      '" value="1" min="0" />' +
      "</div>";
  });

  listEl.innerHTML = html;

  // Bind checkbox and input events
  var checkboxes = listEl.querySelectorAll(".merge-sheet-checkbox");
  checkboxes.forEach(function (cb) {
    cb.onchange = function () {
      var idx = parseInt(cb.dataset.index, 10);
      allSheets[idx].checked = cb.checked;
    };
  });

  var inputs = listEl.querySelectorAll(".merge-sheet-header-input");
  inputs.forEach(function (input) {
    input.onchange = function () {
      var idx = parseInt(input.dataset.index, 10);
      allSheets[idx].headerRow = parseInt(input.value, 10) || 0;
    };
  });
}

function escapeHtml(text) {
  var div = document.createElement("div");
  div.textContent = text;
  return div.innerHTML;
}

function setStatus(message, type) {
  var el = document.getElementById("status");
  el.textContent = message;
  el.className = "merge-status merge-status--" + (type || "idle");
}

function showConfirmBox(message) {
  var confirmBox = document.getElementById("confirmBox");
  var confirmMsgEl = document.getElementById("confirmMsg");
  confirmMsgEl.textContent = message;
  confirmBox.style.display = "block";
  setStatus("状态：等待确认...", "idle");

  return new Promise(function (resolve) {
    pendingConfirmation = { resolve: resolve };
  });
}

function hideConfirmBox() {
  document.getElementById("confirmBox").style.display = "none";
  pendingConfirmation = null;
}

function handleConfirm(action) {
  hideConfirmBox();
  if (pendingConfirmation) {
    pendingConfirmation.resolve(action);
    pendingConfirmation = null;
  }
}

function getSelectedSheets() {
  return allSheets.filter(function (sheet) {
    return sheet.checked;
  });
}

function executeMerge() {
  var executeBtn = document.getElementById("executeBtn");
  executeBtn.disabled = true;
  setStatus("处理中...", "loading");

  var selectedSheets = getSelectedSheets();

  if (selectedSheets.length < 2) {
    executeBtn.disabled = false;
    setStatus("请至少选择两个 Sheet 进行合并", "error");
    return;
  }

  Excel.run(async function (context) {
    // Step 1: Read column count for each selected sheet
    var sheetsWithMeta = [];
    var baseIndex = 0;

    for (var i = 0; i < allSheets.length; i++) {
      if (allSheets[i].checked) {
        var ws = context.workbook.worksheets.getItem(allSheets[i].name);
        var usedRange = ws.getUsedRange();
        usedRange.load("columnCount");
        await context.sync();

        sheetsWithMeta.push({
          name: allSheets[i].name,
          headerRow: allSheets[i].headerRow,
          columnCount: usedRange.columnCount,
          index: baseIndex
        });
        baseIndex++;
      }
    }

    // Step 2: Validate column consistency
    var validation = mergeSheetsUtils.validateColumnConsistency(sheetsWithMeta);
    if (!validation.valid) {
      throw new Error(validation.error);
    }

    // Step 3: Read data from each sheet
    var mergedData = [];
    var hasHeader = false;

    for (var j = 0; j < sheetsWithMeta.length; j++) {
      var sheetInfo = sheetsWithMeta[j];
      var worksheet = context.workbook.worksheets.getItem(sheetInfo.name);
      var usedRange = worksheet.getUsedRange();
      usedRange.load(["values", "rowCount", "columnCount"]);
      await context.sync();

      var data = usedRange.values;
      if (!data || data.length === 0) {
        throw new Error("工作表 '" + sheetInfo.name + "' 无可用数据");
      }

      var headerRowIndex = sheetInfo.headerRow > 0 ? sheetInfo.headerRow - 1 : -1;

      // First sheet: include header if exists
      if (j === 0) {
        if (headerRowIndex >= 0 && headerRowIndex < data.length) {
          mergedData.push(data[headerRowIndex]);
          hasHeader = true;
        }
        // Add data rows after header
        for (var r = headerRowIndex + 1; r < data.length; r++) {
          mergedData.push(data[r]);
        }
      } else {
        // Other sheets: skip header, add all data rows
        var startRow = headerRowIndex >= 0 ? headerRowIndex + 1 : 0;
        for (var r2 = startRow; r2 < data.length; r2++) {
          mergedData.push(data[r2]);
        }
      }
    }

    // Step 4: Check if "合并结果" sheet exists
    var existingSheets = context.workbook.worksheets;
    var targetSheet = existingSheets.getItemOrNullObject("合并结果");
    targetSheet.load("name");
    await context.sync();

    var targetName = "合并结果";
    var sheetExists = targetSheet.name === targetName;

    if (sheetExists) {
      var confirmAction = await showConfirmBox(
        "工作表 '" + targetName + "' 已存在，选择操作："
      );

      if (confirmAction === "cancel") {
        executeBtn.disabled = false;
        setStatus("状态：已取消", "idle");
        return;
      }

      if (confirmAction === "rename") {
        var allSheetNames = existingSheets.items.map(function (ws) {
          return ws.name;
        });
        targetName = mergeSheetsUtils.generateUniqueSheetName("合并结果", allSheetNames);
      }
      // If overwrite, use same name
    }

    // Step 5: Create or clear target sheet and write data
    var finalSheet;
    try {
      finalSheet = context.workbook.worksheets.getItem(targetName);
      finalSheet.delete();
      await context.sync();
    } catch (e) {
      // Sheet doesn't exist, will be created
    }

    finalSheet = context.workbook.worksheets.add(targetName);
    await context.sync();

    if (mergedData.length > 0) {
      var targetRange = finalSheet.getRange(
        "A1:" +
        mergeSheetsUtils.getColumnLetter(mergedData[0].length - 1) +
        mergedData.length
      );
      targetRange.values = mergedData;
      await context.sync();
    }

    var totalRows = mergedData.length;
    var headerInfo = hasHeader ? "（含表头）" : "（无表头）";
    setStatus(
      "完成! 已合并 " + selectedSheets.length + " 个工作表，共 " + totalRows + " 行数据" + headerInfo,
      "success"
    );
  }).catch(function (error) {
    setStatus("错误: " + error.message, "error");
  }).finally(function () {
    executeBtn.disabled = false;
  });
}
```

- [ ] **Step 2: 提交**

```bash
git add src/taskpane/merge-sheets-taskpane.js
git commit -m "feat: 添加多Sheet合并工具主逻辑"
```

---

## Task 5: 更新 webpack.config.js 添加入口和插件

**Files:**
- Modify: `webpack.config.js`

- [ ] **Step 1: 添加 entry**

在 `entry` 对象的 `data-cleaning-taskpane` 后添加:

```javascript
"merge-sheets-taskpane": [
  "./src/taskpane/merge-sheets-taskpane.js",
  "./src/taskpane/merge-sheets-taskpane.html",
],
```

- [ ] **Step 2: 添加 HtmlWebpackPlugin**

在 `data-cleaning-taskpane.html` HtmlWebpackPlugin 配置后添加:

```javascript
new HtmlWebpackPlugin({
  filename: "merge-sheets-taskpane.html",
  template: "./src/taskpane/merge-sheets-taskpane.html",
  chunks: ["polyfill", "merge-sheets-taskpane"],
}),
```

- [ ] **Step 3: 提交**

```bash
git add webpack.config.js
git commit -m "feat: 添加多Sheet合并工具webpack配置"
```

---

## Task 6: 更新 manifest.xml 添加按钮和资源

**Files:**
- Modify: `manifest.xml`

- [ ] **Step 1: 在 DataProcessingGroup 中添加按钮**

在 `<Control xsi:type="Button" id="DataCleaningButton">...</Control>` 之后添加:

```xml
<Control xsi:type="Button" id="MergeSheetsButton">
  <Label resid="MergeSheetsButton.Label"/>
  <Supertip>
    <Title resid="MergeSheetsButton.Label"/>
    <Description resid="MergeSheetsButton.Tooltip"/>
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="MergeSheetsIcon.16x16"/>
    <bt:Image size="32" resid="MergeSheetsIcon.32x32"/>
    <bt:Image size="80" resid="MergeSheetsIcon.80x80"/>
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <TaskpaneId>MergeSheetsTaskpaneId</TaskpaneId>
    <SourceLocation resid="MergeSheetsTaskpane.Url"/>
  </Action>
</Control>
```

- [ ] **Step 2: 添加资源URL**

在 `<bt:Url id="DataCleaningTaskpane.Url"...>` 后添加:

```xml
<bt:Url id="MergeSheetsTaskpane.Url" DefaultValue="https://localhost:3000/merge-sheets-taskpane.html"/>
```

- [ ] **Step 3: 添加图标资源**

在 `<bt:Image id="DataCleaningIcon.80x80"...>` 后添加:

```xml
<bt:Image id="MergeSheetsIcon.16x16" DefaultValue="https://localhost:3000/assets/merge-sheets-16.png"/>
<bt:Image id="MergeSheetsIcon.32x32" DefaultValue="https://localhost:3000/assets/merge-sheets-32.png"/>
<bt:Image id="MergeSheetsIcon.80x80" DefaultValue="https://localhost:3000/assets/merge-sheets-80.png"/>
```

- [ ] **Step 4: 添加短字符串**

在 `<bt:String id="DataCleaningButton.Label"...>` 后添加:

```xml
<bt:String id="MergeSheetsButton.Label" DefaultValue="合并 Sheet"/>
```

- [ ] **Step 5: 添加长字符串**

在 `<bt:String id="DataCleaningButton.Tooltip"...>` 后添加:

```xml
<bt:String id="MergeSheetsButton.Tooltip" DefaultValue="将多个工作表合并为一个新工作表，支持自定义表头行号"/>
```

- [ ] **Step 6: 提交**

```bash
git add manifest.xml
git commit -m "feat: 在manifest中添加多Sheet合并工具按钮"
```

---

## Task 7: 添加图标资源

**Files:**
- Create: `assets/merge-sheets-16.png`
- Create: `assets/merge-sheets-32.png`
- Create: `assets/merge-sheets-80.png`

- [ ] **Step 1: 复制现有图标作为占位符**

```bash
cp assets/split-sheet-16.png assets/merge-sheets-16.png
cp assets/split-sheet-32.png assets/merge-sheets-32.png
cp assets/split-sheet-80.png assets/merge-sheets-80.png
```

- [ ] **Step 2: 提交**

```bash
git add assets/merge-sheets-*.png
git commit -m "feat: 添加多Sheet合并工具图标资源"
```

---

## Task 8: 验证构建

- [ ] **Step 1: 运行构建**

```bash
npm run build
```

预期: 构建成功，merge-sheets-taskpane.html 生成到 dist 目录

- [ ] **Step 2: 验证 manifest.xml 语法**

```bash
npm run validate
```

预期: 验证通过

---

## 自检清单

- [ ] 所有 checkbox 能正常勾选/取消
- [ ] 表头行号输入框默认值为 1，支持 0
- [ ] 选中少于 2 个 Sheet 时提示正确
- [ ] 列数不一致时提示正确
- [ ] "合并结果"Sheet 存在时弹出确认框
- [ ] 覆盖/重命名/取消三种操作正常
- [ ] 合并结果只保留第一个 Sheet 的表头
- [ ] 所有 Sheet 表头行号都为 0 时，合并结果也无表头
- [ ] 状态提示清晰显示进度和结果