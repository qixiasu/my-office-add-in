# 增强查找侧边面板实施计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 将增强查找从对话框迁移到侧边面板，与连接列/导入CSV保持一致的UI风格

**Architecture:** 新建 vlookup-taskpane.html/js/css 三个文件，修改 manifest.xml 将按钮从 ExecuteFunction 改为 ShowTaskpane，修改 webpack.config.js 添加入口

**Tech Stack:** Office JavaScript API, Excel.run, webpack

---

## 文件变更总览

| 操作 | 文件路径 |
|------|---------|
| 新建 | `src/taskpane/vlookup-taskpane.html` |
| 新建 | `src/taskpane/vlookup-taskpane.css` |
| 新建 | `src/taskpane/vlookup-taskpane.js` |
| 修改 | `manifest.xml` — EnhancedVlookupButton 从 ExecuteFunction 改为 ShowTaskpane |
| 修改 | `webpack.config.js` — 添加 vlookup-taskpane 入口和 HtmlWebpackPlugin |
| 删除 | `src/commands/vlookup-dialog.html` — 迁移后废弃 |
| 修改 | `src/commands/commands.html` — 移除 enhancedVlookup action |

---

## Task 1: 创建 vlookup-taskpane.css 样式文件

**Files:**
- Create: `src/taskpane/vlookup-taskpane.css`

- [ ] **Step 1: 创建样式文件**

```css
/* 增强查找侧边栏样式 - 与连接列/导入CSV统一风格 */

.vlookup-container {
  padding: 16px;
}

.vlookup-title {
  font-size: 16px;
  font-weight: 700;
  color: #0078d4;
  margin-bottom: 16px;
  display: flex;
  align-items: center;
  gap: 8px;
}

.vlookup-title svg {
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

.vlookup-form-group {
  margin-bottom: 14px;
}

.vlookup-label {
  display: block;
  font-size: 13px;
  color: #666;
  margin-bottom: 6px;
}

.vlookup-input {
  width: 100%;
  padding: 10px 12px;
  border: 2px solid #0078d4;
  border-radius: 6px;
  font-size: 15px;
  outline: none;
  box-sizing: border-box;
  transition: box-shadow 0.2s;
}

.vlookup-input:focus {
  box-shadow: 0 0 0 3px rgba(0,120,212,0.15);
}

.vlookup-input[readonly] {
  background: #f5f5f5;
}

.vlookup-input-row {
  display: flex;
  gap: 8px;
  align-items: center;
}

.vlookup-input-row .vlookup-input {
  flex: 1;
}

.vlookup-btn {
  padding: 6px 12px;
  background: transparent;
  color: #0078d4;
  border: 2px solid #0078d4;
  border-radius: 6px;
  font-size: 13px;
  font-weight: 600;
  cursor: pointer;
  white-space: nowrap;
  transition: all 0.2s;
}

.vlookup-btn:hover {
  background: #0078d4;
  color: white;
}

.vlookup-btn:active {
  background: #005a9e;
  border-color: #005a9e;
}

.vlookup-btn:disabled {
  background: #f5f5f5;
  color: #999;
  border-color: #999;
  cursor: not-allowed;
}

.vlookup-section-title {
  font-size: 13px;
  font-weight: 600;
  color: #333;
  margin-bottom: 8px;
  margin-top: 16px;
}

.vlookup-checkbox-group {
  border: 1px solid #e0e0e0;
  border-radius: 6px;
  padding: 10px;
  max-height: 120px;
  overflow-y: auto;
}

.vlookup-checkbox-group label {
  display: inline-block;
  margin-right: 16px;
  margin-bottom: 4px;
  font-size: 13px;
  cursor: pointer;
}

.vlookup-checkbox-group label input {
  margin-right: 4px;
}

.vlookup-radio-group {
  margin-top: 8px;
}

.vlookup-radio-group label {
  display: inline-block;
  margin-right: 16px;
  font-size: 13px;
  cursor: pointer;
}

.vlookup-radio-group label input {
  margin-right: 4px;
}

.vlookup-preview {
  margin-top: 16px;
  border: 1px solid #e0e0e0;
  border-radius: 6px;
  max-height: 200px;
  overflow: auto;
  display: none;
}

.vlookup-preview.visible {
  display: block;
}

.vlookup-preview table {
  border-collapse: collapse;
  font-size: 12px;
  width: 100%;
}

.vlookup-preview th,
.vlookup-preview td {
  border: 1px solid #ddd;
  padding: 6px 10px;
  text-align: left;
}

.vlookup-preview th {
  background: #f5f5f5;
  font-weight: 600;
  position: sticky;
  top: 0;
}

.vlookup-btn-row {
  display: flex;
  gap: 8px;
  justify-content: flex-end;
  margin-top: 20px;
}

.vlookup-button {
  padding: 10px 20px;
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

.vlookup-button:disabled {
  opacity: 0.5;
  cursor: not-allowed;
}

.vlookup-button--primary {
  background: #0078d4;
  color: white;
}

.vlookup-button--primary:hover {
  background: #005a9e;
}

.vlookup-button--secondary {
  background: transparent;
  color: #666;
  border-color: #999;
}

.vlookup-button--secondary:hover {
  background: #f0f0f0;
  color: #333;
  border-color: #999;
}

.status-message {
  padding: 10px;
  border-radius: 6px;
  font-size: 13px;
  margin-top: 12px;
  text-align: center;
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

- [ ] **Step 2: Commit**

```bash
git add src/taskpane/vlookup-taskpane.css
git commit -m "feat(vlookup): add vlookup-taskpane.css with unified styling"
```

---

## Task 2: 创建 vlookup-taskpane.html 页面

**Files:**
- Create: `src/taskpane/vlookup-taskpane.html`

- [ ] **Step 1: 创建 HTML 页面**

```html
<!-- Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT License. -->

<!DOCTYPE html>
<html>

<head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>增强查找</title>

    <!-- Office JavaScript API -->
    <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>

    <!-- For more information on Fluent UI, visit https://developer.microsoft.com/fluentui#/. -->
    <link rel="stylesheet" href="https://res-1.cdn.office.net/files/fabric-cdn-prod_20230815.002/office-ui-fabric-core/11.1.0/css/fabric.min.css"/>

    <!-- Template styles -->
    <link href="taskpane.css" rel="stylesheet" type="text/css" />
    <link href="vlookup-taskpane.css" rel="stylesheet" type="text/css" />
</head>

<body class="ms-font-m ms-welcome ms-Fabric">
    <div class="vlookup-container">
        <!-- Header -->
        <h1 class="vlookup-title">
            <svg width="20" height="20" viewBox="0 0 20 20" fill="#0078d4">
                <path d="M8 2a6 6 0 1 0 0 12A6 6 0 0 0 8 2zm0 10.5a4.5 4.5 0 1 1 0-9 4.5 4.5 0 0 1 0 9zm.75-4.25v2.5l2 1.25-.75 1.25-2.5-1.5V8.25h1.25z"/>
            </svg>
            增强查找
        </h1>

        <!-- Guide -->
        <div class="guide-card">
            <p>用 INDEX/MATCH 一次完成多列查找，匹配列可在任意位置</p>
        </div>

        <!-- Form -->
        <div class="vlookup-form-group">
            <label for="lookupValue" class="vlookup-label">查找值区域</label>
            <div class="vlookup-input-row">
                <input type="text" id="lookupValue" class="vlookup-input" readonly placeholder="点击刷新按钮获取" />
                <button id="refreshLookupValueBtn" class="vlookup-btn">刷新选择</button>
            </div>
        </div>

        <div class="vlookup-form-group">
            <label for="lookupRange" class="vlookup-label">查找表区域</label>
            <div class="vlookup-input-row">
                <input type="text" id="lookupRange" class="vlookup-input" placeholder="例如: Sheet1!A1:D500 或 A1:D500" />
                <button id="refreshSelectionBtn" class="vlookup-btn">刷新选择</button>
            </div>
        </div>

        <div class="vlookup-form-group">
            <label for="headerRow" class="vlookup-label">表头行号</label>
            <div class="vlookup-input-row">
                <input type="number" id="headerRow" class="vlookup-input" value="1" min="1" style="width: 100px;" />
                <button id="readHeadersBtn" class="vlookup-btn">读取表头</button>
            </div>
        </div>

        <div class="vlookup-form-group">
            <label for="matchCol" class="vlookup-label">匹配列</label>
            <select id="matchCol" class="vlookup-input" style="width: 100%;"></select>
        </div>

        <div class="vlookup-section-title">返回列（勾选需要的列）</div>
        <div class="vlookup-form-group" style="margin-top: 8px;">
            <button id="selectAllBtn" class="vlookup-btn" style="padding: 4px 10px; font-size: 12px;">全选</button>
            <button id="deselectAllBtn" class="vlookup-btn" style="padding: 4px 10px; font-size: 12px;">取消全选</button>
            <div class="vlookup-checkbox-group" id="returnCols"></div>
        </div>

        <div class="vlookup-section-title">匹配模式</div>
        <div class="vlookup-radio-group">
            <label><input type="radio" name="matchMode" value="0" checked /> 精确匹配</label>
            <label><input type="radio" name="matchMode" value="1" /> 近似匹配</label>
        </div>

        <div class="vlookup-section-title" style="margin-top: 12px;">输出类型</div>
        <div class="vlookup-radio-group">
            <label><input type="radio" name="outputType" value="formula" checked /> 公式</label>
            <label><input type="radio" name="outputType" value="static" /> 静态值（大数据推荐）</label>
        </div>

        <!-- Preview -->
        <div class="vlookup-preview" id="previewSection">
            <table id="previewTable">
                <thead id="previewHead"></thead>
                <tbody id="previewBody"></tbody>
            </table>
        </div>

        <!-- Buttons -->
        <div class="vlookup-btn-row">
            <button id="cancelBtn" class="vlookup-button vlookup-button--secondary">取消</button>
            <button id="execBtn" class="vlookup-button vlookup-button--primary" disabled>执行查找</button>
        </div>

        <!-- Status -->
        <div id="status" class="status-message status-idle">状态：等待操作...</div>
    </div>
</body>

</html>
```

- [ ] **Step 2: Commit**

```bash
git add src/taskpane/vlookup-taskpane.html
git commit -m "feat(vlookup): add vlookup-taskpane.html"
```

---

## Task 3: 创建 vlookup-taskpane.js 逻辑

**Files:**
- Create: `src/taskpane/vlookup-taskpane.js`

- [ ] **Step 1: 创建 JS 逻辑文件**

```javascript
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

var MAX_ROWS = 1050000;

var {
  parseRangeAddress,
  buildColRange,
  buildIndexMatchFormula,
  staticLookup,
} = require("../utils/vlookup-utils");

var { getColumnLetter } = require("../utils/concat-utils");

Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    initEventListeners();
    loadInitialSelection();
  }
});

function initEventListeners() {
  document.getElementById("refreshSelectionBtn").onclick = function () {
    refreshSelection("lookupRange");
  };
  document.getElementById("refreshLookupValueBtn").onclick = function () {
    refreshSelection("lookupValue");
  };
  document.getElementById("readHeadersBtn").onclick = readHeaders;
  document.getElementById("selectAllBtn").onclick = function () {
    var cbs = document.querySelectorAll("#returnCols input[type='checkbox']");
    for (var i = 0; i < cbs.length; i++) { cbs[i].checked = true; }
  };
  document.getElementById("deselectAllBtn").onclick = function () {
    var cbs = document.querySelectorAll("#returnCols input[type='checkbox']");
    for (var i = 0; i < cbs.length; i++) { cbs[i].checked = false; }
  };
  document.getElementById("lookupRange").oninput = function () {
    clearTimeout(window._refreshColumnsTimer);
    window._refreshColumnsTimer = setTimeout(refreshColumns, 300);
  };
  document.getElementById("headerRow").onchange = refreshColumns;
  document.getElementById("cancelBtn").onclick = function () {
    Office.context.ui.close();
  };
  document.getElementById("execBtn").onclick = executeLookup;
}

function loadInitialSelection() {
  Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    range.load("address");
    return context.sync().then(function () {
      if (range.address) {
        document.getElementById("lookupRange").value = range.address;
        document.getElementById("lookupValue").value = extractFirstColAddress(range.address);
        refreshColumns();
      }
    });
  }).catch(function () {});
}

function refreshSelection(target) {
  setStatus("正在更新选区...", "loading");
  Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    range.load("address");
    return context.sync().then(function () {
      var addr = range.address || "";
      if (target === "lookupRange") {
        document.getElementById("lookupRange").value = addr;
        if (!document.getElementById("lookupValue").value.trim()) {
          document.getElementById("lookupValue").value = extractFirstColAddress(addr);
        }
        refreshColumns();
      } else {
        document.getElementById("lookupValue").value = addr;
      }
      setStatus("就绪", "idle");
    });
  }).catch(function (error) {
    setStatus("错误: " + error.message, "error");
  });
}

function refreshColumns() {
  var rangeStr = document.getElementById("lookupRange").value.trim();
  var matchColEl = document.getElementById("matchCol");
  var returnColsEl = document.getElementById("returnCols");
  var execBtn = document.getElementById("execBtn");

  var oldMatchCol = matchColEl.value;
  var oldChecked = {};
  var oldCheckboxes = document.querySelectorAll("#returnCols input[type='checkbox']");
  for (var ci = 0; ci < oldCheckboxes.length; ci++) {
    oldChecked[oldCheckboxes[ci].value] = oldCheckboxes[ci].checked;
  }

  matchColEl.innerHTML = "";
  returnColsEl.innerHTML = "";

  if (!rangeStr) {
    execBtn.disabled = true;
    return;
  }

  var parsed;
  try {
    parsed = parseRangeAddress(rangeStr);
  } catch (e) {
    execBtn.disabled = true;
    return;
  }

  if (parsed.colCount < 1) {
    execBtn.disabled = true;
    return;
  }

  for (var i = 0; i < parsed.colCount; i++) {
    var letter = getColumnLetter(parsed.startCol + i);
    var opt = document.createElement("option");
    opt.value = i;
    opt.textContent = letter;
    matchColEl.appendChild(opt);

    var label = document.createElement("label");
    var cb = document.createElement("input");
    cb.type = "checkbox";
    cb.value = i;
    cb.id = "retCol_" + i;
    if (oldChecked[i] === true) { cb.checked = true; }
    label.htmlFor = "retCol_" + i;
    label.appendChild(cb);
    label.appendChild(document.createTextNode(" " + letter));
    returnColsEl.appendChild(label);
  }

  if (oldMatchCol !== undefined && oldMatchCol !== "" && matchColEl.options.length > 0) {
    var oldVal = parseInt(oldMatchCol, 10);
    if (oldVal >= 0 && oldVal < parsed.colCount) {
      matchColEl.value = oldMatchCol;
    }
  }

  execBtn.disabled = false;
}

function readHeaders() {
  var rangeStr = document.getElementById("lookupRange").value.trim();
  var headerRow = parseInt(document.getElementById("headerRow").value, 10) || 1;

  if (!rangeStr) {
    setStatus("请先输入查找表区域", "error");
    return;
  }

  var parsed;
  try {
    parsed = parseRangeAddress(rangeStr);
  } catch (e) {
    setStatus("区域地址格式无效", "error");
    return;
  }

  var headerRowOffset = headerRow - parsed.startRow;
  if (headerRowOffset < 0 || headerRowOffset >= parsed.rowCount) {
    setStatus("表头行号超出查找表范围", "error");
    return;
  }

  var readBtn = document.getElementById("readHeadersBtn");
  readBtn.disabled = true;
  setStatus("正在读取表头...", "loading");

  Excel.run(function (context) {
    var sheet;
    if (parsed.sheet) {
      sheet = context.workbook.worksheets.getItem(parsed.sheet);
    } else {
      sheet = context.workbook.worksheets.getActiveWorksheet();
    }

    var headerRange = sheet.getRangeByIndexes(headerRow - 1, parsed.startCol, 1, parsed.colCount);
    headerRange.load("values");
    return context.sync().then(function () {
      var headers = headerRange.values[0];
      var capturedStartCol = parsed.startCol;
      updateHeaderLabels(headers, capturedStartCol);

      var previewRows = Math.min(5, parsed.rowCount);
      var previewRange = sheet.getRangeByIndexes(parsed.startRow - 1, parsed.startCol, previewRows, parsed.colCount);
      previewRange.load("values");
      return context.sync().then(function () {
        showPreview(previewRange.values, headers);
        setStatus("表头读取成功，共 " + parsed.colCount + " 列", "success");
        readBtn.disabled = false;
      });
    });
  }).catch(function (error) {
    setStatus("读取失败: " + error.message, "error");
    readBtn.disabled = false;
  });
}

function updateHeaderLabels(headers, startCol) {
  var labels = document.querySelectorAll("#returnCols label");
  for (var i = 0; i < labels.length && i < headers.length; i++) {
    var headerText = headers[i];
    if (headerText === null || headerText === undefined) headerText = "";
    headerText = String(headerText);
    var letter = getColumnLetter(startCol + i);
    var textNode = labels[i].childNodes[1];
    if (textNode) {
      textNode.textContent = " " + letter + " - " + (headerText || "(空)");
    }
  }
}

function showPreview(values, headers) {
  var section = document.getElementById("previewSection");
  section.classList.add("visible");

  var thead = document.getElementById("previewHead");
  var tbody = document.getElementById("previewBody");
  thead.innerHTML = "";
  tbody.innerHTML = "";

  var headerRow = document.createElement("tr");
  for (var i = 0; i < headers.length; i++) {
    var th = document.createElement("th");
    th.textContent = headers[i] || "";
    headerRow.appendChild(th);
  }
  thead.appendChild(headerRow);

  for (var r = 0; r < values.length; r++) {
    var tr = document.createElement("tr");
    for (var c = 0; c < values[r].length; c++) {
      var td = document.createElement("td");
      var v = values[r][c];
      if (v === null || v === undefined) v = "";
      td.textContent = String(v);
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }
}

function getCheckedReturnCols() {
  var checkboxes = document.querySelectorAll("#returnCols input[type='checkbox']");
  var indices = [];
  for (var i = 0; i < checkboxes.length; i++) {
    if (checkboxes[i].checked) {
      indices.push(parseInt(checkboxes[i].value, 10));
    }
  }
  return indices;
}

function extractFirstColAddress(rangeStr) {
  var parsed = parseRangeAddress(rangeStr);
  var firstColLetter = getColumnLetter(parsed.startCol);
  var prefix = parsed.sheet ? parsed.sheet + "!" : "";
  return prefix + "$" + firstColLetter + "$" + parsed.startRow + ":$" + firstColLetter + "$" + parsed.endRow;
}

function setStatus(message, type) {
  var el = document.getElementById("status");
  el.textContent = message;
  el.className = "status-message status-" + (type || "idle");
}

function executeLookup() {
  var statusEl = document.getElementById("status");
  var execBtn = document.getElementById("execBtn");

  var rangeStr = document.getElementById("lookupRange").value.trim();
  if (!rangeStr) {
    setStatus("请输入查找表区域", "error");
    return;
  }

  var parsed;
  try {
    parsed = parseRangeAddress(rangeStr);
  } catch (e) {
    setStatus("区域地址格式无效", "error");
    return;
  }

  var returnCols = getCheckedReturnCols();
  if (returnCols.length === 0) {
    setStatus("请至少勾选一列作为返回列", "error");
    return;
  }

  var matchColIndex = parseInt(document.getElementById("matchCol").value, 10);
  var matchMode = parseInt(document.querySelector("input[name='matchMode']:checked").value, 10);
  var outputType = document.querySelector("input[name='outputType']:checked").value;

  if (isNaN(matchColIndex)) {
    setStatus("请选择匹配列", "error");
    return;
  }

  if (parsed.rowCount > MAX_ROWS) {
    setStatus("数据量过大（" + parsed.rowCount + " 行），最多支持 " + MAX_ROWS + " 行", "error");
    return;
  }

  execBtn.disabled = true;
  setStatus("正在处理...", "loading");

  var lookupValueStr = document.getElementById("lookupValue").value.trim();
  var config = {
    type: "vlookup",
    outputType: outputType,
    matchMode: matchMode,
    lookupRange: rangeStr,
    lookupValue: lookupValueStr,
    parsed: parsed,
    matchColIndex: matchColIndex,
    returnColIndices: returnCols,
    headerRow: parseInt(document.getElementById("headerRow").value, 10) || 1,
  };

  Excel.run(function (context) {
    var resultCount = performLookup(context, config);
    return context.sync().then(function () {
      setStatus("完成! 已写入 " + resultCount + " 列数据", "success");
      execBtn.disabled = false;
    });
  }).catch(function (error) {
    setStatus("错误: " + error.message, "error");
    execBtn.disabled = false;
  });
}

function performLookup(context, config) {
  var sheet;
  if (config.parsed.sheet) {
    sheet = context.workbook.worksheets.getItem(config.parsed.sheet);
  } else {
    sheet = context.workbook.worksheets.getActiveWorksheet();
  }

  var lookupValueRange = sheet.getRange(config.lookupValue);
  lookupValueRange.load(["values", "rowCount", "columnCount"]);
  var lookupTableRange = sheet.getRange(config.lookupRange);
  lookupTableRange.load(["values", "rowCount", "columnCount"]);

  return context.sync().then(function () {
    var lookupValues = lookupValueRange.values;
    var lookupTable = lookupTableRange.values;
    var resultCount = config.returnColIndices.length;

    if (config.outputType === "formula") {
      var flatValues = [];
      for (var r = 0; r < lookupValues.length; r++) {
        for (var c = 0; c < resultCount; c++) {
          var lookupCellRef = "$" + getColumnLetter(lookupValueRange.columnIndex) + "$" + (lookupValueRange.rowIndex + r + 1);
          var lookupColRange = buildColRange(config.parsed, config.matchColIndex);
          var returnColRange = buildColRange(config.parsed, config.returnColIndices[c]);
          var formula = buildIndexMatchFormula(lookupCellRef, lookupColRange, returnColRange, config.matchMode);
          flatValues.push([formula]);
        }
      }

      var targetStartRow = lookupValueRange.rowIndex + 1;
      var targetStartCol = config.parsed.startCol + config.parsed.colCount;
      var targetRange = sheet.getRangeByIndexes(targetStartRow, targetStartCol, lookupValues.length, resultCount);
      targetRange.values = flatValues;
    } else {
      var staticResults = staticLookup(lookupValues, lookupTable, config.matchColIndex, config.returnColIndices, config.matchMode);
      var flatStatic = [];
      for (var sr = 0; sr < staticResults.length; sr++) {
        for (var sc = 0; sc < staticResults[sr].length; sc++) {
          flatStatic.push([staticResults[sr][sc]]);
        }
      }

      var staticTargetStartRow = lookupValueRange.rowIndex + 1;
      var staticTargetStartCol = config.parsed.startCol + config.parsed.colCount;
      var staticTargetRange = sheet.getRangeByIndexes(staticTargetStartRow, staticTargetStartCol, staticResults.length, resultCount);
      staticTargetRange.values = flatStatic;
    }

    return resultCount;
  });
}
```

- [ ] **Step 2: Commit**

```bash
git add src/taskpane/vlookup-taskpane.js
git commit -m "feat(vlookup): add vlookup-taskpane.js with taskpane logic"
```

---

## Task 4: 修改 webpack.config.js 添加 vlookup-taskpane 入口

**Files:**
- Modify: `webpack.config.js`

- [ ] **Step 1: 在 entry 对象中添加 vlookup-taskpane**

```javascript
"vlookup-taskpane": ["./src/taskpane/vlookup-taskpane.js", "./src/taskpane/vlookup-taskpane.html"],
```

- [ ] **Step 2: 在 plugins 中添加 HtmlWebpackPlugin**

```javascript
new HtmlWebpackPlugin({
  filename: "vlookup-taskpane.html",
  template: "./src/taskpane/vlookup-taskpane.html",
  chunks: ["polyfill", "vlookup-taskpane"],
}),
```

- [ ] **Step 3: Commit**

```bash
git add webpack.config.js
git commit -m "feat(vlookup): add vlookup-taskpane entry to webpack"
```

---

## Task 5: 修改 manifest.xml 将 EnhancedVlookupButton 改为 ShowTaskpane

**Files:**
- Modify: `manifest.xml`

- [ ] **Step 1: 将 ExecuteFunction 改为 ShowTaskpane**

将:
```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>enhancedVlookup</FunctionName>
</Action>
```

改为:
```xml
<Action xsi:type="ShowTaskpane">
  <TaskpaneId>VlookupTaskpaneId</TaskpaneId>
  <SourceLocation resid="VlookupTaskpane.Url"/>
</Action>
```

- [ ] **Step 2: 添加 Url 资源**

在 `<bt:Urls>` 中添加:
```xml
<bt:Url id="VlookupTaskpane.Url" DefaultValue="https://localhost:3000/vlookup-taskpane.html"/>
```

- [ ] **Step 3: Commit**

```bash
git add manifest.xml
git commit -m "feat(vlookup): convert EnhancedVlookupButton to ShowTaskpane"
```

---

## Task 6: 清理 commands.html 中的 enhancedVlookup action

**Files:**
- Modify: `src/commands/commands.html`

- [ ] **Step 1: 找到并删除 `Office.actions.associate("enhancedVlookup", function enhancedVlookup(event) {...}` 代码块**

- [ ] **Step 2: Commit**

```bash
git add src/commands/commands.html
git commit -m "feat(vlookup): remove enhancedVlookup ExecuteFunction from commands"
```

---

## Task 7: 删除废弃的 vlookup-dialog.html

**Files:**
- Delete: `src/commands/vlookup-dialog.html`

- [ ] **Step 1: 删除文件**

```bash
rm src/commands/vlookup-dialog.html
```

- [ ] **Step 2: Commit**

```bash
git add -A
git commit -m "feat(vlookup): remove deprecated vlookup-dialog.html"
```

---

## Task 8: 构建并测试

- [ ] **Step 1: 运行构建**

```bash
npm run build:dev
```

- [ ] **Step 2: 验证文件生成**

检查 `dist/vlookup-taskpane.html` 是否存在。

- [ ] **Step 3: 测试功能**

1. 点击"增强查找"按钮，侧边面板正常打开
2. 选中 Excel 中的查找表，点击"刷新选择"，输入框自动填入地址
3. 输入表头行号，点击"读取表头"，列名下拉和预览正常显示
4. 选择匹配列、返回列，点击"执行查找"，数据正确写入
5. 取消按钮关闭面板

- [ ] **Step 4: 提交测试**

```bash
git add -A
git commit -m "test(vlookup): verify vlookup taskpane works correctly"
```

---

## 自检清单

- [ ] 所有文件路径与 spec 一致
- [ ] vlookup-taskpane.js 正确依赖 vlookup-utils.js 和 concat-utils.js
- [ ] manifest.xml 中 VlookupTaskpane.Url 指向正确的 URL
- [ ] webpack.config.js 中 entry 和 HtmlWebpackPlugin 都已添加
- [ ] commands.html 中已移除 enhancedVlookup action
- [ ] vlookup-dialog.html 已删除