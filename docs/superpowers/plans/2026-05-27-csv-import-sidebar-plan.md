# CSV 导入侧边栏实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 将 CSV 导入功能从弹窗改为侧边栏面板，向导式 3 步骤流程，与连接列面板视觉风格一致

**Architecture:** 新建独立 taskpane（HTML/CSS/JS），通过 manifest.xml 修改触发方式为 ShowTaskpane，数据写入固定 A1 位置

**Tech Stack:** Office JavaScript API, Excel.run, FileReader API, RFC 4180 CSV parsing

---

## 文件结构

```
src/
├── taskpane/
│   ├── csv-import-taskpane.html   # 新建 — 向导 HTML 结构
│   ├── csv-import-taskpane.css    # 新建 — 样式（与连接列一致）
│   └── csv-import-taskpane.js    # 新建 — 步骤状态机 + Excel 写入
└── utils/
    └── csv-utils.js              # 复用（parseCSV）
```

---

## Task 1: 创建 csv-import-taskpane.html

**Files:**
- Create: `src/taskpane/csv-import-taskpane.html`

```html
<!-- Taskpane header -->
<div class="taskpane-header">
  <svg width="20" height="20" viewBox="0 0 20 20" fill="#0078d4">
    <path d="M14 3a1 1 0 0 1 1 1v5.586l2.707 2.707A1 1 0 0 1 17 14.414V15a1 1 0 0 1-2 0v-1a1 1 0 0 0-1-1H9a1 1 0 0 0-1 1v1a1 1 0 0 1-2 0v-.586a1 1 0 0 1 .293-.707L11 9.586V4a1 1 0 0 1 1-1h2z"/>
  </svg>
  <h1>导入 CSV 文件</h1>
</div>

<!-- Stepper indicator -->
<div class="stepper">
  <div class="step" id="step1">
    <div class="step-circle">1</div>
    <div class="step-label">选择文件</div>
  </div>
  <div class="step" id="step2">
    <div class="step-circle">2</div>
    <div class="step-label">设置分隔符</div>
  </div>
  <div class="step" id="step3">
    <div class="step-circle">3</div>
    <div class="step-label">确认导入</div>
  </div>
</div>

<!-- Step panels -->
<div class="step-content">
  <!-- Step 1: File selection -->
  <div id="panel1" class="step-panel">
    <div class="guide-card">
      <p>选择 CSV/TXT 文件，支持逗号/制表符分隔的数据文件</p>
    </div>
    <div class="form-group">
      <label>选择文件：</label>
      <input type="file" id="fileInput" accept=".csv,.tsv,.txt" />
      <div id="fileName" class="file-name"></div>
    </div>
  </div>

  <!-- Step 2: Delimiter setting -->
  <div id="panel2" class="step-panel" style="display:none;">
    <div class="guide-card">
      <p>设置分隔符，默认为逗号（,），支持任意单字符</p>
    </div>
    <div class="form-group">
      <label>分隔符：</label>
      <input type="text" id="delimiterInput" value="," maxlength="1" />
    </div>
  </div>

  <!-- Step 3: Confirmation -->
  <div id="panel3" class="step-panel" style="display:none;">
    <div class="guide-card">
      <p id="summaryText"></p>
    </div>
  </div>
</div>

<!-- Navigation buttons -->
<div class="btn-row">
  <button id="prevBtn" class="concat-button concat-button--secondary" disabled>上一步</button>
  <button id="nextBtn" class="concat-button">下一步</button>
</div>

<!-- Status message -->
<div id="status" class="status-message">状态：等待操作...</div>
```

---

## Task 2: 创建 csv-import-taskpane.css

**Files:**
- Create: `src/taskpane/csv-import-taskpane.css`

- [ ] **Step 1: Write the CSS**

```css
/* === Taskpane shell === */
.taskpane-header {
  display: flex;
  align-items: center;
  gap: 8px;
  margin-bottom: 16px;
}

.taskpane-header h1 {
  font-size: 16px;
  font-weight: 700;
  color: #0078d4;
  margin: 0;
}

/* === Stepper === */
.stepper {
  display: flex;
  margin-bottom: 20px;
  position: relative;
}

.step {
  flex: 1;
  text-align: center;
  font-size: 12px;
  color: #999;
  position: relative;
  z-index: 1;
}

.step::before {
  content: "";
  position: absolute;
  top: 10px;
  left: 50%;
  width: 100%;
  height: 2px;
  background: #ddd;
  z-index: 0;
}

.step:last-child::before {
  display: none;
}

.step.active {
  color: #0078d4;
  font-weight: 600;
}

.step.active::before,
.step.active ~ .step::before {
  background: #0078d4;
}

.step.done {
  color: #28a745;
}

.step.done::before {
  background: #28a745;
}

.step-circle {
  width: 20px;
  height: 20px;
  border-radius: 50%;
  background: #ddd;
  color: white;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  font-size: 11px;
  margin-bottom: 4px;
  position: relative;
  z-index: 2;
}

.step.active .step-circle {
  background: #0078d4;
}

.step.done .step-circle {
  background: #28a745;
}

/* === Step content === */
.step-content {
  background: white;
  border: 2px solid #0078d4;
  border-radius: 6px;
  padding: 16px;
  margin-bottom: 16px;
}

.step-panel {
  /* active panel shown by JS */
}

/* === Guide card === */
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

/* === Form controls === */
.form-group {
  margin-bottom: 12px;
}

.form-group label {
  display: block;
  font-size: 13px;
  color: #666;
  margin-bottom: 6px;
}

input[type="text"],
input[type="file"] {
  width: 100%;
  padding: 10px 12px;
  border: 2px solid #0078d4;
  border-radius: 6px;
  font-size: 15px;
  box-sizing: border-box;
}

input[type="text"]:focus {
  outline: none;
  box-shadow: 0 0 0 3px rgba(0, 120, 212, 0.2);
}

/* === File name === */
.file-name {
  font-size: 12px;
  color: #28a745;
  margin-top: 6px;
  min-height: 16px;
}

/* === Buttons === */
.concat-button {
  padding: 10px 20px;
  background: transparent;
  color: #0078d4;
  border: 2px solid #0078d4;
  border-radius: 6px;
  font-size: 14px;
  font-weight: 600;
  cursor: pointer;
  margin-left: 8px;
}

.concat-button:first-child {
  margin-left: 0;
}

.concat-button:hover {
  background: #0078d4;
  color: white;
}

.concat-button:disabled {
  opacity: 0.5;
  cursor: not-allowed;
}

.concat-button--primary {
  background: #0078d4;
  color: white;
}

.concat-button--primary:hover {
  background: #005a9e;
}

.concat-button--secondary {
  background: transparent;
}

.btn-row {
  display: flex;
  gap: 8px;
  justify-content: flex-end;
  margin-top: 16px;
}

/* === Status === */
.status-message {
  padding: 10px;
  border-radius: 6px;
  font-size: 13px;
  background: #f5f5f5;
  color: #666;
  margin-top: 12px;
  text-align: center;
}

.status-message.status-success {
  background: #d4edda;
  color: #155724;
  border: 1px solid #c3e6cb;
}

.status-message.status-error {
  background: #f8d7da;
  color: #721c24;
  border: 1px solid #f5c6cb;
}

.status-message.status-loading {
  background: #d1ecf1;
  color: #0c5460;
  border: 1px solid #bee5eb;
}
```

- [ ] **Step 2: Verify file created**

---

## Task 3: 创建 csv-import-taskpane.js

**Files:**
- Create: `src/taskpane/csv-import-taskpane.js`

- [ ] **Step 1: Write the JS**

```javascript
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

var MAX_ROWS = 1050000;

var csvUtils = require("../utils/csv-utils");

var currentStep = 1;
var selectedFile = null;
var parsedData = null;
var rowCount = 0;
var colCount = 0;

Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    var fileInput = document.getElementById("fileInput");
    var delimiterInput = document.getElementById("delimiterInput");
    var prevBtn = document.getElementById("prevBtn");
    var nextBtn = document.getElementById("nextBtn");

    fileInput.addEventListener("change", onFileSelected);
    delimiterInput.addEventListener("input", onDelimiterChanged);
    prevBtn.addEventListener("click", onPrev);
    nextBtn.addEventListener("click", onNext);
  }
});

function setStatus(message, type) {
  var el = document.getElementById("status");
  el.textContent = message;
  el.className = "status-message status-" + (type || "idle");
}

function updateStepper() {
  for (var i = 1; i <= 3; i++) {
    var step = document.getElementById("step" + i);
    var panel = document.getElementById("panel" + i);
    step.classList.remove("active", "done");
    if (i < currentStep) {
      step.classList.add("done");
    } else if (i === currentStep) {
      step.classList.add("active");
    }
    panel.style.display = i === currentStep ? "block" : "none";
  }
  document.getElementById("prevBtn").disabled = currentStep === 1;
  updateNextButton();
}

function updateNextButton() {
  var nextBtn = document.getElementById("nextBtn");
  var fileName = document.getElementById("fileName");
  if (currentStep === 1) {
    nextBtn.textContent = "下一步";
    nextBtn.disabled = !selectedFile;
  } else if (currentStep === 2) {
    nextBtn.textContent = "下一步";
    nextBtn.disabled = false;
  } else if (currentStep === 3) {
    nextBtn.textContent = "导入";
    nextBtn.disabled = false;
  }
}

function showPanel(step) {
  currentStep = step;
  updateStepper();
}

function onFileSelected(e) {
  var file = e.target.files[0];
  if (!file) return;

  selectedFile = file;
  document.getElementById("fileName").textContent = file.name;
  updateNextButton();

  var reader = new FileReader();
  reader.onload = function (event) {
    var delimiter = document.getElementById("delimiterInput").value || ",";
    try {
      parsedData = csvUtils.parseCSV(event.target.result, delimiter);
      rowCount = parsedData.length;
      colCount = parsedData.length > 0 ? parsedData[0].length : 0;
    } catch (err) {
      setStatus("错误: 文件读取失败", "error");
      selectedFile = null;
      document.getElementById("fileName").textContent = "";
      updateNextButton();
      return;
    }
    if (rowCount === 0) {
      setStatus("错误: 文件为空", "error");
      selectedFile = null;
      document.getElementById("fileName").textContent = "";
      updateNextButton();
      return;
    }
    // Auto-advance to step 2
    showPanel(2);
    setStatus("状态：等待操作...", "idle");
  };
  reader.onerror = function () {
    setStatus("错误: 文件读取失败", "error");
    selectedFile = null;
    document.getElementById("fileName").textContent = "";
    updateNextButton();
  };
  reader.readAsText(file);
}

function onDelimiterChanged() {
  if (currentStep === 2 && selectedFile) {
    // Re-parse with new delimiter
    var fileInput = document.getElementById("fileInput");
    if (fileInput.files[0]) {
      var reader = new FileReader();
      reader.onload = function (event) {
        var delimiter = document.getElementById("delimiterInput").value || ",";
        try {
          parsedData = csvUtils.parseCSV(event.target.result, delimiter);
          rowCount = parsedData.length;
          colCount = parsedData.length > 0 ? parsedData[0].length : 0;
        } catch (err) {
          // silent fail — user will see issue on step 3
        }
      };
      reader.readAsText(fileInput.files[0]);
    }
  }
}

function onPrev() {
  if (currentStep === 2) {
    showPanel(1);
  } else if (currentStep === 3) {
    showPanel(2);
  }
}

function onNext() {
  if (currentStep === 1) {
    showPanel(2);
  } else if (currentStep === 2) {
    updateSummary();
    showPanel(3);
  } else if (currentStep === 3) {
    doImport();
  }
}

function updateSummary() {
  var fileName = selectedFile ? selectedFile.name : "";
  var delimiter = document.getElementById("delimiterInput").value || ",";
  var delimiterLabel = delimiter === "\t" ? "Tab" : delimiter;
  var text = fileName + "，共 " + rowCount + " 行 × " + colCount + " 列，分隔符：【" + delimiterLabel + "】";
  document.getElementById("summaryText").textContent = text;
}

function doImport() {
  var nextBtn = document.getElementById("nextBtn");
  nextBtn.disabled = true;
  setStatus("处理中...", "loading");

  if (rowCount > MAX_ROWS) {
    setStatus("错误: 数据量过大（" + rowCount + " 行），单次最多支持 " + MAX_ROWS + " 行", "error");
    nextBtn.disabled = false;
    return;
  }

  Excel.run(function (context) {
    var worksheet = context.workbook.worksheets.getActiveWorksheet();
    var startCell = worksheet.getRange("A1");
    var endCell = worksheet.getRange(
      getColumnLetter(colCount - 1) + rowCount
    );
    var targetRange = worksheet.getRange(startCell.address + ":" + endCell.address);
    targetRange.values = parsedData;
    return context.sync().then(function () {
      setStatus("完成! 已写入 " + rowCount + " 行 × " + colCount + " 列", "success");
      nextBtn.disabled = false;
    });
  }).catch(function (error) {
    setStatus("错误: " + error.message, "error");
    nextBtn.disabled = false;
  });
}

function getColumnLetter(colIndex) {
  var letter = "";
  var remaining = colIndex;
  do {
    letter = String.fromCharCode(65 + (remaining % 26)) + letter;
    remaining = Math.floor(remaining / 26) - 1;
  } while (remaining >= 0);
  return letter;
}
```

- [ ] **Step 2: Verify file created**

---

## Task 4: 修改 manifest.xml — ImportCsvButton 改为 ShowTaskpane

**Files:**
- Modify: `manifest.xml`（找到 ImportCsvButton Control，将 ExecuteFunction 改为 ShowTaskpane）

- [ ] **Step 1: Read current ImportCsvButton control**

找到 ImportCsvButton 附近的 XML，提交 ChangeId / Action 改动

- [ ] **Step 2: Modify to ShowTaskpane + add bt:Url entry**

```xml
<!-- 1. 修改 Control 的 Action -->
<Action xsi:type="ShowTaskpane">
  <TaskpaneId>CSVImportTaskpaneId</TaskpaneId>
  <SourceLocation resid="CSVImportTaskpane.Url"/>
</Action>

<!-- 2. 在 <bt:Url id="Taskpane.Url"/> 附近添加新的 bt:Url -->
<bt:Url id="CSVImportTaskpane.Url" DefaultValue="https://localhost:3000/csv-import-taskpane.html"/>
```

---

## Task 5: 最终验证

**Files:**
- Read: `src/taskpane/csv-import-taskpane.html`（验证结构完整）
- Read: `src/taskpane/csv-import-taskpane.css`（验证样式匹配连接列）
- Read: `src/taskpane/csv-import-taskpane.js`（验证步骤逻辑）
- Read: `manifest.xml`（验证 ShowTaskpane + bt:Url）
- Read: `src/utils/csv-utils.js`（验证 parseCSV 导出）

- [ ] **Step 1: Verify HTML structure**
- [ ] **Step 2: Verify CSS visual consistency with concat-taskpane**
- [ ] **Step 3: Verify JS step state machine**
- [ ] **Step 4: Verify manifest.xml change**
- [ ] **Step 5: Run lint check**

Run: `npm run lint`
Expected: No new errors

- [ ] **Step 6: Commit**

```bash
git add src/taskpane/csv-import-taskpane.html src/taskpane/csv-import-taskpane.css src/taskpane/csv-import-taskpane.js manifest.xml
git commit -m "feat: add csv-import sidebar with wizard UI"
```