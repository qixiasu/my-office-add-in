# 跨文件 Sheet 合并功能实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 实现跨文件 Sheet 合并功能——选择多个外部 Excel 文件，按统一表头行号合并，添加来源文件名标识列，写入当前工作簿新 sheet。

**Architecture:** 使用 SheetJS (xlsx) 库在浏览器端解析外部 Excel 文件，通过 Office JavaScript API 写入当前工作簿。新建独立 taskpane 页面和工具函数模块。

**Tech Stack:** SheetJS (xlsx) CDN, Office JavaScript API, 原生 JavaScript (ES5 兼容 IE11)

---

## 文件结构

```
新增文件:
- src/taskpane/cross-file-merge-taskpane.html   # 任务面板页面
- src/taskpane/cross-file-merge-taskpane.js     # 任务面板逻辑
- src/taskpane/cross-file-merge-taskpane.css   # 样式（复用 merge-sheets 风格）
- src/utils/cross-file-merge-utils.js           # 工具函数
- src/utils/cross-file-merge-utils.test.js      # 单元测试

修改文件:
- webpack.config.js                              # 添加入口和 HtmlWebpackPlugin
- manifest.xml                                   # 添加按钮和资源 URL
```

---

## Task 1: 创建工具函数模块 cross-file-merge-utils.js

**文件:**
- 创建: `src/utils/cross-file-merge-utils.js`

### 步骤

- [ ] **Step 1: 编写基础测试框架**

```javascript
// src/utils/cross-file-merge-utils.test.js
var crossFileMergeUtils = require("../utils/cross-file-merge-utils");

describe("crossFileMergeUtils", function() {

  describe("getColumnLetter", function() {
    it("returns A for column 0", function() {
      expect(crossFileMergeUtils.getColumnLetter(0)).toBe("A");
    });
    it("returns Z for column 25", function() {
      expect(crossFileMergeUtils.getColumnLetter(25)).toBe("Z");
    });
    it("returns AA for column 26", function() {
      expect(crossFileMergeUtils.getColumnLetter(26)).toBe("AA");
    });
  });

  describe("generateUniqueSheetName", function() {
    it("returns baseName when not in existing names", function() {
      expect(crossFileMergeUtils.generateUniqueSheetName("合并结果", ["Sheet1"]))
        .toBe("合并结果");
    });
    it("returns baseName_1 when baseName exists", function() {
      expect(crossFileMergeUtils.generateUniqueSheetName("合并结果", ["合并结果"]))
        .toBe("合并结果_1");
    });
    it("returns baseName_2 when baseName and _1 exist", function() {
      expect(crossFileMergeUtils.generateUniqueSheetName("合并结果", ["合并结果", "合并结果_1"]))
        .toBe("合并结果_2");
    });
  });

});
```

- [ ] **Step 2: 运行测试验证失败**

Run: `npm test -- --testPathPattern=cross-file-merge-utils`
Expected: FAIL — module not found

- [ ] **Step 3: 编写最小实现**

```javascript
// src/utils/cross-file-merge-utils.js
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
 * 生成唯一的 Sheet 名称
 * @param {string} baseName - 基础名称
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

/**
 * 使用 SheetJS 解析 Excel 文件的当前激活 sheet
 * @param {File} file - Browser File 对象
 * @returns {Promise<{data: Array, sheetName: string}>}
 */
function parseExcelFile(file) {
  return new Promise(function(resolve, reject) {
    var reader = new FileReader();
    reader.onload = function(e) {
      try {
        var data = new Uint8Array(e.target.result);
        var workbook = XLSX.read(data, { type: "array", bookSheets: true });
        var sheetName = workbook.SheetNames[0]; // 当前激活 sheet
        var sheet = workbook.Sheets[sheetName];
        var jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 }); // 2D array
        resolve({ data: jsonData, sheetName: sheetName });
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = function(e) {
      reject(new Error("文件读取失败"));
    };
    reader.readAsArrayBuffer(file);
  });
}

/**
 * 验证所有文件的列数是否一致
 * @param {Array<{name: string, columnCount: number}>} files
 * @returns {{valid: boolean, error: string|null}}
 */
function validateColumnConsistency(files) {
  if (!files || files.length < 2) {
    return { valid: false, error: "请至少选择两个文件" };
  }
  var expected = files[0].columnCount;
  for (var i = 1; i < files.length; i++) {
    if (files[i].columnCount !== expected) {
      return {
        valid: false,
        error: "文件 '" + files[i].name + "' 列数(" + files[i].columnCount + ")与第一个文件(" + expected + ")不一致"
      };
    }
  }
  return { valid: true, error: null };
}

/**
 * 合并多个文件的数据
 * @param {Array<{data: Array, name: string}>} fileDataList - 文件数据列表
 * @param {number} headerRowNumber - 表头行号（0=无表头）
 * @returns {{mergedData: Array, columnCount: number}}
 */
function mergeExcelData(fileDataList, headerRowNumber) {
  var mergedData = [];
  var hasHeader = headerRowNumber > 0;
  var columnCount = 0;

  for (var i = 0; i < fileDataList.length; i++) {
    var fileData = fileDataList[i];
    var data = fileData.data;
    var fileName = fileData.name;

    if (!data || data.length === 0) {
      continue; // 跳过空文件
    }

    var headerRowIndex = hasHeader ? headerRowNumber - 1 : -1;

    // 第一文件：取表头
    if (i === 0) {
      if (hasHeader && headerRowIndex >= 0 && headerRowIndex < data.length) {
        var headerRow = data[headerRowIndex].slice();
        headerRow.unshift("来源文件"); // 添加来源文件列
        mergedData.push(headerRow);
        columnCount = headerRow.length;
      }
      // 添加第一文件的数据行
      var startRow = hasHeader ? headerRowIndex + 1 : 0;
      for (var r = startRow; r < data.length; r++) {
        var dataRow = data[r].slice();
        dataRow.unshift(fileName);
        mergedData.push(dataRow);
      }
    } else {
      // 后续文件：跳过表头（如果存在）
      var startRow = hasHeader ? headerRowIndex + 1 : 0;
      for (var r2 = startRow; r2 < data.length; r2++) {
        var dataRow2 = data[r2].slice();
        dataRow2.unshift(fileName);
        mergedData.push(dataRow2);
      }
    }
  }

  return { mergedData: mergedData, columnCount: columnCount };
}

module.exports = {
  getColumnLetter: getColumnLetter,
  generateUniqueSheetName: generateUniqueSheetName,
  parseExcelFile: parseExcelFile,
  validateColumnConsistency: validateColumnConsistency,
  mergeExcelData: mergeExcelData
};
```

- [ ] **Step 4: 运行测试验证通过**

Run: `npm test -- --testPathPattern=cross-file-merge-utils`
Expected: PASS

- [ ] **Step 5: 提交**

```bash
git add src/utils/cross-file-merge-utils.js src/utils/cross-file-merge-utils.test.js
git commit -m "feat: add cross-file-merge-utils module

Co-Authored-By: Claude <noreply@anthropic.com>"
```

---

## Task 2: 创建任务面板页面 cross-file-merge-taskpane.html

**文件:**
- 创建: `src/taskpane/cross-file-merge-taskpane.html`

### 步骤

- [ ] **Step 1: 编写 HTML 文件**

```html
<!-- Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT License. -->

<!DOCTYPE html>
<html>

<head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>跨文件合并</title>

    <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    <script type="text/javascript" src="https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js"></script>
    <link rel="stylesheet" href="https://res-1.cdn.office.net/files/fabric-cdn-prod_20230815.002/office-ui-fabric-core/11.1.0/css/fabric.min.css"/>
    <link href="taskpane.css" rel="stylesheet" type="text/css" />
    <link href="cross-file-merge-taskpane.css" rel="stylesheet" type="text/css" />
</head>

<body class="ms-font-m ms-welcome ms-Fabric">
    <div class="cfm-container">
        <h1 class="cfm-title">
            <svg width="20" height="20" viewBox="0 0 20 20" fill="#0078d4">
                <path d="M16 10a1 1 0 0 1-1 1H5a1 1 0 0 1 0-2h10a1 1 0 0 1 1 1z"/>
                <path d="M8 5a1 1 0 0 1 1 1v3.586l2.707 2.707A1 1 0 0 1 12 14.414V15a1 1 0 0 1-2 0v-1a1 1 0 0 0-1-1H4a1 1 0 0 0-1 1v1a1 1 0 0 1-2 0V9a1 1 0 0 1 1-1h5z"/>
            </svg>
            跨文件合并
        </h1>

        <div class="guide-card">
            <p>选择多个 Excel 文件，将每个文件的当前激活 sheet 合并到一个新 sheet 中。</p>
        </div>

        <div class="cfm-file-section">
            <input type="file" id="fileInput" multiple accept=".xlsx,.xls" style="display: none;" />
            <button id="selectFilesBtn" class="cfm-button">
                选择 Excel 文件
            </button>
        </div>

        <div id="fileList" class="cfm-file-list">
            <div class="cfm-empty">未选择文件</div>
        </div>

        <div class="cfm-header-setting">
            <label for="headerRowInput">表头行号：</label>
            <input type="number" id="headerRowInput" value="1" min="0" max="100" />
            <span class="cfm-hint">（0=无表头）</span>
        </div>

        <button id="executeBtn" class="cfm-button cfm-button--primary" disabled>执行合并</button>

        <div id="status" role="status" aria-live="polite" class="cfm-status cfm-status--idle">
            状态：等待操作...
        </div>

        <div id="successModal" class="cfm-modal" style="display: none;">
            <div class="cfm-modal-content">
                <div class="cfm-modal-icon">✅</div>
                <div id="successMessage" class="cfm-modal-message"></div>
                <button id="closeModalBtn" class="cfm-button">确定</button>
            </div>
        </div>
    </div>
</body>

</html>
```

- [ ] **Step 2: 提交**

```bash
git add src/taskpane/cross-file-merge-taskpane.html
git commit -m "feat: add cross-file-merge-taskpane.html

Co-Authored-By: Claude <noreply@anthropic.com>"
```

---

## Task 3: 创建任务面板样式 cross-file-merge-taskpane.css

**文件:**
- 创建: `src/taskpane/cross-file-merge-taskpane.css`

### 步骤

- [ ] **Step 1: 编写样式文件**

复用 merge-sheets-taskpane.css 风格：

```css
/* Cross-File Merge Taskpane Styles */

.cfm-container {
  padding: 20px;
}

.cfm-title {
  font-size: 16px;
  font-weight: 700;
  color: #0078d4;
  margin-bottom: 16px;
  display: flex;
  align-items: center;
  gap: 8px;
}

.cfm-title svg {
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

.cfm-file-section {
  margin-bottom: 16px;
}

.cfm-file-list {
  max-height: 250px;
  overflow-y: auto;
  border: 1px solid #edebe9;
  border-radius: 6px;
  margin-bottom: 16px;
}

.cfm-file-item {
  display: flex;
  align-items: center;
  padding: 10px 12px;
  border-bottom: 1px solid #edebe9;
}

.cfm-file-item:last-child {
  border-bottom: none;
}

.cfm-file-item:hover {
  background: #f5f5f5;
}

.cfm-file-name {
  flex: 1;
  font-size: 14px;
  color: #333;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.cfm-file-remove {
  background: transparent;
  border: none;
  color: #0078d4;
  cursor: pointer;
  font-size: 14px;
  padding: 4px 8px;
  border-radius: 4px;
}

.cfm-file-remove:hover {
  background: #f0f0f0;
}

.cfm-empty {
  padding: 20px;
  text-align: center;
  color: #999;
  font-size: 13px;
}

.cfm-header-setting {
  display: flex;
  align-items: center;
  gap: 8px;
  margin-bottom: 16px;
}

.cfm-header-setting label {
  font-size: 14px;
  color: #333;
}

#headerRowInput {
  width: 60px;
  padding: 6px 8px;
  border: 1px solid #edebe9;
  border-radius: 4px;
  font-size: 13px;
  text-align: center;
}

#headerRowInput:focus {
  border-color: #0078d4;
  outline: none;
}

.cfm-hint {
  font-size: 12px;
  color: #666;
}

.cfm-button {
  padding: 8px 16px;
  background: transparent;
  color: #0078d4;
  border: 2px solid #0078d4;
  border-radius: 6px;
  font-size: 14px;
  font-weight: 600;
  cursor: pointer;
  transition: all 0.2s;
}

.cfm-button:hover {
  background: #0078d4;
  color: white;
}

.cfm-button:disabled {
  background: #f5f5f5;
  color: #999;
  border-color: #999;
  cursor: not-allowed;
}

.cfm-button--primary {
  width: 100%;
}

.cfm-status {
  padding: 10px;
  border-radius: 6px;
  font-size: 13px;
  margin-top: 12px;
}

.cfm-status--idle {
  background: #f5f5f5;
  color: #666;
}

.cfm-status--success {
  background: #d4edda;
  border: 1px solid #c3e6cb;
  color: #155724;
}

.cfm-status--error {
  background: #f8d7da;
  border: 1px solid #f5c6cb;
  color: #721c24;
}

.cfm-status--loading {
  background: #fff3cd;
  border: 1px solid #ffeeba;
  color: #856404;
}

/* Success Modal */
.cfm-modal {
  position: fixed;
  top: 0;
  left: 0;
  width: 100%;
  height: 100%;
  background: rgba(0, 0, 0, 0.5);
  display: flex;
  align-items: center;
  justify-content: center;
  z-index: 1000;
}

.cfm-modal-content {
  background: white;
  padding: 32px;
  border-radius: 12px;
  text-align: center;
  max-width: 400px;
  box-shadow: 0 4px 20px rgba(0, 0, 0, 0.15);
}

.cfm-modal-icon {
  font-size: 48px;
  margin-bottom: 16px;
}

.cfm-modal-message {
  font-size: 14px;
  color: #333;
  line-height: 1.6;
  margin-bottom: 24px;
}
```

- [ ] **Step 2: 提交**

```bash
git add src/taskpane/cross-file-merge-taskpane.css
git commit -m "feat: add cross-file-merge-taskpane.css

Co-Authored-By: Claude <noreply@anthropic.com>"
```

---

## Task 4: 创建任务面板逻辑 cross-file-merge-taskpane.js

**文件:**
- 创建: `src/taskpane/cross-file-merge-taskpane.js`

### 步骤

- [ ] **Step 1: 编写核心逻辑**

```javascript
/* global console, document, Excel, Office, XLSX */

var crossFileMergeUtils = require("../utils/cross-file-merge-utils");

var selectedFiles = [];
var headerRowNumber = 1;

function initOffice() {
  Office.onReady(function(info) {
    if (info.host === Office.HostType.Excel) {
      initEventListeners();
    }
  });
}

if (typeof Office !== "undefined") {
  initOffice();
} else {
  var checkOffice = setInterval(function() {
    if (typeof Office !== "undefined") {
      clearInterval(checkOffice);
      initOffice();
    }
  }, 50);
  setTimeout(function() { clearInterval(checkOffice); }, 10000);
}

function initEventListeners() {
  document.getElementById("selectFilesBtn").onclick = function() {
    document.getElementById("fileInput").click();
  };

  document.getElementById("fileInput").onchange = onFileSelected;
  document.getElementById("headerRowInput").onchange = onHeaderRowChanged;
  document.getElementById("executeBtn").onclick = executeMerge;
  document.getElementById("closeModalBtn").onclick = closeSuccessModal;
}

function onFileSelected(e) {
  var files = e.target.files;
  if (!files || files.length === 0) return;

  selectedFiles = Array.from(files).map(function(file) {
    return { file: file, name: file.name };
  });

  renderFileList();
  updateExecuteButton();
}

function onHeaderRowChanged(e) {
  headerRowNumber = parseInt(e.target.value, 10) || 0;
  if (headerRowNumber < 0) headerRowNumber = 0;
  if (headerRowNumber > 100) headerRowNumber = 100;
}

function renderFileList() {
  var listEl = document.getElementById("fileList");

  if (selectedFiles.length === 0) {
    listEl.innerHTML = '<div class="cfm-empty">未选择文件</div>';
    return;
  }

  var html = "";
  selectedFiles.forEach(function(item, index) {
    html +=
      '<div class="cfm-file-item" data-index="' + index + '">' +
      '<span class="cfm-file-name" title="' + escapeHtml(item.name) + '">' +
      escapeHtml(item.name) +
      '</span>' +
      '<button class="cfm-file-remove" data-index="' + index + '">移除</button>' +
      '</div>';
  });

  listEl.innerHTML = html;

  // Bind remove buttons
  var removeButtons = listEl.querySelectorAll(".cfm-file-remove");
  removeButtons.forEach(function(btn) {
    btn.onclick = function() {
      var idx = parseInt(btn.dataset.index, 10);
      selectedFiles.splice(idx, 1);
      renderFileList();
      updateExecuteButton();
    };
  });
}

function escapeHtml(text) {
  var div = document.createElement("div");
  div.textContent = text;
  return div.innerHTML;
}

function updateExecuteButton() {
  var btn = document.getElementById("executeBtn");
  btn.disabled = selectedFiles.length < 2;
}

function setStatus(message, type) {
  var el = document.getElementById("status");
  el.textContent = message;
  el.className = "cfm-status cfm-status--" + (type || "idle");
}

function showSuccessModal(message) {
  document.getElementById("successMessage").innerHTML = message;
  document.getElementById("successModal").style.display = "flex";
}

function closeSuccessModal() {
  document.getElementById("successModal").style.display = "none";
}

function executeMerge() {
  var executeBtn = document.getElementById("executeBtn");
  executeBtn.disabled = true;
  setStatus("处理中...", "loading");

  if (selectedFiles.length < 2) {
    executeBtn.disabled = false;
    setStatus("请至少选择两个文件", "error");
    return;
  }

  headerRowNumber = parseInt(document.getElementById("headerRowInput").value, 10) || 0;

  // 步骤1：解析所有 Excel 文件
  var fileDataList = [];
  var parsePromises = selectedFiles.map(function(item) {
    return crossFileMergeUtils.parseExcelFile(item.file)
      .then(function(result) {
        fileDataList.push({
          data: result.data,
          name: item.name,
          columnCount: result.data.length > 0 ? result.data[0].length : 0
        });
        setStatus("已解析 " + fileDataList.length + "/" + selectedFiles.length + " 个文件...", "loading");
      })
      .catch(function(err) {
        throw new Error("解析文件 '" + item.name + "' 失败: " + err.message);
      });
  });

  Promise.all(parsePromises)
    .then(function() {
      // 步骤2：验证列数一致性
      var validation = crossFileMergeUtils.validateColumnConsistency(fileDataList);
      if (!validation.valid) {
        throw new Error(validation.error);
      }

      // 步骤3：合并数据
      var mergeResult = crossFileMergeUtils.mergeExcelData(fileDataList, headerRowNumber);

      if (mergeResult.mergedData.length === 0) {
        throw new Error("没有可合并的数据");
      }

      // 步骤4：写入当前工作簿
      return Excel.run(function(context) {
        // 获取所有 sheet 名称
        var worksheets = context.workbook.worksheets;
        worksheets.load("items/name");
        return context.sync().then(function() {
          var existingNames = worksheets.items.map(function(ws) { return ws.name; });
          var targetSheetName = crossFileMergeUtils.generateUniqueSheetName("合并结果", existingNames);

          // 创建新 sheet
          var newSheet = worksheets.add(targetSheetName);
          context.sync();

          // 写入数据
          var columnLetter = crossFileMergeUtils.getColumnLetter(mergeResult.columnCount - 1);
          var dataRange = newSheet.getRange("A1:" + columnLetter + mergeResult.mergedData.length);
          dataRange.values = mergeResult.mergedData;
          context.sync();

          return {
            sheetName: targetSheetName,
            fileCount: fileDataList.length,
            rowCount: mergeResult.mergedData.length,
            columnCount: mergeResult.columnCount
          };
        });
      });
    })
    .then(function(result) {
      setStatus("完成！已合并 " + result.fileCount + " 个文件，共 " + result.rowCount + " 行数据", "success");

      var message = "✅ 合并成功！<br><br>" +
        "共合并 " + result.fileCount + " 个文件，" + result.rowCount + " 行数据（" + result.columnCount + " 列）<br>" +
        "输出工作表：" + result.sheetName;
      showSuccessModal(message);
    })
    .catch(function(error) {
      setStatus("错误: " + error.message, "error");
    })
    .finally(function() {
      executeBtn.disabled = false;
    });
}
```

- [ ] **Step 2: 提交**

```bash
git add src/taskpane/cross-file-merge-taskpane.js
git commit -m "feat: add cross-file-merge-taskpane.js

Co-Authored-By: Claude <noreply@anthropic.com>"
```

---

## Task 5: 更新 webpack.config.js

**文件:**
- 修改: `webpack.config.js`

### 步骤

- [ ] **Step 1: 添加入口和 HtmlWebpackPlugin 配置**

在 `entry` 对象的 `merge-sheets-taskpane` 条目后添加：

```javascript
"cross-file-merge-taskpane": [
  "./src/taskpane/cross-file-merge-taskpane.js",
  "./src/taskpane/cross-file-merge-taskpane.html",
],
```

在 `plugins` 数组中，`merge-sheets-taskpane` 的 HtmlWebpackPlugin 后添加：

```javascript
new HtmlWebpackPlugin({
  filename: "cross-file-merge-taskpane.html",
  template: "./src/taskpane/cross-file-merge-taskpane.html",
  chunks: ["polyfill", "cross-file-merge-taskpane"],
}),
```

- [ ] **Step 2: 验证构建**

Run: `npm run build:dev`
Expected: 成功构建，无错误

- [ ] **Step 3: 提交**

```bash
git add webpack.config.js
git commit -m "feat: add cross-file-merge-taskpane webpack entry

Co-Authored-By: Claude <noreply@anthropic.com>"
```

---

## Task 6: 更新 manifest.xml

**文件:**
- 修改: `manifest.xml`

### 步骤

- [ ] **Step 1: 添加按钮和资源 URL**

在 `VersionOverrides` 中找到现有按钮定义，参考 `MergeSheetsButton` 添加新按钮：

```xml
<Control xsi:type="Button" id="CrossFileMergeButton">
  <Label resid="CrossFileMergeButton.Label"/>
  <Supertip>
    <Title resid="CrossFileMergeButton.Label"/>
    <Description resid="CrossFileMergeButton.Tooltip"/>
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="CrossFileMergeIcon.16x16"/>
    <bt:Image size="32" resid="CrossFileMergeIcon.32x32"/>
    <bt:Image size="80" resid="CrossFileMergeIcon.80x80"/>
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <TaskpaneId>CrossFileMergeTaskpaneId</TaskpaneId>
    <SourceLocation resid="CrossFileMergeTaskpane.Url"/>
  </Action>
</Control>
```

在 `Resources` 部分添加：

```xml
<bt:String id="CrossFileMergeButton.Label" DefaultValue="跨文件合并"/>
<bt:String id="CrossFileMergeButton.Tooltip" DefaultValue="选择多个 Excel 文件，合并它们的当前激活 sheet"/>

<bt:Url id="CrossFileMergeTaskpane.Url" DefaultValue="https://localhost:3000/cross-file-merge-taskpane.html"/>

<bt:Image id="CrossFileMergeIcon.16x16" DefaultValue="https://localhost:3000/assets/tools-32.png"/>
<bt:Image id="CrossFileMergeIcon.32x32" DefaultValue="https://localhost:3000/assets/tools-32.png"/>
<bt:Image id="CrossFileMergeIcon.80x80" DefaultValue="https://localhost:3000/assets/tools-32.png"/>
```

- [ ] **Step 2: 提交**

```bash
git add manifest.xml
git commit -m "feat: add CrossFileMerge button to manifest

Co-Authored-By: Claude <noreply@anthropic.com>"
```

---

## Task 7: 单元测试覆盖

**文件:**
- 修改: `src/utils/cross-file-merge-utils.test.js`

### 步骤

- [ ] **Step 1: 添加 parseExcelFile 和 mergeExcelData 的测试**

```javascript
describe("parseExcelFile", function() {
  it("should reject for invalid file", function(done) {
    crossFileMergeUtils.parseExcelFile(null)
      .then(function() {
        done(new Error("Should have rejected"));
      })
      .catch(function(err) {
        done();
      });
  });
});

describe("validateColumnConsistency", function() {
  it("returns invalid for less than 2 files", function() {
    var result = crossFileMergeUtils.validateColumnConsistency([
      { name: "a.xlsx", columnCount: 3 }
    ]);
    expect(result.valid).toBe(false);
  });

  it("returns valid when column counts match", function() {
    var result = crossFileMergeUtils.validateColumnConsistency([
      { name: "a.xlsx", columnCount: 3 },
      { name: "b.xlsx", columnCount: 3 }
    ]);
    expect(result.valid).toBe(true);
  });

  it("returns invalid when column counts differ", function() {
    var result = crossFileMergeUtils.validateColumnConsistency([
      { name: "a.xlsx", columnCount: 3 },
      { name: "b.xlsx", columnCount: 5 }
    ]);
    expect(result.valid).toBe(false);
    expect(result.error).toContain("不一致");
  });
});

describe("mergeExcelData", function() {
  it("merges two files with header row 1", function() {
    var fileDataList = [
      { data: [["Name", "Age"], ["Alice", 25], ["Bob", 30]], name: "a.xlsx" },
      { data: [["Name", "Age"], ["Carol", 35]], name: "b.xlsx" }
    ];
    var result = crossFileMergeUtils.mergeExcelData(fileDataList, 1);
    expect(result.mergedData.length).toBe(3); // 1 header + 2 data rows
    expect(result.mergedData[0][0]).toBe("来源文件");
    expect(result.mergedData[1][0]).toBe("a.xlsx");
    expect(result.mergedData[2][0]).toBe("b.xlsx");
  });

  it("merges two files with no header", function() {
    var fileDataList = [
      { data: [["Alice", 25], ["Bob", 30]], name: "a.xlsx" },
      { data: [["Carol", 35]], name: "b.xlsx" }
    ];
    var result = crossFileMergeUtils.mergeExcelData(fileDataList, 0);
    expect(result.mergedData.length).toBe(3);
    expect(result.mergedData[0][0]).toBe("a.xlsx");
  });
});
```

- [ ] **Step 2: 运行所有测试**

Run: `npm test -- --testPathPattern=cross-file-merge`
Expected: ALL PASS

- [ ] **Step 3: 提交**

```bash
git add src/utils/cross-file-merge-utils.test.js
git commit -m "test: add cross-file-merge-utils unit tests

Co-Authored-By: Claude <noreply@anthropic.com>"
```

---

## Task 8: 集成测试

### 步骤

- [ ] **Step 1: 启动开发服务器**

Run: `npm run dev-server`

- [ ] **Step 2: 测试完整流程**

1. 在 Excel 中打开"跨文件合并"任务面板
2. 选择 2 个 Excel 文件
3. 设置表头行号
4. 点击"执行合并"
5. 验证结果 sheet 中的数据和来源文件列

- [ ] **Step 3: 测试错误处理**

1. 选择列数不一致的两个文件
2. 验证错误提示

---

## 依赖项检查

- SheetJS CDN: `https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js`
- 无需添加到 package.json（通过 CDN 加载）
