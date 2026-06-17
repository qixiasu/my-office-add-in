# 跨文件合并：手动选择 Sheet 实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 用户选中文件后，每个文件旁显示下拉框列出所有 sheet，默认选中第一个，用户可自由选择要合并的 sheet。

**Architecture:** 新增 `getSheetNames()` utility 函数；修改 `parseExcelFile()` 支持指定 sheet；修改 taskpane 的 `selectedFiles` 数据结构和渲染逻辑，文件选中后异步加载 sheet 列表并渲染下拉框。

**Tech Stack:** Plain JavaScript, SheetJS (XLSX), Jest

---

## 文件变更总览

| 文件 | 改动类型 |
|------|---------|
| `src/utils/cross-file-merge-utils.js` | 新增 `getSheetNames()`，修改 `parseExcelFile()` 支持可选 `sheetName` 参数 |
| `src/taskpane/cross-file-merge-taskpane.js` | 重构 `selectedFiles` 结构、`onFileSelected`、`renderFileList`、`executeMerge` |
| `src/taskpane/cross-file-merge-taskpane.css` | 新增 sheet 选择器相关样式 |
| `src/utils/cross-file-merge-utils.test.js` | 新增 `getSheetNames` 和 `parseExcelFile` 带 sheetName 参数的测试 |

---

## Task 1: 新增 getSheetNames 和修改 parseExcelFile

**Files:**
- Modify: `src/utils/cross-file-merge-utils.js`

- [ ] **Step 1: 新增 getSheetNames 函数（在 parseExcelFile 之前添加）**

在 `parseExcelFile` 函数之前添加：

```javascript
/**
 * 获取 Excel 文件的所有 sheet 名称
 * @param {File} file - Browser File 对象
 * @returns {Promise<string[]>} sheet 名称数组
 */
function getSheetNames(file) {
  return new Promise(function (resolve, reject) {
    var reader = new FileReader();
    reader.onload = function (e) {
      try {
        var data = new Uint8Array(e.target.result);
        var workbook = XLSX.read(data, { type: "array" });
        resolve(workbook.SheetNames);
      } catch (err) {
        reject(new Error("读取文件 sheet 失败: " + err.message));
      }
    };
    reader.onerror = function () {
      reject(new Error("文件读取失败"));
    };
    reader.readAsArrayBuffer(file);
  });
}
```

- [ ] **Step 2: 修改 parseExcelFile 支持可选 sheetName 参数**

将函数签名和内部逻辑从：

```javascript
function parseExcelFile(file) {
  return new Promise(function (resolve, reject) {
    var reader = new FileReader();
    reader.onload = function (e) {
      try {
        var data = new Uint8Array(e.target.result);
        var workbook = XLSX.read(data, { type: "array" });
        var sheetName = workbook.SheetNames[0]; // 当前激活 sheet
        var sheet = workbook.Sheets[sheetName];
        if (!sheet) {
          return reject(new Error("工作表 '" + sheetName + "' 不存在或为空"));
        }
        var jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        resolve({ data: jsonData, sheetName: sheetName });
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = function () {
      reject(new Error("文件读取失败"));
    };
    reader.readAsArrayBuffer(file);
  });
}
```

修改为：

```javascript
/**
 * 使用 SheetJS 解析 Excel 文件的指定 sheet
 * @param {File} file - Browser File 对象
 * @param {string} [sheetName] - 可选，指定 sheet 名称，不传则使用第一个
 * @returns {Promise<{data: Array, sheetName: string}>}
 */
function parseExcelFile(file, sheetName) {
  return new Promise(function (resolve, reject) {
    var reader = new FileReader();
    reader.onload = function (e) {
      try {
        var data = new Uint8Array(e.target.result);
        var workbook = XLSX.read(data, { type: "array" });
        var targetSheetName = sheetName || workbook.SheetNames[0];
        var sheet = workbook.Sheets[targetSheetName];
        if (!sheet) {
          return reject(new Error("工作表 '" + targetSheetName + "' 不存在或为空"));
        }
        var jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        resolve({ data: jsonData, sheetName: targetSheetName });
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = function () {
      reject(new Error("文件读取失败"));
    };
    reader.readAsArrayBuffer(file);
  });
}
```

- [ ] **Step 3: 更新 module.exports**

在 `module.exports` 中添加 `getSheetNames`：

```javascript
module.exports = {
  getColumnLetter: getColumnLetter,
  generateUniqueSheetName: generateUniqueSheetName,
  getSheetNames: getSheetNames,
  parseExcelFile: parseExcelFile,
  validateColumnConsistency: validateColumnConsistency,
  mergeExcelData: mergeExcelData,
};
```

- [ ] **Step 4: 运行测试验证**

```bash
npm test -- --testPathPatterns="cross-file-merge"
```

预期：原有测试仍通过（`parseExcelFile` 不传 sheetName 时行为不变）

- [ ] **Step 5: 提交**

```bash
git add src/utils/cross-file-merge-utils.js
git commit -m "feat(cross-file-merge): add getSheetNames and make parseExcelFile accept optional sheetName"
```

---

## Task 2: 修改 taskpane 数据结构和事件处理

**Files:**
- Modify: `src/taskpane/cross-file-merge-taskpane.js`

- [ ] **Step 1: 修改 selectedFiles 数据结构**

约第 5 行，`selectedFiles = []` 的声明保持不变，修改 `onFileSelected` 函数中的初始化逻辑。

找到 `onFileSelected` 函数（约第 39-49 行），将：

```javascript
selectedFiles = Array.from(files).map(function(file) {
  return { file: file, name: file.name };
});
```

修改为：

```javascript
selectedFiles = Array.from(files).map(function(file) {
  return {
    file: file,
    name: file.name,
    sheetName: null,
    sheetNames: [],
    loadingSheets: true,
    loadError: null
  };
});
```

- [ ] **Step 2: 添加异步获取 sheet 列表逻辑**

在同一 `onFileSelected` 函数末尾（`renderFileList(); updateExecuteButton();` 之后）添加：

```javascript
// 异步获取每个文件的 sheet 列表
selectedFiles.forEach(function(item, index) {
  crossFileMergeUtils.getSheetNames(item.file)
    .then(function(names) {
      item.sheetNames = names;
      item.sheetName = names[0]; // 默认选第一个
      item.loadingSheets = false;
      renderFileList(); // 重新渲染以显示下拉框
    })
    .catch(function(err) {
      item.sheetNames = [];
      item.sheetName = null;
      item.loadingSheets = false;
      item.loadError = err.message;
      renderFileList();
    });
});
```

- [ ] **Step 3: 修改 renderFileList 渲染 sheet 下拉框**

找到 `renderFileList` 函数（约第 57-88 行），将内部的 html 生成部分修改为：

```javascript
var html = "";
selectedFiles.forEach(function(item, index) {
  var sheetSelect = "";
  if (item.loadingSheets) {
    sheetSelect = '<span class="cfm-sheet-loading">加载中...</span>';
  } else if (item.loadError) {
    sheetSelect = '<span class="cfm-sheet-error" title="' + escapeHtml(item.loadError) + '">加载失败</span>';
  } else if (item.sheetNames.length === 0) {
    sheetSelect = '<span class="cfm-sheet-error">无 sheet</span>';
  } else {
    var options = item.sheetNames.map(function(name) {
      var selected = name === item.sheetName ? ' selected' : '';
      return '<option value="' + escapeHtml(name) + '"' + selected + '>' + escapeHtml(name) + '</option>';
    }).join('');
    sheetSelect = '<select class="cfm-sheet-select" data-index="' + index + '">' + options + '</select>';
  }

  html +=
    '<div class="cfm-file-item" data-index="' + index + '">' +
      '<span class="cfm-file-name" title="' + escapeHtml(item.name) + '">' + escapeHtml(item.name) + '</span>' +
      '<span class="cfm-file-sheet">' + sheetSelect + '</span>' +
      '<button class="cfm-file-remove" data-index="' + index + '">移除</button>' +
    '</div>';
});
```

然后在 `listEl.innerHTML = html;` 之后、移除按钮绑定之前，添加下拉框事件绑定：

```javascript
// 绑定下拉框 change 事件
var sheetSelects = listEl.querySelectorAll(".cfm-sheet-select");
sheetSelects.forEach(function(select) {
  select.onchange = function() {
    var idx = parseInt(select.dataset.index, 10);
    selectedFiles[idx].sheetName = select.value;
  };
});
```

- [ ] **Step 4: 修改 executeMerge 传递用户选择的 sheet**

找到 `executeMerge` 函数中 `parseExcelFile` 的调用处（约第 132 行），将：

```javascript
return crossFileMergeUtils.parseExcelFile(item.file)
```

修改为：

```javascript
return crossFileMergeUtils.parseExcelFile(item.file, item.sheetName)
```

同时将 `.catch` 中的错误消息从：

```javascript
throw new Error("解析文件 '" + item.name + "' 失败: " + err.message);
```

修改为：

```javascript
throw new Error("解析文件 '" + item.name + "' 的 Sheet '" + item.sheetName + "' 失败: " + err.message);
```

- [ ] **Step 5: 运行测试验证**

```bash
npm test -- --testPathPatterns="cross-file-merge"
```

预期：原有测试仍通过

- [ ] **Step 6: 提交**

```bash
git add src/taskpane/cross-file-merge-taskpane.js
git commit -m "feat(cross-file-merge): add sheet selector dropdown per file"
```

---

## Task 3: 添加 CSS 样式

**Files:**
- Modify: `src/taskpane/cross-file-merge-taskpane.css`

- [ ] **Step 1: 添加 sheet 选择器相关样式**

在文件末尾添加：

```css
/* Sheet 选择器 */
.cfm-file-item {
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 6px 8px;
  border-bottom: 1px solid #f0f0f0;
}

.cfm-file-name {
  flex: 1;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.cfm-file-sheet {
  flex-shrink: 0;
}

.cfm-sheet-select {
  padding: 2px 6px;
  border: 1px solid #d0d0d0;
  border-radius: 4px;
  font-size: 12px;
  min-width: 80px;
  max-width: 150px;
}

.cfm-sheet-loading {
  font-size: 12px;
  color: #888;
}

.cfm-sheet-error {
  font-size: 12px;
  color: #d83b01;
}
```

- [ ] **Step 2: 提交**

```bash
git add src/taskpane/cross-file-merge-taskpane.css
git commit -m "feat(cross-file-merge): add styles for sheet selector dropdown"
```

---

## Task 4: 单元测试

**Files:**
- Modify: `src/utils/cross-file-merge-utils.test.js`

- [ ] **Step 1: 添加 getSheetNames 测试**

在 `describe('getSheetNames')` block 中添加（如果文件已有多个 describe 块，可选择合适位置）：

```javascript
describe("getSheetNames", function () {
  beforeEach(function () {
    // Mock FileReader and XLSX
    global.FileReader = function () {};
    global.XLSX = {
      read: function (data, opts) {
        return {
          SheetNames: ["Sheet1", "Sheet2", "数据表"]
        };
      }
    };
  });

  it("returns sheet names array", function () {
    var file = { name: "test.xlsx" };
    return crossFileMergeUtils.getSheetNames(file).then(function (names) {
      expect(names).toEqual(["Sheet1", "Sheet2", "数据表"]);
    });
  });
});
```

- [ ] **Step 2: 添加 parseExcelFile 带 sheetName 参数的测试**

在 `describe('parseExcelFile')` 中新增测试用例：

```javascript
it("parses specified sheet when sheetName is provided", function () {
  var file = { name: "test.xlsx" };
  var mockSheet = { A1: { v: "test" } };
  global.XLSX = {
    read: function (data, opts) {
      return {
        SheetNames: ["Sheet1", "Sheet2"],
        Sheets: {
          "Sheet2": mockSheet
        }
      };
    }
  };
  global.XLSX.utils = {
    sheet_to_json: function (sheet, opts) {
      return [["A1"], ["B1"]];
    }
  };

  return crossFileMergeUtils.parseExcelFile(file, "Sheet2").then(function (result) {
    expect(result.sheetName).toBe("Sheet2");
    expect(result.data).toEqual([["A1"], ["B1"]]);
  });
});
```

- [ ] **Step 3: 运行所有测试验证**

```bash
npm test -- --testPathPatterns="cross-file-merge"
```

预期：所有测试通过（包括原有测试 + 新增测试）

- [ ] **Step 4: 提交**

```bash
git add src/utils/cross-file-merge-utils.test.js
git commit -m "test(cross-file-merge): add getSheetNames and parseExcelFile sheetName tests"
```

---

## Task 5: 最终集成验证

- [ ] **Step 1: 运行完整测试套件**

```bash
npm test -- --testPathPatterns="cross-file-merge"
```

- [ ] **Step 2: 检查 manifest.xml 中的按钮描述更新（如需要）**

当前描述"将每个文件的当前激活 sheet 合并"可更新为"将每个文件的指定 sheet 合并"，但这不是必选项。

预期：全部测试通过
