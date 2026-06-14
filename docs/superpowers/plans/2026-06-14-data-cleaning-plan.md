# 数据清洗工具 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 为 Excel 工具箱新增一个「数据清洗」任务面板，提供 5 个高频数据清理操作（修剪空格、删除空行、大小写转换、清除不可见字符、移除重复行）。

**Architecture:** 单一任务面板 + 分区布局。面板 JS 通过 Office JS API 读取选区数据，调用 utils 层纯函数在内存中执行清洗，再将结果一次性写回 Excel。撤销功能通过原始数据快照实现。

**Tech Stack:** JavaScript (ES5+, Office JS API), Jest, webpack + Babel (IE11 compatible)

---

### Task 1: 创建 utils 骨架 + 编写全部测试用例

**Files:**
- Create: `src/utils/data-cleaning-utils.js`
- Create: `src/utils/data-cleaning-utils.test.js`

- [ ] **Step 1: 创建 data-cleaning-utils.js（仅导出空函数）**

```javascript
// src/utils/data-cleaning-utils.js

/**
 * 修剪空格
 * @param {Array<Array>} values - 二维数组
 * @param {string} mode - 'both' | 'all' | 'leading' | 'trailing'
 * @returns {Array<Array>} 新二维数组
 */
function trimSpaces(values, mode) {
  // TODO: implement
  return values;
}

/**
 * 删除空行
 * @param {Array<Array>} values - 二维数组
 * @param {string} mode - 'all' | 'column' | 'ratio'
 * @param {number|null} columnIndex - 按列删除时指定的列索引
 * @param {number|null} ratioThreshold - 空值率阈值 (0-100)
 * @returns {Array<Array>} 新二维数组
 */
function removeEmptyRows(values, mode, columnIndex, ratioThreshold) {
  // TODO: implement
  return values;
}

/**
 * 大小写转换
 * @param {Array<Array>} values - 二维数组
 * @param {string} mode - 'upper' | 'lower' | 'capitalize' | 'camel'
 * @returns {Array<Array>} 新二维数组
 */
function convertCase(values, mode) {
  // TODO: implement
  return values;
}

/**
 * 清除不可见字符
 * @param {Array<Array>} values - 二维数组
 * @param {string} mode - 'control' | 'whitespace' | 'zero-width' | 'all'
 * @returns {Array<Array>} 新二维数组
 */
function removeInvisible(values, mode) {
  // TODO: implement
  return values;
}

/**
 * 移除重复行
 * @param {Array<Array>} values - 二维数组
 * @param {Array<number>|null} keyColumns - 依据列索引数组，null=所有列
 * @param {string} keep - 'first' | 'last'
 * @returns {Array<Array>} 新二维数组
 */
function removeDuplicates(values, keyColumns, keep) {
  // TODO: implement
  return values;
}

module.exports = {
  trimSpaces: trimSpaces,
  removeEmptyRows: removeEmptyRows,
  convertCase: convertCase,
  removeInvisible: removeInvisible,
  removeDuplicates: removeDuplicates,
};
```

- [ ] **Step 2: 编写全部测试用例**

```javascript
// src/utils/data-cleaning-utils.test.js

var utils = require("./data-cleaning-utils");

// ===== trimSpaces =====

test("trimSpaces removes leading and trailing spaces", function () {
  var input = [["  hello  "], ["  world  "]];
  var result = utils.trimSpaces(input, "both");
  expect(result[0][0]).toBe("hello");
  expect(result[1][0]).toBe("world");
});

test("trimSpaces removes all extra spaces", function () {
  var input = [["hello   world"], ["a    b"]];
  var result = utils.trimSpaces(input, "all");
  expect(result[0][0]).toBe("hello world");
  expect(result[1][0]).toBe("a b");
});

test("trimSpaces removes all spaces", function () {
  var input = [["hello world"], [" a b "]];
  var result = utils.trimSpaces(input, "leading");
  expect(result[0][0]).toBe("hello world");
  expect(result[1][0]).toBe("a b ");
});

test("trimSpaces trailing mode", function () {
  var input = [["hello  "], ["world "]];
  var result = utils.trimSpaces(input, "trailing");
  expect(result[0][0]).toBe("hello");
  expect(result[1][0]).toBe("world");
});

test("trimSpaces handles null and empty cells", function () {
  var input = [[null], [""], ["  test  "]];
  var result = utils.trimSpaces(input, "both");
  expect(result[0][0]).toBeNull();
  expect(result[1][0]).toBe("");
  expect(result[2][0]).toBe("test");
});

test("trimSpaces handles single cell", function () {
  var input = [["  only one  "]];
  var result = utils.trimSpaces(input, "both");
  expect(result[0][0]).toBe("only one");
});

test("trimSpaces does not affect non-string values", function () {
  var input = [[123], [true], [null]];
  var result = utils.trimSpaces(input, "both");
  expect(result[0][0]).toBe(123);
  expect(result[1][0]).toBe(true);
  expect(result[2][0]).toBeNull();
});

// ===== removeEmptyRows =====

test("removeEmptyRows removes completely empty rows", function () {
  var input = [
    ["a", "b"],
    [null, null],
    ["c", "d"],
    ["", ""],
  ];
  var result = utils.removeEmptyRows(input, "all", null, null);
  expect(result.length).toBe(2);
  expect(result[0][0]).toBe("a");
  expect(result[1][0]).toBe("c");
});

test("removeEmptyRows no empty rows returns same data", function () {
  var input = [["a"], ["b"], ["c"]];
  var result = utils.removeEmptyRows(input, "all", null, null);
  expect(result.length).toBe(3);
});

test("removeEmptyRows removes rows where specific column is empty", function () {
  var input = [
    ["a", "x"],
    ["b", null],
    ["c", "y"],
  ];
  var result = utils.removeEmptyRows(input, "column", 1, null);
  expect(result.length).toBe(2);
  expect(result[0][0]).toBe("a");
  expect(result[1][0]).toBe("c");
});

test("removeEmptyRows removes rows with empty ratio above threshold", function () {
  // 每行3列，阈值 50% → 空值 >=2 的行被删除
  var input = [
    ["a", "b", "c"],   // 0 empty → keep
    [null, "e", null], // 2 empty (66%) → remove
    ["g", null, null], // 2 empty (66%) → remove
    ["j", "k", null],  // 1 empty (33%) → keep
  ];
  var result = utils.removeEmptyRows(input, "ratio", null, 50);
  expect(result.length).toBe(2);
  expect(result[0][0]).toBe("a");
  expect(result[1][0]).toBe("j");
});

test("removeEmptyRows handles all rows empty", function () {
  var input = [[null, null], [null, null]];
  var result = utils.removeEmptyRows(input, "all", null, null);
  expect(result.length).toBe(0);
});

test("removeEmptyRows does not mutate original", function () {
  var input = [["a"], [null], ["b"]];
  var copy = [["a"], [null], ["b"]];
  utils.removeEmptyRows(input, "all", null, null);
  expect(input).toEqual(copy);
});

// ===== convertCase =====

test("convertCase converts to upper case", function () {
  var input = [["hello"], ["World"]];
  var result = utils.convertCase(input, "upper");
  expect(result[0][0]).toBe("HELLO");
  expect(result[1][0]).toBe("WORLD");
});

test("convertCase converts to lower case", function () {
  var input = [["HELLO"], ["World"]];
  var result = utils.convertCase(input, "lower");
  expect(result[0][0]).toBe("hello");
  expect(result[1][0]).toBe("world");
});

test("convertCase capitalizes first letter", function () {
  var input = [["hello world"], ["javaScript"]];
  var result = utils.convertCase(input, "capitalize");
  expect(result[0][0]).toBe("Hello world");
  expect(result[1][0]).toBe("JavaScript");
});

test("convertCase does not affect Chinese characters", function () {
  var input = [["hello 你好"], ["WORLD 世界"]];
  var result = utils.convertCase(input, "upper");
  expect(result[0][0]).toBe("HELLO 你好");
  expect(result[1][0]).toBe("WORLD 世界");
});

test("convertCase handles non-string values", function () {
  var input = [[123], [true], [null]];
  var result = utils.convertCase(input, "upper");
  expect(result[0][0]).toBe(123);
  expect(result[1][0]).toBe(true);
  expect(result[2][0]).toBeNull();
});

// ===== removeInvisible =====

test("removeInvisible removes control characters", function () {
  var input = [["hello\nworld"], ["tab\there"]];
  var result = utils.removeInvisible(input, "control");
  expect(result[0][0]).toBe("helloworld");
  expect(result[1][0]).toBe("tabhere");
});

test("removeInvisible removes zero-width characters", function () {
  var input = [["hello​world"], ["a‌b"]];
  var result = utils.removeInvisible(input, "zero-width");
  expect(result[0][0]).toBe("helloworld");
  expect(result[1][0]).toBe("ab");
});

test("removeInvisible removes all invisible characters", function () {
  var input = [["hello\nworld​!"]];
  var result = utils.removeInvisible(input, "all");
  expect(result[0][0]).toBe("helloworld!");
});

test("removeInvisible preserves normal text", function () {
  var input = [["hello world"], ["你好世界"]];
  var result = utils.removeInvisible(input, "all");
  expect(result[0][0]).toBe("hello world");
  expect(result[1][0]).toBe("你好世界");
});

test("removeInvisible handles null cells", function () {
  var input = [[null], ["hello"]];
  var result = utils.removeInvisible(input, "all");
  expect(result[0][0]).toBeNull();
  expect(result[1][0]).toBe("hello");
});

// ===== removeDuplicates =====

test("removeDuplicates removes duplicate rows based on all columns", function () {
  var input = [
    ["a", 1],
    ["b", 2],
    ["a", 1],
    ["c", 3],
  ];
  var result = utils.removeDuplicates(input, null, "first");
  expect(result.length).toBe(3);
});

test("removeDuplicates keeps first occurrence", function () {
  var input = [["a"], ["b"], ["a"], ["c"]];
  var result = utils.removeDuplicates(input, null, "first");
  expect(result[0][0]).toBe("a");
  expect(result[1][0]).toBe("b");
  expect(result[2][0]).toBe("c");
});

test("removeDuplicates keeps last occurrence", function () {
  var input = [["a", 1], ["b", 2], ["a", 99]];
  var result = utils.removeDuplicates(input, [0], "last");
  expect(result.length).toBe(2);
  // "a" 行保留最后一行的值
  expect(result[0][0]).toBe("b");
  expect(result[1][0]).toBe("a");
  expect(result[1][1]).toBe(99);
});

test("removeDuplicates with specific key columns", function () {
  var input = [
    ["a", "x", 1],
    ["a", "y", 2],  // col0 重复但 col1 不同 → 视为不同行
    ["b", "x", 3],
  ];
  var result = utils.removeDuplicates(input, [0], "first");
  expect(result.length).toBe(2); // 行0和行2（行1的col0重复）
});

test("removeDuplicates no duplicates returns same data", function () {
  var input = [["a"], ["b"], ["c"]];
  var result = utils.removeDuplicates(input, null, "first");
  expect(result.length).toBe(3);
});

test("removeDuplicates handles null values in key columns", function () {
  var input = [[null, "a"], [null, "b"]];
  var result = utils.removeDuplicates(input, [0], "first");
  // null 被视为相同值
  expect(result.length).toBe(1);
});

test("removeDuplicates does not mutate original", function () {
  var input = [["a"], ["b"], ["a"]];
  var copy = [["a"], ["b"], ["a"]];
  utils.removeDuplicates(input, null, "first");
  expect(input).toEqual(copy);
});
```

- [ ] **Step 3: 运行测试验证全部失败**

Run:
```bash
npx jest src/utils/data-cleaning-utils.test.js --verbose
```
Expected: 25 tests, all FAIL (因为函数都返回原值，部分断言会通过，但逻辑测试会失败)

---

### Task 2: 实现 trimSpaces

**Files:**
- Modify: `src/utils/data-cleaning-utils.js` (trimSpaces 函数)

- [ ] **Step 1: 实现 trimSpaces**

```javascript
function trimSpaces(values, mode) {
  return values.map(function (row) {
    return row.map(function (cell) {
      if (typeof cell !== "string") return cell;
      switch (mode) {
        case "leading":
          return cell.replace(/^\s+/, "");
        case "trailing":
          return cell.replace(/\s+$/, "");
        case "all":
          return cell.replace(/\s+/g, " ").trim();
        case "both":
        default:
          return cell.trim();
      }
    });
  });
}
```

- [ ] **Step 2: 运行 trimSpaces 测试确认通过**

Run: `npx jest src/utils/data-cleaning-utils.test.js -t "trimSpaces" --verbose`

Expected: 7 tests, all PASS

- [ ] **Step 3: 提交**

Run:
```bash
git add src/utils/data-cleaning-utils.js src/utils/data-cleaning-utils.test.js
git commit -m "feat: add trimSpaces function for data cleaning"
```

---

### Task 3: 实现 removeEmptyRows

**Files:**
- Modify: `src/utils/data-cleaning-utils.js` (removeEmptyRows 函数)

- [ ] **Step 1: 实现 removeEmptyRows**

```javascript
function removeEmptyRows(values, mode, columnIndex, ratioThreshold) {
  return values.filter(function (row) {
    if (mode === "all") {
      return row.some(function (cell) {
        return cell !== null && cell !== undefined && cell !== "";
      });
    }
    if (mode === "column") {
      var val = row[columnIndex];
      return val !== null && val !== undefined && val !== "";
    }
    if (mode === "ratio") {
      var emptyCount = 0;
      for (var i = 0; i < row.length; i++) {
        if (row[i] === null || row[i] === undefined || row[i] === "") {
          emptyCount++;
        }
      }
      var emptyRatio = (emptyCount / row.length) * 100;
      return emptyRatio < ratioThreshold;
    }
    return true;
  });
}
```

- [ ] **Step 2: 运行 removeEmptyRows 测试确认通过**

Run: `npx jest src/utils/data-cleaning-utils.test.js -t "removeEmptyRows" --verbose`

Expected: 5 tests, all PASS

- [ ] **Step 3: 提交**

Run:
```bash
git add src/utils/data-cleaning-utils.js
git commit -m "feat: add removeEmptyRows function for data cleaning"
```

---

### Task 4: 实现 convertCase

**Files:**
- Modify: `src/utils/data-cleaning-utils.js` (convertCase 函数)

- [ ] **Step 1: 实现 convertCase**

```javascript
function convertCase(values, mode) {
  return values.map(function (row) {
    return row.map(function (cell) {
      if (typeof cell !== "string") return cell;
      switch (mode) {
        case "upper":
          return cell.toUpperCase();
        case "lower":
          return cell.toLowerCase();
        case "capitalize":
          return cell.charAt(0).toUpperCase() + cell.slice(1);
        case "camel":
          return cell.replace(/[^a-zA-Z0-9]+(.)/g, function (_, chr) {
            return chr.toUpperCase();
          });
        default:
          return cell;
      }
    });
  });
}
```

- [ ] **Step 2: 运行 convertCase 测试确认通过**

Run: `npx jest src/utils/data-cleaning-utils.test.js -t "convertCase" --verbose`

Expected: 5 tests, all PASS

- [ ] **Step 3: 提交**

Run:
```bash
git add src/utils/data-cleaning-utils.js
git commit -m "feat: add convertCase function for data cleaning"
```

---

### Task 5: 实现 removeInvisible

**Files:**
- Modify: `src/utils/data-cleaning-utils.js` (removeInvisible 函数)

- [ ] **Step 1: 实现 removeInvisible**

```javascript
function removeInvisible(values, mode) {
  var patterns = {
    control: /[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g,
    whitespace: /[\t\n\r\x0B\x0C\x1F]/g,
    "zero-width": /[​‌‍﻿⁠‎‏]/g,
    all: /[\x00-\x08\x0B\x0C\x0E-\x1F\x7F\t\n\r​‌‍﻿⁠‎‏]/g,
  };

  var regex = patterns[mode] || patterns.all;

  return values.map(function (row) {
    return row.map(function (cell) {
      if (typeof cell !== "string") return cell;
      return cell.replace(regex, "");
    });
  });
}
```

- [ ] **Step 2: 运行 removeInvisible 测试确认通过**

Run: `npx jest src/utils/data-cleaning-utils.test.js -t "removeInvisible" --verbose`

Expected: 5 tests, all PASS

- [ ] **Step 3: 提交**

Run:
```bash
git add src/utils/data-cleaning-utils.js
git commit -m "feat: add removeInvisible function for data cleaning"
```

---

### Task 6: 实现 removeDuplicates

**Files:**
- Modify: `src/utils/data-cleaning-utils.js` (removeDuplicates 函数)

- [ ] **Step 1: 实现 removeDuplicates**

```javascript
function removeDuplicates(values, keyColumns, keep) {
  var seen = [];
  var result = [];

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var key = keyColumns
      ? keyColumns.map(function (col) { return row[col]; }).join("|||")
      : row.join("|||");

    var found = false;
    for (var j = 0; j < seen.length; j++) {
      if (seen[j] === key) {
        found = true;
        if (keep === "last") {
          result[j] = row;
        }
        break;
      }
    }

    if (!found) {
      seen.push(key);
      result.push(row);
    }
  }

  return result;
}
```

- [ ] **Step 2: 运行 removeDuplicates 测试确认通过**

Run: `npx jest src/utils/data-cleaning-utils.test.js -t "removeDuplicates" --verbose`

Expected: 6 tests, all PASS

- [ ] **Step 3: 运行全部测试确认全部通过**

Run: `npx jest src/utils/data-cleaning-utils.test.js --verbose`

Expected: 28 tests, all PASS

- [ ] **Step 4: 提交**

Run:
```bash
git add src/utils/data-cleaning-utils.js
git commit -m "feat: add removeDuplicates function for data cleaning"
```

---

### Task 7: 创建 taskpane HTML + CSS

**Files:**
- Create: `src/taskpane/data-cleaning-taskpane.html`
- Create: `src/taskpane/data-cleaning-taskpane.css`

- [ ] **Step 1: 创建 data-cleaning-taskpane.html**

```html
<!-- src/taskpane/data-cleaning-taskpane.html -->
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8" />
  <meta http-equiv="X-UA-Compatible" content="IE=edge" />
  <title>数据清洗</title>
  <link rel="stylesheet" href="https://localhost:3000/taskpane.css" />
  <link rel="stylesheet" href="https://localhost:3000/data-cleaning-taskpane.css" />
</head>
<body>
  <div id="container" class="ms-welcome">
    <header class="ms-welcome__header ms-bgColor-neutralLighter">
      <h1 class="ms-font-xl">🧹 数据清洗</h1>
    </header>
    <section class="ms-welcome__main">
      <!-- 选区信息 -->
      <div id="selection-bar" class="section">
        <div class="flex-row">
          <span id="selection-address" class="label">选区：未选择</span>
          <button id="refresh-btn" class="btn-secondary" title="刷新选区">↻</button>
        </div>
      </div>

      <!-- 操作选择 -->
      <div id="operation-section" class="section">
        <h3 class="section-title">🔧 选择操作</h3>
        <div id="operation-grid" class="op-grid">
          <button class="op-btn" data-op="trimSpaces">修剪空格</button>
          <button class="op-btn" data-op="removeEmpty">删除空行</button>
          <button class="op-btn" data-op="convertCase">大小写转换</button>
          <button class="op-btn" data-op="removeInvisible">清除不可见</button>
          <button class="op-btn" data-op="removeDuplicates">移除重复行</button>
        </div>
      </div>

      <!-- 参数设置 -->
      <div id="params-section" class="section">
        <h3 class="section-title">⚙️ 参数设置</h3>
        <div id="params-container"></div>
      </div>

      <!-- 预览 -->
      <div id="preview-section" class="section">
        <h3 class="section-title">👁 预览（前 5 行）</h3>
        <div id="preview-container">
          <table id="preview-table">
            <thead><tr id="preview-header"><th>原值</th><th>清洗后</th></tr></thead>
            <tbody id="preview-body"></tbody>
          </table>
        </div>
      </div>

      <!-- 执行 & 撤销 -->
      <div id="action-bar" class="section flex-row">
        <button id="execute-btn" class="btn-primary" disabled>🚀 执行清洗</button>
        <button id="undo-btn" class="btn-secondary" disabled>↩ 撤销</button>
      </div>

      <!-- 状态 -->
      <div id="status" class="status-message status-idle"></div>
    </section>
  </div>
</body>
</html>
```

- [ ] **Step 2: 创建 data-cleaning-taskpane.css**

```css
/* src/taskpane/data-cleaning-taskpane.css */

#container {
  padding: 10px;
}

.section {
  margin-bottom: 14px;
}

.section-title {
  font-size: 14px;
  font-weight: 600;
  margin: 0 0 8px 0;
  color: #333;
}

.flex-row {
  display: flex;
  align-items: center;
  gap: 8px;
}

.label {
  font-size: 12px;
  color: #666;
}

/* 选区栏 */
#selection-bar {
  padding: 8px;
  background: #f5f5f5;
  border-radius: 4px;
}

/* 操作网格 */
.op-grid {
  display: flex;
  flex-wrap: wrap;
  gap: 6px;
}

.op-btn {
  padding: 8px 14px;
  border: 1px solid #ccc;
  border-radius: 4px;
  background: #fff;
  cursor: pointer;
  font-size: 13px;
  transition: all 0.15s;
}

.op-btn:hover {
  border-color: #0078d4;
  background: #e8f4ff;
}

.op-btn.selected {
  border-color: #0078d4;
  background: #0078d4;
  color: #fff;
}

/* 参数设置 */
#params-container {
  padding: 10px;
  background: #fafafa;
  border: 1px solid #eee;
  border-radius: 4px;
  min-height: 40px;
}

.param-row {
  display: flex;
  align-items: center;
  gap: 10px;
  margin-bottom: 8px;
}

.param-row label {
  font-size: 12px;
  color: #555;
  min-width: 70px;
}

.param-row select,
.param-row input[type="number"] {
  flex: 1;
  padding: 4px 8px;
  border: 1px solid #ccc;
  border-radius: 3px;
  font-size: 12px;
}

.param-row select:focus,
.param-row input:focus {
  outline: none;
  border-color: #0078d4;
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

/* 预览 */
#preview-container {
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
  padding: 4px 8px;
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

/* 按钮 */
.btn-primary {
  padding: 8px 20px;
  background: #0078d4;
  color: #fff;
  border: none;
  border-radius: 4px;
  cursor: pointer;
  font-size: 13px;
  flex: 1;
}

.btn-primary:disabled {
  background: #ccc;
  cursor: not-allowed;
}

.btn-primary:hover:not(:disabled) {
  background: #106ebe;
}

.btn-secondary {
  padding: 8px 16px;
  background: #fff;
  color: #333;
  border: 1px solid #ccc;
  border-radius: 4px;
  cursor: pointer;
  font-size: 13px;
}

.btn-secondary:disabled {
  color: #ccc;
  cursor: not-allowed;
}

.btn-secondary:hover:not(:disabled) {
  background: #f0f0f0;
}

/* 状态 */
.status-message {
  padding: 8px;
  border-radius: 4px;
  font-size: 12px;
  text-align: center;
}

.status-idle { display: none; }
.status-loading { background: #fff3cd; color: #856404; display: block; }
.status-success { background: #d4edda; color: #155724; display: block; }
.status-error { background: #f8d7da; color: #721c24; display: block; }

/* 空状态 */
.empty-state {
  padding: 20px;
  text-align: center;
  color: #999;
  font-size: 13px;
}
```

- [ ] **Step 3: 提交**

Run:
```bash
git add src/taskpane/data-cleaning-taskpane.html src/taskpane/data-cleaning-taskpane.css
git commit -m "feat: add data cleaning taskpane HTML and CSS"
```

---

### Task 8: 创建 taskpane JS

**Files:**
- Create: `src/taskpane/data-cleaning-taskpane.js`

- [ ] **Step 1: 创建 taskpane JS 文件**

```javascript
// src/taskpane/data-cleaning-taskpane.js

/* global console, document, Excel, Office */

var cleaningUtils = require("../utils/data-cleaning-utils");

var MAX_PREVIEW_ROWS = 5;
var LARGE_DATA_THRESHOLD = 10000;

// 当前状态
var state = {
  selectionAddress: null,
  rawValues: null,      // 原始数据（二维数组）— 用于撤销
  headerRow: null,      // 表头数组（如果有多行数据）
  dataRows: null,       // 数据行（不含表头）
  selectedOp: null,     // 当前选中的操作
  undoData: null,       // 撤销快照
};

// ===== 操作参数配置 =====
var OPERATION_PARAMS = {
  trimSpaces: {
    label: "修剪空格",
    template: function () {
      return (
        '<div class="param-row">' +
          '<label>清洗模式：</label>' +
          '<select id="trim-mode">' +
            '<option value="both">首尾空格</option>' +
            '<option value="all">所有多余空格</option>' +
            '<option value="leading">开头空格</option>' +
            '<option value="trailing">结尾空格</option>' +
          "</select>" +
        "</div>"
      );
    },
    getParams: function () {
      return { mode: document.getElementById("trim-mode").value };
    },
    execute: function (values, params) {
      return cleaningUtils.trimSpaces(values, params.mode);
    },
  },
  removeEmpty: {
    label: "删除空行",
    template: function () {
      return (
        '<div class="param-row">' +
          '<label>判定方式：</label>' +
          '<select id="empty-mode">' +
            '<option value="all">完全空行</option>' +
            '<option value="column">指定列为空</option>' +
            '<option value="ratio">空值率超过</option>' +
          "</select>" +
        "</div>" +
        '<div class="param-row" id="empty-column-row" style="display:none">' +
          '<label>列索引：</label>' +
          '<input type="number" id="empty-column" value="0" min="0" />' +
        "</div>" +
        '<div class="param-row" id="empty-ratio-row" style="display:none">' +
          '<label>阈值(%)：</label>' +
          '<input type="number" id="empty-ratio" value="50" min="0" max="100" />' +
        "</div>"
      );
    },
    bindEvents: function () {
      var modeSelect = document.getElementById("empty-mode");
      if (modeSelect) {
        modeSelect.addEventListener("change", function () {
          var val = this.value;
          document.getElementById("empty-column-row").style.display = val === "column" ? "flex" : "none";
          document.getElementById("empty-ratio-row").style.display = val === "ratio" ? "flex" : "none";
        });
      }
    },
    getParams: function () {
      var mode = document.getElementById("empty-mode").value;
      return {
        mode: mode,
        columnIndex: mode === "column" ? parseInt(document.getElementById("empty-column").value) || 0 : null,
        ratioThreshold: mode === "ratio" ? parseInt(document.getElementById("empty-ratio").value) || 50 : null,
      };
    },
    execute: function (values, params) {
      // removeEmptyRows 整行删除，对 dataRows 操作后合并表头
      return cleaningUtils.removeEmptyRows(values, params.mode, params.columnIndex, params.ratioThreshold);
    },
  },
  convertCase: {
    label: "大小写转换",
    template: function () {
      return (
        '<div class="param-row">' +
          '<label>转换模式：</label>' +
          '<select id="case-mode">' +
            '<option value="upper">全部大写</option>' +
            '<option value="lower">全部小写</option>' +
            '<option value="capitalize">首字母大写</option>' +
          "</select>" +
        "</div>"
      );
    },
    getParams: function () {
      return { mode: document.getElementById("case-mode").value };
    },
    execute: function (values, params) {
      return cleaningUtils.convertCase(values, params.mode);
    },
  },
  removeInvisible: {
    label: "清除不可见字符",
    template: function () {
      return (
        '<div class="param-row">' +
          '<label>清除类型：</label>' +
          '<select id="invisible-mode">' +
            '<option value="control">控制字符</option>' +
            '<option value="whitespace">空白字符(\\t\\n)</option>' +
            '<option value="zero-width">零宽字符</option>' +
            '<option value="all">全部非打印字符</option>' +
          "</select>" +
        "</div>"
      );
    },
    getParams: function () {
      return { mode: document.getElementById("invisible-mode").value };
    },
    execute: function (values, params) {
      return cleaningUtils.removeInvisible(values, params.mode);
    },
  },
  removeDuplicates: {
    label: "移除重复行",
    template: function () {
      // 列选择器在 bindEvents 时根据实际数据填充
      return (
        '<div class="param-row">' +
          '<label>依据列：</label>' +
          '<div id="dup-columns" class="checkbox-group"></div>' +
        "</div>" +
        '<div class="param-row">' +
          '<label>保留：</label>' +
          '<select id="dup-keep">' +
            '<option value="first">首行</option>' +
            '<option value="last">末行</option>' +
          "</select>" +
        "</div>"
      );
    },
    bindEvents: function () {
      var container = document.getElementById("dup-columns");
      if (!container || !state.headerRow) return;
      container.innerHTML = "";
      // 生成列选择复选框
      for (var i = 0; i < state.headerRow.length; i++) {
        var label = document.createElement("label");
        var cb = document.createElement("input");
        cb.type = "checkbox";
        cb.value = i;
        cb.checked = true;
        label.appendChild(cb);
        label.appendChild(document.createTextNode(" " + state.headerRow[i]));
        container.appendChild(label);
      }
    },
    getParams: function () {
      var checkboxes = document.querySelectorAll("#dup-columns input[type=checkbox]:checked");
      var cols = [];
      for (var i = 0; i < checkboxes.length; i++) {
        cols.push(parseInt(checkboxes[i].value));
      }
      return {
        keyColumns: cols.length > 0 ? cols : null,
        keep: document.getElementById("dup-keep").value,
      };
    },
    execute: function (values, params) {
      return cleaningUtils.removeDuplicates(values, params.keyColumns, params.keep);
    },
  },
};

// ===== Office 初始化 =====

Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("refresh-btn").addEventListener("click", loadSelection);
    document.getElementById("execute-btn").addEventListener("click", executeClean);
    document.getElementById("undo-btn").addEventListener("click", undoClean);

    // 操作按钮点击
    var opBtns = document.querySelectorAll(".op-btn");
    for (var i = 0; i < opBtns.length; i++) {
      opBtns[i].addEventListener("click", function () {
        selectOperation(this.dataset.op);
      });
    }

    // 打开面板时自动加载选区
    loadSelection();
  }
});

// ===== 选区加载 =====

function loadSelection() {
  setStatus("读取选区...", "loading");

  Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    range.load(["address", "values"]);
    return context.sync().then(function () {
      var address = range.address;
      var values = range.values;

      if (!values || !Array.isArray(values) || values.length === 0) {
        state.rawValues = null;
        state.headerRow = null;
        state.dataRows = null;
        state.selectionAddress = address || "未知";
        updateUIForEmpty();
        setStatus("请选中包含数据的目标区域", "error");
        return;
      }

      state.selectionAddress = address;
      state.rawValues = values;

      // 多行时首行为表头
      if (values.length > 1) {
        state.headerRow = values[0].map(String);
        state.dataRows = values.slice(1);
      } else {
        state.headerRow = null;
        state.dataRows = values;
      }

      document.getElementById("selection-address").textContent = "选区：" + address;
      enableExecuteIfReady();
      updatePreview();
      setStatus("已读取 " + state.dataRows.length + " 行数据", "success");
    });
  }).catch(function (error) {
    setStatus("读取选区失败: " + error.message, "error");
  });
}

// ===== 操作选择 =====

function selectOperation(opKey) {
  state.selectedOp = opKey;

  // 高亮选中按钮
  var btns = document.querySelectorAll(".op-btn");
  for (var i = 0; i < btns.length; i++) {
    btns[i].classList.remove("selected");
  }
  var activeBtn = document.querySelector('.op-btn[data-op="' + opKey + '"]');
  if (activeBtn) activeBtn.classList.add("selected");

  // 渲染参数
  var opConfig = OPERATION_PARAMS[opKey];
  if (opConfig) {
    document.getElementById("params-container").innerHTML = opConfig.template();
    if (opConfig.bindEvents) opConfig.bindEvents();
  }

  enableExecuteIfReady();
  updatePreview();
}

// ===== 预览 =====

function updatePreview() {
  var tbody = document.getElementById("preview-body");
  tbody.innerHTML = "";

  if (!state.dataRows || state.dataRows.length === 0) {
    tbody.innerHTML = '<tr><td colspan="2" class="empty-state">暂无数据</td></tr>';
    return;
  }

  if (!state.selectedOp) {
    tbody.innerHTML = '<tr><td colspan="2" class="empty-state">请先选择操作</td></tr>';
    return;
  }

  var opConfig = OPERATION_PARAMS[state.selectedOp];
  if (!opConfig) return;

  var params = opConfig.getParams();
  var previewData = state.dataRows.slice(0, MAX_PREVIEW_ROWS);
  var resultData = opConfig.execute(previewData, params);

  for (var i = 0; i < previewData.length; i++) {
    var tr = document.createElement("tr");
    var originalStr = previewData[i].map(formatCell).join(" | ");
    var cleanedStr = (resultData[i] || []).map(formatCell).join(" | ");
    tr.innerHTML =
      "<td>" + escapeHtml(originalStr) + "</td>" +
      "<td>" + escapeHtml(cleanedStr) + "</td>";
    tbody.appendChild(tr);
  }
}

function formatCell(val) {
  if (val === null || val === undefined) return "";
  return String(val);
}

function escapeHtml(str) {
  if (typeof str !== "string") return String(str);
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

// ===== 执行 =====

function executeClean() {
  if (!state.selectedOp || !state.dataRows || state.dataRows.length === 0) {
    return;
  }

  if (state.dataRows.length > LARGE_DATA_THRESHOLD) {
    var confirmed = confirm(
      "数据量较大（" + state.dataRows.length + " 行），确定继续执行？"
    );
    if (!confirmed) return;
  }

  var opConfig = OPERATION_PARAMS[state.selectedOp];
  var params = opConfig.getParams();

  // 备份原始数据（撤销用）
  state.undoData = state.rawValues;

  setStatus("执行中...", "loading");
  document.getElementById("execute-btn").disabled = true;

  Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    range.load("address");
    return context.sync().then(function () {
      // 用选区数据执行清洗（含表头）
      var allValues = state.headerRow
        ? [state.headerRow].concat(state.dataRows)
        : state.dataRows;
      var cleaned = opConfig.execute(allValues, params);

      // 写回 Excel
      var writeRange = context.workbook.getSelectedRange();
      writeRange.values = cleaned;
      return context.sync();
    });
  }).then(function () {
    setStatus("清洗完成！已处理 " + state.dataRows.length + " 行", "success");
    document.getElementById("undo-btn").disabled = false;
    // 更新本地数据
    loadSelection();
  }).catch(function (error) {
    setStatus("执行失败: " + error.message, "error");
    document.getElementById("execute-btn").disabled = false;
  });
}

// ===== 撤销 =====

function undoClean() {
  if (!state.undoData) return;

  setStatus("撤销中...", "loading");

  Excel.run(function (context) {
    var writeRange = context.workbook.getSelectedRange();
    writeRange.values = state.undoData;
    return context.sync();
  }).then(function () {
    state.undoData = null;
    document.getElementById("undo-btn").disabled = true;
    setStatus("已撤销", "success");
    loadSelection();
  }).catch(function (error) {
    setStatus("撤销失败: " + error.message, "error");
  });
}

// ===== UI 辅助 =====

function updateUIForEmpty() {
  document.getElementById("selection-address").textContent = "选区：无数据";
  document.getElementById("execute-btn").disabled = true;
  document.getElementById("undo-btn").disabled = true;
  document.getElementById("preview-body").innerHTML =
    '<tr><td colspan="2" class="empty-state">请选中数据区域后点击刷新</td></tr>';
}

function enableExecuteIfReady() {
  document.getElementById("execute-btn").disabled =
    !state.selectedOp || !state.dataRows || state.dataRows.length === 0;
}

function setStatus(message, type) {
  var el = document.getElementById("status");
  el.textContent = message;
  el.className = "status-message status-" + type;
}
```

- [ ] **Step 2: 提交**

Run:
```bash
git add src/taskpane/data-cleaning-taskpane.js
git commit -m "feat: add data cleaning taskpane JS logic"
```

---

### Task 9: webpack 入口 + manifest 注册

**Files:**
- Modify: `webpack.config.js`
- Modify: `manifest.xml`

- [ ] **Step 1: webpack.config.js — 添加 data-cleaning-taskpane 入口**

Add after the `ai-assistant-taskpane` entry and its HtmlWebpackPlugin:

```javascript
// In entry object (after ai-assistant-taskpane):
"data-cleaning-taskpane": [
  "./src/taskpane/data-cleaning-taskpane.js",
  "./src/taskpane/data-cleaning-taskpane.html",
],
```

```javascript
// In plugins, after the ai-assistant HtmlWebpackPlugin:
new HtmlWebpackPlugin({
  filename: "data-cleaning-taskpane.html",
  template: "./src/taskpane/data-cleaning-taskpane.html",
  chunks: ["polyfill", "data-cleaning-taskpane"],
}),
```

- [ ] **Step 2: manifest.xml — 添加按钮到「数据处理」组**

In the `DataProcessingGroup`, add after the SplitSheetButton:

```xml
<Control xsi:type="Button" id="DataCleaningButton">
  <Label resid="DataCleaningButton.Label"/>
  <Supertip>
    <Title resid="DataCleaningButton.Label"/>
    <Description resid="DataCleaningButton.Tooltip"/>
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="ToolsIcon.16x16"/>
    <bt:Image size="32" resid="ToolsIcon.32x32"/>
    <bt:Image size="80" resid="ToolsIcon.80x80"/>
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <TaskpaneId>DataCleaningTaskpaneId</TaskpaneId>
    <SourceLocation resid="DataCleaningTaskpane.Url"/>
  </Action>
</Control>
```

In the `<bt:Urls>` section, add:
```xml
<bt:Url id="DataCleaningTaskpane.Url" DefaultValue="https://localhost:3000/data-cleaning-taskpane.html"/>
```

In the `<bt:ShortStrings>` section, add:
```xml
<bt:String id="DataCleaningButton.Label" DefaultValue="数据清洗"/>
```

In the `<bt:LongStrings>` section, add:
```xml
<bt:String id="DataCleaningButton.Tooltip" DefaultValue="修剪空格、删除空行、大小写转换、清除不可见字符、移除重复行"/>
```

- [ ] **Step 3: 提交**

Run:
```bash
git add webpack.config.js manifest.xml
git commit -m "feat: register data cleaning taskpane in webpack and manifest"
```

---

### Task 10: 构建验证

- [ ] **Step 1: 运行全部单元测试**

Run: `npx jest --verbose`

Expected: All existing tests + 28 new data cleaning tests all PASS

- [ ] **Step 2: 验证 webpack 构建**

Run: `npm run build:dev`

Expected: webpack exits without error, dist/ contains `data-cleaning-taskpane.html` and `data-cleaning-taskpane.js`

- [ ] **Step 3: 提交最终构建就绪状态**

Run:
```bash
git add -A
git commit -m "chore: data cleaning tool ready for build verification"
```
