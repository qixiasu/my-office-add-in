# 连接列功能多列扩展实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**目标:** 将连接列功能从仅支持2列扩展为支持任意多列，超过3列时执行前弹出确认框

**架构:** 改动集中在 `concat-utils.js`（新增函数）、`concat-taskpane.js`（核心逻辑改造）、`concat-taskpane.html`（文案更新）、`concat-utils.test.js`（新增测试用例）。公式构造沿用 `&` 拼接方式，通过循环支持 N 列。

**技术栈:** 纯 JavaScript + Excel JavaScript API，无需新增依赖

---

## 文件结构

```
src/
├── utils/
│   ├── concat-utils.js        # 新增 buildNConcatFormula()
│   └── concat-utils.test.js   # 新增多列测试用例
└── taskpane/
    ├── concat-taskpane.html  # 指引文案更新
    └── concat-taskpane.js    # 核心逻辑改造：循环构建公式、多列支持、确认框
```

---

## Task 1: 修改 concat-utils.js — 新增 buildNConcatFormula 函数

**文件:**
- Modify: `src/utils/concat-utils.js`

- [ ] **Step 1: 添加 buildNConcatFormula 函数**

在 `concat-utils.js` 文件末尾，在现有 `module.exports` 之后添加新函数（注意：`module.exports` 已在第 25 行，需要在其之前添加）：

```javascript
/**
 * Build concatenation formula for N columns
 * @param {string[]} columns - Array of column letters, e.g. ['A', 'B', 'C']
 * @param {string} connector - Connector string, e.g. '_'
 * @returns {string} Excel formula string
 */
function buildNConcatFormula(columns, connector) {
  var escaped = escapeFormulaText(connector);
  var rowRef = "1";

  // Build empty-check string: "A1&B1&C1"
  var emptyCheck = columns.map(function(col) {
    return col + rowRef;
  }).join('&');

  // Build concat string: 'A1&"_"&B1&"_"&C1'
  var concatParts = columns.map(function(col, index) {
    if (index === 0) {
      return col + rowRef;
    }
    return '"' + escaped + '"&' + col + rowRef;
  });
  var concatStr = concatParts.join('&');

  return '=IF(' + emptyCheck + '="","",' + concatStr + ')';
}
```

- [ ] **Step 2: 更新 module.exports**

将 `module.exports` 从：
```javascript
module.exports = { getColumnLetter, escapeFormulaText, buildConcatFormula };
```
更新为：
```javascript
module.exports = { getColumnLetter, escapeFormulaText, buildConcatFormula, buildNConcatFormula };
```

- [ ] **Step 3: 提交**

```bash
git add src/utils/concat-utils.js
git commit -m "feat(concat): add buildNConcatFormula for N-column support"
```

---

## Task 2: 修改 concat-utils.test.js — 新增多列测试用例

**文件:**
- Modify: `src/utils/concat-utils.test.js`

- [ ] **Step 1: 添加 3 列和 4 列的测试用例**

在 `concat-utils.test.js` 文件末尾添加：

```javascript
describe("buildNConcatFormula", function () {
  it("builds formula for 3 columns with underscore connector", function () {
    var formula = buildNConcatFormula(['A', 'B', 'C'], '_');
    expect(formula).toBe('=IF(A1&B1&C1="","",A1&"_"&B1&"_"&C1)');
  });

  it("builds formula for 4 columns with dash connector", function () {
    var formula = buildNConcatFormula(['A', 'B', 'C', 'D'], '-');
    expect(formula).toBe('=IF(A1&B1&C1&D1="","",A1&"-"&B1&"-"&C1&"-"&D1)');
  });

  it("escapes double quotes in connector", function () {
    var formula = buildNConcatFormula(['A', 'B'], 'a"b');
    expect(formula).toBe('=IF(A1&B1="","",A1&"a""b"&B1)');
  });

  it("works with multi-letter column references", function () {
    var formula = buildNConcatFormula(['AA', 'AB', 'AC'], '|');
    expect(formula).toBe('=IF(AA1&AB1&AC1="","",AA1&"|"&AB1&"|"&AC1)');
  });

  it("returns empty string formula for single column", function () {
    var formula = buildNConcatFormula(['A'], '_');
    expect(formula).toBe('=IF(A1="","",A1)');
  });
});
```

- [ ] **Step 2: 更新 require 语句**

将文件开头的 `require` 从：
```javascript
var { getColumnLetter, escapeFormulaText, buildConcatFormula } = require("./concat-utils");
```
更新为：
```javascript
var { getColumnLetter, escapeFormulaText, buildConcatFormula, buildNConcatFormula } = require("./concat-utils");
```

- [ ] **Step 3: 运行测试验证**

Run: `npm test -- --testPathPattern=concat-utils`
Expected: PASS (所有新测试用例通过)

- [ ] **Step 4: 提交**

```bash
git add src/utils/concat-utils.test.js
git commit -m "test(concat): add multi-column test cases for buildNConcatFormula"
```

---

## Task 3: 修改 concat-taskpane.html — 更新指引文案

**文件:**
- Modify: `src/taskpane/concat-taskpane.html`

- [ ] **Step 1: 更新指引文案**

将第 34-35 行从：
```html
                <span class="guide-step">1. 选中 Excel 中要连接的两列</span><br>
                <span class="guide-step">2. 输入连接符，点击执行</span>
```
更新为：
```html
                <span class="guide-step">1. 选中 Excel 中要连接的多列（2列或更多）</span><br>
                <span class="guide-step">2. 输入连接符，点击执行</span>
```

- [ ] **Step 2: 提交**

```bash
git add src/taskpane/concat-taskpane.html
git commit -m "docs(concat): update guide text for multi-column support"
```

---

## Task 4: 修改 concat-taskpane.js — 核心逻辑改造

**文件:**
- Modify: `src/taskpane/concat-taskpane.js`

- [ ] **Step 1: 重构 runConcat 函数中的公式构建部分**

将第 24-32 行的 `buildConcatFormula` 函数改为接受多列：

```javascript
function buildConcatFormula(firstColLetter, secondColLetter, connector) {
  var escaped = escapeFormulaText(connector);
  return (
    "=IF(" +
    firstColLetter + "1&" + secondColLetter + '1="","",' +
    firstColLetter + '1&"' + escaped + '"&' + secondColLetter +
    "1)"
  );
}
```

**替换为**（保持向后兼容，但核心逻辑改为调用新函数）：

```javascript
// Keep old 2-column function for backward compatibility
function buildConcatFormula(firstColLetter, secondColLetter, connector) {
  return buildNConcatFormula([firstColLetter, secondColLetter], connector);
}
```

- [ ] **Step 2: 重构 runConcat 函数，支持多列**

将第 49-130 行的 `runConcat` 函数中，第 65-78 行的选区处理部分：

**原来的逻辑（第 65-78 行）:**
```javascript
    range.load(["address", "columnCount", "columnIndex"]);
    return context.sync().then(function () {
      if (range.columnCount < 2) {
        executeBtn.disabled = false;
        setStatus("错误: 请至少选择两列", "error");
        return;
      }

      var worksheet = context.workbook.worksheets.getActiveWorksheet();
      var colIndex = range.columnIndex;
      var firstColLetter = getColumnLetter(colIndex);
      var secondColLetter = getColumnLetter(colIndex + 1);
```

**替换为:**
```javascript
    range.load(["address", "columnCount", "columnIndex"]);
    return context.sync().then(function () {
      if (range.columnCount < 2) {
        executeBtn.disabled = false;
        setStatus("错误: 请至少选择两列", "error");
        return;
      }

      var colCount = range.columnCount;
      var colIndex = range.columnIndex;
      var worksheet = context.workbook.worksheets.getActiveWorksheet();

      // Build columns array: [getColumnLetter(colIndex), getColumnLetter(colIndex+1), ...]
      var columns = [];
      for (var i = 0; i < colCount; i++) {
        columns.push(getColumnLetter(colIndex + i));
      }
```

- [ ] **Step 3: 重构行数检查和目标列逻辑**

将第 80-100 行：

**原来的逻辑:**
```javascript
      var usedInSelection = range.getUsedRange();
      usedInSelection.load("rowCount");
      return context.sync().then(function () {
        var rowCount = usedInSelection.rowCount;

        if (rowCount === 0) {
          executeBtn.disabled = false;
          setStatus("错误: 没有数据", "error");
          return;
        }

        if (rowCount > MAX_ROWS) {
          executeBtn.disabled = false;
          setStatus(
            "错误: 数据量过大（" + rowCount + " 行），单次最多支持 " + MAX_ROWS + " 行。",
            "error"
          );
          return;
        }

        var targetColLetter = getColumnLetter(colIndex + 2);
```

**替换为:**
```javascript
      var usedInSelection = range.getUsedRange();
      usedInSelection.load("rowCount");
      return context.sync().then(function () {
        var rowCount = usedInSelection.rowCount;

        if (rowCount === 0) {
          executeBtn.disabled = false;
          setStatus("错误: 没有数据", "error");
          return;
        }

        if (rowCount > MAX_ROWS) {
          executeBtn.disabled = false;
          setStatus(
            "错误: 数据量过大（" + rowCount + " 行），单次最多支持 " + MAX_ROWS + " 行。",
            "error"
          );
          return;
        }

        // 超过3列显示确认框
        if (colCount > 3) {
          var startColLetter = getColumnLetter(colIndex);
          var endColLetter = getColumnLetter(colIndex + colCount - 1);
          var confirmMsg = "将连接第 " + startColLetter + " 列到第 " + endColLetter + " 列，使用连接符【" + connector + "】";
          if (!window.confirm(confirmMsg)) {
            executeBtn.disabled = false;
            setStatus("状态：等待操作...", "idle");
            return;
          }
        }

        // 目标列插入到所选范围最右侧
        var targetColLetter = getColumnLetter(colIndex + colCount);
```

- [ ] **Step 4: 重构公式构建和填充逻辑**

将第 104-123 行：

**原来的逻辑:**
```javascript
          var formula = buildConcatFormula(firstColLetter, secondColLetter, connector);

          var startCell = worksheet.getRange(targetColLetter + "1");
          startCell.formulas = [[formula]];
          return context.sync().then(function () {
            if (rowCount > 1) {
              var fillRange = worksheet.getRange(
                targetColLetter + "1:" + targetColLetter + rowCount
              );
              startCell.autoFill(fillRange, Excel.AutoFillType.fillDefault);
              return context.sync();
            }
          }).then(function () {
            setStatus(
              "完成! 已在第 " + targetColLetter + " 列写入 " + rowCount + " 行公式",
              "success"
            );
            executeBtn.disabled = false;
          });
```

**替换为:**
```javascript
          var formula = buildNConcatFormula(columns, connector);

          var startCell = worksheet.getRange(targetColLetter + "1");
          startCell.formulas = [[formula]];
          return context.sync().then(function () {
            if (rowCount > 1) {
              var fillRange = worksheet.getRange(
                targetColLetter + "1:" + targetColLetter + rowCount
              );
              startCell.autoFill(fillRange, Excel.AutoFillType.fillDefault);
              return context.sync();
            }
          }).then(function () {
            setStatus(
              "完成! 已在第 " + targetColLetter + " 列写入 " + rowCount + " 行公式",
              "success"
            );
            executeBtn.disabled = false;
          });
```

- [ ] **Step 5: 提交**

```bash
git add src/taskpane/concat-taskpane.js
git commit -m "feat(concat): support N-column concatenation with confirmation for >3 columns"
```

---

## Task 5: 整体验证

**验证步骤:**
- [ ] 运行 `npm run lint` 确保无 lint 错误
- [ ] 运行 `npm test` 确保所有测试通过
- [ ] 手动测试场景：
  - 选 2 列 → 直接执行，结果正确
  - 选 4 列 → 弹出确认框，确认后执行，取消则不执行
  - 连接符含双引号 → 公式中双引号正确转义

---

## 实现顺序

1. Task 1 → Task 2 → Task 3 → Task 4 → Task 5
2. 每个 Task 完成后立即提交
3. Task 5 整体验证通过后，再进行一次 commit 记录验证结果