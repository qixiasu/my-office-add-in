# 字段计数功能 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 在 Excel 侧边栏添加字段计数工具，用户选中列后统计每个字段值出现的次数，生成的新列插入到选中列后面

**Architecture:** 单 taskpane 架构，遵循现有项目的文件结构和模式。核心逻辑：获取选中列 → 确定数据范围（表头行下方）→ 为每列生成 COUNTIF 公式 → 插入新列并写入

**Tech Stack:** Office JavaScript API, Excel.run, webpack (Babel transpilation)

---

## 文件结构

| 文件 | 用途 |
|------|------|
| `src/taskpane/count-values-taskpane.html` | UI 结构：标题、表头行号输入框、执行按钮、状态消息 |
| `src/taskpane/count-values-taskpane.css` | 样式：容器、标题、输入框、按钮、状态消息 |
| `src/taskpane/count-values-taskpane.js` | 主逻辑：获取选中列、计算范围、生成公式、插入列 |
| `src/taskpane/taskpane.css` | 共享基础样式（Fluent UI） |
| `src/utils/concat-utils.js` | 参考：现有 `getColumnLetter` 辅助函数 |
| `manifest.xml` | 添加按钮、资源映射（图标、URL、字符串） |
| `webpack.config.js` | 添加新的 entry point 和 HtmlWebpackPlugin 配置 |

---

## Task 1: 创建 count-values-taskpane.html

**Files:**
- Create: `src/taskpane/count-values-taskpane.html`

- [ ] **Step 1: 创建 HTML 文件**

```html
<!-- Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT License. -->

<!DOCTYPE html>
<html>

<head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>字段计数工具</title>

    <!-- Office JavaScript API -->
    <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>

    <!-- Fluent UI -->
    <link rel="stylesheet" href="https://res-1.cdn.office.net/files/fabric-cdn-prod_20230815.002/office-ui-fabric-core/11.1.0/css/fabric.min.css"/>

    <!-- Template styles -->
    <link href="taskpane.css" rel="stylesheet" type="text/css" />
    <link href="count-values-taskpane.css" rel="stylesheet" type="text/css" />
</head>

<body class="ms-font-m ms-welcome ms-Fabric">
    <div class="count-container">
        <h1 class="count-title">
            <svg width="20" height="20" viewBox="0 0 20 20" fill="#0078d4">
                <path d="M3 3h14v2H3V3zm0 4h10v2H3V7zm0 4h14v2H3v-2zm0 4h10v2H3v-2z"/>
            </svg>
            字段计数工具
        </h1>

        <div class="guide-card">
            <p>
                <span class="guide-step">1. 选中要统计的一列或多列</span><br>
                <span class="guide-step">2. 设置表头行号（默认为第1行）</span><br>
                <span class="guide-step">3. 点击执行按钮</span>
            </p>
        </div>

        <div class="count-form-group">
            <label for="headerRow" class="count-label">表头行号</label>
            <input type="number" id="headerRow" class="count-input" value="1" min="1" />
        </div>

        <button id="executeBtn" class="count-button">执行计数</button>

        <div id="status" class="status-message status-idle">状态：等待操作...</div>
    </div>
</body>

</html>
```

- [ ] **Step 2: Commit**

```bash
git add src/taskpane/count-values-taskpane.html
git commit -m "feat: add count-values-taskpane HTML structure"
```

---

## Task 2: 创建 count-values-taskpane.css

**Files:**
- Create: `src/taskpane/count-values-taskpane.css`

- [ ] **Step 1: 创建 CSS 文件**

```css
/* 容器 */
.count-container {
  padding: 16px;
}

/* 标题 */
.count-title {
  font-size: 16px;
  font-weight: 700;
  color: #0078d4;
  display: flex;
  align-items: center;
  gap: 8px;
  margin: 0 0 16px 0;
}

/* 步骤引导卡片 */
.guide-card {
  background: #f0f7ff;
  border-left: 3px solid #0078d4;
  padding: 12px;
  border-radius: 4px;
  margin-bottom: 16px;
}

/* 表单组 */
.count-form-group {
  margin-bottom: 16px;
}

/* 标签 */
.count-label {
  display: block;
  font-size: 14px;
  color: #333;
  margin-bottom: 6px;
  font-weight: 500;
}

/* 输入框 */
.count-input {
  width: 100%;
  padding: 10px 12px;
  border: 2px solid #0078d4;
  border-radius: 6px;
  font-size: 15px;
  outline: none;
  transition: box-shadow 0.2s;
  box-sizing: border-box;
}

.count-input:focus {
  box-shadow: 0 0 0 3px rgba(0, 120, 212, 0.15);
}

/* 按钮 */
.count-button {
  width: 100%;
  padding: 12px 24px;
  background: #0078d4;
  color: white;
  border: none;
  border-radius: 6px;
  font-size: 15px;
  font-weight: 600;
  cursor: pointer;
  transition: background 0.2s;
}

.count-button:hover {
  background: #106ebe;
}

.count-button:disabled {
  background: #c8c8c8;
  cursor: not-allowed;
}

/* 状态消息 */
.status-message {
  margin-top: 16px;
  padding: 12px;
  border-radius: 6px;
  font-size: 14px;
  text-align: center;
}

.status-idle {
  background: #f5f5f5;
  color: #999;
}

.status-loading {
  background: #fff3cd;
  color: #856404;
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
```

- [ ] **Step 2: Commit**

```bash
git add src/taskpane/count-values-taskpane.css
git commit -m "feat: add count-values-taskpane styles"
```

---

## Task 3: 创建 count-values-taskpane.js

**Files:**
- Create: `src/taskpane/count-values-taskpane.js`

- [ ] **Step 1: 创建 JS 文件**

```javascript
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

var MAX_ROWS = 1050000;

/**
 * 将列索引（0-based）转换为 Excel 列字母
 * @param {number} colIndex - 列索引（0-based）
 * @returns {string} 列字母（如 A, B, ..., Z, AA, AB, ...）
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

Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    var headerRowInput = document.getElementById("headerRow");
    var executeBtn = document.getElementById("executeBtn");

    // 自动聚焦到表头行号输入框
    headerRowInput.focus();

    // 执行按钮点击事件
    executeBtn.addEventListener("click", function () {
      runCount();
    });
  }
});

/**
 * 设置状态消息
 * @param {string} message - 状态文本
 * @param {string} type - 状态类型：idle, loading, success, error
 */
function setStatus(message, type) {
  var el = document.getElementById("status");
  el.textContent = message;
  el.className = "status-message status-" + type;
}

/**
 * 主执行函数：获取选中列，统计计数
 */
function runCount() {
  var headerRowInput = document.getElementById("headerRow");
  var executeBtn = document.getElementById("executeBtn");
  var headerRow = parseInt(headerRowInput.value) || 1;

  // 禁用按钮，显示加载状态
  executeBtn.disabled = true;
  setStatus("处理中...", "loading");

  Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    range.load(["address", "columnCount", "columnIndex", "rowCount"]);
    return context.sync().then(function () {
      if (range.columnCount < 1) {
        executeBtn.disabled = false;
        setStatus("错误: 请先选择一列", "error");
        return;
      }

      var colIndex = range.columnIndex; // 0-based
      var colCount = range.columnCount;
      var worksheet = context.workbook.worksheets.getActiveWorksheet();

      // 获取每个选中列的表头和数据范围
      var columns = [];
      for (var i = 0; i < colCount; i++) {
        var currentColIndex = colIndex + i;
        var colLetter = getColumnLetter(currentColIndex);

        // 获取表头值
        var headerCell = worksheet.getRange(colLetter + headerRow);
        headerCell.load("value");
        columns.push({
          colLetter: colLetter,
          colIndex: currentColIndex,
          headerCell: headerCell
        });
      }

      return context.sync().then(function () {
        // 获取每个列的数据范围（表头行下方到最后一行的数据）
        var dataStartRow = headerRow + 1;

        // 获取所有列的最后一行的最大值
        var maxLastRow = 0;
        var columnDataRanges = [];

        for (var j = 0; j < columns.length; j++) {
          var col = columns[j];
          var colLetter = col.colLetter;

          // 使用 getUsedRange 获取实际使用范围
          var usedRange = worksheet.getRange(colLetter + ":" + colLetter);
          usedRange.load("rowCount");
          usedRange.load("columnCount");

          // 获取特定列的UsedRange
          var singleColUsed = worksheet.getRange(colLetter + dataStartRow + ":" + colLetter + 1000000);
          singleColUsed.load("values");
          singleColUsed.load("rowCount");

          columnDataRanges.push({
            col: col,
            dataStartRow: dataStartRow,
            usedRange: singleColUsed
          });
        }

        return context.sync().then(function () {
          // 计算每个列的实际数据范围
          for (var k = 0; k < columnDataRanges.length; k++) {
            var item = columnDataRanges[k];
            var values = item.usedRange.values;
            var rowCount = item.usedRange.rowCount;

            // 找到最后有数据的行
            var lastRow = dataStartRow;
            for (var r = 0; r < rowCount; r++) {
              if (values[r] && values[r][0] !== null && values[r][0] !== "") {
                lastRow = dataStartRow + r;
              }
            }
            item.lastRow = lastRow;
            item.headerValue = item.col.headerCell.value;
            if (item.headerValue === null || item.headerValue === "") {
              item.headerValue = "Column" + (item.col.colIndex + 1);
            }
          }

          // 执行计数操作
          return executeCount(worksheet, columnDataRanges, headerRow, executeBtn);
        });
      });
    });
  }).catch(function (error) {
    setStatus("错误: " + error.message, "error");
    executeBtn.disabled = false;
  });
}

/**
 * 执行计数操作：插入新列并写入 COUNTIF 公式
 * @param {object} worksheet - Excel worksheet 对象
 * @param {Array} columnDataRanges - 列数据范围信息数组
 * @param {number} headerRow - 表头行号
 * @param {object} executeBtn - 按钮对象
 * @returns {Promise}
 */
function executeCount(worksheet, columnDataRanges, headerRow, executeBtn) {
  return Excel.run(function (context) {
    var results = [];

    for (var i = 0; i < columnDataRanges.length; i++) {
      var item = columnDataRanges[i];
      var colLetter = item.col.colLetter;
      var colIndex = item.col.colIndex;
      var dataStartRow = item.dataStartRow;
      var lastRow = item.lastRow;
      var headerValue = item.headerValue;

      // 在当前列后插入新列
      var resultColLetter = getColumnLetter(colIndex + 1);
      worksheet.getRange(resultColLetter + ":" + resultColLetter).insert(Excel.InsertShiftDirection.right);

      // 写入表头
      var headerCell = worksheet.getRange(resultColLetter + headerRow);
      headerCell.values = [[headerValue + "-计数"]];

      // 生成 COUNTIF 公式
      var formula = "=COUNTIF($" + colLetter + "$" + dataStartRow + ":$" + colLetter + "$" + lastRow + "," + colLetter + dataStartRow + ")";

      // 写入第一行公式
      var startCell = worksheet.getRange(resultColLetter + dataStartRow);
      startCell.formulas = [[formula]];

      // 自动填充到最后一行
      if (lastRow > dataStartRow) {
        var fillRange = worksheet.getRange(resultColLetter + dataStartRow + ":" + resultColLetter + lastRow);
        startCell.autoFill(fillRange, Excel.AutoFillType.fillDefault);
      }

      results.push({
        colLetter: colLetter,
        resultColLetter: resultColLetter,
        rowCount: lastRow - dataStartRow + 1
      });
    }

    return context.sync().then(function () {
      // 汇总结果信息
      var resultParts = [];
      for (var j = 0; j < results.length; j++) {
        var res = results[j];
        resultParts.push("第" + res.resultColLetter + "列(" + res.rowCount + "行)");
      }
      setStatus("完成！已在 " + resultParts.join(", ") + " 写入计数结果", "success");
      executeBtn.disabled = false;
    });
  }).catch(function (error) {
    setStatus("错误: " + error.message, "error");
    executeBtn.disabled = false;
  });
}
```

- [ ] **Step 2: Commit**

```bash
git add src/taskpane/count-values-taskpane.js
git commit -m "feat: add count-values-taskpane main logic"
```

---

## Task 4: 更新 manifest.xml 添加按钮和资源

**Files:**
- Modify: `manifest.xml`

**需要添加的内容：**
1. 新的 Control 条目（在 FillSeriesButton 之后）
2. 新的 bt:Image 条目（图标）
3. 新的 bt:Url 条目（Taskpane URL）
4. 新的 bt:String 条目（按钮标签和 Tooltip）

- [ ] **Step 1: 添加 Control 条目**

在 `</Group>` 前添加新的 Control：

```xml
<Control xsi:type="Button" id="CountValuesButton">
  <Label resid="CountValuesButton.Label"/>
  <Supertip>
    <Title resid="CountValuesButton.Label"/>
    <Description resid="CountValuesButton.Tooltip"/>
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="CountValuesIcon.16x16"/>
    <bt:Image size="32" resid="CountValuesIcon.32x32"/>
    <bt:Image size="80" resid="CountValuesIcon.80x80"/>
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <TaskpaneId>CountValuesTaskpaneId</TaskpaneId>
    <SourceLocation resid="CountValuesTaskpane.Url"/>
  </Action>
</Control>
```

- [ ] **Step 2: 添加 bt:Image 条目**

在 `</bt:Images>` 前添加：

```xml
<bt:Image id="CountValuesIcon.16x16" DefaultValue="https://localhost:3000/assets/count-values-16.png"/>
<bt:Image id="CountValuesIcon.32x32" DefaultValue="https://localhost:3000/assets/count-values-32.png"/>
<bt:Image id="CountValuesIcon.80x80" DefaultValue="https://localhost:3000/assets/count-values-80.png"/>
```

- [ ] **Step 3: 添加 bt:Url 条目**

在 `</bt:Urls>` 前添加：

```xml
<bt:Url id="CountValuesTaskpane.Url" DefaultValue="https://localhost:3000/count-values-taskpane.html"/>
```

- [ ] **Step 4: 添加 bt:String 条目**

在 `<bt:ShortStrings>` 中添加：

```xml
<bt:String id="CountValuesButton.Label" DefaultValue="字段计数"/>
```

在 `<bt:LongStrings>` 中添加：

```xml
<bt:String id="CountValuesButton.Tooltip" DefaultValue="统计选中列每个字段值出现的次数"/>
```

- [ ] **Step 5: Commit**

```bash
git add manifest.xml
git commit -m "feat: add CountValues button to manifest"
```

---

## Task 5: 更新 webpack.config.js 添加 entry 和 plugin

**Files:**
- Modify: `webpack.config.js`

- [ ] **Step 1: 添加 entry**

在 `entry` 对象的 `fill-series-taskpane` 后添加：

```javascript
"count-values-taskpane": ["./src/taskpane/count-values-taskpane.js", "./src/taskpane/count-values-taskpane.html"],
```

- [ ] **Step 2: 添加 HtmlWebpackPlugin 配置**

在 `fill-series-taskpane` plugin 配置后添加：

```javascript
new HtmlWebpackPlugin({
  filename: "count-values-taskpane.html",
  template: "./src/taskpane/count-values-taskpane.html",
  chunks: ["polyfill", "count-values-taskpane"],
}),
```

- [ ] **Step 3: Commit**

```bash
git add webpack.config.js
git commit -m "feat: add count-values-taskpane entry to webpack"
```

---

## Task 6: 创建图标文件

**Files:**
- Create: `assets/count-values-16.png`
- Create: `assets/count-values-32.png`
- Create: `assets/count-values-80.png`

> **注意:** 图标文件需要手动创建或复制现有图标重命名。建议复制 `concat-*.png` 并重命名作为基础，然后可以用图像编辑软件微调。

- [ ] **Step 1: 复制现有图标作为基础**

```bash
cp assets/concat-16.png assets/count-values-16.png
cp assets/concat-32.png assets/count-values-32.png
cp assets/concat-80.png assets/count-values-80.png
```

- [ ] **Step 2: Commit**

```bash
git add assets/count-values-*.png
git commit -m "feat: add count-values icons"
```

---

## Task 7: 测试完整流程

- [ ] **Step 1: 启动开发服务器**

```bash
npm run dev-server
```

- [ ] **Step 2: 启动 Excel 调试**

```bash
npm run start
```

- [ ] **Step 3: 测试场景**

1. 在 Excel 中选中一列数据（第1行是表头）
2. 点击"字段计数"按钮
3. 验证表头行号输入框默认为 1
4. 点击"执行计数"
5. 验证新列已插入，表头为"原表头-计数"，每行有 COUNTIF 公式

- [ ] **Step 4: 测试多列场景**

1. 选中连续三列
2. 点击"字段计数"
3. 验证每列后都插入了计数列

- [ ] **Step 5: 测试表头行号自定义**

1. 选中一列，表头在第3行
2. 在侧边栏输入表头行号为 3
3. 执行计数
4. 验证只统计第4行开始的数据

---

## 验证清单

- [ ] 编译无错误
- [ ] Excel 侧边栏显示"字段计数工具"
- [ ] 输入框、按钮、状态消息样式正常
- [ ] 单列计数功能正常
- [ ] 多列计数功能正常（每列后插入一列）
- [ ] 表头行号可自定义
- [ ] 表头行号位置正确写入"原表头-计数"
- [ ] COUNTIF 公式正确生成
- [ ] 错误处理正常（未选中列、无数据等）