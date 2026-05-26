# 连接列侧边栏实施计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 将连接列功能从弹窗模式改造为独立的 Excel 侧边栏操作界面

**Architecture:** 新建独立的连接列侧边栏页面 (`concat-taskpane.html/js`)，修改 manifest.xml 将 ConcatenateButton 从 ExecuteFunction 改为 ShowTaskpane

**Tech Stack:** Office JavaScript API, Fluent UI, webpack

---

## 文件结构

```
新增文件:
- src/taskpane/concat-taskpane.html  # 连接列侧边栏页面
- src/taskpane/concat-taskpane.js    # 连接列侧边栏逻辑

修改文件:
- manifest.xml                       # 添加 ConcatTaskpane.Url，修改 ConcatenateButton Action
- src/taskpane/taskpane.html        # 移除连接列相关 UI
- src/taskpane/taskpane.js          # 移除 runConcat 函数
```

---

## Task 1: 创建 concat-taskpane.html

**Files:**
- Create: `src/taskpane/concat-taskpane.html`

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
</head>

<body class="ms-font-m ms-welcome ms-Fabric">
    <header class="ms-welcome__header ms-bgColor-neutralLighter">
        <h1 class="ms-font-su">连接列工具</h1>
    </header>
    <main id="app-body" class="ms-welcome__main">
        <h2 class="ms-font-l">选中 Excel 中的两列</h2>
        <p class="ms-font-m">输入连接符后按 Enter 键执行合并</p>
        <div style="margin: 20px 0;">
            <label for="connector" class="ms-font-m">连接符：</label>
            <input type="text" id="connector" value="_" autofocus
                   style="width: 150px; height: 30px; font-size: 16px; margin-top: 8px;"/>
        </div>
        <p id="status" style="margin-top: 20px; color: green;"></p>
    </main>
</body>

</html>
```

---

## Task 2: 创建 concat-taskpane.js

**Files:**
- Create: `src/taskpane/concat-taskpane.js`

```javascript
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

var MAX_ROWS = 1050000;

function getColumnLetter(colIndex) {
  var letter = "";
  var remaining = colIndex;
  do {
    letter = String.fromCharCode(65 + (remaining % 26)) + letter;
    remaining = Math.floor(remaining / 26) - 1;
  } while (remaining >= 0);
  return letter;
}

function escapeFormulaText(text) {
  return text.replace(/"/g, '""');
}

function buildConcatFormula(firstColLetter, secondColLetter, connector) {
  var escaped = escapeFormulaText(connector);
  return (
    "=IF(" +
    firstColLetter + "1&" + secondColLetter + '1="","",' +
    firstColLetter + '1&"' + escaped + '"&' + secondColLetter +
    "1)"
  );
}

Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    var connectorInput = document.getElementById("connector");

    // 自动聚焦输入框
    connectorInput.focus();

    // Enter 键触发执行
    connectorInput.addEventListener("keydown", function (event) {
      if (event.key === "Enter") {
        event.preventDefault();
        runConcat();
      }
    });
  }
});

function runConcat() {
  var statusEl = document.getElementById("status");
  var connectorInput = document.getElementById("connector");
  var connector = connectorInput.value || "_";

  statusEl.textContent = "处理中...";
  statusEl.style.color = "green";

  Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    range.load(["address", "columnCount", "columnIndex"]);
    return context.sync().then(function () {
      if (range.columnCount < 2) {
        statusEl.textContent = "错误: 请至少选择两列";
        statusEl.style.color = "red";
        return;
      }

      var worksheet = context.workbook.worksheets.getActiveWorksheet();
      var colIndex = range.columnIndex;
      var firstColLetter = getColumnLetter(colIndex);
      var secondColLetter = getColumnLetter(colIndex + 1);

      var usedInSelection = range.getUsedRange();
      usedInSelection.load("rowCount");
      return context.sync().then(function () {
        var rowCount = usedInSelection.rowCount;

        if (rowCount === 0) {
          statusEl.textContent = "错误: 没有数据";
          statusEl.style.color = "red";
          return;
        }

        if (rowCount > MAX_ROWS) {
          statusEl.textContent =
            "错误: 数据量过大（" + rowCount + " 行），单次最多支持 " + MAX_ROWS + " 行。";
          statusEl.style.color = "red";
          return;
        }

        var targetColLetter = getColumnLetter(colIndex + 2);
        worksheet
          .getRange(targetColLetter + ":" + targetColLetter)
          .insert(Excel.InsertShiftDirection.right);
        return context.sync().then(function () {
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
            statusEl.textContent =
              "完成! 已在第 " + targetColLetter + " 列写入 " + rowCount + " 行公式";
            statusEl.style.color = "green";
          });
        });
      });
    });
  }).catch(function (error) {
    statusEl.textContent = "错误: " + error.message;
    statusEl.style.color = "red";
  });
}
```

---

## Task 3: 修改 manifest.xml

**Files:**
- Modify: `manifest.xml`

**变更点：**
1. 在 `<Resources>` 的 `<bt:Urls>` 中添加新 URL
2. 修改 `ConcatenateButton` 的 `<Action>` 从 `ExecuteFunction` 改为 `ShowTaskpane`

在 `bt:Urls` 中添加：
```xml
<bt:Url id="ConcatTaskpane.Url" DefaultValue="https://localhost:3000/concat-taskpane.html"/>
```

修改 `ConcatenateButton` 的 Action 部分：
```xml
<Action xsi:type="ShowTaskpane">
  <TaskpaneId>ConcatTaskpaneId</TaskpaneId>
  <SourceLocation resid="ConcatTaskpane.Url"/>
</Action>
```

---

## Task 4: 从 taskpane.html 移除连接列相关 UI

**Files:**
- Modify: `src/taskpane/taskpane.html`

在 `<main id="app-body">` 中移除以下内容（第 27-36 行）：
```html
<h2 class="ms-font-l">选中 Excel 中的两列，点击下方按钮进行连接</h2>
<p class="ms-font-m">连接符默认为 <b>_</b></p>
<div style="margin: 20px 0;">
    <label for="connector" class="ms-font-m">自定义连接符（可选）：</label><br/>
    <input type="text" id="connector" value="_" style="width: 100px; height: 30px; font-size: 16px; margin-top: 8px;"/>
</div>
<div role="button" id="concatBtn" class="ms-welcome__action ms-Button ms-Button--hero ms-font-xl">
    <span class="ms-Button-label">连接选中列</span>
</div>
<p id="status" style="margin-top: 20px; color: green;"></p>
```

保留 CSV 导入和增强查找部分。

---

## Task 5: 从 taskpane.js 移除 runConcat 相关代码

**Files:**
- Modify: `src/taskpane/taskpane.js`

**变更点：**
1. 移除 `var { getColumnLetter, escapeFormulaText, buildConcatFormula } = require("../utils/concat-utils");`
2. 移除 `document.getElementById("concatBtn").onclick = runConcat;`
3. 移除整个 `runConcat` 函数（从第 26 行到第 104 行）
4. 移除 `var isProcessing = false;`（如果只用于 runConcat）

最终 taskpane.js 应该只保留 CSV 导入和增强查找功能。

---

## Task 6: 验证 webpack 配置

**Files:**
- Check: `webpack.config.js`

确认 webpack 配置会处理新的 `concat-taskpane.html` 和 `concat-taskpane.js` 文件。通常 `HtmlWebpackPlugin` 会自动包含所有 HTML 文件在 `src` 目录下。

---

## Task 7: 测试完整流程

1. 运行 `npm run dev-server` 启动开发服务器
2. 在 Excel 中点击功能区「连接列」按钮
3. 验证侧边栏打开并自动聚焦连接符输入框
4. 输入自定义连接符，按 Enter 键
5. 验证合并结果正确
6. 验证状态显示正确

---

## Task 8: 删除废弃文件

**Files:**
- Delete: `src/commands/connector-dialog.html`（不再需要）

---

## 实施检查清单

- [ ] Task 1: 创建 concat-taskpane.html
- [ ] Task 2: 创建 concat-taskpane.js
- [ ] Task 3: 修改 manifest.xml
- [ ] Task 4: 从 taskpane.html 移除连接列 UI
- [ ] Task 5: 从 taskpane.js 移除 runConcat
- [ ] Task 6: 验证 webpack 配置
- [ ] Task 7: 测试完整流程
- [ ] Task 8: 删除 connector-dialog.html
- [ ] Git commit 所有变更
