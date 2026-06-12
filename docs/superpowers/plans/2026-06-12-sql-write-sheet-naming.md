# SQL 查询写入新工作表命名功能 实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 用户点击"写入新工作表"时可自定义工作表名称，不指定时自动使用不重复的序号名。

**Architecture:** 在 `writeResultToSheet()` 函数内添加 `prompt()` 获取用户输入，然后查询已有 sheet 名称列表，通过循环生成不重复的名称，最后用该名称创建 sheet。

**Tech Stack:** Office.js API（`worksheets.load("items/name")`）、浏览器原生 `prompt()`

**设计文档:** `docs/superpowers/specs/2026-06-12-sql-write-sheet-naming-design.md`

---

### Task 1: 修改 `writeResultToSheet()` 添加命名和去重逻辑

**文件:**
- 修改: `src/taskpane/sql-query-taskpane.js:467-494`

- [ ] **Step 1: 确认当前代码状态**

确认当前 `writeResultToSheet()` 函数内容：

```javascript
function writeResultToSheet() {
  if (!currentQueryResult) return;

  var rows = currentQueryResult.rows;
  var columns = currentQueryResult.columns;

  Excel.run(function (context) {
    var sheetCollection = context.workbook.worksheets;
    var newSheet = sheetCollection.add("查询结果");
    newSheet.position = 0;

    var totalRows = rows.length + 1;
    var range = newSheet.getRangeByIndexes(0, 0, totalRows, columns.length);
    var values = [columns];
    for (var r = 0; r < rows.length; r++) {
      values.push(rows[r]);
    }
    range.values = values;
    range.format.autofitColumns();
    newSheet.activate();

    return context.sync();
  }).then(function () {
    setStatusText("queryStatus", "已将 " + rows.length + " 行结果写入新工作表", "success");
  }).catch(function (error) {
    setStatusText("queryStatus", "写入失败: " + error.message, "error");
  });
}
```

- [ ] **Step 2: 替换为带 prompt 和去重的新实现**

将整个函数替换为：

```javascript
function writeResultToSheet() {
  if (!currentQueryResult) return;

  var defaultName = "查询结果";
  var sheetName = prompt("请输入工作表名称:", defaultName);
  if (sheetName === null || sheetName.trim() === "") {
    sheetName = defaultName;
  }
  var rows = currentQueryResult.rows;
  var columns = currentQueryResult.columns;

  Excel.run(function (context) {
    var sheetCollection = context.workbook.worksheets;
    sheetCollection.load("items/name");
    return context.sync().then(function () {
      // 生成不重复的名称
      var finalName = sheetName;
      var counter = 1;
      var exists = true;
      while (exists) {
        exists = false;
        for (var i = 0; i < sheetCollection.items.length; i++) {
          if (sheetCollection.items[i].name === finalName) {
            exists = true;
            finalName = sheetName + " (" + counter + ")";
            counter++;
            break;
          }
        }
      }

      var newSheet = sheetCollection.add(finalName);
      newSheet.position = 0;

      var totalRows = rows.length + 1;
      var range = newSheet.getRangeByIndexes(0, 0, totalRows, columns.length);
      var values = [columns];
      for (var r = 0; r < rows.length; r++) {
        values.push(rows[r]);
      }
      range.values = values;
      range.format.autofitColumns();
      newSheet.activate();

      return context.sync();
    });
  }).then(function () {
    setStatusText("queryStatus", "已将 " + rows.length + " 行结果写入新工作表", "success");
  }).catch(function (error) {
    setStatusText("queryStatus", "写入失败: " + error.message, "error");
  });
}
```

- [ ] **Step 3: 开发构建验证**

```bash
npm run build:dev
```

预期输出：`webpack 5.x.x compiled successfully in XXXX ms`

- [ ] **Step 4: 提交**

```bash
git add src/taskpane/sql-query-taskpane.js
git commit -m "feat: add custom sheet naming with duplicate detection for SQL query results"
```
