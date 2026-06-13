# 查询结果写入工作表 — 分块进度反馈实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 将 `writeResultToSheet()` 改为分块写入模式，每批 5000 行，逐批更新进度条和状态文字，让用户清晰看到写入进度。

**Architecture:** 利用分批 `Excel.run` 调用间的异步间隙更新 DOM 进度条。每批写入后更新进度百分比和行数文字。CSS transition 动画平滑过渡进度条宽度。

**Tech Stack:** Office.js `Excel.run`、CSS3 transition、递归异步批量写入模式

---

### 任务 1：添加进度条 HTML 元素

**文件：**
- 修改：`src/taskpane/sql-query-taskpane.html`

- [ ] **Step 1: 在 resultActions 和 resultDisplay 之间添加进度条**

在 `resultActions` div（第 133 行 `</div>`）之后、`resultDisplay` div（第 134 行 `<div id="resultDisplay">`）之前插入：

```html
            <!-- 写入进度条 -->
            <div id="writeProgress" class="write-progress" style="display:none">
              <div class="write-progress-bar">
                <div id="writeProgressFill" class="write-progress-fill"></div>
              </div>
              <span id="writeProgressText" class="write-progress-text">正在写入...</span>
            </div>
```

所以这块区域变为：
```html
            <div id="resultActions" class="result-actions" style="display:none">
                <button id="writeSheetBtn" class="btn-secondary" title="将查询结果写入新的工作表">📝 写入新工作表</button>
                <button id="copyResultBtn" class="btn-secondary" title="复制结果为文本">📋 复制结果</button>
            </div>
            <!-- 写入进度条 -->
            <div id="writeProgress" class="write-progress" style="display:none">
              <div class="write-progress-bar">
                <div id="writeProgressFill" class="write-progress-fill"></div>
              </div>
              <span id="writeProgressText" class="write-progress-text">正在写入...</span>
            </div>
            <div id="resultDisplay" class="result-display" style="display:none">
```

- [ ] **Step 2: 提交**

```bash
git add src/taskpane/sql-query-taskpane.html
git commit -m "feat: add progress bar HTML for write result to sheet"
```

---

### 任务 2：添加进度条 CSS 样式

**文件：**
- 修改：`src/taskpane/sql-query-taskpane.css`

- [ ] **Step 1: 在 CSS 末尾添加进度条样式**

在文件末尾（上次添加的 `@keyframes sql-spin` 之后）添加：

```css
/* —— 写入进度条 —— */

.write-progress {
  margin: 8px 0;
  display: flex;
  align-items: center;
  gap: 8px;
}

.write-progress-bar {
  flex: 1;
  height: 8px;
  background: #e0e0e0;
  border-radius: 4px;
  overflow: hidden;
}

.write-progress-fill {
  height: 100%;
  width: 0%;
  background: #0078d4;
  border-radius: 4px;
  transition: width 0.3s ease;
}

.write-progress-text {
  font-size: 11px;
  color: #666;
  white-space: nowrap;
  min-width: 120px;
}
```

- [ ] **Step 2: 提交**

```bash
git add src/taskpane/sql-query-taskpane.css
git commit -m "feat: add progress bar CSS styles for write result to sheet"
```

---

### 任务 3：重构 `writeResultToSheet()` — 分块写入 + 进度更新

**文件：**
- 修改：`src/taskpane/sql-query-taskpane.js`

当前 `writeResultToSheet()` 函数（约第 604-638 行）全部替换。

- [ ] **Step 1: 添加辅助函数**

在 `writeResultToSheet()` 之前添加新辅助函数：

```javascript
/**
 * 生成不重复的工作表名
 * @param {object} sheetCollection - Office.js worksheet 集合（已加载 items/name）
 * @param {string} baseName - 基础名称
 * @returns {string}
 */
function generateUniqueSheetName(sheetCollection, baseName) {
  var finalName = baseName;
  var counter = 1;
  var exists = true;
  while (exists) {
    exists = false;
    for (var i = 0; i < sheetCollection.items.length; i++) {
      if (sheetCollection.items[i].name === finalName) {
        exists = true;
        finalName = baseName + " (" + counter + ")";
        counter++;
        break;
      }
    }
  }
  return finalName;
}

/**
 * 更新进度条和文字
 * @param {HTMLElement} fillEl - 进度条填充元素
 * @param {HTMLElement} textEl - 进度文字元素
 * @param {number} current - 当前已写入行数
 * @param {number} total - 总行数
 */
function updateProgress(fillEl, textEl, current, total) {
  var pct = Math.round((current / total) * 100);
  fillEl.style.width = pct + "%";
  textEl.textContent = "已写入 " + current + "/" + total + " 行";
}
```

- [ ] **Step 2: 从 `writeResultToSheet()` 中提取出唯一名称生成的逻辑**

原函数中生成唯一工作表名的逻辑（`while (exists)` 循环）将被移到 `generateUniqueSheetName` 中。

- [ ] **Step 3: 替换整个 `writeResultToSheet()` 函数**

将原函数（从 `function writeResultToSheet() {` 到末尾 `}`）替换为新实现：

```javascript
function writeResultToSheet() {
  if (!currentQueryResult) return;

  showSheetNameDialog(function (sheetName) {
    var rows = currentQueryResult.rows;
    var columns = currentQueryResult.columns;
    if (!rows || rows.length === 0) {
      setStatusText("queryStatus", "没有数据可写入", "error");
      return;
    }

    var CHUNK_SIZE = 5000;
    var totalRows = rows.length;
    var totalCols = columns.length;

    var writeBtn = document.getElementById("writeSheetBtn");
    var progressContainer = document.getElementById("writeProgress");
    var progressFill = document.getElementById("writeProgressFill");
    var progressText = document.getElementById("writeProgressText");

    // 禁用按钮、显示 spinner、显示进度条
    writeBtn.disabled = true;
    writeBtn.textContent = "⏳ 写入中...";
    writeBtn.classList.add("sql-button-loading");
    progressContainer.style.display = "flex";
    updateProgress(progressFill, progressText, 0, totalRows);

    var batchIndex = 0;
    var finalSheetName = sheetName;
    var sheetCreated = false;

    function writeNextBatch() {
      var startRow = batchIndex * CHUNK_SIZE;
      if (startRow >= totalRows) {
        // 全部写入完成 → 执行 autofitColumns
        doAutoFit();
        return;
      }
      var endRow = Math.min(startRow + CHUNK_SIZE, totalRows);
      var batchSize = endRow - startRow;

      Excel.run(function (context) {
        if (!sheetCreated) {
          // 第一批：创建表 + 写表头 + 写数据
          var sheetCollection = context.workbook.worksheets;
          sheetCollection.load("items/name");
          return context.sync().then(function () {
            finalSheetName = generateUniqueSheetName(sheetCollection, sheetName);
            var newSheet = sheetCollection.add(finalSheetName);
            newSheet.position = 0;
            sheetCreated = true;

            // 写入表头 + 第一批数据行
            var rangeRows = batchSize + 1; // +1 是表头行
            var range = newSheet.getRangeByIndexes(0, 0, rangeRows, totalCols);
            var values = [columns];
            for (var r = startRow; r < endRow; r++) {
              values.push(rows[r]);
            }
            range.values = values;
            return context.sync();
          });
        } else {
          // 后续批次：通过已记录的表名获取工作表，追加数据
          var sheet = context.workbook.worksheets.getItem(finalSheetName);
          // +1 跳过表头行
          var range = sheet.getRangeByIndexes(startRow + 1, 0, batchSize, totalCols);
          var values = [];
          for (var r = startRow; r < endRow; r++) {
            values.push(rows[r]);
          }
          range.values = values;
          return context.sync();
        }
      }).then(function () {
        batchIndex++;
        updateProgress(progressFill, progressText, Math.min(endRow, totalRows), totalRows);
        writeNextBatch(); // 递归开始下一批
      }).catch(function (error) {
        restoreWriteButton(writeBtn, progressContainer);
        var msg = (error && error.message) ? error.message : String(error || "未知错误");
        setStatusText("queryStatus", "写入失败: " + msg, "error");
      });
    }

    function doAutoFit() {
      Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getItem(finalSheetName);
        var range = sheet.getUsedRange();
        range.format.autofitColumns();
        return context.sync();
      }).then(function () {
        restoreWriteButton(writeBtn, progressContainer);
        setStatusText("queryStatus", "已将 " + totalRows + " 行结果写入新工作表", "success");
      }).catch(function () {
        // autofitColumns 失败不影响数据写入，只提示
        restoreWriteButton(writeBtn, progressContainer);
        setStatusText("queryStatus", "已将 " + totalRows + " 行结果写入新工作表", "success");
      });
    }

    function restoreWriteButton(btn, progressEl) {
      btn.disabled = false;
      btn.textContent = "📝 写入新工作表";
      btn.classList.remove("sql-button-loading");
      progressEl.style.display = "none";
    }

    // 启动写入
    writeNextBatch();
  });
}
```

**重要说明：**
- `restoreWriteButton` 作为 `writeResultToSheet()` 内部函数，捕获 `writeBtn` 和 `progressContainer` 闭包变量
- 使用递归 `writeNextBatch()` 而非循环，因为每个批次是异步 `Excel.run`
- `generateUniqueSheetName` 作为独立的顶层函数（不依赖 `writeResultToSheet` 闭包），便于测试
- `updateProgress` 作为独立的顶层函数，便于测试
- `doAutoFit` 和 `restoreWriteButton` 作为 `writeResultToSheet` 内部函数，捕获闭包变量

- [ ] **Step 4: 提交**

```bash
git add src/taskpane/sql-query-taskpane.js
git commit -m "feat: chunked write with progress bar for sql query results"
```

---

### 任务 4：验证

- [ ] **Step 1: 验证构建**

```bash
npm run build:dev
```
预期：编译成功，无报错。

- [ ] **Step 2: 验证测试**

```bash
npm test
```
预期：所有测试通过。

- [ ] **Step 3: 验证 lint**

```bash
node node_modules/eslint/bin/eslint.js -c eslint.config.mjs "src/taskpane/sql-query-taskpane.js"
```
预期：无新增 lint 错误（已有 prettier 格式错误可接受）。
