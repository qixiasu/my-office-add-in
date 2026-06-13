# SQL 导出分页流式查询实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 消除 SQL 预览模式导出时全量查询导致的 UI 冻结，改为分页流式查询 + 逐批写入。

**Architecture:** 在 `writeResultToSheet()` 的预览模式分支中，将单次 `dbManager.exec(原始SQL)` 替换为 `SELECT * FROM (原始SQL) LIMIT 5000 OFFSET ?` 循环，每批查询后立即写入 Excel，通过 `setTimeout(0)` yield 给浏览器更新 UI。

**Tech Stack:** sql.js (同步 WebAssembly SQLite), Office.js Excel API, 纯 JavaScript (ES5, 无 polyfill 需求)

**改动文件:**
- 修改: `src/taskpane/sql-query-taskpane.js`
- 修改: `src/taskpane/sql-query-taskpane.css`

---

### Task 1: 添加 `buildPaginationQuery()` 辅助函数

**Files:**
- Modify: `src/taskpane/sql-query-taskpane.js` (在 `buildPreviewQuery` 函数旁，约第 1033 行附近)

- [ ] **Step 1: 在 `buildPreviewQuery` 之后添加 `buildPaginationQuery` 函数**

在 `buildPreviewQuery` 函数之后（约 1038 行）插入新函数。此函数将原始 SQL 包裹为子查询以支持 LIMIT/OFFSET：

```javascript
/**
 * 为原始 SQL 构建分页查询语句
 * 将原始 SQL 包裹为子查询以支持 LIMIT/OFFSET
 * @param {string} originalSQL - 原始 SQL
 * @param {number} limit - 每批行数
 * @param {number} offset - 偏移量
 * @returns {string}
 */
function buildPaginationQuery(originalSQL, limit, offset) {
  var sql = originalSQL.replace(/[;\s]*$/, '');
  return 'SELECT * FROM (' + sql + ') LIMIT ' + limit + ' OFFSET ' + offset;
}
```

- [ ] **Step 2: 确认函数位置正确**

确认 `buildPaginationQuery` 的放置位置在 `buildPreviewQuery` 和 `setStatusText` 之间，函数边界清晰，无语法错误。

---

### Task 2: 添加 `setWriteButtonLoading()` 和 `updateWriteProgress()` 辅助函数

**Files:**
- Modify: `src/taskpane/sql-query-taskpane.js` (在 `updateProgress` 函数旁，约第 725 行附近)

- [ ] **Step 1: 在 `updateProgress` 之后添加 `setWriteButtonLoading`**

在 `updateProgress` 函数（约 725 行）之后插入：

```javascript
/**
 * 设置导出按钮的 loading 状态
 * @param {boolean} isLoading - 是否处于加载中
 */
function setWriteButtonLoading(isLoading) {
  var btn = document.getElementById("writeSheetBtn");
  var progressContainer = document.getElementById("writeProgress");
  if (isLoading) {
    btn.disabled = true;
    btn.textContent = "⏳ 正在导出数据...";
    btn.classList.add("sql-button-loading");
    progressContainer.style.display = "flex";
    document.getElementById("writeProgressText").textContent = "正在导出数据...";
  } else {
    btn.disabled = false;
    btn.textContent = "📝 写入新工作表";
    btn.classList.remove("sql-button-loading");
    progressContainer.style.display = "none";
  }
}
```

- [ ] **Step 2: 在 `setWriteButtonLoading` 之后添加 `updateWriteProgress`**

```javascript
/**
 * 更新导出进度（仅显示绝对行数，无百分比）
 * @param {number} written - 已写入行数
 */
function updateWriteProgress(written) {
  var fillEl = document.getElementById("writeProgressFill");
  var textEl = document.getElementById("writeProgressText");
  textEl.textContent = "已写入 " + written + " 行";
}
```

注意：进度条填充 (`fillEl`) 不再设置宽度百分比——CSS 中会改为不定动画。

---

### Task 3: 添加不定进度条 CSS

**Files:**
- Modify: `src/taskpane/sql-query-taskpane.css` (在 `.write-progress-fill` 规则后，约 480 行)

- [ ] **Step 1: 为进度条填充添加不定动画**

将现有的 `.write-progress-fill` 规则从固定的宽度过渡改为不定动画：

```css
.write-progress-fill {
  height: 100%;
  width: 100%;
  background: linear-gradient(90deg, #0078d4 0%, #4da8e8 50%, #0078d4 100%);
  background-size: 200% 100%;
  animation: write-progress-indeterminate 1.5s ease infinite;
}

@keyframes write-progress-indeterminate {
  0% { background-position: 200% 0; }
  100% { background-position: -200% 0; }
}
```

注意：由于我们不知道总行数，进度条不再表示精确百分比，而是作为一个装饰性的活动指示器。后续如果任务 4 需要逐批更新进度条（如每批完成后微调），可以在 `updateWriteProgress` 中通过 `fillEl.style.width` 做装饰性调整。

---

### Task 4: 重构 `writeResultToSheet()` 中预览模式的分支

**Files:**
- Modify: `src/taskpane/sql-query-taskpane.js` (约 727-852 行)

这是核心改动。将预览模式（`isPreviewResult && currentOriginalSQL`）下的单次 `dbManager.exec()` 重写为分页循环。

- [ ] **Step 1: 删除旧的预览模式重新查询代码**

删除 `writeResultToSheet` 函数中预览模式分支内的以下代码（约 733-754 行）：

```javascript
if (isPreviewResult && currentOriginalSQL) {
  // 预览模式：重新查询全部数据
  var writeBtn = document.getElementById("writeSheetBtn");
  writeBtn.disabled = true;
  writeBtn.textContent = "⏳ 查询全部数据...";
  writeBtn.classList.add("sql-button-loading");

  var fullResult = dbManager.exec(currentOriginalSQL);
  if (fullResult.type === "error") {
    setStatusText("queryStatus", "导出失败: " + fullResult.message, "error");
    writeBtn.disabled = false;
    writeBtn.textContent = "📝 写入新工作表";
    writeBtn.classList.remove("sql-button-loading");
    return;
  }
  rows = fullResult.rows;
  columns = fullResult.columns;

  // 重置按钮状态
  writeBtn.disabled = false;
  writeBtn.textContent = "📝 写入新工作表";
  writeBtn.classList.remove("sql-button-loading");
}
```

替换为新的分页循环代码。注意保持 `else` 分支（非预览模式）完全不变。

- [ ] **Step 2: 添加分页循环代码**

将预览模式分支替换为：

```javascript
if (isPreviewResult && currentOriginalSQL) {
  // 分页流式查询 + 逐批写入（避免 UI 冻结）
  var CHUNK_SIZE = 5000;
  var offset = 0;
  var totalWritten = 0;
  var columns = null;
  var sheetCreated = false;
  var finalSheetName = sheetName;

  // 设置 loading 状态
  setWriteButtonLoading(true);
  document.getElementById("writeProgressText").textContent = "正在导出数据...";

  /**
   * 处理下一批数据：查询 5000 行 → 写入 Excel → yield → 继续
   */
  function processNextChunk() {
    // yield 给浏览器更新 UI（在同步的 sql.js 执行前）
    setTimeout(function () {
      // 1) 查询当前批
      var paginatedSQL = buildPaginationQuery(currentOriginalSQL, CHUNK_SIZE, offset);
      var result = dbManager.exec(paginatedSQL);

      if (result.type === "error") {
        setWriteButtonLoading(false);
        setStatusText("queryStatus", "导出失败: " + result.message, "error");
        return;
      }

      if (result.rowCount === 0) {
        // 2a) 没有更多数据 → 完成
        setWriteButtonLoading(false);
        setStatusText("queryStatus", "已将 " + totalWritten + " 行结果写入新工作表", "success");
        return;
      }

      if (!columns && result.columns) {
        columns = result.columns;
      }

      var chunkRows = result.rows;
      var batchStartRow = totalWritten;

      // 2b) 写入当前批到 Excel
      Excel.run(function (context) {
        if (!sheetCreated) {
          // 第一批：创建工作表 + 写入表头 + 写入数据
          var sheetCollection = context.workbook.worksheets;
          sheetCollection.load("items/name");
          return context.sync().then(function () {
            finalSheetName = generateUniqueSheetName(sheetCollection, sheetName);
            var newSheet = sheetCollection.add(finalSheetName);
            newSheet.position = 0;
            sheetCreated = true;

            var rangeRows = chunkRows.length + 1; // +1 表头
            var range = newSheet.getRangeByIndexes(0, 0, rangeRows, columns.length);
            var values = [columns];
            for (var r = 0; r < chunkRows.length; r++) {
              values.push(chunkRows[r]);
            }
            range.values = values;
            return context.sync();
          });
        } else {
          // 后续批次：追加数据
          var sheet = context.workbook.worksheets.getItem(finalSheetName);
          var range = sheet.getRangeByIndexes(batchStartRow + 1, 0, chunkRows.length, columns.length);
          range.values = chunkRows;
          return context.sync();
        }
      }).then(function () {
        // 3) 更新进度，继续下一批
        totalWritten += chunkRows.length;
        offset += CHUNK_SIZE;
        updateWriteProgress(totalWritten);
        processNextChunk();
      }).catch(function (error) {
        // 4) 写入失败
        setWriteButtonLoading(false);
        var msg = error && error.message ? error.message : String(error || "未知错误");
        setStatusText("queryStatus", "写入失败: " + msg, "error");
      });
    }, 0);
  }

  // 启动循环
  processNextChunk();
  // 注意：processNextChunk 是异步的，此处直接 return，后续流程由回调驱动
  return;
}
```

- [ ] **Step 3: 确认非预览模式分支未受改动影响**

确认 `else` 分支（约原 755-759 行）保持原样不动：

```javascript
} else {
  // 非预览模式：直接使用现有结果
  rows = currentQueryResult.rows;
  columns = currentQueryResult.columns;
}
```

注意预览分支结束后使用 `return` 退出函数，避免执行到非预览分支的后续写入代码。

---

### Task 5: 验证非预览模式分支不受影响

**Files:**
- Review: `src/taskpane/sql-query-taskpane.js` (约 755-852 行区域)

- [ ] **Step 1: 检查非预览模式代码路径**

确认 `writeResultToSheet()` 中 `else` 分支 (`isPreviewResult` 为 false 时) 的代码路径完全不变。验证：
- `rows` 和 `columns` 从 `currentQueryResult` 赋值
- `totalRows` 和 `totalCols` 计算正确
- `restoreWriteButton` 函数使用正确
- 进度条仍然显示百分比（因为非预览模式有 `totalRows`）

实际确认这部分的代码从 `writeBtn = document.getElementById("writeSheetBtn");` 开始一直到 `writeNextBatch()` 没有被修改。

- [ ] **Step 2: 检查 `restoreWriteButton` 和 `updateProgress` 未被误删**

确认 `restoreWriteButton`（约 843-848 行）和 `updateProgress`（约 720-725 行）函数仍然保留，因为非预览模式分支仍需使用它们。

---

### Task 6: 最终代码审查和提交

**Files:**
- All modified files

- [ ] **Step 1: 逐项审查改动**

审查要点：
1. `buildPaginationQuery` — SQL 包裹语法是否正确？是否处理了末尾 `;`？
2. `setWriteButtonLoading` — 按钮状态切换是否正确？进度容器显示/隐藏是否正确？
3. `updateWriteProgress` — 进度文本格式是否符合设计文档？
4. 预览分支流程 — `processNextChunk` 是否形成正确的异步循环？
5. 非预览分支 — 是否原样保留？`restoreWriteButton` 和 `updateProgress` 是否仍被非预览分支使用？
6. `setTimeout(fn, 0)` — 是否在每批查询前 yield？
7. 预览分支最后的 `return` — 是否避免了执行到非预览分支的代码？

- [ ] **Step 2: 提交改动**

```bash
git add src/taskpane/sql-query-taskpane.js src/taskpane/sql-query-taskpane.css
git commit -m "feat: implement chunked streaming query for SQL export

Replace single synchronous dbManager.exec() with paginated
LIMIT/OFFSET sub-queries. Each 5000-row chunk is queried and
written to Excel before yielding to the browser via setTimeout(0),
eliminating UI freeze during large dataset export."
```
