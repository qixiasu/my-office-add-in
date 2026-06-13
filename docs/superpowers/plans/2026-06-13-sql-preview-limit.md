# SQL 查询预览限制 — 实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 在 SQL 查询功能中实现自动 LIMIT 200 预览，导出时重新查询全部数据，复制时仅复制预览行并明确告知用户。

**Architecture:** 核心逻辑集中在 `sql-query-taskpane.js`，通过新增 SQL 判断工具函数 + 修改三个主要函数（`runQuery`, `writeResultToSheet`, `copyResult`）实现。sql.js 的 `exec()` 方法不变，`sql-utils.js` 无需修改。

**Tech Stack:** 纯 JavaScript + sql.js (WASM) + Office.js

**相关文档:** [docs/superpowers/specs/2026-06-13-sql-preview-limit-design.md](../specs/2026-06-13-sql-preview-limit-design.md)

---

### Task 1: 添加 SQL 判断工具函数

**Files:**
- Modify: `src/taskpane/sql-query-taskpane.js` — 在 `renderQueryHistory()` 之后、文件末尾之前添加三个辅助函数

这组函数不会被导出，仅作为模块内部工具。它们封装了"是否 SELECT"、"是否已有 LIMIT"、"构建预览 SQL"三个操作。

- [ ] **Step 1: 添加 `isSelectQuery` 函数**

```javascript
/**
 * 判断 SQL 是否为 SELECT 查询
 * @param {string} sql - 原始 SQL
 * @returns {boolean}
 */
function isSelectQuery(sql) {
  return /^\s*SELECT\b/i.test(sql.trim());
}
```

- [ ] **Step 2: 添加 `hasExplicitLimit` 函数**

```javascript
/**
 * 判断 SQL 是否已包含 LIMIT 子句
 * 注意：不解析字符串字面量中的 LIMIT，实际使用中覆盖 99% 场景
 * @param {string} sql - 原始 SQL
 * @returns {boolean}
 */
function hasExplicitLimit(sql) {
  return /\bLIMIT\b/i.test(sql);
}
```

- [ ] **Step 3: 添加 `buildPreviewQuery` 函数**

```javascript
/**
 * 为 SELECT 查询构建预览版本（自动追加 LIMIT）
 * 如果已有 LIMIT 或非 SELECT，返回原 SQL
 * @param {string} sql - 原始 SQL
 * @param {number} limit - 预览行数上限
 * @returns {string}
 */
function buildPreviewQuery(sql, limit) {
  if (isSelectQuery(sql) && !hasExplicitLimit(sql)) {
    return sql.replace(/;?\s*$/, '') + ' LIMIT ' + limit;
  }
  return sql;
}
```

- [ ] **Step 4: 提交**

```bash
git add src/taskpane/sql-query-taskpane.js
git commit -m "feat: add SQL helper functions for preview LIMIT detection"
```

---

### Task 2: 新增预览状态变量并修改 `runQuery()`

**Files:**
- Modify: `src/taskpane/sql-query-taskpane.js` — 变量声明区（~第 12 行附近）+ `runQuery()` 函数（~第 440-528 行）

**改动概要：**
1. 新增两个状态变量 `currentOriginalSQL` 和 `isPreviewResult`
2. 修改 `runQuery()`：执行前判断是否需要追加 LIMIT，执行后标记预览状态

- [ ] **Step 1: 新增状态变量**

将原有变量声明区（~第 11-12 行）：

```javascript
var currentQueryResult = null;
```

替换为：

```javascript
var currentQueryResult = null;
var currentOriginalSQL = null;   // 存储原始 SQL（无 LIMIT 追加），供导出时重新查询
var isPreviewResult = false;     // 标记当前结果是否为截断预览
```

- [ ] **Step 2: 修改 `runQuery()`——执行前追加 LIMIT**

找到 `runQuery()` 函数中 `try` 块内的 SQL 执行部分（~第 469 行）：

```javascript
try {
    var result = dbManager.exec(sql);
```

替换为：

```javascript
try {
    // 预览模式：SELECT 无 LIMIT 时自动追加 LIMIT 200
    var previewSQL = sql;
    currentOriginalSQL = sql;
    isPreviewResult = false;

    if (isSelectQuery(sql) && !hasExplicitLimit(sql)) {
      previewSQL = buildPreviewQuery(sql, 200);
    }

    var result = dbManager.exec(previewSQL);
```

- [ ] **Step 3: 修改 `runQuery()`——结果处理后标记预览状态**

找到 SELECT 结果处理段，在 `currentQueryResult = result;` 之后（~第 494 行）添加预览标记逻辑。

原代码：

```javascript
    currentQueryResult = result;

    // 渲染结果表格
```

改为：

```javascript
    currentQueryResult = result;

    // 判断是否为截断预览：自动加了 LIMIT 且结果恰好等于上限
    if (isSelectQuery(sql) && !hasExplicitLimit(sql) && result.rowCount === 200) {
      isPreviewResult = true;
    }

    // 渲染结果表格
```

- [ ] **Step 4: 修改 `runQuery()`——状态消息区分预览模式**

找到状态消息设置部分（~第 491 行）：

```javascript
    statusEl.textContent = "查询完成，返回 " + result.rowCount + " 行 (" + result.elapsed.toFixed(2) + " 秒)";
    statusEl.className = "status-message status-success";
```

替换为：

```javascript
    if (isPreviewResult) {
      statusEl.textContent = "预览前 " + result.rowCount + " 行 (" + result.elapsed.toFixed(2) + " 秒)";
    } else if (hasExplicitLimit(sql)) {
      statusEl.textContent = "查询完成，返回 " + result.rowCount + " 行（用户指定 LIMIT）(" + result.elapsed.toFixed(2) + " 秒)";
    } else {
      statusEl.textContent = "查询完成，返回 " + result.rowCount + " 行 (" + result.elapsed.toFixed(2) + " 秒)";
    }
    statusEl.className = "status-message status-success";
```

- [ ] **Step 5: 修改 `runQuery()`——清空时重置状态变量**

找到 `runQuery()` 中当 `result.type === "modification"` 时的重置代码（~第 482 行）：

```javascript
      currentQueryResult = null;
```

在 `clearSql()` 函数中也找到类似的 `currentQueryResult = null;`（~第 536 行）。

确认两处都重置了新增变量：

`runQuery()` 中 modification 分支（~第 482 行）：
```javascript
      currentQueryResult = null;
      currentOriginalSQL = null;
      isPreviewResult = false;
```

`clearSql()` 函数中（~第 536 行）：
```javascript
      currentQueryResult = null;
      currentOriginalSQL = null;
      isPreviewResult = false;
```

- [ ] **Step 6: 提交**

```bash
git add src/taskpane/sql-query-taskpane.js
git commit -m "feat: auto-append LIMIT 200 to SELECT queries without explicit LIMIT"
```

---

### Task 3: 修改 `writeResultToSheet()`——预览时重新查询全量

**Files:**
- Modify: `src/taskpane/sql-query-taskpane.js` — `writeResultToSheet()` 函数（~第 642-738 行）

**改动概要：** 在写入逻辑开始前，判断 `isPreviewResult` 是否为 true。如果是，先重新执行原始 SQL 获取全部数据，再写入。

- [ ] **Step 1: 在写入前插入重新查询逻辑**

找到 `writeResultToSheet()` 函数中对话框 callback 开始处（~第 645 行）：

```javascript
  showSheetNameDialog(function (sheetName) {
    var rows = currentQueryResult.rows;
    var columns = currentQueryResult.columns;
```

替换为：

```javascript
  showSheetNameDialog(function (sheetName) {
    var rows, columns;

    if (isPreviewResult && currentOriginalSQL) {
      // 预览模式：重新查询全部数据
      var writeBtn = document.getElementById("writeSheetBtn");
      var prevText = writeBtn.textContent;
      writeBtn.disabled = true;
      writeBtn.textContent = "⏳ 查询全部数据...";
      writeBtn.classList.add("sql-button-loading");

      var fullResult = dbManager.exec(currentOriginalSQL);
      if (fullResult.type === "error") {
        setStatusText("queryStatus", "导出失败: " + fullResult.message, "error");
        writeBtn.disabled = false;
        writeBtn.textContent = prevText;
        writeBtn.classList.remove("sql-button-loading");
        return;
      }
      rows = fullResult.rows;
      columns = fullResult.columns;

      writeBtn.textContent = prevText;
      writeBtn.classList.remove("sql-button-loading");
    } else {
      // 非预览模式：直接使用现有结果
      rows = currentQueryResult.rows;
      columns = currentQueryResult.columns;
    }
```

注意：后面原有代码中 `rows = currentQueryResult.rows` 和 `columns = currentQueryResult.columns` 已被提前赋值，确保后续代码中的 `rows`/`columns` 变量可访问。将原来 `showSheetNameDialog` 回调中 `var rows = currentQueryResult.rows; var columns = currentQueryResult.columns;` 那一行去掉，替换为 `var rows, columns;`，然后在 if/else 中分别赋值。

- [ ] **Step 2: 确认后续代码无需修改**

`writeResultToSheet()` 中后续的写入逻辑（分块 5000 行/批、创建 sheet、进度条等）使用的是 `rows` 和 `columns` 变量，现在无论预览还是非预览模式都已正确赋值，后续代码无需改动。

- [ ] **Step 3: 提交**

```bash
git add src/taskpane/sql-query-taskpane.js
git commit -m "feat: re-query full dataset on export when result is preview-limited"
```

---

### Task 4: 修改 `copyResult()`——预览行数提示

**Files:**
- Modify: `src/taskpane/sql-query-taskpane.js` — `copyResult()` 函数（~第 740-760 行）

**改动概要：** 复制的内容不变（始终是 `currentQueryResult.rows`，预览模式下就是 200 行），但状态提示消息区分预览和非预览场景。

- [ ] **Step 1: 替换状态提示消息**

找到 `copyResult()` 中的成功回调：

```javascript
  navigator.clipboard.writeText(text).then(function () {
    setStatusText("queryStatus", "已复制 " + currentQueryResult.rows.length + " 行到剪贴板", "success");
  }).catch(function () {
```

替换为：

```javascript
  navigator.clipboard.writeText(text).then(function () {
    var copyMsg;
    if (isPreviewResult) {
      copyMsg = "已复制前 " + currentQueryResult.rows.length + " 行，如需全部数据请使用「写入新工作表」";
    } else {
      copyMsg = "已复制 " + currentQueryResult.rows.length + " 行到剪贴板";
    }
    setStatusText("queryStatus", copyMsg, "success");
  }).catch(function () {
```

- [ ] **Step 2: 同样修改 fallback 分支**

在同一个函数中找到 fallback 分支（~第 755 行）：

```javascript
    document.execCommand("copy");
    document.body.removeChild(textarea);
    setStatusText("queryStatus", "已复制到剪贴板", "success");
```

替换为：

```javascript
    document.execCommand("copy");
    document.body.removeChild(textarea);
    var copyMsg;
    if (isPreviewResult) {
      copyMsg = "已复制前 " + currentQueryResult.rows.length + " 行，如需全部数据请使用「写入新工作表」";
    } else {
      copyMsg = "已复制 " + currentQueryResult.rows.length + " 行到剪贴板";
    }
    setStatusText("queryStatus", copyMsg, "success");
```

- [ ] **Step 3: 提交**

```bash
git add src/taskpane/sql-query-taskpane.js
git commit -m "feat: show preview-limited copy notification with full-data hint"
```

---

### Task 5: 自测清单

- [ ] **Step 1: 执行 `SELECT * FROM small_table`（< 200 行）**
  - 预期：正常返回全部行，状态显示 `查询完成，返回 N 行`
  - 导出：直接使用现有结果（不重新查询）
  - 复制：`已复制 N 行到剪贴板`

- [ ] **Step 2: 执行 `SELECT * FROM large_table`（> 200 行）**
  - 预期：快速返回，显示 200 行，状态显示 `预览前 200 行`
  - 导出：先显示 `正在查询全部数据...`，然后写入进度条
  - 复制：`已复制前 200 行，如需全部数据请使用「写入新工作表」`

- [ ] **Step 3: 执行 `SELECT * FROM large_table LIMIT 500`**
  - 预期：返回 500 行（用户指定 LIMIT），状态显示 `查询完成，返回 500 行（用户指定 LIMIT）`
  - 导出：直接使用现有结果（不重新查询，因为不是预览截断）
  - 复制：`已复制 500 行到剪贴板`

- [ ] **Step 4: 执行 INSERT/UPDATE/DELETE/DROP**
  - 预期：行为不变，`currentOriginalSQL` 和 `isPreviewResult` 重置
  - 导出按钮应隐藏（现有行为）

- [ ] **Step 5: 清空后重新查询**
  - 点「清空」按钮 → 确认所有状态变量重置
  - 再点执行 → 正常

- [ ] **Step 6: 构建验证**
  ```bash
  npm run build:dev
  ```
  - 预期：构建成功，无报错
