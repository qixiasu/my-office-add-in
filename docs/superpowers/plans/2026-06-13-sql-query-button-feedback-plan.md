# SQL 查询执行按钮反馈优化 — 实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 在 SQL 查询执行时让按钮变灰禁用、显示旋转动画和文字变化，查询结束后恢复并显示耗时。

**Architecture:** 利用 CSS 合成线程动画不受主线程同步阻塞影响的特性，在执行 dbManager.exec() 前立即修改按钮状态并让浏览器渲染，执行结束后恢复。改少量仅 2 个文件，不改动业务逻辑。

**Tech Stack:** CSS3 `@keyframes` 动画、`async/await` + `setTimeout(0)` 让出控制权模式

---

### 任务 1：添加 CSS 旋转动画样式

**文件：**
- 修改：`src/taskpane/sql-query-taskpane.css`

- [ ] **Step 1: 在 CSS 末尾添加 loading 状态样式和动画**

在 `sql-query-taskpane.css` 末尾（`.confirm-buttons .sql-button` 选择器之后）添加以下内容：

```css
/* —— 执行按钮 loading 状态 —— */

.sql-button-loading {
  position: relative;
  padding-left: 28px;
}

.sql-button-loading::before {
  content: '';
  position: absolute;
  left: 10px;
  top: 50%;
  width: 14px;
  height: 14px;
  margin-top: -7px;
  border: 2px solid rgba(255, 255, 255, 0.3);
  border-top-color: #fff;
  border-radius: 50%;
  animation: sql-spin 0.8s linear infinite;
}

@keyframes sql-spin {
  to { transform: rotate(360deg); }
}
```

- [ ] **Step 2: 提交**

```bash
git add src/taskpane/sql-query-taskpane.css
git commit -m "feat: add CSS spinner animation for sql query execute button"
```

---

### 任务 2：修改 `runQuery()` 函数 — 按钮禁用与恢复

**文件：**
- 修改：`src/taskpane/sql-query-taskpane.js`

关键逻辑：
1. 在二次确认之后、`exec()` 之前 → 禁用按钮、改文字、加 CSS 类
2. 用 `await new Promise(r => setTimeout(r, 0))` 让浏览器渲染按钮状态
3. 在所有 3 个退出路径（error、modification、query）上恢复按钮

- [ ] **Step 1: 获取按钮引用并添加禁用逻辑**

找到 `runQuery()` 函数（约第 440 行）。在 `statusEl.className = "status-message status-loading";` 这一行之前，添加按钮的获取和禁用：

原始代码：
```javascript
async function runQuery() {
  var sqlInput = document.getElementById("sqlInput");
  var sql = sqlInput.value.trim();
  if (!sql) return;

  var statusEl = document.getElementById("queryStatus");

  // DROP/DELETE/UPDATE 二次确认
  var upperSql = sql.toUpperCase().trim();
  if (upperSql.startsWith("DROP") || upperSql.startsWith("DELETE") || upperSql.startsWith("UPDATE")) {
    var confirmed = await showConfirm("确定要执行危险操作吗？\n\n" + sql);
    if (!confirmed) {
      return;
    }
  }

  statusEl.className = "status-message status-loading";
  statusEl.textContent = "执行中...";

  var result = dbManager.exec(sql);
  // ... 后续代码
```

修改为：
```javascript
async function runQuery() {
  var sqlInput = document.getElementById("sqlInput");
  var sql = sqlInput.value.trim();
  if (!sql) return;

  var statusEl = document.getElementById("queryStatus");
  var executeBtn = document.getElementById("executeBtn");

  // DROP/DELETE/UPDATE 二次确认
  var upperSql = sql.toUpperCase().trim();
  if (upperSql.startsWith("DROP") || upperSql.startsWith("DELETE") || upperSql.startsWith("UPDATE")) {
    var confirmed = await showConfirm("确定要执行危险操作吗？\n\n" + sql);
    if (!confirmed) {
      return;
    }
  }

  // —— 禁用按钮、提供视觉反馈 ——
  executeBtn.disabled = true;
  executeBtn.textContent = "⏳ 执行中...";
  executeBtn.classList.add("sql-button-loading");

  statusEl.className = "status-message status-loading";
  statusEl.textContent = "执行中...";

  // 让浏览器渲染按钮的 disabled 状态后再执行同步查询
  await new Promise(function (resolve) { setTimeout(resolve, 0); });

  var result = dbManager.exec(sql);
  // ... 后续代码
```

- [ ] **Step 2: 在所有退出路径恢复按钮**

现在代码中有 3 个地方需要恢复按钮状态。

**路径 1：error 类型结果**（约第 461-465 行）

原始代码：
```javascript
  if (result.type === "error") {
    statusEl.className = "status-message status-error";
    statusEl.textContent = "错误: " + result.message;
    return;
  }
```

修改为：
```javascript
  if (result.type === "error") {
    executeBtn.disabled = false;
    executeBtn.textContent = "▶ 执行";
    executeBtn.classList.remove("sql-button-loading");
    statusEl.className = "status-message status-error";
    statusEl.textContent = "错误: " + result.message;
    return;
  }
```

**路径 2：modification 类型结果**（约第 467-478 行）

原始代码：
```javascript
  if (result.type === "modification") {
    statusEl.textContent = "完成，影响 " + result.rowsAffected + " 行 (" + result.elapsed.toFixed(2) + " 秒)";
    statusEl.className = "status-message status-success";
    // ...
    return;
  }
```

修改为：
```javascript
  if (result.type === "modification") {
    executeBtn.disabled = false;
    executeBtn.textContent = "▶ 执行";
    executeBtn.classList.remove("sql-button-loading");
    statusEl.textContent = "完成，影响 " + result.rowsAffected + " 行 (" + result.elapsed.toFixed(2) + " 秒)";
    statusEl.className = "status-message status-success";
    // ...
    return;
  }
```

**路径 3：query 结果**（约第 480-513 行）。在函数末尾（`addQueryHistory` 之后）添加恢复代码：

原始代码：
```javascript
  addQueryHistory(sql, result.type, result.elapsed, result.rowCount);
}
```

修改为（在 `addQueryHistory` 之后、函数末尾前）：
```javascript
  addQueryHistory(sql, result.type, result.elapsed, result.rowCount);

  executeBtn.disabled = false;
  executeBtn.textContent = "▶ 执行";
  executeBtn.classList.remove("sql-button-loading");
}
```

- [ ] **Step 3: 提交**

```bash
git add src/taskpane/sql-query-taskpane.js
git commit -m "feat: disable execute button during sql query and show loading feedback"
```

---

### 任务 3：验证

- [ ] **Step 1: 验证构建无报错**

```bash
npm run build:dev
```

- [ ] **Step 2: 验证 lint**

```bash
npm run lint
```

- [ ] **Step 3: 验证测试**

```bash
npm test
```
