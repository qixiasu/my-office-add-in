# SQL 查询结果按钮位置调整 实施计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 将"写入新工作表"和"复制结果"按钮从结果表格底部移到执行按钮下方、查询结果上方，方便用户即时点击。

**架构:** 纯前端 DOM 结构调整 + JS 联动。按钮容器 `resultActions` 从 `resultDisplay` 内部抽出，独立放在二者之间，与 `resultDisplay` 同步显示/隐藏。

**Tech Stack:** HTML + JavaScript (Office Add-in)

---

## 文件结构

| 文件 | 操作 | 说明 |
|------|------|------|
| `src/taskpane/sql-query-taskpane.html` | 修改 | 将 `.result-actions` 从 `#resultDisplay` 内部移出 |
| `src/taskpane/sql-query-taskpane.js` | 修改 | 3 处 JS 增加对 `#resultActions` 的显示/隐藏控制 |
| `src/taskpane/sql-query-taskpane.css` | 无需修改 | `.result-actions` / `.result-actions button` 样式已存在 |

---

### Task 1: 调整 HTML 按钮位置

**Files:**
- Modify: `src/taskpane/sql-query-taskpane.html:129-139`

- [ ] **Step 1: 将 result-actions 从 resultDisplay 内部移出**

    当前代码（第 129-139 行）：

    ```html
            <div class="section-divider">查询结果</div>
            <div id="resultDisplay" class="result-display" style="display:none">
                <table class="data-table">
                    <thead id="resultHead"></thead>
                    <tbody id="resultBody"></tbody>
                </table>
                <div class="result-actions">
                    <button id="writeSheetBtn" class="btn-secondary" title="将查询结果写入新的工作表">📝 写入新工作表</button>
                    <button id="copyResultBtn" class="btn-secondary" title="复制结果为文本">📋 复制结果</button>
                </div>
            </div>
    ```

    改为：

    ```html
            <div class="section-divider">查询结果</div>
            <div id="resultActions" class="result-actions" style="display:none">
                <button id="writeSheetBtn" class="btn-secondary" title="将查询结果写入新的工作表">📝 写入新工作表</button>
                <button id="copyResultBtn" class="btn-secondary" title="复制结果为文本">📋 复制结果</button>
            </div>
            <div id="resultDisplay" class="result-display" style="display:none">
                <table class="data-table">
                    <thead id="resultHead"></thead>
                    <tbody id="resultBody"></tbody>
                </table>
            </div>
    ```

    关键变化：
    - 按钮容器从 `resultDisplay` 内部提到 `section-divider`（"查询结果"标题）之后、`resultDisplay` 之前
    - 容器 id 从无改为 `resultActions`（以便 JS 定位）
    - 保留 `style="display:none"` 初始隐藏

- [ ] **Step 2: 验证 HTML 结构**

    检查项目：`writeSheetBtn` 和 `copyResultBtn` 的 id 保持不变（JS 绑定不中断），`resultActions` 位于 `queryStatus` 和 `resultDisplay` 之间。

---

### Task 2: JS 联动控制按钮区域显示/隐藏

**Files:**
- Modify: `src/taskpane/sql-query-taskpane.js:414,448,459`

- [ ] **Step 1: 在修改类操作后隐藏按钮区域（第 414 行附近）**

    当前代码（第 414 行附近）：

    ```js
        document.getElementById("resultDisplay").style.display = "none";
        currentQueryResult = null;
    ```

    改为：

    ```js
        document.getElementById("resultActions").style.display = "none";
        document.getElementById("resultDisplay").style.display = "none";
        currentQueryResult = null;
    ```

    说明：DROUP/DELETE/UPDATE 执行后不产生查询结果表格，按钮区也应隐藏。

- [ ] **Step 2: 在 SELECT 查询后显示按钮区域（第 448 行附近）**

    当前代码（第 448 行附近）：

    ```js
        document.getElementById("resultDisplay").style.display = "block";
    ```

    改为：

    ```js
        document.getElementById("resultActions").style.display = "block";
        document.getElementById("resultDisplay").style.display = "block";
    ```

    说明：SELECT 查询渲染结果表格后，同时显示上方的操作按钮。

- [ ] **Step 3: 在清空时隐藏按钮区域（第 459 行附近）**

    当前代码（第 459 行附近）：

    ```js
        document.getElementById("resultDisplay").style.display = "none";
    ```

    改为：

    ```js
        document.getElementById("resultActions").style.display = "none";
        document.getElementById("resultDisplay").style.display = "none";
    ```

    说明：点击"清空"或切换标签页清理 SQL 面板时，按钮区也应隐藏。

---

### Task 3: 验证改动

- [ ] **Step 1: 检查 `resultActions` 初始状态**

    确认 `resultActions` 的初始 `style="display:none"` 与 `resultDisplay` 一致，打开任务面板时按钮不显示。

- [ ] **Step 2: 检查 JS 绑定不中断**

    确认 `writeSheetBtn` 和 `copyResultBtn` 的 id 未变化，事件绑定（`bindEvents` 中第 74-75 行）正常。

- [ ] **Step 3: 构建验证**

    ```bash
    npm run build:dev
    ```

    确认无构建错误。

- [ ] **Step 4: 提交**

    ```bash
    git add src/taskpane/sql-query-taskpane.html src/taskpane/sql-query-taskpane.js
    git commit -m "feat: move query result action buttons above result table"
    ```

    提交信息说明：将"写入新工作表"和"复制结果"按钮从结果表格底部移到表格上方，方便用户立即点击操作。
