# SQL 查询结果按钮位置调整设计

## 概述

将 SQL 查询面板中「写入新工作表」和「复制结果」两个按钮从结果表格底部移至执行按钮下方、查询结果上方，减少用户操作步骤。

## 改动范围

只涉及两个文件：
- `src/taskpane/sql-query-taskpane.html` — 调整 DOM 结构
- `src/taskpane/sql-query-taskpane.css` — 可能无需改动（`result-actions` 样式已在全局生效）

不涉及 JS 逻辑改动。

## 具体方案

### HTML 结构调整

**当前结构（简化）：**
```html
<div class="query-actions">
  <button id="executeBtn">▶ 执行</button>
  <button id="clearSqlBtn">🧹 清空</button>
</div>

<div id="queryStatus" class="status-message"></div>

<div class="section-divider">查询结果</div>
<div id="resultDisplay" class="result-display" style="display:none">
  <table class="data-table">...</table>
  <div class="result-actions">
    <button id="writeSheetBtn">📝 写入新工作表</button>
    <button id="copyResultBtn">📋 复制结果</button>
  </div>
</div>
```

**改动后结构：**
```html
<div class="query-actions">
  <button id="executeBtn">▶ 执行</button>
  <button id="clearSqlBtn">🧹 清空</button>
</div>

<div id="queryStatus" class="status-message"></div>

<!-- 结果操作按钮移到结果表格上方 -->
<div id="resultActions" class="result-actions" style="display:none">
  <button id="writeSheetBtn">📝 写入新工作表</button>
  <button id="copyResultBtn">📋 复制结果</button>
</div>

<div class="section-divider">查询结果</div>
<div id="resultDisplay" class="result-display" style="display:none">
  <table class="data-table">...</table>
</div>
```

### JS 联动改动

在 `sql-query-taskpane.js` 的 `runQuery()` 函数中，当显示查询结果时，需要同时显示 `resultActions` 区域：

```js
// 当前：只显示 resultDisplay
document.getElementById("resultDisplay").style.display = "block";

// 改动后：同时显示 resultActions
document.getElementById("resultActions").style.display = "block";
document.getElementById("resultDisplay").style.display = "block";
```

在 `clearSql()` 函数中，需要同时隐藏 `resultActions`。

## 注意事项

1. 按钮容器 `resultActions` 的 `display:none` 初始状态与 `resultDisplay` 保持一致
2. 当无结果时（如 DROP/DELETE/UPDATE 操作后），按钮区域应自动隐藏
3. `result-actions` 的 CSS 样式（flex 布局、gap、flex:1）已存在，无需新增
