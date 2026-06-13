# SQL 查询结果写入工作表 — 分块进度反馈设计

**日期**: 2026-06-13
**状态**: 已批准

## 问题描述

在「数据库查询」功能中执行 SQL 查询后，点击「写入新工作表」将大量结果写入 Excel 时，界面没有任何反馈。按钮不变灰、不显示进度，用户不确定写入是否正在进行还是已经完成。

**根因**:
- `writeResultToSheet()` 一次构建整个大数组并一次性写入 Excel
- 对大结果集（如 42 万行），数组构建和 Office.js 序列化传输耗时数秒
- 期间没有任何 UI 反馈

## 设计目标

1. **防止重复点击**：写入按钮在执行期间变灰禁用
2. **即时视觉反馈**：按钮在写入开始后立即改变状态（禁用 + 文字变化 + 旋转动画）
3. **行数进度显示**：逐批写入时，状态栏实时显示「已写入 X/Y 行」
4. **进度条**：视觉进度条直观反映完成百分比

## 原理

将写入操作拆分为多个批次（每批 5000 行），每个批次使用独立的 `Excel.run` 调用。批次之间主线程空闲，可以更新 DOM（进度条、状态文字）。CSS `@keyframes` 动画在 Office.js 异步操作期间持续运行。

```
Excel.run #1 → 创建表 + 写入 5000 行 → 更新进度 → Excel.run #2 → 写入 5000 行 → 更新进度 → ...
```

## 改动范围

### 1. `src/taskpane/sql-query-taskpane.html` — 新增进度条元素

在「查询结果」区域添加进度条组件：

```html
<!-- 在 resultActions 之后、resultDisplay 之前 -->
<div id="writeProgress" class="write-progress" style="display:none">
  <div class="write-progress-bar">
    <div id="writeProgressFill" class="write-progress-fill"></div>
  </div>
  <span id="writeProgressText" class="write-progress-text">正在写入...</span>
</div>
```

### 2. `src/taskpane/sql-query-taskpane.css` — 进度条样式

```css
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

复用已有 `.sql-button-loading` spinner 类（前次改动已添加）。

### 3. `src/taskpane/sql-query-taskpane.js` — 重写 `writeResultToSheet()`

核心变更：

**3a. 新增辅助函数：**
- `generateUniqueSheetName(sheetCollection, baseName)` — 生成不重复的工作表名
- `updateProgress(fillEl, textEl, current, total)` — 更新进度条和文字
- `restoreWriteButton()` — 恢复写入按钮状态、隐藏进度条

**3b. 重构 `writeResultToSheet()`：**
- 在 `showSheetNameDialog` callback 中获取按钮和进度条元素引用
- 禁用按钮、改文字、添加 CSS 类、显示进度条
- 使用递归 `writeNextBatch()` 函数逐批写入
- 第一批 `Excel.run`：加载工作表列表 → 创建新表 → 写表头 + 第一批数据
- 后续批次 `Excel.run`：通过 `getItem` 获取已创建的表 → 追加数据
- 每批完成后更新进度条
- 全部完成后执行最后一次 `Excel.run` 调用 `autofitColumns()`
- 完成后恢复按钮、隐藏进度条、显示成功信息

**3c. 分块参数：**
- `CHUNK_SIZE = 5000`（每批 5000 行，可根据实际性能调整）
- 最后一批自动处理不足 5000 行的剩余数据

## 边界情况处理

| 场景 | 处理方式 |
|------|----------|
| 用户取消命名对话框 | 在 callback 前 return，不触发按钮变化 |
| 写入中途出错 | catch 中恢复按钮 + 隐藏进度条 + 显示错误 |
| 结果集为空 | 现有 `if (!currentQueryResult) return;` 已处理 |
| 自动列宽失败 | 单独 catch，不影响数据写入结果 |
| 快速连续点击 | 按钮已 disabled，不会重复触发 |
| 超大数据行数 | 数据来自查询结果，受 SQLite 和 Excel 行数限制 |

## 测试要点

1. 点击写入 → 弹出命名对话框 → 确定 → 按钮禁用 + spinner + 进度条显示
2. 进度条随分块写入逐步增长
3. 写入完成后进度条隐藏、按钮恢复、显示成功信息
4. 命名对话框取消 → 按钮不变化
5. 写入出错 → 按钮恢复、进度条隐藏、显示错误
