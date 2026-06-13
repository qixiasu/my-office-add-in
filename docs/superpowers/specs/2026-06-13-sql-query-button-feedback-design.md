# SQL 查询执行按钮反馈优化设计

**日期**: 2026-06-13
**状态**: 已批准

## 问题描述

在「数据库查询」功能中执行 SQL 查询时，如果查询的表较大，点击「执行」按钮后会有明显的延迟（主线程被 sql.js 同步 API 阻塞），但执行按钮**没有变化**（不变灰、不改变文字），导致用户疑惑是否真正点击了按钮。

**根因**:
- `dbManager.exec(sql)` 调用 sql.js 的 `db.exec()`，这是一个**同步阻塞**调用
- 在同步执行期间，JavaScript 主线程被占用，浏览器无法处理 UI 更新
- 当前代码没有在执行前修改按钮的状态

## 设计目标

1. **防止重复点击**：执行按钮在执行期间变灰禁用
2. **即时视觉反馈**：按钮在点击后立即改变状态（禁用 + 文字变化 + 旋转动画）
3. **执行耗时显示**：查询结束后显示总耗时

## 原理

CSS `@keyframes` 动画运行在浏览器的**合成线程**（Compositor Thread），不受主线程 JavaScript 阻塞影响。即使 `db.exec()` 同步阻塞主线程数秒，按钮上的旋转动画仍然持续渲染。

结合 `await new Promise(r => setTimeout(r, 0))` 模式，在执行同步查询前先让浏览器完成一次渲染（将按钮的 disabled 状态绘制出来），然后才开始阻塞操作。

## 改动范围

涉及 3 个文件，均只需少量修改：

### 1. `src/taskpane/sql-query-taskpane.js` — `runQuery()` 函数

在执行 `dbManager.exec()` 之前：
- 获取执行按钮元素引用
- 禁用按钮：`executeBtn.disabled = true`
- 改变按钮文字：`executeBtn.textContent = "⏳ 执行中..."`
- 添加 CSS 类激活旋转动画：`executeBtn.classList.add("sql-button-loading")`
- 添加 `await new Promise(r => setTimeout(r, 0))` 让浏览器先渲染按钮状态
- 记录开始时间：`var startTime = performance.now()`

执行之后、在所有返回路径上：
- 恢复按钮：`executeBtn.disabled = false; executeBtn.textContent = "▶ 执行"`
- 移除 CSS 类：`executeBtn.classList.remove("sql-button-loading")`
- 在状态栏显示实际耗时：`result.elapsed.toFixed(2) + "秒"`

**注意**：危险操作二次确认（DROP/DELETE/UPDATE 的 `showConfirm`）必须在按钮禁用之前执行，当前代码的顺序已经保证这一点。

### 2. `src/taskpane/sql-query-taskpane.css` — 新增 loading 样式和动画

```css
/* 按钮内嵌旋转动画 */
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

/* 已有 .sql-button-primary:disabled 不动，保持灰色 */
```

### 3. 无需改动的文件

- `src/taskpane/sql-query-taskpane.html` — 不需要改动（按钮已存在）
- `src/utils/sql-utils.js` — 不需要改动（数据库逻辑不变）

## 边界情况处理

| 场景 | 处理方式 |
|------|----------|
| 快速连续按执行 | 按钮已 disabled，不会重复进入 |
| SQL 空输入 | 在按钮禁用前已 return |
| DROP/DELETE/UPDATE 二次确认 | 确认弹窗在按钮禁用之前执行 |
| 查询执行出错 | catch 块中恢复按钮状态 |
| 查询途中切换 Tab | 同步执行结束后按钮恢复，状态正常 |
| 浏览器页面关闭 | 无需特殊处理 |

## 测试要点

1. 点击执行按钮 → 按钮立即变灰、文字变 "⏳ 执行中..."、旋转动画出现
2. 按钮禁用期间无法再次点击
3. 查询结束后按钮恢复、文字恢复、显示耗时
4. 空输入不触发任何变化
5. DROP/DELETE 二次确认取消后，按钮不变化
6. 查询出错时按钮恢复
