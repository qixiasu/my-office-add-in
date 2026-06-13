# SQL 查询预览限制设计

**日期**: 2026-06-13
**状态**: 已批准
**涉及项目**: My Office Add-in

## 背景

SQL 查询功能目前执行完整查询后才返回结果。当查询大表（如 40 万行、93 列）时，执行耗时很长，
用户需要等待全部数据返回才能看到任何结果。这严重拖慢了调试 SQL 脚本的效率。

**目标**: 类似 DBeaver 等数据库工具的行为——自动限制预览行数，让用户快速看到结果，
同时全量导出不受限制。

## 设计决策

| 决策 | 结论 |
|------|------|
| 预览限制行数 | 200 行 |
| 用户自写 LIMIT | 尊重原文，不做追加 |
| 导出策略 | 总是导出全部数据 |
| 总行数显示 | 不显示（避免额外 COUNT 查询阻塞 UI） |
| 复制行为 | 复制预览的 200 行，明确告知用户 |

## 流程设计

### 1. 查询执行流

```
用户点击「执行」
  │
  ├─ 危险操作（DROP/DELETE/UPDATE）? → 二次确认
  │
  ├─ SELECT 语句?
  │     ├─ 已有 LIMIT 子句 → 按原文执行，标记 `isPreview = false`
  │     └─ 无 LIMIT 子句 → 自动追加 `LIMIT 200`，标记 `isPreview = true`
  │
  ├─ 执行结果
  │     ├─ 错误 → 显示错误
  │     ├─ 修改语句（INSERT/UPDATE/DELETE）→ 显示影响行数，无变化
  │     └─ 查询结果 →
  │           ├─ 渲染结果表格（最多 200 行）
  │           ├─ 存储原始 SQL → `currentOriginalSQL`
  │           └─ 更新状态栏：仅提示预览前 200 行，不额外查询总行数
  │
  └─ 完成
```

### 2. LIMIT 追加逻辑

```javascript
// 判断是否为 SELECT
function isSelectQuery(sql) {
  return /^\s*SELECT\b/i.test(sql.trim());
}

// 判断是否已有 LIMIT（简化实现）
function hasExplicitLimit(sql) {
  // 注意：此实现不解析字符串字面量和子查询中的 LIMIT
  // 但实际使用中已足够覆盖 99% 的场景
  return /\bLIMIT\b/i.test(sql);
}

// 构建预览 SQL
function buildPreviewQuery(sql, limit) {
  if (isSelectQuery(sql) && !hasExplicitLimit(sql)) {
    return sql.replace(/;?\s*$/, '') + ' LIMIT ' + limit;
  }
  return sql;
}
```

### 3. 导出（写入新工作表）流

```
用户点击「写入新工作表」
  │
  ├─ 输入工作表名称
  │
  ├─ isPreview === true?
  │     ├─ 否 → 直接使用 currentQueryResult.rows（全部数据已在手）
  │     └─ 是 → 需重新查询全量数据
  │           ├─ 状态：'正在查询全部数据...'
  │           ├─ 执行 currentOriginalSQL（原始 SQL，无 LIMIT）
  │           └─ 获取完整 result.rows
  │
  └─ 写入 Excel（复用现有分块逻辑，5000 行/批，进度条不变）
```

### 4. 复制结果

```javascript
function copyResult() {
  // 只复制当前显示的行（最多 200 行）
  var displayRows = currentQueryResult.rows;
  // ... 复制逻辑不变 ...
  
  // 状态提示区分场景
  if (isPreviewResult) {
    // '已复制前 200 行，如需全部数据请使用「写入新工作表」'
  } else {
    // '已复制 83 行到剪贴板'
  }
}
```

## 状态消息规范

| 场景 | 条件 | 消息 |
|------|------|------|
| 自动截断预览 | SELECT 无 LIMIT + 结果 == 200 行 | `✅ 预览前 200 行 (0.50 秒)` |
| 用户自写 LIMIT | SQL 含 LIMIT | `✅ 查询完成，返回 1000 行（用户指定 LIMIT）(1.20 秒)` |
| 结果不足 200 | 结果行 < 200 | `✅ 查询完成，返回 83 行 (0.02 秒)` |
| 修改语句 | INSERT/UPDATE/DELETE/DROP/CREATE | 不变: `完成，影响 500 行 (0.30 秒)` |
| 复制预览 | isPreview === true | `📋 已复制前 200 行，如需全部数据请使用「写入新工作表」` |
| 复制非预览 | isPreview === false | `📋 已复制 83 行到剪贴板` |
| 导出开始 | isPreview === true | `🔄 正在查询全部数据...` |
| 写入进度 | 写入中 | `📊 已写入 5000/154283 行` |
| 导出完成 | 写入完成 | `✅ 已将 154,283 行结果写入新工作表` |

## 数据结构变化

```javascript
// 新增/修改的变量
var currentQueryResult = null;    // 不变：存储当前查询结果
var currentOriginalSQL = null;    // 新增：存储原始 SQL（无 LIMIT）
var isPreviewResult = false;      // 新增：标记是否为截断预览
```

## 涉及文件

| 文件 | 改动量 | 说明 |
|------|--------|------|
| `src/taskpane/sql-query-taskpane.js` | ~80 行 | 核心逻辑：LIMIT 追加、预览标记、导出前重新查询、复制提示 |
| `src/taskpane/sql-query-taskpane.css` | 0 行 | 无样式变更 |
| `src/taskpane/sql-query-taskpane.html` | 0 行 | 无 UI 变更 |
| `src/utils/sql-utils.js` | 0 行 | `exec()` 方法不变 |

## 边界情况

- **查询不足 200 行** → 正常显示，`isPreview = false`，导出无需重新查询
- **用户写 `LIMIT 5`** → 尊重，显示 5 行，`isPreview = false`
- **聚合查询**（COUNT/GROUP BY）→ 通常不足 200 行，正常行为
- **用户写 `LIMIT 300`** → 尊重，显示 300 行，`isPreview = false`
- **多条语句** → 只处理 `results[0]`，行为不变
- **修改语句** → 不受 LIMIT 逻辑影响

## 未纳入范围

- 分页浏览（翻页）
- 后台渐进加载
- LIMIT 值可配置（当前固定 200，未来可通过 UI 设置）
