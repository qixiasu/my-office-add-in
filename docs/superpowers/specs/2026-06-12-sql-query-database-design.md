# SQL 数据库查询功能设计文档

> 为 Excel 数据处理工具 Add-in 添加 SQLite 数据库查询能力，支持将工作表数据导入 SQLite、使用 SQL 进行多表关联查询、结果预览与导出。

**日期**: 2026-06-12
**状态**: 设计稿 v1

---

## 1. 概述

### 1.1 目标

在 Excel Office Add-in 中集成 SQLite 数据库引擎（通过 sql.js WASM），使用户能够：
- 将工作表中的数据导入 SQLite 数据库自动建表
- 使用 SQL 对多个表进行复杂查询和关联分析
- 在任务面板中预览查询结果
- 可选择将结果写回 Excel 工作表

### 1.2 非目标

- 不提供 PostgreSQL / MySQL 支持（浏览器环境下无法直连）
- 不实现可视化 SQL 构建器（保留纯文本 SQL 编辑器）
- 不实现实时同步或协同功能
- 不替代 Excel 原生功能，而是补充其复杂查询能力

### 1.3 关键技术约束

- 运行环境：Office WebView（Chromium 内核）
- 无服务端，全部在浏览器端完成
- 数据库引擎：sql.js（SQLite3 的 WebAssembly 编译版本）
- 持久化：IndexedDB 存储数据库二进制文件

---

## 2. 系统架构

```
┌─────────────────────────────────────────────────────────┐
│                   Excel (Desktop / Web)                  │
│  ┌──────────────────────────────────────────────────┐   │
│  │            Office WebView (Edge Chromium)         │   │
│  │                                                   │   │
│  │  ┌──────────────────────┐  ┌──────────────────┐  │   │
│  │  │  Office JavaScript API │  │  sql.js (WASM)  │  │   │
│  │  │  (getSelectedRange,   │  │  SQLite3 Engine  │  │   │
│  │  │   values read/write)  │  │  db.exec()       │  │   │
│  │  └────────┬─────────────┘  └────────┬─────────┘  │   │
│  │           │                         │            │   │
│  │  ┌────────▼─────────────────────────▼─────────┐  │   │
│  │  │        数据库查询任务面板 (sql-query-taskpane)  │  │   │
│  │  │  ┌─────────┐ ┌─────────┐ ┌────────────┐   │  │   │
│  │  │  │ 导入视图  │ │ 表浏览器 │ │ SQL查询视图 │   │  │   │
│  │  │  └─────────┘ └─────────┘ └────────────┘   │  │   │
│  │  └───────────────────────────────────────────┘  │   │
│  │                                                   │   │
│  │  ┌──────────────────────────────────────────┐   │   │
│  │  │      IndexedDB (自动持久化)               │   │   │
│  │  │      + .db 文件导出/导入                  │   │   │
│  │  └──────────────────────────────────────────┘   │   │
│  └──────────────────────────────────────────────────┘   │
└─────────────────────────────────────────────────────────┘
```

### 2.1 核心模块

| 模块 | 职责 | 文件 |
|------|------|------|
| SQL 引擎 | 封装 sql.js 初始化、数据库创建/加载、SQL 执行 | `src/utils/sql-utils.js` |
| 持久化层 | IndexedDB 读写、防抖自动保存、.db 文件导入导出 | `src/utils/sql-utils.js` |
| 导入引擎 | 读取 Excel 选中区域、推断类型、建表插数据 | `src/utils/sql-utils.js` |
| UI 控制器 | 三标签页面板、事件绑定、状态管理 | `src/taskpane/sql-query-taskpane.js` |
| 视图模板 | HTML 模板、CSS 样式 | `.html` / `.css` |

---

## 3. 数据导入

### 3.1 导入流程

```
用户选中 Excel 区域 → 点击「导入选中区域」
    → Office.context.workbook.getSelectedRange()
    → range.values 获取二维数组
    → 提取列名（第一行，可配置）
    → 列名清理（空格→下划线、去重、空值填充）
    → 列类型推断（INTEGER → REAL → TEXT）
    → CREATE TABLE IF NOT EXISTS
    → 事务批量插入（每 500 行一批提交）
    → 自动保存数据库到 IndexedDB
    → 刷新表列表
```

### 3.2 列名处理规则

| 原始列名 | 处理后 |
|---------|--------|
| `商品名称` | `商品名称`（Unicode 直接支持）|
| `Sales Amount (2024)` | `Sales_Amount_2024` |
| (空) | `col_1` |
| `name`, `name` | `name`, `name_2` |

### 3.3 类型推断优先级

对每列逐行扫描，按以下优先级判定：

1. 所有值均为整数（或 null）→ `INTEGER`
2. 所有值均为数字（或 null）→ `REAL`
3. 有任何非数字值 → `TEXT`

特殊情况：
- 全 null 列 → `TEXT`
- 空表（只有列名无数据）→ 全部 `TEXT`

### 3.4 表名生成规则

- 用户可手动输入表名
- 留空时自动生成：以 Sheet 名称为基础，如 `Sheet1_Import`
- 只允许字母、数字、下划线
- 表名重复时追加序号 `_2`, `_3`

---

## 4. SQL 查询

### 4.1 SQL 执行

```javascript
// sql.js 执行入口
function executeSQL(db, sql) {
    try {
        const startTime = performance.now();
        const results = db.exec(sql);  // 支持多语句
        const elapsed = performance.now() - startTime;

        if (results.length === 0) {
            // 非 SELECT 语句（INSERT/UPDATE/DELETE/DDL）
            const rowsAffected = db.getRowsModified();
            return { type: 'modification', rowsAffected, elapsed };
        }

        // SELECT 结果
        return {
            type: 'query',
            columns: results[0].columns,
            rows: results[0].values,
            elapsed,
            rowCount: results[0].values.length
        };
    } catch (error) {
        return { type: 'error', message: error.message };
    }
}
```

### 4.2 查询历史

- 自动记录最近 50 条执行的 SQL
- 存储：SQL 文本、执行时间、耗时、行数
- 存储位置：localStorage（轻量，不涉及数据库二进制）
- 点击历史条目可重新填充到 SQL 编辑器中

### 4.3 安全措施

| 限制 | 实现 |
|------|------|
| 危险操作确认 | DROP TABLE、DELETE FROM、UPDATE 需二次确认 |
| 查询超时 | 暂不实现（sql.js 同步执行，长查询会阻塞 UI，需通过 Web Worker 优化）|
| SQL 注入 | 不适用（本地执行，用户自己的数据）|

> ⚠️ **注意**：sql.js 的 `exec()` 是同步操作。对于大表查询（百万行级别）会阻塞 UI 线程。后续优化方向是使用 Web Worker 在后台线程执行。

---

## 5. 持久化

### 5.1 IndexedDB 自动保存

```
db.export() → Uint8Array 二进制 → IndexedDB.put(key='sqliteDb')
```

- **触发时机**：数据导入后、DDL/DML 执行后
- **防抖**：1 秒内多次修改合并为一次写入
- **初始加载**：面板打开时从 IndexedDB 恢复
- **容量**：IndexedDB 通常有几十 MB 到几百 MB 的配额，对 Excel 数据足够

### 5.2 .db 文件导出/导入

- **导出**：`db.export()` → `new Blob([buffer], { type: 'application/vnd.sqlite3' })` → 通过 `<a>` 标签下载
- **导入**：`<input type="file" accept=".db">` → `FileReader.readAsArrayBuffer()` → `new Uint8Array()` → 重建 sql.js 实例

### 5.3 数据库生命周期

```
面板打开 → [IndexedDB 中有数据?]
    → 是: 从 IndexedDB 加载 → 初始化 sql.js
    → 否: 新建空数据库 → 初始化 sql.js

面板关闭 → (IndexedDB 已保存，无需额外操作)
```

---

## 6. 用户界面

### 6.1 整体布局

三标签页布局，顶部导航栏切换：

| 标签 | 图标 | 功能 |
|------|------|------|
| 导入数据 | 📥 | 导入选中区域、管理已导入的表 |
| 表浏览器 | 📋 | 查看表结构、预览数据 |
| SQL 查询 | 💬 | 编写执行 SQL、查看结果 |

### 6.2 导入视图

```
[导入数据] 标签页:
- 操作说明文字
- 表名输入框（可选）
- 「第一行为列名」复选框（默认勾选）
- 「导入选中区域」按钮
- 已导入的表列表（表名 + 列信息）
  - 每行：删除按钮、查看按钮
- 底部工具栏：保存 .db 文件 / 加载 .db 文件
```

### 6.3 表浏览器视图

```
[表浏览器] 标签页:
- 表名下拉选择器
- 表结构表格（列名、类型、约束）
- 数据预览表格（前 N 行）
- 总行数显示
- 操作按钮：清空表、删除表
```

### 6.4 SQL 查询视图

```
[SQL 查询] 标签页:
- 多行 SQL 文本输入框（placeholder: "输入 SQL 语句..."）
- 操作按钮：执行、清空
- 状态栏（执行耗时、返回行数、错误信息）
- 结果数据表格（列名 + 数据行）
- 导出按钮：写入新工作表、复制结果
- 查询历史列表（可点击回填）
```

### 6.5 样式

与现有 taskpane 保持一致的视觉风格：
- 使用与 `concat-taskpane.css` / `expand-taskpane.css` 相同的设计语言
- Microsoft Fabric 设计体系色板
- 响应式布局适配不同面板宽度

---

## 7. manifest.xml 变更

### 7.1 新增命令组

```xml
<Group id="DatabaseQueryGroup">
    <Label resid="DatabaseQueryGroup.Label"/>
    <Icon>
        <bt:Image size="16" resid="DbIcon.16x16"/>
        <bt:Image size="32" resid="DbIcon.32x32"/>
        <bt:Image size="80" resid="DbIcon.80x80"/>
    </Icon>
    <Control xsi:type="Button" id="SqlQueryButton">
        <Label resid="SqlQueryButton.Label"/>
        <Supertip>
            <Title resid="SqlQueryButton.Label"/>
            <Description resid="SqlQueryButton.Tooltip"/>
        </Supertip>
        <Icon>
            <bt:Image size="16" resid="DbIcon.16x16"/>
            <bt:Image size="32" resid="DbIcon.32x32"/>
            <bt:Image size="80" resid="DbIcon.80x80"/>
        </Icon>
        <Action xsi:type="ShowTaskpane">
            <TaskpaneId>SqlQueryTaskpaneId</TaskpaneId>
            <SourceLocation resid="SqlQueryTaskpane.Url"/>
        </Action>
    </Control>
</Group>
```

### 7.2 新增资源

| 资源类型 | ID | 值 |
|---------|-----|-----|
| Image | `DbIcon.16x16` | `assets/database-16.png` |
| Image | `DbIcon.32x32` | `assets/database-32.png` |
| Image | `DbIcon.80x80` | `assets/database-80.png` |
| Url | `SqlQueryTaskpane.Url` | `sql-query-taskpane.html` |
| ShortString | `DatabaseQueryGroup.Label` | `数据库` |
| ShortString | `SqlQueryButton.Label` | `数据库查询` |
| LongString | `SqlQueryButton.Tooltip` | `将工作表数据导入 SQLite 数据库，使用 SQL 进行多表关联查询与分析` |

### 7.3 组顺序

```
DataProcessingGroup  →  数据处理
DataConversionGroup  →  数据转换
DatabaseQueryGroup   →  数据库         ← 新增
QuickSelectGroup     →  快速选区
LookupImportGroup    →  查找与导入
```

---

## 8. 新增文件清单

| 文件 | 用途 |
|------|------|
| `src/taskpane/sql-query-taskpane.html` | 任务面板 HTML |
| `src/taskpane/sql-query-taskpane.js` | 面板逻辑（UI 控制、事件绑定、Excel 交互） |
| `src/taskpane/sql-query-taskpane.css` | 面板样式 |
| `src/utils/sql-utils.js` | 数据库核心工具（sql.js 封装、导入引擎、持久化） |
| `src/utils/sql-utils.test.js` | 单元测试 |
| `assets/database-16.png` | 16x16 图标 |
| `assets/database-32.png` | 32x32 图标 |
| `assets/database-80.png` | 80x80 图标 |

### 8.1 webpack 配置变更

```javascript
entry: {
    // ... 现有入口
    "sql-query-taskpane": [
        "./src/taskpane/sql-query-taskpane.js",
        "./src/taskpane/sql-query-taskpane.html"
    ],
},
plugins: [
    // ... 现有插件
    new HtmlWebpackPlugin({
        filename: "sql-query-taskpane.html",
        template: "./src/taskpane/sql-query-taskpane.html",
        chunks: ["polyfill", "sql-query-taskpane"],
    }),
],
// sql.js 的 WASM 文件需要复制到输出目录
// resolve: { alias: { 'sql.js': ... } }
```

---

## 9. 依赖

### 9.1 新增 npm 依赖

| 包名 | 用途 | 类型 |
|------|------|------|
| `sql.js` | SQLite3 WASM 引擎 | dependencies |

安装命令：
```bash
npm install sql.js
```

### 9.2 sql.js 的 WASM 部署

sql.js 的 WASM 文件 `sql-wasm.wasm`（约 1.2MB）需要能够在浏览器中加载。

**推荐方案**：通过 `copy-webpack-plugin` 将 WASM 文件复制到输出目录：
```javascript
// webpack.config.js
new CopyWebpackPlugin({
    patterns: [
        // ... 现有规则
        {
            from: "node_modules/sql.js/dist/sql-wasm.wasm",
            to: "assets/",
        },
    ],
}),
```

然后在代码中配置 WASM 路径：
```javascript
const initSqlJs = require('sql.js');
const SQL = await initSqlJs({
    locateFile: file => `assets/${file}`
});
```

备选方案：使用 CDN `https://cdn.jsdelivr.net/npm/sql.js/dist/sql-wasm.wasm`

---

## 10. 测试策略

### 10.1 单元测试（sql-utils.test.js）

| 测试项 | 内容 |
|--------|------|
| 数据库创建 | 新建空数据库、从 buffer 加载数据库 |
| 表创建与导入 | CREATE TABLE、批量 INSERT 数据 |
| 类型推断 | INTEGER / REAL / TEXT 判定逻辑 |
| 列名清理 | 特殊字符替换、空列名、重复列名 |
| SQL 执行 | SELECT、INSERT、UPDATE、DELETE、多语句 |
| 持久化 | IndexedDB 读写、db.export() 验证 |

### 10.2 手动测试

| 场景 | 条件 |
|------|------|
| 导入数据 | 选中不同数据类型（数字、文本、日期、混合） |
| 多表关联 | 导入 2-3 个表，执行 JOIN 查询 |
| 数据库持久化 | 关闭面板 → 重新打开 → 数据仍在 |
| .db 导出/导入 | 导出 .db 文件 → 重新加载 → 数据完整 |
| 结果写入 Excel | 执行查询 → 写入新工作表 → 验证数据正确 |

---

## 11. 未来优化方向

- **Web Worker 支持**：将 sql.js 移到 Web Worker 中执行，避免大查询阻塞 UI
- **渐进式结果加载**：对超大结果集使用分页或流式加载
- **多数据库管理**：创建多个数据库文件，在不同项目间切换
- **SQL 语法高亮**：为编辑器添加基础语法高亮
- **可视化表关系**：自动检测外键关系并可视化展示
- **计划任务**：定时从 Excel 同步数据到数据库
