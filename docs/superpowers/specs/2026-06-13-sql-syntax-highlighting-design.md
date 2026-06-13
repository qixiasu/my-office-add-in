# SQL 语法高亮设计方案：CodeMirror 6 集成

**日期：** 2026-06-13
**项目：** My Office Add-in — sql-query-taskpane
**状态：** 已批准设计

## 1. 目标

为 Office Add-in 的 SQL 查询输入框提供 IDE 级别的编辑体验：语法高亮、括号匹配、代码折叠、SQL 关键字补全和数据库 schema 感知的自动补全。

## 2. 技术选型

| 方案 | 选型 | 理由 |
|------|------|------|
| 编辑器内核 | **CodeMirror 6** | 树摇优化、200KB gzip、模块化、纯 ESM |
| SQL 语言 | `@codemirror/lang-sql` | 官方 SQL 支持，schema 参数注入表/列名 |
| 基础功能 | `basicSetup` + 按需扩展 | 括号匹配、折叠、历史、选中、行号 |
| IE11 策略 | sql-query-taskpane 入口的 babel target 放宽 | 其他 taskpane 不受影响 |

**不选 Monaco / Ace 的理由：** 文件过大（15-25MB / 800KB），在 Office 侧边栏受限的渲染环境中增加不必要的内存和加载开销。

## 3. 架构

### 3.1 文件结构

```
src/
  utils/
    cm6-sql-editor.js    ★ 新增：CM6 编辑器工厂模块
  taskpane/
    sql-query-taskpane.js  ★ 修改：textarea 交互替换为 CM6 API
    sql-query-taskpane.html ★ 修改：<textarea> → <div id="sqlEditorContainer">
    sql-query-taskpane.css  ★ 修改：CM6 容器样式 + 无 textarea 样式清理
```

### 3.2 `cm6-sql-editor.js` API

```javascript
createSqlEditor(container, {
  getSchema: () => ({ tableName: ["col1", "col2"], ... }),  // 动态 schema
  darkMode: false,                                            // 暗色模式
})
// 返回：
{
  view,            // EditorView 实例（低级访问，一般不需要）
  getValue(),      // → string              读取 SQL
  setValue(sql),   // → void                写入 SQL（自动高亮）
  updateSchema(),  // → void                热更新自动补全
  onExecute(fn),   // → void                注册 Ctrl+Enter 回调
  destroy(),       // → void                清理
}
```

### 3.3 数据流

```
┌─────────┐  输入   ┌──────────────────┐  点击执行  ┌──────────────┐
│ 用户    │ ──────→ │  CM6 EditorView   │ ───────→  │ dbManager    │
│ 键盘    │         │  (语法高亮+补全)   │            │ .exec(sql)   │
└─────────┘         └──────────────────┘            └──────────────┘
                        ↑       │                        │
                        │       │ 查询历史回填            │
                        │       └─────────────────────── │ getTables()
                        │                                │ getTableSchema()
                        └────────────────────────────────┘
                         updateSchema()（导入新表后调用）
```

## 4. 集成细节

### 4.1 HTML 改动

```diff
- <textarea id="sqlInput" class="sql-editor" placeholder="..." rows="6"></textarea>
+ <div id="sqlEditorContainer" class="sql-editor"></div>
```

### 4.2 JS 改动点

| 功能点 | 现在（textarea） | 之后（CM6） |
|--------|------------------|-------------|
| 读取 SQL | `el.value` | `editor.getValue()` |
| 写入 SQL | `el.value = sql` | `editor.setValue(sql)` |
| 清空 | `.value = ""` | `editor.setValue("")` |
| 执行快捷键 | 无 | Ctrl+Enter → `onExecute(runQuery)` |
| 历史回填 | `el.value = sql` | `editor.setValue(sql)`（自动着色） |
| placeholder | `textarea.placeholder` | CM6 原生 placeholder 扩展 |

### 4.3 CSS 改动

```css
/* 容器样式替换原 textarea 样式 */
.sql-editor {
  border: 1px solid #d0d0d0;
  border-radius: 2px;
  min-height: 120px;
}
/* CM6 内部样式通过 EditorView.theme() 自定义 */
```

### 4.4 webpack 构建

babel-loader 增加 override，仅对 sql-query-taskpane 入口放宽 target：

```js
{
  test: /\.js$/,
  exclude: /node_modules/,
  use: {
    loader: "babel-loader",
    options: {
      overrides: [{
        test: /sql-query-taskpane/,
        presets: [["@babel/preset-env", { targets: "last 2 Chrome versions" }]]
      }]
    }
  }
}
```

此 override **不影响** 其他入口（concat-taskpane、vlookup-taskpane 等仍以 IE11 为目标）。CM6 的 node_modules 代码原本就不经过 babel，保持原生 ESM。

### 4.5 新增依赖（npm install）

```
@codemirror/state @codemirror/view @codemirror/language
@codemirror/lang-sql @codemirror/autocomplete @codemirror/commands
@codemirror/basic-setup
```

## 5. 主题配色

| Token | 颜色 | 用途 |
|-------|------|------|
| 关键字 | `#0078d4` | SELECT, FROM, WHERE, JOIN... |
| 字符串 | `#d32f2f` | '带引号的字符串' |
| 数字   | `#098658` | 42, 3.14 |
| 注释   | `#808080 italic` | -- 单行注释, /* 块注释 */ |
| 运算符 | `#333333` | =, <>, LIKE, IN |
| 背景   | `#ffffff` | 编辑器底色 |
| 光标行 | `#f5f5f5` | 当前行背景 |
| 选中   | `#0078d420` | 文本选中背景 |

通过 `HighlightStyle` + `EditorView.theme()` 定义，不依赖第三方主题包。

## 6. 自动补全

### 6.1 两级补全

1. **SQL 关键字补全**（开箱即用）：输入 `SEL` → 补全 `SELECT`，`WH` → `WHERE`
2. **Schema 感知补全**（定制开发）：`FROM [表名列表]`，`WHERE [该表列名列表]`

### 6.2 Schema 热替换

通过 CM6 的 `Compartment` 机制实现 `updateSchema()`：

```javascript
const schemaCompartment = new Compartment();

// 初始创建
const extensions = [
  sql({ schema: getSchema() }),
  schemaCompartment.of([]),
  // ... 其他扩展
];

// 热替换（导入新表后调用）
view.dispatch({
  effects: schemaCompartment.reconfigure(sql({ schema: getSchema() }))
});
```

每次 `updateSchema()` 调用时：
- `getSchema()` 读取 `dbManager.getTables()` + `dbManager.getTableSchema(name)`
- 生成 `{ tableName: ["col1", "col2", ...] }` 格式
- SQL 关键字和 schema 名称由 CM6 自动合并显示

## 7. 错误处理与边界情况

| 场景 | 处理方式 |
|------|----------|
| CM6 加载失败 | `try-catch` 包裹，降级为普通 `<textarea>`（回退策略） |
| 无效 SQL | 保持原样，由 `dbManager.exec()` 返回错误（和现有流程一致） |
| 编辑器被销毁 | taskpane 卸载时调用 `editor.destroy()` |
| 大 SQL | CM6 自带虚拟渲染，无性能问题 |
| 快速双层 Ctrl+Enter | CM6 transaction 同步，getValue() 始终最新 |
| Schema 为空 | 仅提供关键字补全，不报错 |

## 8. 最小改动原则

- 不改动其他 taskpane（concat/vlookup/csv-import...）
- 不改动 `dbManager` 或 `sql-utils.js`
- 不引入新框架（React/Vue/Angular）
- 不改动 `queryHistory` 的持久化和渲染逻辑
- 不改动结果表格渲染和导出流程

## 9. 实施步骤

1. `npm install` 新增 CM6 依赖
2. 创建 `src/utils/cm6-sql-editor.js`（核心工厂模块）
3. 修改 `sql-query-taskpane.html`（textarea → div）
4. 修改 `sql-query-taskpane.css`（移除 textarea 样式，添加 CM6 容器样式）
5. 修改 `sql-query-taskpane.js`（集成 editor 对象，替换所有 textarea 交互）
6. 修改 `webpack.config.js`（添加 babel override）
7. 修改 `sql-query-taskpane.js` 的 `renderQueryHistory` 点击事件（改用 `setValue`）
8. 导入新表后的 `refreshTableList()` 旁添加 `editor.updateSchema()` 调用
9. 验证构建和功能

## 10. 未涵盖/后续扩展

- **暗色模式**：当前仅实现浅色主题，暗色模式可后续通过 system preference media query 检测
- **行号**：`basicSetup` 包含行号，在小屏侧边栏中可能占用空间，可配置隐藏
- **SQL 格式化**：CM6 无内置格式化，可通过外部库（sql-formatter）扩展
