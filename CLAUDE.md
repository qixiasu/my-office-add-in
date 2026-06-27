# CLAUDE.md

此文件为 Claude Code 提供本仓库的指导信息。

## 项目概述

Excel 数据处理工具 — Microsoft Office Excel 加载项，基于 Office JavaScript API 构建。提供多个任务面板和功能区按钮，用于操作 Excel 工作表数据。纯 JavaScript 项目，使用 webpack 打包 + Babel 转译（兼容 IE11）。

## 命令

```bash
# 开发
npm run dev-server    # 启动 webpack 开发服务器 (https://localhost:3000)
npm run start         # 启动 Excel 桌面版并旁加载加载项
npm run stop          # 停止调试会话
npm run watch         # webpack 监听模式开发

# 构建
npm run build         # 生产构建（将 localhost URL 替换为生产域名）
npm run build:dev     # 开发构建

# 测试
npm run test          # 运行 Jest 测试

# 登录/登出（Microsoft 365 账号）
npm run signin        # 登录 M365 账号
npm run signout       # 登出 M365 账号

# 代码检查与格式化
npm run lint          # 检查 lint 问题
npm run lint:fix      # 自动修复 lint 问题
npm run prettier      # 运行 prettier 格式化

# 验证
npm run validate      # 验证 manifest.xml
```

VS Code 调试：F5 → 选择 "Excel Desktop (Edge Chromium)"。会先执行 "Debug: Excel Desktop" 预启动任务启动 Excel 并旁加载加载项，然后在端口 9229 上附加 Edge 调试器。

## 架构

```
manifest.xml              # 加载项清单 — 定义功能区按钮、权限、资源 URL
webpack.config.js         # 多入口：多个 taskpane + commands
src/
  index.html              # 入口页面（空白壳页面）
  taskpane/               # 每个工具一个独立 taskpane
    concat-taskpane.js/html
    csv-import-taskpane.js/html
    vlookup-taskpane.js/html
    fill-series-taskpane.js/html
    count-values-taskpane.js/html
    expand-taskpane.js/html
    split-sheet-taskpane.js/html
    sql-query-taskpane.js/html
    ai-assistant-taskpane.js/html
  commands/
    commands.html         # 命令宿主页（空白壳页面，仅加载 office.js）
    commands.js           # ExecuteFunction 处理器，通过 Office.actions.associate 注册
    csv-import-dialog.html  # CSV 导入对话框
  utils/                  # 共享工具模块
    concat-utils.js       # 连接列工具
    expand-utils.js       # 展开列工具
    split-sheet-utils.js  # 按列拆分工具
    vlookup-utils.js      # 增强查找工具
    csv-utils.js          # CSV 解析工具
    sql-utils.js          # SQL 查询引擎
    cm6-sql-editor.js     # CodeMirror 6 SQL 编辑器封装
    ai-utils.js           # AI 助手工具（选区分析、AI 对话）
    *.test.js             # 对应各工具的 Jest 测试
assets/                   # 清单引用的图标文件
```

### 两种执行路径

1. **任务面板**（`ShowTaskpane` 动作）：在 Excel 中打开侧边面板。每个工具都是独立 HTML 页面，加载 `office.js` 后在 `Office.onReady` 中绑定事件。所有 Excel 操作在 `Excel.run` 回调中执行。

2. **命令函数**（`ExecuteFunction` 动作）：功能区按钮触发通过 `Office.actions.associate("name", handler)` 注册的函数。运行在隐藏浏览器进程中，`console.log` 输出到 VS Code 调试控制台。处理器接收 `event` 对象，**必须**调用 `event.completed()`，否则加载项会挂起。

### webpack 要点

- `HtmlWebpackPlugin` 为每个 taskpane 生成独立 HTML 文件，分别注入对应 chunk
- `CopyWebpackPlugin` 复制 `assets/*`、`manifest*.xml`、`csv-import-dialog.html`、`sql-wasm-browser.wasm` 到输出目录；生产模式下将 `localhost:3000` 替换为生产域名
- 开发服务器使用自签名 HTTPS 证书（Office 要求）
- Polyfill chunk（core-js + regenerator-runtime）包含在每个页面中，兼容 IE11
- 多入口配置：polyfill + commands + 9 个 taskpane 入口

### 功能区按钮（manifest.xml）

**Excel 工具箱** 自定义选项卡，包含 6 个功能组：

| 组 | 按钮 | 类型 |
|---|---|---|
| 数据处理 | 连接列、字段计数、按列拆分 | ShowTaskpane |
| 数据转换 | 展开列、填充序列 | ShowTaskpane |
| 查找与导入 | 增强查找、导入CSV | ShowTaskpane |
| 数据库 | 数据库查询 | ShowTaskpane |
| 快速选区 | 向下选区、向右选区 | ExecuteFunction |
| AI 助手 | AI 助手 | ShowTaskpane |

所有开发环境资源 URL 使用 `https://localhost:3000/`，生产构建时替换为 `https://qinxuan1989.gitee.io/my-office-add-in/`。

## 关键依赖

- **sql.js** + **CodeMirror 6** — SQL 查询编辑器（SELECT 语句，支持多表关联）
- **jest** — 单元测试框架
- **core-js / regenerator-runtime** — IE11 Polyfill
- **office-addin-*** — Office 加载项调试、证书、lint、清单验证工具链

## 参考

官方 Excel JavaScript API 参考文档：`javascript-api-office-js-docs-reference-excel-js-preview.pdf`（项目根目录）

## 常用 API

- `Office.onReady()` — 初始化入口
- `Office.actions.associate()` — 注册功能区命令处理器
- `Excel.run()` — 所有 Excel 对象模型操作的必需包装器
- `context.workbook.getSelectedRange()` — 获取用户当前选区
- `worksheet.getUsedRange()` — 获取工作表已使用区域
- `worksheet.getRange().insert()` — 插入单元格/列（带移位方向）
- `worksheet.getRange().values` — 读写区域数据（二维数组）
- `range.getRow()/getColumn()` — 获取区域行列信息
- `range.getResizedRange()` / `getAbsoluteResizedRange()` — 调整选区范围
