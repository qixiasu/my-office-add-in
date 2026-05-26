# 连接列功能侧边栏改造设计

## 概述

将连接列功能从弹窗模式改为独立的 Excel 侧边栏操作界面。用户点击功能区「连接列」按钮后，打开专用侧边栏，输入连接符后按 Enter 键执行合并。

## 背景

**当前实现：**
- 功能区按钮 "连接列" → 弹出 `connector-dialog.html` 对话框 → 用户输入连接符 → 执行

**问题：**
- 弹窗体验割裂，操作不流畅
- 连接符输入和执行分离，用户需要两次确认

**目标：**
- 侧边栏内完成全部操作
- 打开侧边栏自动聚焦输入框
- Enter 键触发执行

## 交互流程

```
用户点击功能区「连接列」按钮
    ↓
打开连接列侧边栏 (concat-taskpane.html)
    ↓
连接符输入框自动聚焦 (autofocus)
    ↓
用户输入连接符 → 按 Enter 键
    ↓
执行合并，在第3列写入公式
    ↓
显示成功状态，侧边栏保持打开
```

## 设计

### 侧边栏 UI

**结构：**
```
┌─────────────────────────────┐
│  连接列工具                 │
├─────────────────────────────┤
│                             │
│  选中 Excel 中的两列，       │
│  输入连接符后按 Enter       │
│                             │
│  连接符: [________]  ← 自动聚焦 │
│                             │
│  状态: ✓ 完成! 已在C列...    │
│                             │
└─────────────────────────────┘
```

### 视觉风格
- 使用现有 Fluent UI 样式（与原 taskpane 一致）
- 简洁聚焦，单一功能

## 技术方案

### 文件变更

**新增文件：**
- `src/taskpane/concat-taskpane.html` — 连接列侧边栏页面
- `src/taskpane/concat-taskpane.js` — 连接列侧边栏逻辑

**修改文件：**
- `manifest.xml` — 添加新 URL 资源，修改 `ConcatenateButton` 为 `ShowTaskpane`

### manifest.xml 变更

```xml
<!-- 新增 URL -->
<bt:Url id="ConcatTaskpane.Url" DefaultValue="https://localhost:3000/concat-taskpane.html"/>

<!-- 修改按钮 Action -->
<Control xsi:type="Button" id="ConcatenateButton">
  ...
  <Action xsi:type="ShowTaskpane">
    <TaskpaneId>ConcatTaskpaneId</TaskpaneId>
    <SourceLocation resid="ConcatTaskpane.Url"/>
  </Action>
</Control>
```

### concat-taskpane.html 结构

```html
<header><h1>连接列工具</h1></header>
<main>
  <p>选中两列后，输入连接符并按 Enter</p>
  <input type="text" id="connector" autofocus />
  <p id="status"></p>
</main>
```

### concat-taskpane.js 逻辑

1. **自动聚焦：** 页面加载时 `document.getElementById("connector").focus()`
2. **Enter 键监听：** `input.onkeydown` 检测 Enter 键触发 `runConcat()`
3. **执行合并：** 核心逻辑复用现有 `performConcat` 函数
4. **状态显示：** 在 `#status` 元素显示成功/错误信息

### 复用代码

核心合并逻辑 (`performConcat`) 来自 `commands.html`，可抽取到独立 utils 文件或直接内联在 `concat-taskpane.js`。

## 移除原功能

由于连接列功能迁移到独立侧边栏，应从原 `taskpane.html` 中移除：
- 连接列 UI 区域（header 说明、connector input、concatBtn、status）
- `runConcat` 函数及 `concat-utils` 引用

原 taskpane 保留 CSV 导入和增强查找功能。

## 错误处理

| 场景 | 处理 |
|------|------|
| 未选中两列 | 显示错误："请至少选择两列" |
| 无数据 | 显示错误："没有数据" |
| 数据量过大 | 显示错误："数据量过大（X行），单次最多支持 1050000 行" |
| 执行失败 | 显示错误信息 |

## 成功状态显示

执行成功后，在侧边栏显示：
```
完成! 已在第 C 列写入 1000 行公式
```

状态文字颜色：成功=绿色，错误=红色。

## 待实施步骤

1. 创建 `src/taskpane/concat-taskpane.html`
2. 创建 `src/taskpane/concat-taskpane.js`
3. 修改 `manifest.xml` 添加新 URL 和按钮 Action
4. 验证功能区按钮行为
5. 从原 taskpane 移除连接列相关代码
6. 测试完整流程
