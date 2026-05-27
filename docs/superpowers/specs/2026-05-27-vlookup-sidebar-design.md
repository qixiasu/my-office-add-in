# 增强查找侧边面板设计

## 概述

将增强查找功能从对话框（Dialog）迁移到侧边面板（Taskpane），与连接列、导入 CSV 保持一致的 UI 风格。

## 文件变更

| 操作 | 文件路径 |
|------|---------|
| 新建 | `src/taskpane/vlookup-taskpane.html` |
| 新建 | `src/taskpane/vlookup-taskpane.css` |
| 新建 | `src/taskpane/vlookup-taskpane.js` |
| 修改 | `manifest.xml` — EnhancedVlookupButton 从 ExecuteFunction 改为 ShowTaskpane |
| 修改 | `webpack.config.js` — 添加 vlookup-taskpane 入口和 HtmlWebpackPlugin |

## UI 布局

一屏展示，无步骤向导：

```
┌─────────────────────────────┐
│ [图标] 增强查找              │  ← 标题区
├─────────────────────────────┤
│ 📋 选中查找值区域和查找表... │  ← 引导卡片
├─────────────────────────────┤
│ 查找值区域： [输入框] [刷新] │  ← 表单区
│ 查找表区域： [输入框] [刷新] │
│ 表头行号：   [输入框] [读取] │
│ 匹配列：     [下拉框]       │
│ 返回列：     ☑A  ☑B  ☑C    │
│ 匹配模式：   ○精确  ○近似   │
│ 输出类型：   ○公式  ○静态值 │
│                             │
│ [数据预览折叠区]            │
├─────────────────────────────┤
│ [取消]            [执行查找]│  ← 操作栏
├─────────────────────────────┤
│ 状态：等待操作...           │  ← 状态消息
└─────────────────────────────┘
```

## 样式统一

与连接列（concat-taskpane.css）、导入 CSV（csv-import-taskpane.css）保持一致：

- 标题区：`display: flex; align-items: center; gap: 8px;`，图标 + 文字
- 引导卡片：`background: #f0f7ff; border-left: 3px solid #0078d4; border-radius: 4px;`
- 输入框：`border: 2px solid #0078d4; border-radius: 6px;`
- 按钮：透明背景 + 蓝色边框，hover 时蓝色填充
- 状态消息：成功绿色、错误红色、加载黄色

## 关键改进点

### 刷新选择直接读取选区

不再依赖 URL 参数传参。点击"刷新选择"按钮时：

```javascript
Excel.run(function (context) {
  var range = context.workbook.getSelectedRange();
  range.load("address");
  return context.sync().then(function () {
    document.getElementById("lookupRange").value = range.address;
    refreshColumns();
  });
});
```

### 数据预览

- 默认折叠，点击"读取表头"后展开
- 使用 `<details>/<summary>` 或 JS 控制显示/隐藏
- 表格可滚动，最大高度 200px

## manifest.xml 变更

将 EnhancedVlookupButton 从 ExecuteFunction 改为 ShowTaskpane：

```xml
<Control xsi:type="Button" id="EnhancedVlookupButton">
  <Action xsi:type="ShowTaskpane">
    <TaskpaneId>VlookupTaskpaneId</TaskpaneId>
    <SourceLocation resid="VlookupTaskpane.Url"/>
  </Action>
</Control>
```

新增 Url 资源：
```xml
<bt:Url id="VlookupTaskpane.Url" DefaultValue="https://localhost:3000/vlookup-taskpane.html"/>
```

## webpack.config.js 变更

```javascript
entry: {
  // ...
  "vlookup-taskpane": ["./src/taskpane/vlookup-taskpane.js", "./src/taskpane/vlookup-taskpane.html"],
},

// plugins 添加
new HtmlWebpackPlugin({
  filename: "vlookup-taskpane.html",
  template: "./src/taskpane/vlookup-taskpane.html",
  chunks: ["polyfill", "vlookup-taskpane"],
}),
```

## 依赖关系

- `vlookup-taskpane.js` 依赖 `src/utils/vlookup-utils.js` 中的工具函数
- 需要在 webpack 中配置 `require` 别名或直接 require

## 测试验证

1. 点击"增强查找"按钮，侧边面板正常打开
2. 选中 Excel 中的查找表，点击"刷新选择"，输入框自动填入地址
3. 输入表头行号，点击"读取表头"，列名下拉和预览正常显示
4. 选择匹配列、返回列，点击"执行查找"，数据正确写入
5. 取消按钮关闭面板