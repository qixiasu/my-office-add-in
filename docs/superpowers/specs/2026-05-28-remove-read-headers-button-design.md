# 增强查找 - 移除读取表头按钮设计

## 日期: 2026-05-28

## 背景

增强查找侧边栏中有个"读取表头"按钮，用户点击后出现错误：
```
Cannot set properties of null (setting 'innerHTML')
```

原因是 `showPreview()` 函数访问了不存在的 DOM 元素 `previewHead` 和 `previewBody`。

另外，用户认为此按钮无意义——用户已在"表头行号"输入框中输入行号，按钮操作冗余。

## 设计决策

**移除"读取表头"按钮，改为自动触发：**

1. 用户在"表头行号"输入框中输入/修改值时，自动读取表头并更新匹配列/返回列选项
2. 同时修复 `showPreview()` 中缺失的 `<thead>` / `<tbody>` 结构

## 变更范围

### 1. HTML 修改
- 移除 `<button id="readHeaders">` 按钮
- 为 `<table id="previewTable">` 添加 `<thead id="previewHead">` 和 `<tbody id="previewBody">` 结构

### 2. JS 修改
- 移除 `document.getElementById("readHeaders").onclick` 事件绑定
- 在 `headerRow` 输入框上添加 `input` 事件监听，输入变化时自动调用 `readHeaders()`
- 确保 `loadInitialSelection()` 之后自动加载表头

## 预期效果

- 用户修改表头行号 → 自动读取并更新下拉选项
- 用户修改查找表区域（刷新选择）→ 自动重新加载表头
- 预览表格正常显示数据

## 测试要点

1. 输入框输入表头行号后，下拉选项立即更新为实际列名
2. 预览表格正确渲染表头和数据行
3. 执行按钮在各状态下的启用/禁用逻辑正常