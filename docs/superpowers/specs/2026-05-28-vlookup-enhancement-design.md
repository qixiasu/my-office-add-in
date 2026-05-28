# 增强查找功能修改设计

**日期：** 2026-05-28
**项目：** My Office Add-in
**需求：** 增强查找功能修改

---

## 1. 简化输出类型

- **删除**"输出类型"选项组（公式/静态值单选框），UI 和 HTML 中的相关代码一并移除
- 固定输出静态值（删除 `performLookup()` 中 `config.outputType === "formula"` 的逻辑分支）
- `executeLookup()` 中移除 `outputType` 字段读取

---

## 2. 删除预览区域

- 移除 HTML 中的 `<div id="vlookupPreview">` 及内部 table 结构
- 删除 JS 中的 `showPreview()` 函数
- 删除 `updateHeaderLabels()` 中对预览区域的调用（保留表头更新逻辑）
- CSS 中 `.vlookup-preview` 相关样式可移除

---

## 3. 新增"查找不到时填充默认值"

- **位置**：独立一行，放置在"匹配模式"下方、"返回列"上方
- **UI**：文本输入框，占位符 `#N/A`
- **默认值**：`#N/A`
- **行为**：用户可修改为任意值（空字符串、0、自定义文字均可）
- **实现**：
  - 读取表单 `document.getElementById("defaultValue").value`
  - 传入 `performLookup()` 的 config 对象
  - 传递至 `staticLookup()` 替代硬编码的 `"#N/A"`

---

## 4. 性能分级策略

### 阈值
- **100,000 行**作为分界线

### 小数据（< 100,000 行）
- 保持现有逻辑：一次性读取 lookupTable → 调用 `staticLookup()` → 一次性写入

### 大数据（≥ 100,000 行）
- **分批处理**，每批 10,000 行
- 每批流程：读取该批 lookupValues → 计算该批结果 → 立即写入 Excel → 再读下一批
- 不在内存中堆积全部结果

### 实现要点
- 在 `staticLookup()` 层面无需修改，只需在上层 `performLookup()` 中控制读取/写入节奏
- 大数据模式下，每批写入后 `context.sync()` 确保数据落地
- 进度状态提示（可选）：每批完成后更新状态文字

---

## 修改文件清单

| 文件 | 改动内容 |
|------|---------|
| `src/taskpane/vlookup-taskpane.html` | 删除 outputType 单选框、删除 vlookupPreview div、新增 defaultValue 输入框 |
| `src/taskpane/vlookup-taskpane.js` | 删除 formula 分支、删除 showPreview()、新增 defaultValue 传参、大数据分批逻辑 |
| `src/taskpane/vlookup-taskpane.css` | 移除 .vlookup-preview 相关样式 |

---

## 关键函数签名变更

```js
// staticLookup 新增 defaultValue 参数
staticLookup(lookupValues, lookupTable, matchColIndex, returnColIndices, matchMode, defaultValue)

// executeLookup config 新增字段
config.defaultValue  // string，默认 "#N/A"
```