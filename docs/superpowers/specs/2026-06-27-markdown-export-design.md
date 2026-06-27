# Markdown 表格导出功能设计

## 功能概述

将 Excel 选中区域转换为 GitHub 风格的 Markdown 表格，通过弹窗预览并支持一键复制。

## 用户需求确认

- **输出方式**：对话框预览 + 一键复制
- **复杂表格**：保留合并单元格的 colspan/rowspan 信息，并显示警告提示
- **表格风格**：GitHub 风格（带对齐标记）
- **数字/日期格式**：保留单元格显示格式

## 架构设计

```
src/
├── taskpane/
│   └── markdown-export-taskpane.{html,js,css}   # 新建：预览对话框
├── utils/
│   ├── markdown-table-utils.js                   # 新建：核心转换逻辑
│   └── markdown-table-utils.test.js              # 新建：单元测试
```

### 数据流向

1. 用户选中区域 → `Excel.run()` 获取 `range.values` + 格式信息
2. 调用 `markdownTableUtils.generate(rangeData, options)` → 返回 Markdown 字符串
3. 渲染到 Taskpane 弹窗中显示预览
4. 用户点击"复制" → 写入剪贴板 → 显示成功提示

## 核心模块设计

### `markdown-table-utils.js` — 转换引擎

```javascript
// 主要函数签名
generateMarkdownTable(rangeData, options) → string

// options 可选项
{
  includeAlignment: true,      // GitHub 风格对齐标记
  preserveFormat: true,        // 保留单元格显示格式
}
```

### 内部逻辑（5个函数）

| 函数 | 职责 |
|------|------|
| `parseRange(values)` | 解析二维数组，过滤空行空列 |
| `detectAlignment(columnValues)` | 根据数值类型推断列对齐方式（数字右对齐，文本左对齐） |
| `formatCell(value, cellFormat)` | 处理数字/日期显示格式 |
| `handleMergedCells(mergedRanges)` | 检测合并单元格，生成 colspan/rowspan |
| `buildTableString(rows, alignment)` | 拼接 Markdown 表格字符串 |

## 错误处理设计

| 场景 | 处理方式 |
|------|----------|
| 选区为空 | 弹窗提示"请先选择数据区域" |
| 选区过大（>10000行） | 弹窗提示"数据量过大，请缩小选区" |
| 含合并单元格 | 在表格相应位置输出 colspan/rowspan，并在弹窗顶部显示警告提示 |

## 新增文件清单

| 文件 | 用途 |
|------|------|
| `src/utils/markdown-table-utils.js` | Markdown 表格转换核心逻辑 |
| `src/utils/markdown-table-utils.test.js` | 单元测试 |
| `src/taskpane/markdown-export-taskpane.html` | 弹窗 HTML 结构 |
| `src/taskpane/markdown-export-taskpane.js` | 弹窗逻辑，调用转换工具 |
| `src/taskpane/markdown-export-taskpane.css` | 弹窗样式 |

## manifest.xml 更新

在功能区添加新按钮：
- 组：数据处理 或 查找与导入
- 按钮：导出 Markdown
- 类型：ShowTaskpane