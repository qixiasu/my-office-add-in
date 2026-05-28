# Remove Legacy Taskpane Button

## Summary

删除 CommandsGroup 中已废弃的"显示面板"按钮及其关联代码。

## Motivation

CommandsGroup 目前有 4 个按钮：显示面板、连接列、导入CSV、增强查找。

"显示面板"按钮打开旧版 `taskpane.html`，其中包含旧版 CSV 导入和旧版增强查找（对话框方式）。这两个功能已分别迁移到独立的专用按钮（导入CSV、增强查找），旧面板不再需要。

## Changes

### manifest.xml

- 删除 `TaskpaneButton` 控件定义（`id="TaskpaneButton"`）
- 删除 `Taskpane.Url` 资源
- 删除 `TaskpaneButton.Label` 字符串
- 删除 `TaskpaneButton.Tooltip` 字符串

### webpack.config.js

- 删除 `taskpane` entry
- 删除 taskpane 对应的 `HtmlWebpackPlugin`

### 删除文件

- `src/taskpane/taskpane.js`
- `src/taskpane/taskpane.html`

### 保留

- `src/taskpane/taskpane.css` — 被 concat、csv-import、vlookup 三个面板共用
- manifest.xml 的 `<SourceLocation>` — 保留不变，VersionOverrides 已覆盖
