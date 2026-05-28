# Custom Logo Icons for Ribbon Buttons

## Summary

为 CommandsGroup 中的三个功能按钮和分组标签设计符合功能含义的自定义图标，替换当前共用的通用图标。

## Current State

所有按钮和分组共用 `icon-16.png / icon-32.png / icon-80.png` 通用图标。

## Design

### 风格

- 简约线条图标，与 Office 原生功能区 UI 风格统一
- 单色深灰 (#333333)，透明背景
- PNG RGBA 格式，三种尺寸：16×16、32×32、80×80

### 图标方案

| 功能 | 图标概念 | 文件名前缀 |
|------|----------|-----------|
| 连接列 | 两个矩形单元格 + 中间加号，表示两列合并 | `concat` |
| 导入CSV | 文档轮廓 + 向下导入箭头 | `import` |
| 增强查找 | 放大镜 + 小单元格，表示表中查找 | `lookup` |
| 我的工具 (分组) | 扳手/工具图形 | `tools` |

### 生成方式

使用 Python Pillow 库程序化绘制矢量风格的线条图标。

## Changes

### manifest.xml

将各按钮和图标的 `resid` 指向新图标：

- `ConcatenateIcon.{size}` → `concat-{size}.png`
- `ImportCsvButton` 的图标 → `import-{size}.png`
- `EnhancedVlookupButton` 的图标 → `lookup-{size}.png`
- 分组图标 → `tools-{size}.png`

### assets/ 目录

新增 12 个 PNG 文件（3 功能 + 1 分组，各 3 种尺寸）。
