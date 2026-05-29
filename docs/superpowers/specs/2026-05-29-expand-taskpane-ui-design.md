# 展开列工具 UI 优化设计

**日期：** 2026-05-29
**状态：** 已批准

## 概述

优化展开列工具侧边栏界面的视觉呈现，主要解决两个问题：刷新按钮缺少文字标识、容器无内边距导致内容贴近边缘。

## 改动详情

### 1. 容器内边距
- **文件：** `src/taskpane/expand-taskpane.css`
- **改动：** `.expand-container` 添加 `padding: 24px`
- **原因：** 原设计无边距，内容紧贴边缘，视觉上过于拥挤

### 2. 刷新按钮改为文字按钮
- **文件：** `src/taskpane/expand-taskpane.html`
- **改动：**
  - 移除按钮内的 SVG 图标 `<svg>...</svg>`
  - 按钮文字改为 "刷新选择"
- **文件：** `src/taskpane/expand-taskpane.css`
- **改动：**
  - `.expand-refresh-btn` 宽度从固定 `40px` 改为自适应 `padding: 8px 12px`
  - 添加 `white-space: nowrap` 防止文字换行
  - 样式与执行按钮保持一致（透明背景 + 蓝色边框 + hover 效果）

### 3. 间距微调
- `.expand-container` 的 `padding: 24px` 已提供足够内边距
- 其他元素的 margin 保持不变，与 concat-taskpane 风格协调

## 设计决策

| 决策项 | 选择 | 原因 |
|--------|------|------|
| 刷新按钮样式 | 文字按钮，无图标 | 更简洁，与执行按钮风格统一 |
| 容器内边距 | 24px | 宽松布局，视觉更舒展 |
| 按钮悬停效果 | 蓝色背景 + 白色文字 | 与现有 hover 行为一致 |

## 涉及文件

- `src/taskpane/expand-taskpane.html`
- `src/taskpane/expand-taskpane.css`