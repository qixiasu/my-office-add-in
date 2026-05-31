# 按列拆分数据 - 设计文档

## 功能概述

将 Excel 数据按某一列（Key 列）的值拆分为多个新工作表，每个 Key 值一个 Sheet，包含表头和对应的数据行。

## 用户交互流程

1. 用户选择数据区域（A1:E6）→ 点击「刷新选择」确认
2. 用户指定表头行号（默认1）
3. Key 列下拉菜单自动加载表头，用户选择拆分依据列
4. 点击「执行拆分」→ 前端读取数据 → 分组 → 创建新 Sheet → 写入数据
5. 若 Sheet 已存在 → 报错停止

## 表单设计

| 字段 | 类型 | 说明 |
|------|------|------|
| 数据区域 | 只读输入框 + 刷新按钮 | 用户选择数据范围 |
| 表头行号 | 数字输入框 | 默认1，支持用户指定 |
| Key 列 | 下拉菜单 | 读取表头后加载列名选项 |
| 执行按钮 | 主按钮 | 验证通过后启用 |

## 数据流程

用户选择区域 → 刷新选择 → 读取数据 → 按 Key 列分组
                                          ↓
    每个 Key 值 → 新建 Sheet → 写入表头 + 对应数据行

## 核心逻辑

### 分组逻辑
- 读取整个数据区域到内存数组
- 跳过表头行，按 Key 列值分类
- Key 为空/空格 → 提示错误，停止操作

### Sheet 创建
- Sheet 名称 = Key 值（最长 31 字符，超出截断）
- 表头行 + 数据行一起写入新 Sheet
- 原工作表保留不变

### 错误处理

| 场景 | 处理 |
|------|------|
| Key 列有空格/空值 | 提示"Key 列不能为空" |
| Sheet 已存在 | 报错停止，提示"Sheet [名称] 已存在" |
| 读取失败 | 提示具体错误信息 |

## 新增文件

| 文件 | 说明 |
|------|------|
| src/taskpane/split-sheet-taskpane.html | UI 页面 |
| src/taskpane/split-sheet-taskpane.css | 样式 |
| src/taskpane/split-sheet-taskpane.js | 主逻辑 |
| src/utils/split-sheet-utils.js | 工具函数（分组逻辑） |
| src/utils/split-sheet-utils.test.js | 单元测试 |

## manifest.xml 改动

新增 Ribbon 按钮，分组归属 DataProcessingGroup 。

## UI 风格

- 与现有 vlookup-taskpane 风格一致
- 使用 Office Fluent UI CSS
- 进度条显示处理状态
- 状态消息区分成功/失败样式
