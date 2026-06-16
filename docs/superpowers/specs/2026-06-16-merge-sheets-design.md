# 多 Sheet 合并工具设计

## 概述

提供 Excel 多 Sheet 合并功能，用户可选择多个 Sheet 并指定表头行号后合并为一个新 Sheet。

## 用户交互流程

### 界面布局

单面板包含：

1. **Sheet 列表区域** — 显示工作簿中所有 Sheet 名称，每行左侧有复选框，右侧有"表头行号"输入框（默认值为 1），行号 0 表示无表头
2. **全选/取消全选** — 列表顶部提供快捷操作
3. **执行按钮** — 底部"执行合并"按钮
4. **状态提示区** — 显示操作进度和结果

### 操作步骤

1. 页面加载时读取工作簿所有 Sheet 名称，生成带复选框和表头行号输入框的列表
2. 用户勾选要合并的 Sheet，并设置每个 Sheet 的表头行号（0 表示无表头）
3. 点击"执行合并"后，验证列数是否一致
   - **列数一致** → 继续执行
   - **列数不一致** → 提示用户重新选择，阻止合并
4. 检查"合并结果"Sheet 是否存在
   - **不存在** → 直接创建并写入
   - **存在** → 弹出确认框：覆盖 / 重命名 / 取消
5. 合并逻辑：
   - 保留第一个选中 Sheet 的表头（除非其表头行号=0）
   - 其他 Sheet 的表头行不参与合并
   - 如果所有 Sheet 表头行号都为 0，则合并结果也无表头

### 特殊规则

- 若存在"合并结果"Sheet 且用户选择重命名：`合并结果_1`（若仍存在则 `_2`，以此类推）
- 列名不要求一致，只比较列数
- 整列选择时需用 getUsedRange 处理（参考 memory: excel-js-selected-range-values-null）

## 错误处理

| 场景 | 处理方式 |
|------|----------|
| 未选择任何 Sheet 或只选择一个 | 提示"请至少选择两个 Sheet 进行合并" |
| 列数不一致 | 提示"所选 Sheet 列数不一致，请重新选择" |
| 所选 Sheet 无可用数据 | 提示"所选 Sheet 无可用数据" |
| 重命名后仍存在同名 Sheet | 继续递增序号尝试 |

## 技术实现

### 文件结构

- `src/taskpane/merge-sheets-taskpane.html` — 界面模板
- `src/taskpane/merge-sheets-taskpane.js` — 主逻辑，Office.onReady + 事件绑定
- `src/taskpane/merge-sheets-taskpane.css` — 样式
- `src/utils/merge-sheets-utils.js` — 共享工具函数

### manifest.xml 更新

在现有功能区按钮组中添加"合并 Sheet"按钮（ShowTaskpane 类型）。

### 核心 API

- `Excel.run()` — 所有操作包装器
- `context.workbook.worksheets` — 获取 Sheet 列表
- `worksheet.getUsedRange()` — 获取实际数据范围
- `worksheet.getRange().values` — 读写区域数据
- `context.workbook.worksheets.add()` — 新建 Sheet

### 合并算法

```
输入: selectedSheets = [{name, headerRow}, ...]

1. 读取每个 Sheet 的列数（从 getUsedRange 获取）
2. 验证列数一致性
3. 读取每个 Sheet 的数据（根据 headerRow 定位起始行）
4. 构建输出数据:
   - 第一行 = 第一个 Sheet 的表头（若 headerRow > 0）
   - 其余行 = 所有 Sheet 的数据行（排除表头）
5. 创建/覆盖"合并结果"Sheet 并写入数据
```

## 验收标准

- [ ] 正确列出工作簿所有 Sheet 名称
- [ ] 复选框可多选，全选/取消全选功能正常
- [ ] 表头行号输入框默认值为 1，支持 0
- [ ] 列数不一致时正确提示
- [ ] "合并结果"Sheet 已存在时弹出覆盖/重命名/取消选择
- [ ] 合并结果只保留第一个 Sheet 的表头
- [ ] 所有 Sheet 表头行号都为 0 时，合并结果也无表头
- [ ] 状态提示清晰显示进度和结果