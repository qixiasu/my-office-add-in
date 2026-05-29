# 字段计数功能设计

## 概述

对用户选中的列统计每个字段值出现的次数，在选中列后面插入新列写入结果，表头名称为「原表头-计数」。

## 需求

- 用户选中一列或多列
- 用户指定表头在第几行（默认第1行，只统计表头行下方区域）
- 统计方式：每行一个 COUNTIF 公式
- 新列插入位置：每列后各插入一列
- 新列表头：原表头名 + "-计数"，写入位置与用户指定表头行同级

## UI

### 侧边栏结构

```
┌─────────────────────────────┐
│ 📊 字段计数工具              │
├─────────────────────────────┤
│ 表头行号 [____] (默认1)     │
│                             │
│ [ 执行计数 ]                │
│                             │
│ 状态: 等待操作...           │
└─────────────────────────────┘
```

### 输入框

- 数字输入框，默认值 1
- 表头行号：1-based（即 Excel 中的行号）

## 核心逻辑

### 数据范围确定

1. 获取用户选中的列
2. 从用户指定的表头行+1 开始，到该列最后有数据的行作为计数范围
3. 使用 `getUsedRange()` 获取实际数据范围

### 公式生成

对数据区域的每一行，生成公式：
```
=COUNTIF($B$3:$B$100, B3)
```

- 计数范围：绝对引用 `$B$3:$B$100`（表头行下方到末尾）
- 计数目标：相对引用 `B3`（当前行）

### 表头处理

- 新列表头在用户指定的表头行位置写入「原表头-计数」
- 如果表头行原已有内容，覆盖写入

## 技术实现

### 文件结构

```
src/taskpane/
  count-values-taskpane.js   # 主逻辑
  count-values-taskpane.html  # UI
  count-values-taskpane.css   # 样式
src/utils/
  count-values-utils.js       # 工具函数
```

### 关键实现

```javascript
// 获取数据范围（表头下方开始）
var headerRow = parseInt(headerRowInput.value) || 1;
var dataStartRow = headerRow + 1;
var dataRange = worksheet.getRange(colLetter + dataStartRow + ":" + colLetter + lastRow);
dataRange.load("values");
dataRange.load("rowCount");

// 插入新列
var resultColLetter = getColumnLetter(colIndex + 1);
worksheet.getRange(resultColLetter + ":" + resultColLetter).insert(Excel.InsertShiftDirection.right);

// 写入表头
var headerCell = worksheet.getRange(resultColLetter + headerRow);
headerCell.values = [[originalHeader + "-计数"]];

// 写入公式
var startCell = worksheet.getRange(resultColLetter + dataStartRow);
var formula = "=COUNTIF($" + colLetter + "$" + dataStartRow + ":$" + colLetter + "$" + lastRow + "," + colLetter + dataStartRow + ")";
startCell.formulas = [[formula]];
startCell.autoFill(fillRange, Excel.AutoFillType.fillDefault);
```

## 错误处理

| 场景 | 显示 |
|------|------|
| 未选中列 | 错误：请先选择一列 |
| 表头行号无效 | 错误：请输入有效的表头行号 |
| 数据为空 | 错误：没有数据 |
| 执行失败 | 错误：[error.message] |
| 成功 | 完成！已在第 X 列写入 Y 行计数结果 |

## 实施步骤

1. 创建 `count-values-taskpane.html`
2. 创建 `count-values-taskpane.css`
3. 创建 `count-values-taskpane.js`
4. 创建 `count-values-utils.js`
5. 更新 `manifest.xml` 添加按钮和资源映射
6. 更新 `webpack.config.js` 添加新的 entry point
7. 测试完整流程