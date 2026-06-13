# 写入工作表列宽优化 — 仅表头自动列宽

**日期**: 2026-06-13
**状态**: 已批准

## 问题

写入 92列 × 42万行 数据到 Excel 后，`doAutoFit()` 中的 `sheet.getUsedRange().format.autofitColumns()` 需要扫描全部 3864 万个单元格以计算最佳列宽，导致 Excel 卡死数秒。

## 方案

将列宽自适应范围从「全部已用区域」改为「仅第 1 行（表头行）」：

**改动前：**
```javascript
var range = sheet.getUsedRange();
range.format.autofitColumns();
```

**改动后：**
```javascript
var headerRange = sheet.getRange("1:1");
headerRange.format.autofitColumns();
```

## 取舍

- 列宽按表头文字长度计算，数据列内容比表头长时可能显示不全
- 用户可通过选中列后双击边缘或全选后统一调整
- 对于超大数据集，这是性能和可用性的最佳平衡

## 改动范围

仅 1 个文件，1 行代码变更：
- `src/taskpane/sql-query-taskpane.js` — `doAutoFit()` 内部
