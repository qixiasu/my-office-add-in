# 跨文件合并：新增 Sheet 名列

## 背景

当前跨文件合并功能在合并结果中已添加"来源文件"列，但用户需要同时记录每行数据来源于哪个 Sheet。

## 目标

在合并结果的表头和数据行中，在"来源文件"列之后新增一列"Sheet名"，记录该行数据来自源文件的哪个工作表。

## 数据流分析

```
parseExcelFile(file)
  → { data: [...], sheetName: "Sheet1" }   ← 已返回 sheetName，但未被使用

executeMerge (taskpane)
  → fileDataList.push({ data, name, columnCount })   ← sheetName 未传入

mergeExcelData(fileDataList, headerRowNumber)
  → 只使用了 fileData.name，未使用 sheetName
```

## 修改点

### 1. cross-file-merge-taskpane.js（步骤1解析后）

在 `fileDataList.push` 时添加 `sheetName` 字段：

```javascript
fileDataList.push({
  data: result.data,
  name: item.name,
  sheetName: result.sheetName,   // ← 新增
  columnCount: result.data.length > 0 ? result.data[0].length : 0
});
```

### 2. cross-file-merge-utils.js — mergeExcelData()

在表头和数据行中添加第二列"Sheet名"：

**表头变化**：`["来源文件"]` → `["来源文件", "Sheet名"]`

**数据行变化**：`[fileName, ...]` → `[fileName, sheetName, ...]`

具体修改：

```javascript
// 第一文件表头
var headerRow = data[headerRowIndex].slice();
headerRow.unshift("来源文件");
headerRow.unshift("Sheet名");     // ← 在最前面插入
mergedData.push(headerRow);
columnCount = headerRow.length;

// 第一文件数据行
dataRow.unshift(fileName);
dataRow.unshift(fileData.sheetName);  // ← 新增
mergedData.push(dataRow);

// 后续文件数据行
dataRow2.unshift(fileName);
dataRow2.unshift(fileData.sheetName); // ← 新增
mergedData.push(dataRow2);
```

### 3. 单元测试 cross-file-merge-utils.test.js

更新所有涉及表头的测试用例，添加 `sheetName` 字段：

- `parseExcelFile` mock 返回值需添加 `sheetName: "Sheet1"`
- `fileDataList` 中每个对象需添加 `sheetName` 字段
- 期望的 `mergedData` 数组中，每个子数组前面多一个 sheet name 元素

## 预期结果示例

假设有两个文件：
- `销售数据.xlsx` → Sheet1 → [["姓名","销售额"],["张三",1000]]
- `销售数据2.xlsx` → Sheet1 → [["姓名","销售额"],["李四",2000]]

表头行设置=1，合并结果：

| Sheet名 | 来源文件 | 姓名 | 销售额 |
|---------|---------|------|--------|
| Sheet1 | 销售数据.xlsx | 张三 | 1000 |
| Sheet1 | 销售数据2.xlsx | 李四 | 2000 |

## 边界情况

- **无表头模式（headerRowNumber=0）**：第一行数据也会被视为数据行，同样在前面插入 `["Sheet名", "来源文件"]` 两列
- **空文件**：已在现有代码中通过 `if (!data || data.length === 0) continue;` 跳过，不受影响
- **单个文件**：`validateColumnConsistency` 会在文件数<2时返回失败，不会执行到合并步骤
