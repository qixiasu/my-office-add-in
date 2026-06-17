# 跨文件合并新增Sheet名列实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 在跨文件合并结果的表头和数据行中，新增一列"Sheet名"，记录每行数据来源于源文件的哪个工作表。

**Architecture:** 修改两个文件：(1) `cross-file-merge-taskpane.js` 透传 `sheetName` 到 `fileDataList`；(2) `cross-file-merge-utils.js` 的 `mergeExcelData` 在每行数据前面插入 `[sheetName, fileName, ...]`，表头变为 `["来源文件", "Sheet名", ...]`。

**Tech Stack:** Plain JavaScript, SheetJS (XLSX), Jest

---

## 任务总览

需要修改 **3 个文件**，涉及 **6 个步骤**（含测试更新）：

| 文件 | 改动 |
|------|------|
| `src/taskpane/cross-file-merge-taskpane.js` | `fileDataList.push` 添加 `sheetName` 字段 |
| `src/utils/cross-file-merge-utils.js` | `mergeExcelData` 添加 sheet 名列 |
| `src/utils/cross-file-merge-utils.test.js` | 更新所有涉及表头的测试用例 |

---

## Task 1: 更新 cross-file-merge-taskpane.js，透传 sheetName

**Files:**
- Modify: `src/taskpane/cross-file-merge-taskpane.js:134-138`

- [ ] **Step 1: 在 fileDataList.push 中添加 sheetName 字段**

找到约第 134 行附近的代码：

```javascript
fileDataList.push({
  data: result.data,
  name: item.name,
  columnCount: result.data.length > 0 ? result.data[0].length : 0
});
```

修改为：

```javascript
fileDataList.push({
  data: result.data,
  name: item.name,
  sheetName: result.sheetName,
  columnCount: result.data.length > 0 ? result.data[0].length : 0
});
```

- [ ] **Step 2: 提交**

```bash
git add src/taskpane/cross-file-merge-taskpane.js
git commit -m "feat(cross-file-merge): pass sheetName to fileDataList"
```

---

## Task 2: 更新 mergeExcelData，添加 Sheet 名列

**Files:**
- Modify: `src/utils/cross-file-merge-utils.js:101-144`

- [ ] **Step 1: 修改表头行插入逻辑**

约第 120-124 行：

```javascript
var headerRow = data[headerRowIndex].slice();
headerRow.unshift("来源文件"); // 添加来源文件列
mergedData.push(headerRow);
columnCount = headerRow.length;
```

修改为：

```javascript
var headerRow = data[headerRowIndex].slice();
headerRow.unshift("来源文件");   // 第二列：来源文件
headerRow.unshift("Sheet名");     // 第一列：Sheet名
mergedData.push(headerRow);
columnCount = headerRow.length;
```

- [ ] **Step 2: 修改第一文件数据行插入逻辑**

约第 127-131 行：

```javascript
for (var r = startRow; r < data.length; r++) {
  var dataRow = data[r].slice();
  dataRow.unshift(fileName);
  mergedData.push(dataRow);
}
```

修改为：

```javascript
for (var r = startRow; r < data.length; r++) {
  var dataRow = data[r].slice();
  dataRow.unshift(fileName);              // 第二列：来源文件
  dataRow.unshift(fileData.sheetName);    // 第一列：Sheet名
  mergedData.push(dataRow);
}
```

- [ ] **Step 3: 修改后续文件数据行插入逻辑**

约第 135-139 行：

```javascript
for (var r2 = startRow; r2 < data.length; r2++) {
  var dataRow2 = data[r2].slice();
  dataRow2.unshift(fileName);
  mergedData.push(dataRow2);
}
```

修改为：

```javascript
for (var r2 = startRow; r2 < data.length; r2++) {
  var dataRow2 = data[r2].slice();
  dataRow2.unshift(fileName);              // 第二列：来源文件
  dataRow2.unshift(fileData.sheetName);    // 第一列：Sheet名
  mergedData.push(dataRow2);
}
```

- [ ] **Step 4: 提交**

```bash
git add src/utils/cross-file-merge-utils.js
git commit -m "feat(cross-file-merge): add sheet name column to merged output"
```

---

## Task 3: 更新单元测试

**Files:**
- Modify: `src/utils/cross-file-merge-utils.test.js`

需要更新的测试用例（预期每行多一个 sheet name 元素）：

- `mergeExcelData` 所有涉及表头的测试用例（约第 148-197 行）
- `parseExcelFile` mock 返回值已包含 `sheetName: "Sheet1"`

具体测试数据变化对照：

| 测试用例 | 旧期望 | 新期望 |
|----------|--------|--------|
| 有表头合并 | `["来源文件", "Name", "Age"]` | `["来源文件", "Sheet名", "Name", "Age"]` |
| 有表头合并-数据行 | `["file1.xlsx", "Alice", 25]` | `["file1.xlsx", "Sheet1", "Alice", 25]` |
| 无表头合并 | `["file1.xlsx", "A1", "B1"]` | `["file1.xlsx", "Sheet1", "A1", "B1"]` |

**示例修改**（以"merges two files without header"为例）：

```javascript
// 旧
expect(result.mergedData).toEqual([
  ["file1.xlsx", "A1", "B1"],
  ["file1.xlsx", "A2", "B2"],
  ["file2.xlsx", "A3", "B3"],
  ["file2.xlsx", "A4", "B4"],
]);

// 新
expect(result.mergedData).toEqual([
  ["file1.xlsx", "Sheet1", "A1", "B1"],
  ["file1.xlsx", "Sheet1", "A2", "B2"],
  ["file2.xlsx", "Sheet1", "A3", "B3"],
  ["file2.xlsx", "Sheet1", "A4", "B4"],
]);
```

- [ ] **Step 1: 逐个修改测试用例中期望的 mergedData 数组**

每个子数组前面插入源文件的 sheetName 元素（均为 `"Sheet1"`，因为测试中 `parseExcelFile` mock 返回的 sheetName 为 `"Sheet1"`）。

- [ ] **Step 2: 运行测试验证**

```bash
npm test -- --testPathPatterns="cross-file-merge"
```

预期：22 tests passed

- [ ] **Step 3: 提交**

```bash
git add src/utils/cross-file-merge-utils.test.js
git commit -m "test(cross-file-merge): update expected output for sheet name column"
```

---

## 验证

所有任务完成后：

```bash
npm test -- --testPathPatterns="cross-file-merge"
```

预期结果：`Tests: 22 passed, 1 total`
