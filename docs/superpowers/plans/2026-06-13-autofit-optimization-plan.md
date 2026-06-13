# autofitColumns 优化 — 仅表头列宽实现计划

**Goal:** 将 `doAutoFit()` 中的列宽自适应范围从全部已用区域改为仅表头行，避免扫描 3864 万单元格导致 Excel 卡死。

**Architecture:** 将 `getUsedRange()` 替换为 `getRange("1:1")`，仅对第一行执行 `autofitColumns`。

**Tech Stack:** Office.js Excel API

---

### 任务 1：修改 `doAutoFit()` 中的列宽范围

**文件：**
- 修改：`src/taskpane/sql-query-taskpane.js` 的 `doAutoFit()` 函数

- [ ] **Step 1: 替换列宽自适应范围**

将：
```javascript
var range = sheet.getUsedRange();
range.format.autofitColumns();
```

改为：
```javascript
// 仅对表头行执行列宽自适应（避免扫描全量数据导致 Excel 卡死）
var headerRange = sheet.getRange("1:1");
headerRange.format.autofitColumns();
```

- [ ] **Step 2: 验证构建和测试**

```bash
npm run build:dev && npm test
```

- [ ] **Step 3: 提交**

```bash
git add src/taskpane/sql-query-taskpane.js
git commit -m "perf: only autofit header row instead of full range to prevent Excel freeze"
```
