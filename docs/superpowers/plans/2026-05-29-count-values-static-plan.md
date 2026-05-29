# Count-Values 静态值实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 将 count-values 功能从 COUNTIF 公式改为 JavaScript Map 计算，直接写入静态值，消除大数据量时 Excel 卡顿问题

**Architecture:** 在单个 `Excel.run` 内完成所有操作：读取数据 → JS Map 统计 → 批量写入静态值。减少 context.sync() 调用次数，避免 formula 重算。

**Tech Stack:** Pure JavaScript + Office JavaScript API

---

## 文件映射

- **Modify:** `src/taskpane/count-values-taskpane.js:165-220` (`executeCount` 函数重构)

---

## 实现步骤

### Task 1: 修改 executeCount 函数逻辑

**Files:**
- Modify: `src/taskpane/count-values-taskpane.js:165-220`

- [ ] **Step 1: 重写 executeCount 函数**

将 `executeCount` 函数替换为以下实现：

```javascript
function executeCount(worksheet, columnDataRanges, headerRow, executeBtn) {
  return Excel.run(function (context) {
    var results = [];

    // 从右到左处理，避免插入列后影响后续列的索引
    for (var i = columnDataRanges.length - 1; i >= 0; i--) {
      var item = columnDataRanges[i];
      var colLetter = item.col.colLetter;
      var colIndex = item.col.colIndex;
      var dataStartRow = item.dataStartRow;
      var lastRow = item.lastRow;
      var headerValue = item.headerValue;
      var values = item.usedRange.values;
      var rowCount = item.usedRange.rowCount;

      // Step 2: 用 Map 统计每个值的出现次数
      var countMap = new Map();
      for (var r = 0; r < rowCount; r++) {
        var key = values[r][0];
        // 跳过空值
        if (key === null || key === undefined || key === "") {
          continue;
        }
        countMap.set(key, (countMap.get(key) || 0) + 1);
      }

      // Step 3: 在当前列后插入新列
      var resultColLetter = getColumnLetter(colIndex + 1);
      worksheet.getRange(resultColLetter + ":" + resultColLetter).insert(Excel.InsertShiftDirection.right);

      // Step 4: 写入表头
      var headerCell = worksheet.getRange(resultColLetter + headerRow);
      headerCell.values = [[headerValue + "-计数"]];

      // Step 5: 生成计数结果数组（2D 数组格式）
      var dataRowCount = lastRow - dataStartRow + 1;
      var countResults = [];
      for (var dr = 0; dr < rowCount; dr++) {
        var cellValue = values[dr][0];
        var count = countMap.get(cellValue) || 0;
        countResults.push([count]);
      }

      // Step 6: 一次性写入所有计数结果
      var dataRange = worksheet.getRange(resultColLetter + dataStartRow + ":" + resultColLetter + lastRow);
      dataRange.values = countResults;

      results.push({
        colLetter: colLetter,
        resultColLetter: resultColLetter,
        rowCount: dataRowCount
      });
    }

    return context.sync().then(function () {
      var resultParts = [];
      for (var j = 0; j < results.length; j++) {
        var res = results[j];
        resultParts.push("第" + res.resultColLetter + "列(" + res.rowCount + "行)");
      }
      setStatus("完成！已在 " + resultParts.join(", ") + " 写入计数结果", "success");
      executeBtn.disabled = false;
    });
  }).catch(function (error) {
    setStatus("错误: " + error.message, "error");
    executeBtn.disabled = false;
  });
}
```

- [ ] **Step 2: 提交代码**

```bash
git add src/taskpane/count-values-taskpane.js
git commit -m "feat: count-values use static values instead of COUNTIF formulas"

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>
```

---

## 自检清单

- [ ] Spec coverage: 设计文档中的所有需求都在 Task 1 中实现
- [ ] Placeholder scan: 无 TBD/TODO/IMPLEMENT LATER 等占位符
- [ ] Type consistency: 函数签名和变量名在前后任务中保持一致
- [ ] 测试验证: 在 Excel 中选择多列数据，执行计数功能，确认结果正确

---

## 执行选项

**1. Subagent-Driven (recommended)** - I dispatch a fresh subagent per task, review between tasks, fast iteration

**2. Inline Execution** - Execute tasks in this session using executing-plans, batch execution with checkpoints

**Which approach?**