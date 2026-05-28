# VLOOKUP 性能优化 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 移除 `staticLookup` 中的调试日志，引入哈希索引缓存避免重复构建，提升大数据量 VLOOKUP 查找速度。

**Architecture:** 将索引构建从 `staticLookup` 中分离为独立的 `buildLookupIndex` 函数，`staticLookup` 增加可选 `indexCache` 参数。调用方（`performLookup`）在首次查找前构建索引并传入后续所有批次调用，确保索引只构建一次。

**Tech Stack:** JavaScript (ES5 var-style), Jest 30, Office JavaScript API

---

## File Map

| File | Role |
|------|------|
| `src/utils/vlookup-utils.js` | 查找核心逻辑：`staticLookup`, `buildLookupIndex`（新增） |
| `src/utils/vlookup-utils.test.js` | Jest 单元测试 |
| `src/taskpane/vlookup-taskpane.js` | 侧边栏执行逻辑，调用 `staticLookup` |

---

### Task 1: 移除 staticLookup 内的 console.log

**Files:**
- Modify: `src/utils/vlookup-utils.js` — 删除 9 处循环内/外的 `console.log`

**Why:** 每条记录查找都打印多条 `console.log`，在大数据量场景下严重拖慢执行速度。这些日志是开发调试用途，不应出现在生产路径中。

- [ ] **Step 1: 删除 staticLookup 中所有 console.log**

`src/utils/vlookup-utils.js` — 删除 `staticLookup` 函数内（第 91-128 行及第 137-158 行）的所有 `console.log` 语句。

删除后的 `staticLookup` 函数（仅展示变更部分，完整函数见后）：

```javascript
function staticLookup(
  lookupValues,
  lookupTable,
  matchColIndex,
  returnColIndices,
  matchMode,
  defaultValue
) {
  if (defaultValue === undefined || defaultValue === null) {
    defaultValue = "#N/A";
  }

  var results = [];

  var index = {};
  if (matchMode === 0) {
    for (var r = 0; r < lookupTable.length; r++) {
      var key = lookupTable[r][matchColIndex];
      if (key === null || key === undefined) {
        key = "";
      }
      key = String(key);
      index[key] = r;
    }
  }

  for (var i = 0; i < lookupValues.length; i++) {
    var row = [];
    var val = lookupValues[i];
    if (val === null || val === undefined) {
      val = "";
    }
    // Null/undefined in approximate mode → defaultValue (avoid Number("") → 0)
    if (matchMode !== 0 && (lookupValues[i] === null || lookupValues[i] === undefined)) {
      for (var q = 0; q < returnColIndices.length; q++) {
        row.push(defaultValue);
      }
      results.push(row);
      continue;
    }
    var valStr = String(val);

    if (matchMode === 0) {
      var matchedRow = index[valStr];
      if (matchedRow !== undefined) {
        for (var j = 0; j < returnColIndices.length; j++) {
          row.push(lookupTable[matchedRow][returnColIndices[j]]);
        }
      } else {
        for (var k = 0; k < returnColIndices.length; k++) {
          row.push(defaultValue);
        }
      }
    } else {
      // Approximate match: find largest value <= lookup
      // Assumes lookupTable is sorted ascending on match column (same as Excel VLOOKUP)
      var bestRow = -1;
      for (var m = 0; m < lookupTable.length; m++) {
        var tableVal = lookupTable[m][matchColIndex];
        if (tableVal === null || tableVal === undefined) continue;
        if (Number(tableVal) <= Number(val)) {
          bestRow = m;
        } else {
          break;
        }
      }
      if (bestRow >= 0) {
        for (var n = 0; n < returnColIndices.length; n++) {
          row.push(lookupTable[bestRow][returnColIndices[n]]);
        }
      } else {
        for (var p = 0; p < returnColIndices.length; p++) {
          row.push(defaultValue);
        }
      }
    }

    results.push(row);
  }

  return results;
}
```

- [ ] **Step 2: 运行现有测试验证功能未受影响**

Run: `npm test -- --testPathPattern="vlookup-utils.test.js"`
Expected: 所有 16 个测试 PASS

- [ ] **Step 3: 提交**

```bash
git add src/utils/vlookup-utils.js
git commit -m "perf(vlookup): remove debug console.log from staticLookup"
```

---

### Task 2: 新增 buildLookupIndex 函数

**Files:**
- Create: (in) `src/utils/vlookup-utils.js` — 新增 `buildLookupIndex` 函数
- Modify: `src/utils/vlookup-utils.test.js` — 新增测试

**Why:** 将索引构建与查找分离，让调用方可以提前构建索引并复用，避免每次调用 `staticLookup` 都重复构建哈希表。

- [ ] **Step 1: 编写 buildLookupIndex 的测试**

在 `src/utils/vlookup-utils.test.js` 文件末尾添加：

```javascript
describe("buildLookupIndex", function () {
  var table = [
    ["张三", "研发部", "A001"],
    ["李四", "市场部", "A002"],
    ["王五", "财务部", "A003"],
  ];

  it("builds index keyed by match column values", function () {
    var index = buildLookupIndex(table, 0);
    expect(index["张三"]).toBe(0);
    expect(index["李四"]).toBe(1);
    expect(index["王五"]).toBe(2);
  });

  it("builds index against non-first column", function () {
    var index = buildLookupIndex(table, 2);
    expect(index["A001"]).toBe(0);
    expect(index["A002"]).toBe(1);
    expect(index["A003"]).toBe(2);
  });

  it("handles null key by converting to empty string", function () {
    var nullTable = [
      [null, "value1"],
      ["key2", "value2"],
    ];
    var index = buildLookupIndex(nullTable, 0);
    expect(index[""]).toBe(0);
    expect(index["key2"]).toBe(1);
  });

  it("later duplicate key overwrites earlier row index", function () {
    var dupTable = [
      ["dup", "first"],
      ["dup", "second"],
    ];
    var index = buildLookupIndex(dupTable, 0);
    expect(index["dup"]).toBe(1);
  });

  it("returns empty object for empty table", function () {
    var index = buildLookupIndex([], 0);
    expect(index).toEqual({});
  });
});
```

同时需要在文件顶部解构中添加 `buildLookupIndex`：

```javascript
var {
  parseRangeAddress,
  parseCellRef,
  buildColRange,
  buildIndexMatchFormula,
  staticLookup,
  buildLookupIndex,
} = require("./vlookup-utils");
```

- [ ] **Step 2: 运行测试验证它们失败**

Run: `npm test -- --testPathPattern="vlookup-utils.test.js"`
Expected: `buildLookupIndex` 相关测试 FAIL，因为函数尚未定义

- [ ] **Step 3: 实现 buildLookupIndex 函数**

在 `src/utils/vlookup-utils.js` 中，`staticLookup` 函数之前添加：

```javascript
function buildLookupIndex(lookupTable, matchColIndex) {
  var index = {};
  for (var r = 0; r < lookupTable.length; r++) {
    var key = lookupTable[r][matchColIndex];
    if (key === null || key === undefined) {
      key = "";
    }
    key = String(key);
    index[key] = r;
  }
  return index;
}
```

同时更新 `module.exports` 导出：

```javascript
module.exports = {
  parseRangeAddress: parseRangeAddress,
  parseCellRef: parseCellRef,
  buildColRange: buildColRange,
  buildIndexMatchFormula: buildIndexMatchFormula,
  staticLookup: staticLookup,
  buildLookupIndex: buildLookupIndex,
};
```

- [ ] **Step 4: 运行测试验证通过**

Run: `npm test -- --testPathPattern="vlookup-utils.test.js"`
Expected: 全部 21 个测试 PASS（16 个原有 + 5 个新增）

- [ ] **Step 5: 提交**

```bash
git add src/utils/vlookup-utils.js src/utils/vlookup-utils.test.js
git commit -m "feat(vlookup): add buildLookupIndex for reusable hash index"
```

---

### Task 3: staticLookup 增加 indexCache 参数

**Files:**
- Modify: `src/utils/vlookup-utils.js` — `staticLookup` 签名增加第 7 个参数 `indexCache`
- Modify: `src/utils/vlookup-utils.test.js` — 新增缓存场景测试

**Why:** 允许调用方传入预构建的索引，跳过重复的索引构建步骤。

- [ ] **Step 1: 编写 indexCache 相关测试**

在 `src/utils/vlookup-utils.test.js` 的 `describe("staticLookup"` 块末尾（`describe("buildLookupIndex"` 之前）添加：

```javascript
it("uses indexCache when provided, giving same results", function () {
  var table = [
    ["张三", "研发部", "A001", 8000],
    ["李四", "市场部", "A002", 6000],
    ["王五", "财务部", "A003", 7000],
  ];
  var indexCache = buildLookupIndex(table, 0);

  var resultWithCache = staticLookup(["张三", "李四", "王五"], table, 0, [1, 3], 0, "#N/A", indexCache);
  var resultWithoutCache = staticLookup(["张三", "李四", "王五"], table, 0, [1, 3], 0, "#N/A");

  expect(resultWithCache).toEqual(resultWithoutCache);
});

it("ignores indexCache in approximate match mode", function () {
  var numTable = [
    [100, "low"],
    [200, "mid"],
    [300, "high"],
  ];
  var indexCache = buildLookupIndex(numTable, 0);

  var result = staticLookup([150, 250, 350], numTable, 0, [1], 1, "#N/A", indexCache);

  expect(result).toEqual([["low"], ["mid"], ["high"]]);
});

it("works with indexCache when lookup value not found", function () {
  var table = [
    ["张三", "研发部"],
    ["李四", "市场部"],
  ];
  var indexCache = buildLookupIndex(table, 0);

  var result = staticLookup(["不存在"], table, 0, [1], 0, "#N/A", indexCache);

  expect(result).toEqual([["#N/A"]]);
});
```

- [ ] **Step 2: 运行测试验证新增测试失败**

Run: `npm test -- --testPathPattern="vlookup-utils.test.js"`
Expected: 3 个新增测试 FAIL（函数尚未接受 `indexCache` 参数，表现可能不同）

- [ ] **Step 3: 修改 staticLookup 支持 indexCache**

将 `staticLookup` 函数签名从：

```javascript
function staticLookup(
  lookupValues,
  lookupTable,
  matchColIndex,
  returnColIndices,
  matchMode,
  defaultValue
) {
```

改为：

```javascript
function staticLookup(
  lookupValues,
  lookupTable,
  matchColIndex,
  returnColIndices,
  matchMode,
  defaultValue,
  indexCache
) {
```

将索引构建逻辑（`if (matchMode === 0) { ... }` 块）从无条件构建改为优先使用缓存：

```javascript
  var index = {};
  if (matchMode === 0) {
    if (indexCache) {
      index = indexCache;
    } else {
      for (var r = 0; r < lookupTable.length; r++) {
        var key = lookupTable[r][matchColIndex];
        if (key === null || key === undefined) {
          key = "";
        }
        key = String(key);
        index[key] = r;
      }
    }
  }
```

- [ ] **Step 4: 运行测试验证全部通过**

Run: `npm test -- --testPathPattern="vlookup-utils.test.js"`
Expected: 全部 24 个测试 PASS（16 原有 + 5 buildLookupIndex + 3 indexCache）

- [ ] **Step 5: 提交**

```bash
git add src/utils/vlookup-utils.js src/utils/vlookup-utils.test.js
git commit -m "feat(vlookup): add indexCache parameter to staticLookup"
```

---

### Task 4: taskpane 使用索引缓存

**Files:**
- Modify: `src/taskpane/vlookup-taskpane.js` — `performLookup` 中构建索引并传入 `staticLookup`

**Why:** 在批量处理模式下，`staticLookup` 被多次调用（每个批次一次），每次都会重建索引。现改为首次前构建一次索引，所有后续调用传入缓存。

- [ ] **Step 1: 更新导入语句**

`src/taskpane/vlookup-taskpane.js` 第 8 行，将 `staticLookup` 改为同时引入 `buildLookupIndex`：

```javascript
var { staticLookup, parseRangeAddress, buildLookupIndex } = require("../utils/vlookup-utils");
```

- [ ] **Step 2: 在 performLookup 中首次调用前构建索引**

在 `performLookup` 函数中，`Excel.run` 回调内，读取完 `g_lookupTableData` 之后、首次调用 `staticLookup` 之前，添加索引构建：

对于**小数据模式**（`if (dataRowCount < LARGE_DATA_THRESHOLD)` 分支），在 `var lookupValues = ...` 之前添加：

```javascript
// Build lookup index once for reuse
var lookupIndex = null;
if (config.matchMode === 0) {
  lookupIndex = buildLookupIndex(g_lookupTableData, config.matchColIndex);
}
```

然后修改该分支中的 `staticLookup` 调用，增加第 7 个参数：

```javascript
var results = staticLookup(
  lookupValues,
  g_lookupTableData,
  config.matchColIndex,
  config.returnColIndices,
  config.matchMode,
  config.defaultValue,
  lookupIndex
);
```

对于**大数据模式**（`else` 分支），在 `while` 循环之前添加同样的索引构建：

```javascript
// Build lookup index once for all batches
var lookupIndex = null;
if (config.matchMode === 0) {
  lookupIndex = buildLookupIndex(g_lookupTableData, config.matchColIndex);
}
```

然后修改 `while` 循环内的 `staticLookup` 调用，增加第 7 个参数：

```javascript
var batchResults = staticLookup(
  batchLookupValues,
  g_lookupTableData,
  config.matchColIndex,
  config.returnColIndices,
  config.matchMode,
  config.defaultValue,
  lookupIndex
);
```

- [ ] **Step 3: 验证构建通过**

Run: `npm run build:dev`
Expected: BUILD SUCCESS，无编译错误

- [ ] **Step 4: 运行测试确认无回归**

Run: `npm test -- --testPathPattern="vlookup-utils.test.js"`
Expected: 全部 24 个测试 PASS

- [ ] **Step 5: 提交**

```bash
git add src/taskpane/vlookup-taskpane.js
git commit -m "perf(vlookup): cache lookup index across batch calls in performLookup"
```

---

### Task 5 (可选): 近似匹配二分查找

**Files:**
- Modify: `src/utils/vlookup-utils.js` — `staticLookup` 近似匹配分支改为二分查找
- Modify: `src/utils/vlookup-utils.test.js` — 新增二分查找验证测试

**Why:** 将近似匹配从 O(m × n) 线性扫描优化为 O(m × log n) 二分查找，进一步加速大数据量近似匹配场景。

**前置说明:** 此优化仅当查找表 match 列已按升序排列且中间不夹杂 null/undefined 值时完全等效于线性扫描。当前代码的线性扫描对 null/undefined 值使用 `continue` 跳过，二分查找版本改为跳过 null 区间继续搜索。

- [ ] **Step 1: 编写二分查找验证测试**

在 `src/utils/vlookup-utils.test.js` 的 `describe("staticLookup"` 块末尾添加：

```javascript
it("approximate match with binary search returns same as linear scan", function () {
  var numTable = [];
  for (var n = 0; n < 1000; n++) {
    numTable.push([n * 10, "val_" + n]);
  }

  var lookupValues = [15, 25, 95, 105, 555, 9995];
  var result = staticLookup(lookupValues, numTable, 0, [1], 1);

  // 15 → largest <= 15 is 10 → val_1
  // 25 → largest <= 25 is 20 → val_2
  // 95 → largest <= 95 is 90 → val_9
  // 105 → largest <= 105 is 100 → val_10
  // 555 → largest <= 555 is 550 → val_55
  // 9995 → largest <= 9995 is 9990 → val_999
  expect(result).toEqual([
    ["val_1"],
    ["val_2"],
    ["val_9"],
    ["val_10"],
    ["val_55"],
    ["val_999"],
  ]);
});

it("approximate match returns defaultValue when all values > lookup", function () {
  var numTable = [
    [100, "low"],
    [200, "mid"],
  ];
  var result = staticLookup([50], numTable, 0, [1], 1);
  expect(result).toEqual([["#N/A"]]);
});

it("approximate match works with nulls in match column (skipped)", function () {
  var numTable = [
    [null, "skip"],
    [100, "low"],
    [200, "mid"],
  ];
  var result = staticLookup([150], numTable, 0, [1], 1);
  expect(result).toEqual([["low"]]);
});
```

- [ ] **Step 2: 运行测试确认现有测试通过、新增测试通过/失败**

Run: `npm test -- --testPathPattern="vlookup-utils.test.js"`
Expected: 已有 24 个测试 PASS；新增 3 个测试中前 2 个应 PASS（线性扫描已有正确行为），第 3 个检查逻辑

- [ ] **Step 3: 将近似匹配线性扫描替换为二分查找**

将 `staticLookup` 中近似匹配分支（`else` 块内的 for 循环，即 `for (var m = 0; m < lookupTable.length; m++) { ... }` 部分）替换为：

```javascript
    } else {
      // Binary search: find largest value <= lookup
      // Assumes lookupTable is sorted ascending on match column (same as Excel VLOOKUP)
      var bestRow = -1;
      var lo = 0;
      var hi = lookupTable.length - 1;

      while (lo <= hi) {
        var mid = Math.floor((lo + hi) / 2);
        var tableVal = lookupTable[mid][matchColIndex];

        if (tableVal === null || tableVal === undefined) {
          // Search right for the next valid value to determine direction
          var foundValid = false;
          for (var scan = mid + 1; scan <= hi; scan++) {
            var scanVal = lookupTable[scan][matchColIndex];
            if (scanVal !== null && scanVal !== undefined) {
              if (Number(scanVal) <= Number(val)) {
                bestRow = scan;
                lo = scan + 1;
              } else {
                hi = mid - 1;
              }
              foundValid = true;
              break;
            }
          }
          if (!foundValid) {
            hi = mid - 1;
          }
        } else if (Number(tableVal) <= Number(val)) {
          bestRow = mid;
          lo = mid + 1;
        } else {
          hi = mid - 1;
        }
      }

      if (bestRow >= 0) {
        for (var n = 0; n < returnColIndices.length; n++) {
          row.push(lookupTable[bestRow][returnColIndices[n]]);
        }
      } else {
        for (var p = 0; p < returnColIndices.length; p++) {
          row.push(defaultValue);
        }
      }
    }
```

- [ ] **Step 4: 运行测试验证全部通过**

Run: `npm test -- --testPathPattern="vlookup-utils.test.js"`
Expected: 全部 27 个测试 PASS

- [ ] **Step 5: 提交**

```bash
git add src/utils/vlookup-utils.js src/utils/vlookup-utils.test.js
git commit -m "perf(vlookup): replace linear scan with binary search for approximate match"
```

---

## 自检

**1. Spec coverage:**
- 优化一（移除调试日志）→ Task 1
- 优化二（哈希索引缓存）→ Task 2 + Task 3 + Task 4
- 优化三（近似匹配二分查找）→ Task 5（可选）

**2. Placeholder scan:** 无 TBD/TODO/implement later 模式

**3. Type consistency:**
- `buildLookupIndex(lookupTable, matchColIndex)` 签名在 Task 2 定义，在 Task 3/4/5 中一致使用
- `indexCache` 作为 `staticLookup` 第 7 个参数，所有调用处一致传递
- 测试文件导入名称与 `module.exports` 导出一致
