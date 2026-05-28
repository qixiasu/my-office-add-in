# 增强查找功能实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal：** 修改增强查找插件：删除输出类型选项、删除预览区、新增默认值填充、添加大数据分批处理

**Architecture：** 修改 vlookup-taskpane.html/js/css，vlookup-utils.js。保持现有结构不变，仅做删减和新增功能点。

**Tech Stack：** Vanilla JavaScript, Office JavaScript API (Excel)

---

## 文件改动概览

| 文件 | 改动 |
|------|------|
| `src/taskpane/vlookup-taskpane.html` | 删除 outputType 单选框、vlookupPreview div；新增 defaultValue 输入框 |
| `src/taskpane/vlookup-taskpane.js` | 删除 formula 分支、showPreview()、outputType 读取；新增 defaultValue 传参、staticLookup 支持默认值参数、大数据分批逻辑 |
| `src/taskpane/vlookup-taskpane.css` | 移除 .vlookup-preview 相关样式 |

---

## Task 1: 修改 HTML — 删除 outputType、删除预览区、新增默认值输入框

**Files:**
- Modify: `src/taskpane/vlookup-taskpane.html:101-114`
- Modify: `src/taskpane/vlookup-taskpane.html:117-123`
- Modify: `src/taskpane/vlookup-taskpane.html:86-99`

- [ ] **Step 1: 删除输出类型单选框（101-114 行）**

删除整块：
```html
<!-- 输出类型 -->
<div class="vlookup-form-group">
    <span class="vlookup-section-title">输出类型</span>
    <div class="vlookup-radio-group">
        <label class="vlookup-radio-label">
            <input type="radio" name="outputType" value="formula" checked />
            <span>公式</span>
        </label>
        <label class="vlookup-radio-label">
            <input type="radio" name="outputType" value="static" />
            <span>静态值</span>
        </label>
    </div>
</div>
```

- [ ] **Step 2: 在"匹配模式"下方（99 行之后）插入默认值输入框，作为独立一行**

```html
<!-- 默认值 -->
<div class="vlookup-form-group">
    <label for="defaultValue" class="vlookup-label">查找不到时填充</label>
    <div class="vlookup-input-row">
        <input type="text" id="defaultValue" class="vlookup-input" placeholder="#N/A" value="#N/A" />
    </div>
</div>
```

- [ ] **Step 3: 删除预览区域（117-123 行）**

删除：
```html
<!-- Preview section (hidden by default) -->
<div id="vlookupPreview" class="vlookup-preview">
    <table id="previewTable">
        <thead id="previewHead"></thead>
        <tbody id="previewBody"></tbody>
    </table>
</div>
```

- [ ] **Step 4: 提交**
```bash
git add src/taskpane/vlookup-taskpane.html
git commit -m "refactor(vlookup): remove outputType radio and vlookupPreview div, add defaultValue input"
```

---

## Task 2: 修改 JS — 删除 formula 输出分支、删除 showPreview、删除 outputType 读取

**Files:**
- Modify: `src/taskpane/vlookup-taskpane.js:50-69`
- Modify: `src/taskpane/vlookup-taskpane.js:222-250`
- Modify: `src/taskpane/vlookup-taskpane.js:284-299`
- Modify: `src/taskpane/vlookup-taskpane.js:301-422`

- [ ] **Step 1: initEventListeners 中删除 executeBtn 之外的 refreshLookupValue/refreshLookupTable/slectAllColumns 等保持不变，此处只需确认 showPreview 调用已被移除**

确认 `loadTableHeaders()` 中不再调用 `showPreview()` — 在 Task 3 中一并删除 showPreview 函数本身。

- [ ] **Step 2: executeLookup() 中删除 outputType 字段读取（295 行）**

将：
```js
outputType: document.querySelector("input[name=\"outputType\"]:checked").value,
```
删除。

- [ ] **Step 3: performLookup() 中删除 formula 分支（321-361 行）**

删除：
```js
if (config.outputType === "formula") {
    // 整块 formula 逻辑
} else {
    // 保留此分支但移除 outputType 引用
}
```
改为直接进入现有 else 分支（静态值逻辑）。

同时删除 `buildIndexMatchFormula` import（不再需要）：
```js
var {
  buildColRange,
  buildIndexMatchFormula,
  staticLookup,
  parseRangeAddress,
} = require("../utils/vlookup-utils");
```
改为：
```js
var {
  buildColRange,
  staticLookup,
  parseRangeAddress,
} = require("../utils/vlookup-utils");
```

- [ ] **Step 4: 删除 showPreview 函数（222-250 行）**

删除整个 `showPreview` 函数。

- [ ] **Step 5: loadTableHeaders() 中删除 showPreview 调用（195 行）**

删除：
```js
showPreview(headers, g_lookupTableData, headerRow);
```

- [ ] **Step 6: 提交**
```bash
git add src/taskpane/vlookup-taskpane.js
git commit -m "refactor(vlookup): remove formula output branch, showPreview, and outputType field"
```

---

## Task 3: 修改 JS — staticLookup 支持默认值参数

**Files:**
- Modify: `src/utils/vlookup-utils.js:77-147`
- Modify: `src/taskpane/vlookup-taskpane.js:387-393`

- [ ] **Step 1: staticLookup 增加 defaultValue 参数，将所有硬编码 "#N/A" 替换为该参数**

将函数签名从：
```js
function staticLookup(lookupValues, lookupTable, matchColIndex, returnColIndices, matchMode)
```
改为：
```js
function staticLookup(lookupValues, lookupTable, matchColIndex, returnColIndices, matchMode, defaultValue)
```

在函数开头添加：
```js
if (defaultValue === undefined || defaultValue === null) {
  defaultValue = "#N/A";
}
```

三处硬编码 `"#N/A"` 替换为 `defaultValue`：
- 近似模式下 null/undefined 查找值（101 行附近）
- 精确匹配未找到（116 行附近）
- 近似匹配无结果（138 行附近）

- [ ] **Step 2: performLookup 中传递 defaultValue 给 staticLookup**

在 `staticLookup()` 调用处（387 行附近）添加第六个参数：
```js
var results = staticLookup(
  lookupValues,
  g_lookupTableData,
  config.matchColIndex,
  config.returnColIndices,
  config.matchMode,
  config.defaultValue  // 新增
);
```

- [ ] **Step 3: performLookup 中从 config 读取 defaultValue**

在 `executeLookup()` config 对象中新增：
```js
defaultValue: document.getElementById("defaultValue").value || "#N/A",
```

- [ ] **Step 4: 提交**
```bash
git add src/utils/vlookup-utils.js src/taskpane/vlookup-taskpane.js
git commit -m "feat(vlookup): add defaultValue parameter to staticLookup for custom not-found fill"
```

---

## Task 4: 修改 JS — 大数据分批处理逻辑

**Files:**
- Modify: `src/taskpane/vlookup-taskpane.js:301-422`

- [ ] **Step 1: 在 performLookup 顶部添加阈值常量**

```js
var BATCH_SIZE = 10000;
var LARGE_DATA_THRESHOLD = 100000;
```

- [ ] **Step 2: 将现有 performLookup 中的逻辑包装为小数据分支（dataRowCount < LARGE_DATA_THRESHOLD）**

现有逻辑保持不变，仅在外面包裹 if 分支：
```js
if (dataRowCount < LARGE_DATA_THRESHOLD) {
  // 现有一次性处理逻辑
} else {
  // 分批处理逻辑（见 Step 3）
}
```

- [ ] **Step 3: 实现分批处理逻辑**

在 else 分支中：
```js
var BATCH_SIZE = 10000;
var totalRows = dataRowCount;
var processedRows = 0;

while (processedRows < totalRows) {
  var currentBatchSize = Math.min(BATCH_SIZE, totalRows - processedRows);
  var batchStartRow = dataStartRow + processedRows;
  var batchEndRow = batchStartRow + currentBatchSize - 1;

  // 读取当前批次 lookupValues
  var batchDataRange = worksheet.getRange(
    dataColLetter + batchStartRow + ":" + dataColLetter + batchEndRow
  );
  batchDataRange.load("values");
  await context.sync();

  var batchLookupValues = [];
  for (var bj = 0; bj < batchDataRange.values.length; bj++) {
    batchLookupValues.push(batchDataRange.values[bj][0]);
  }

  // 计算当前批次
  var batchResults = staticLookup(
    batchLookupValues,
    g_lookupTableData,
    config.matchColIndex,
    config.returnColIndices,
    config.matchMode,
    config.defaultValue
  );

  // 写入当前批次结果
  var batchTargetRange = worksheet.getRange(
    getColumnLetter(insertPos2) + batchStartRow + ":" +
    getColumnLetter(insertPos2 + returnColCount2 - 1) + batchEndRow
  );
  batchTargetRange.values = batchResults;
  await context.sync();

  processedRows += currentBatchSize;

  // 更新进度状态
  setStatus("处理中... 已完成 " + processedRows + " / " + totalRows + " 行", "info");
}
```

注意：insertPos2 和 returnColCount2 在判断大小数据之前已定义（396-401 行），在 if 分支之前声明即可。

- [ ] **Step 4: 提交**
```bash
git add src/taskpane/vlookup-taskpane.js
git commit -m "feat(vlookup): add batch processing for large datasets (>100k rows)"
```

---

## Task 5: 修改 CSS — 移除 .vlookup-preview 相关样式

**Files:**
- Modify: `src/taskpane/vlookup-taskpane.css`

- [ ] **Step 1: 查找并删除 .vlookup-preview 相关样式**

搜索文件中 `.vlookup-preview` 相关内容并删除（包括 `.vlookup-preview.visible` 等）。

- [ ] **Step 2: 提交**
```bash
git add src/taskpane/vlookup-taskpane.css
git commit -m "refactor(vlookup): remove unused .vlookup-preview styles"
```

---

## Task 6: 整体测试

**Files:**
- Test: `src/taskpane/vlookup-taskpane.js`

- [ ] **Step 1: 使用 /verify 启动 Excel Desktop 测试以下场景：**

1. 精确匹配 + 默认值 #N/A → 确认未找到时填充 #N/A
2. 自定义默认值（如 "未找到"）→ 确认输出为该自定义值
3. 大数据（>100,000 行）→ 确认分批处理正常，进度状态更新
4. 近似匹配模式 → 确认默认值逻辑正确应用
5. 删除功能验证：无 outputType 单选框、无预览区

- [ ] **Step 2: 运行 lint 检查**
```bash
npm run lint
```

---

## Task 7: 合并提交（如所有测试通过）**

- [ ] 将 Task 1-6 的所有改动合并为一个有意义的 commit：
```bash
git add -A
git commit -m "feat(vlookup): simplify output type, add default value fill, add batch processing for large datasets"
```

---

## 实现顺序

1. Task 1（HTML 修改）
2. Task 2（JS 删除逻辑）
3. Task 3（默认值参数）
4. Task 4（大数据分批）
5. Task 5（CSS 清理）
6. Task 6（测试验证）
7. Task 7（合并提交）