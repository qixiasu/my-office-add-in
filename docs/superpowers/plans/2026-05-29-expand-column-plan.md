# 展开列工具 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 新建独立侧边栏工具「展开列」，将宽表数据转换为长表格式。

**Architecture:** 纯 JavaScript 实现，遵循现有项目模式。选中区域 → 读取数据 → 展开算法 → 创建新工作表输出结果。

**Tech Stack:** Office JavaScript API, plain JavaScript, Jest for tests

---

## File Structure

**新建文件：**
- `src/taskpane/expand-taskpane.html` — 侧边栏 HTML 结构（参考 concat-taskpane.html）
- `src/taskpane/expand-taskpane.css` — 样式（复用 concat-taskpane.css 视觉风格）
- `src/taskpane/expand-taskpane.js` — 业务逻辑（参考 concat-taskpane.js）
- `src/utils/expand-utils.js` — 展开算法工具函数
- `src/utils/expand-utils.test.js` — 单元测试
- `assets/expand-16.png`, `assets/expand-32.png`, `assets/expand-80.png` — 图标

**修改文件：**
- `manifest.xml` — 注册按钮和资源
- `webpack.config.js` — 添加 entry point 和 HtmlWebpackPlugin

---

### Task 1: 创建 expand-utils.js 展开算法

**Files:**
- Create: `src/utils/expand-utils.js`
- Test: `src/utils/expand-utils.test.js`

- [ ] **Step 1: 创建测试文件**

```javascript
// src/utils/expand-utils.test.js
var { expandData } = require("./expand-utils");

describe("expandData", function () {
  it("expands two data rows correctly", function () {
    var values = [
      ["省份", "地市", "地市", "地市"],
      ["河南", "郑州", "南阳", "许昌"],
      ["河北", "石家庄", "保定", ""],
    ];
    var result = expandData(values);
    expect(result).toEqual([
      ["河南", "郑州"],
      ["河南", "南阳"],
      ["河南", "许昌"],
      ["河北", "石家庄"],
      ["河北", "保定"],
    ]);
  });

  it("handles single data row", function () {
    var values = [
      ["省份", "地市", "地市"],
      ["河南", "郑州", "南阳"],
    ];
    var result = expandData(values);
    expect(result).toEqual([
      ["河南", "郑州"],
      ["河南", "南阳"],
    ]);
  });

  it("skips empty cells", function () {
    var values = [
      ["省份", "地市", "地市"],
      ["河南", "郑州", ""],
    ];
    var result = expandData(values);
    expect(result).toEqual([
      ["河南", "郑州"],
    ]);
  });

  it("returns empty array for header-only data", function () {
    var values = [["省份", "地市"]];
    var result = expandData(values);
    expect(result).toEqual([]);
  });
});
```

- [ ] **Step 2: 运行测试验证失败**

Run: `npm test -- --testPathPattern=expand-utils.test.js`
Expected: FAIL with "Cannot find module './expand-utils'"

- [ ] **Step 3: 编写最小实现**

```javascript
// src/utils/expand-utils.js

/**
 * 展开宽表数据为长表格式
 * @param {Array<Array>} values - 2D数组，第一行是表头
 * @returns {Array<Array>} 展开后的2D数组，每行 [键列值, 展开列值]
 */
function expandData(values) {
  if (!values || values.length <= 1) {
    return [];
  }

  var result = [];

  // 从第1行开始是数据（第0行是表头）
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var key = row[0];

    // 遍历其他列（从索引1开始）
    for (var j = 1; j < row.length; j++) {
      var val = row[j];
      // 跳过空单元格
      if (val !== null && val !== "" && val !== undefined) {
        result.push([key, val]);
      }
    }
  }

  return result;
}

module.exports = {
  expandData: expandData,
};
```

- [ ] **Step 4: 运行测试验证通过**

Run: `npm test -- --testPathPattern=expand-utils.test.js`
Expected: PASS

- [ ] **Step 5: 提交**

```bash
git add src/utils/expand-utils.js src/utils/expand-utils.test.js
git commit -m "feat: add expand-utils with expandData function"
```

---

### Task 2: 创建 expand-taskpane.html 侧边栏界面

**Files:**
- Create: `src/taskpane/expand-taskpane.html`
- Reference: `src/taskpane/concat-taskpane.html`

- [ ] **Step 1: 创建 HTML 文件**

```html
<!-- Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT License. -->

<!DOCTYPE html>
<html>

<head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>展开列工具</title>

    <!-- Office JavaScript API -->
    <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>

    <!-- For more information on Fluent UI, visit https://developer.microsoft.com/fluentui#/. -->
    <link rel="stylesheet" href="https://res-1.cdn.office.net/files/fabric-cdn-prod_20230815.002/office-ui-fabric-core/11.1.0/css/fabric.min.css"/>

    <!-- Template styles -->
    <link href="taskpane.css" rel="stylesheet" type="text/css" />
    <link href="expand-taskpane.css" rel="stylesheet" type="text/css" />
</head>

<body class="ms-font-m ms-welcome ms-Fabric">
    <div class="concat-container">
        <h1 class="concat-title">
            <svg width="20" height="20" viewBox="0 0 20 20" fill="#0078d4">
                <path d="M16 5H4a1 1 0 0 0-1 1v8a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1V6a1 1 0 0 0-1-1zm-1 8H5V7h10v6z"/>
                <path d="M7 9h2v2H7zM11 9h2v2h-2zM7 12h2v2H7zM11 12h2v2h-2z"/>
            </svg>
            展开列工具
        </h1>

        <div class="guide-card">
            <p>
                <span class="guide-step">1. 在 Excel 中选中包含数据的区域</span><br>
                <span class="guide-step">2. 点击「执行展开」按钮</span>
            </p>
        </div>

        <div class="concat-form-group">
            <label class="concat-label">当前选中</label>
            <div id="selectionInfo" class="selection-info">等待选择数据...</div>
        </div>

        <button id="executeBtn" class="concat-button">执行展开</button>

        <div id="status" class="status-message status-idle">状态：等待操作...</div>
    </div>
</body>

</html>
```

- [ ] **Step 2: 提交**

```bash
git add src/taskpane/expand-taskpane.html
git commit -m "feat: create expand-taskpane.html skeleton"
```

---

### Task 3: 创建 expand-taskpane.css 样式

**Files:**
- Create: `src/taskpane/expand-taskpane.css`
- Reference: `src/taskpane/concat-taskpane.css`

- [ ] **Step 1: 创建 CSS 文件**

```css
/* 选中信息显示 */
.selection-info {
  background: #f5f5f5;
  border: 1px solid #e0e0e0;
  border-radius: 4px;
  padding: 8px 12px;
  font-size: 13px;
  color: #666;
  word-break: break-all;
}

.selection-info.has-selection {
  background: #f0f7ff;
  border-color: #0078d4;
  color: #333;
}
```

- [ ] **Step 2: 提交**

```bash
git add src/taskpane/expand-taskpane.css
git commit -m "feat: create expand-taskpane.css styles"
```

---

### Task 4: 创建 expand-taskpane.js 核心逻辑

**Files:**
- Create: `src/taskpane/expand-taskpane.js`
- Reference: `src/taskpane/concat-taskpane.js`

- [ ] **Step 1: 创建 JS 文件**

```javascript
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

var expandUtils = require("../utils/expand-utils");

var MAX_ROWS = 1050000;

Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    var executeBtn = document.getElementById("executeBtn");

    // 按钮点击触发展开
    executeBtn.addEventListener("click", function () {
      runExpand();
    });

    // 初始化时显示当前选中区域
    updateSelectionInfo();
  }
});

function setStatus(message, type) {
  var el = document.getElementById("status");
  el.textContent = message;
  el.className = "status-message status-" + type;
}

function updateSelectionInfo() {
  Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    range.load(["address"]);
    return context.sync().then(function () {
      var selectionInfo = document.getElementById("selectionInfo");
      selectionInfo.textContent = range.address;
      selectionInfo.className = "selection-info has-selection";
    });
  }).catch(function (error) {
    console.error("Failed to get selection:", error);
  });
}

function getColumnLetter(colIndex) {
  var letter = "";
  var remaining = colIndex;
  do {
    letter = String.fromCharCode(65 + (remaining % 26)) + letter;
    remaining = Math.floor(remaining / 26) - 1;
  } while (remaining >= 0);
  return letter;
}

function runExpand() {
  var executeBtn = document.getElementById("executeBtn");

  // 禁用按钮，显示加载状态
  executeBtn.disabled = true;
  setStatus("处理中...", "loading");

  Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    range.load(["address", "columnCount", "rowCount"]);
    return context.sync().then(function () {
      // 校验：至少2列
      if (range.columnCount < 2) {
        executeBtn.disabled = false;
        setStatus("错误：只有一列数据，无需展开", "error");
        return;
      }

      // 读取数据
      var values = range.values;
      if (values.length <= 1) {
        executeBtn.disabled = false;
        setStatus("错误：选中区域没有数据", "error");
        return;
      }

      // 执行展开算法
      var result = expandUtils.expandData(values);

      if (result.length === 0) {
        executeBtn.disabled = false;
        setStatus("错误：没有可展开的数据", "error");
        return;
      }

      // 获取原工作表名称
      var originalSheet = context.workbook.worksheets.getActiveWorksheet();
      var originalSheetName = originalSheet.name;
      var newSheetName = originalSheetName + "_展开";

      // 创建新工作表（在原工作表之后）
      var sheetCollection = context.workbook.worksheets;
      sheetCollection.load("items");
      return context.sync().then(function () {
        // 检查是否存在同名工作表，如果存在则添加后缀
        var existingNames = sheetCollection.items.map(function (s) {
          return s.name;
        });

        var finalSheetName = newSheetName;
        var counter = 1;
        while (existingNames.indexOf(finalSheetName) !== -1) {
          finalSheetName = newSheetName + " (" + counter + ")";
          counter++;
        }

        // 添加新工作表
        var newSheet = sheetCollection.add(originalSheetName + "_展开");
        newSheet.position = originalSheet.index + 1;

        // 写入表头
        var headerRange = newSheet.getRange("A1:B1");
        headerRange.values = [[values[0][0], values[0][1] || "展开值"]];

        // 写入展开后的数据
        if (result.length > 0) {
          var dataRange = newSheet.getRange("A2:B" + (result.length + 1));
          dataRange.values = result;
        }

        return context.sync().then(function () {
          setStatus(
            "完成! 已在「" + finalSheetName + "」工作表写入 " + result.length + " 行数据",
            "success"
          );
          executeBtn.disabled = false;
        });
      });
    });
  }).catch(function (error) {
    setStatus("错误: " + error.message, "error");
    executeBtn.disabled = false;
  });
}
```

- [ ] **Step 2: 提交**

```bash
git add src/taskpane/expand-taskpane.js
git commit -m "feat: create expand-taskpane.js core logic"
```

---

### Task 5: 更新 webpack.config.js 添加 entry point

**Files:**
- Modify: `webpack.config.js:22-27` (entry section)
- Modify: `webpack.config.js:107-111` (HtmlWebpackPlugin section)

- [ ] **Step 1: 添加 entry point**

在 entry 对象的 `"count-values-taskpane"` 后添加：
```javascript
"expand-taskpane": ["./src/taskpane/expand-taskpane.js", "./src/taskpane/expand-taskpane.html"],
```

- [ ] **Step 2: 添加 HtmlWebpackPlugin**

在 `new HtmlWebpackPlugin` for `count-values-taskpane.html` 后添加：
```javascript
new HtmlWebpackPlugin({
  filename: "expand-taskpane.html",
  template: "./src/taskpane/expand-taskpane.html",
  chunks: ["polyfill", "expand-taskpane"],
}),
```

- [ ] **Step 3: 提交**

```bash
git add webpack.config.js
git commit -m "feat: add expand-taskpane webpack entry point"
```

---

### Task 6: 更新 manifest.xml 注册按钮和资源

**Files:**
- Modify: `manifest.xml:84-85` (新增按钮到 DataProcessingGroup)
- Modify: `manifest.xml:147-150` (添加图标资源)
- Modify: `manifest.xml:157-158` (添加 Url 资源)
- Modify: `manifest.xml:164-165` (添加 ShortStrings)
- Modify: `manifest.xml:176-177` (添加 LongStrings)

- [ ] **Step 1: 在 DataProcessingGroup 添加按钮**

在 `CountValuesButton` 按钮后添加：
```xml
<Control xsi:type="Button" id="ExpandColumnButton">
  <Label resid="ExpandColumnButton.Label"/>
  <Supertip>
    <Title resid="ExpandColumnButton.Label"/>
    <Description resid="ExpandColumnButton.Tooltip"/>
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="ExpandIcon.16x16"/>
    <bt:Image size="32" resid="ExpandIcon.32x32"/>
    <bt:Image size="80" resid="ExpandIcon.80x80"/>
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <TaskpaneId>ExpandTaskpaneId</TaskpaneId>
    <SourceLocation resid="ExpandTaskpane.Url"/>
  </Action>
</Control>
```

- [ ] **Step 2: 在 bt:Images 中添加图标资源**

在 `CountValuesIcon.80x80` 后添加：
```xml
<bt:Image id="ExpandIcon.16x16" DefaultValue="https://localhost:3000/assets/expand-16.png"/>
<bt:Image id="ExpandIcon.32x32" DefaultValue="https://localhost:3000/assets/expand-32.png"/>
<bt:Image id="ExpandIcon.80x80" DefaultValue="https://localhost:3000/assets/expand-80.png"/>
```

- [ ] **Step 3: 在 bt:Urls 中添加 Url 资源**

在 `CountValuesTaskpane.Url` 后添加：
```xml
<bt:Url id="ExpandTaskpane.Url" DefaultValue="https://localhost:3000/expand-taskpane.html"/>
```

- [ ] **Step 4: 在 bt:ShortStrings 中添加标签**

在 `CountValuesButton.Label` 后添加：
```xml
<bt:String id="ExpandColumnButton.Label" DefaultValue="展开列"/>
```

- [ ] **Step 5: 在 bt:LongStrings 中添加提示**

在 `CountValuesButton.Tooltip` 后添加：
```xml
<bt:String id="ExpandColumnButton.Tooltip" DefaultValue="将宽表数据转换为长表格式，第一列作为键，其他列的值展开为独立行"/>
```

- [ ] **Step 6: 提交**

```bash
git add manifest.xml
git commit -m "feat: register expand column tool in manifest"
```

---

### Task 7: 生成图标文件

**Files:**
- Create: `assets/expand-16.png`
- Create: `assets/expand-32.png`
- Create: `assets/expand-80.png`

- [ ] **Step 1: 检查现有图标生成脚本**

查看 `scripts/generate_icons.py` 是否可以复用或需要添加新图标的绘制函数。

- [ ] **Step 2: 生成图标**

如果脚本支持，直接运行生成。如果不支持，创建简单的展开图标（双向箭头表示展开方向）。

- [ ] **Step 3: 提交**

```bash
git add assets/expand-*.png
git commit -m "feat: add expand tool icons"
```

---

### Task 8: 完整流程测试

**Testing Scope:**
- [ ] 选中包含表头和多列数据的区域，点击执行展开
- [ ] 验证新工作表创建位置正确（原始工作表之后）
- [ ] 验证表头正确（第一列列名 + "展开值"）
- [ ] 验证展开结果正确（每个键值对一行）
- [ ] 验证空白单元格被正确跳过
- [ ] 验证单列数据时报错提示
- [ ] 验证空数据时报错提示
- [ ] 验证同名工作表存在时自动重命名

---

### Spec Coverage Check

| Spec Requirement | Task |
|-----------------|------|
| 独立侧边栏工具 | Task 2, 3, 4 |
| 用户选中区域 | Task 4 |
| 表头处理（跳过第一行） | Task 1 |
| 输出到新工作表 | Task 4 |
| 自动命名（xxx_展开） | Task 4 |
| 展开算法 | Task 1 |
| Ribbon 按钮注册 | Task 6 |
| 错误处理 | Task 4 |

### Type Consistency Check

- `expand-utils.js` 导出 `expandData` 函数 ✓
- `expand-taskpane.js` 引入 `expandUtils` 模块 ✓
- manifest.xml 中 `ExpandTaskpane.Url` 与 `webpack.config.js` 中 `expand-taskpane.html` 对应 ✓