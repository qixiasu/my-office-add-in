# 增强查找 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 实现 INDEX/MATCH 替代 VLOOKUP 的增强查找功能，通过弹窗一次配置完成多列查找，解决多列 VLOOKUP 繁琐和匹配列必须在左侧的痛点。

**Architecture:** 新增 `vlookup-utils.js` 工具模块（公式构建 + 静态值匹配），新增 `vlookup-dialog.html` 弹窗（配置 UI），修改 `commands.html` 注册 ribbon 命令，修改 `taskpane` 增加调试按钮，修改 `manifest.xml` 和 `webpack.config.js`。

**Tech Stack:** JavaScript (IE11 兼容), Office.js Excel API, Jest, webpack

---

### 文件结构

| 文件 | 操作 | 职责 |
|------|------|------|
| `src/utils/vlookup-utils.js` | 新增 | 公式构建、静态值匹配、地址解析 |
| `src/utils/vlookup-utils.test.js` | 新增 | 单元测试 |
| `src/commands/vlookup-dialog.html` | 新增 | 弹窗 UI + 配置逻辑 |
| `src/commands/commands.html` | 修改 | 注册 enhancedVlookup 命令 |
| `src/taskpane/taskpane.html` | 修改 | 新增"增强查找"按钮 |
| `src/taskpane/taskpane.js` | 修改 | 新增弹窗打开/消息处理逻辑 |
| `manifest.xml` | 修改 | 新增 ribbon 按钮和资源字符串 |
| `webpack.config.js` | 修改 | 新增 vlookup-dialog 复制到输出 |

---

### Task 1: vlookup-utils 工具函数

**Files:**
- Create: `src/utils/vlookup-utils.js`

- [ ] **Step 1: 创建 vlookup-utils.js**

```javascript
var { getColumnLetter } = require("./concat-utils");

var SEPA = "!";
var ABS = "$";

/**
 * Parse a range address like "Sheet2!A1:D500" or "A1:D500" into components.
 * Returns { sheet, startCol, startRow, endCol, endRow, colCount, rowCount }.
 */
function parseRangeAddress(address) {
  var sheet = "";
  var rangePart = address;

  var bangIdx = address.indexOf(SEPA);
  if (bangIdx !== -1) {
    sheet = address.substring(0, bangIdx);
    rangePart = address.substring(bangIdx + 1);
  }

  var parts = rangePart.split(":");
  var start = parseCellRef(parts[0]);
  var end = parseCellRef(parts[1] || parts[0]);

  return {
    sheet: sheet,
    startCol: start.col,
    startRow: start.row,
    endCol: end.col,
    endRow: end.row,
    colCount: end.col - start.col + 1,
    rowCount: end.row - start.row + 1
  };
}

/**
 * Parse "A1" → { col: 0, row: 1 }
 */
function parseCellRef(ref) {
  var col = 0;
  var row = 0;
  var i = 0;

  while (i < ref.length) {
    var ch = ref.charCodeAt(i);
    if (ch >= 65 && ch <= 90) {
      col = col * 26 + (ch - 65);
      i++;
    } else {
      break;
    }
  }

  row = parseInt(ref.substring(i), 10) || 1;
  return { col: col, row: row };
}

/**
 * Build an absolute column range string from a parsed range.
 * e.g., buildColRange({ sheet:"Sheet2", startCol:0, startRow:1, endRow:500 }, 1) → "Sheet2!$B$1:$B$500"
 */
function buildColRange(parsed, colIndex) {
  var colLetter = getColumnLetter(parsed.startCol + colIndex);
  var prefix = parsed.sheet ? parsed.sheet + SEPA : "";
  var colAbs = ABS + colLetter + ABS;
  return prefix + colAbs + parsed.startRow + ":" + colAbs + parsed.endRow;
}

/**
 * Build INDEX/MATCH formula.
 * @param {string} lookupCellRef - Relative reference to lookup value cell, e.g., "A1"
 * @param {string} lookupColRange - Absolute range of lookup column, e.g., "Sheet2!$A$1:$A$500"
 * @param {string} returnColRange - Absolute range of return column, e.g., "Sheet2!$B$1:$B$500"
 * @param {number} matchMode - 0 for exact match, 1 for approximate (less than or equal)
 */
function buildIndexMatchFormula(lookupCellRef, lookupColRange, returnColRange, matchMode) {
  return "=INDEX(" + returnColRange + ", MATCH(" + lookupCellRef + ", " + lookupColRange + ", " + matchMode + "))";
}

/**
 * Perform static (in-memory) lookup.
 * @param {Array} lookupValues - 1D array of values to look up
 * @param {Array} lookupTable - 2D array [row][col] of lookup table data
 * @param {number} matchColIndex - Column index in lookupTable used for matching (0-based)
 * @param {number[]} returnColIndices - Column indices in lookupTable to return (0-based)
 * @param {number} matchMode - 0 for exact, 1 for approximate (sorted ascending assumed)
 * @returns {Array} 2D array [valueIndex][returnColIndex] of results, #N/A as string for not found
 */
function staticLookup(lookupValues, lookupTable, matchColIndex, returnColIndices, matchMode) {
  var results = [];

  // Build hash map for exact match
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
    var val = lookupValues[i];
    if (val === null || val === undefined) {
      val = "";
    }
    var valStr = String(val);
    var row = [];

    if (matchMode === 0) {
      var matchedRow = index[valStr];
      if (matchedRow !== undefined) {
        for (var j = 0; j < returnColIndices.length; j++) {
          row.push(lookupTable[matchedRow][returnColIndices[j]]);
        }
      } else {
        for (var k = 0; k < returnColIndices.length; k++) {
          row.push("#N/A");
        }
      }
    } else {
      // Approximate match: find largest value <= lookup
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
          row.push("#N/A");
        }
      }
    }

    results.push(row);
  }

  return results;
}

module.exports = {
  parseRangeAddress: parseRangeAddress,
  parseCellRef: parseCellRef,
  buildColRange: buildColRange,
  buildIndexMatchFormula: buildIndexMatchFormula,
  staticLookup: staticLookup
};
```

---

### Task 2: vlookup-utils 单元测试

**Files:**
- Create: `src/utils/vlookup-utils.test.js`

- [ ] **Step 1: 编写测试**

```javascript
var {
  parseRangeAddress,
  parseCellRef,
  buildColRange,
  buildIndexMatchFormula,
  staticLookup
} = require("./vlookup-utils");

describe("parseCellRef", function () {
  it("parses A1 correctly", function () {
    var result = parseCellRef("A1");
    expect(result.col).toBe(0);
    expect(result.row).toBe(1);
  });

  it("parses Z1 correctly", function () {
    var result = parseCellRef("Z1");
    expect(result.col).toBe(25);
    expect(result.row).toBe(1);
  });

  it("parses AA1 correctly", function () {
    var result = parseCellRef("AA1");
    expect(result.col).toBe(26);
    expect(result.row).toBe(1);
  });

  it("parses AB10 correctly", function () {
    var result = parseCellRef("AB10");
    expect(result.col).toBe(27);
    expect(result.row).toBe(10);
  });
});

describe("parseRangeAddress", function () {
  it("parses simple range without sheet", function () {
    var result = parseRangeAddress("A1:D500");
    expect(result.sheet).toBe("");
    expect(result.startCol).toBe(0);
    expect(result.startRow).toBe(1);
    expect(result.endCol).toBe(3);
    expect(result.endRow).toBe(500);
    expect(result.colCount).toBe(4);
    expect(result.rowCount).toBe(500);
  });

  it("parses range with sheet name", function () {
    var result = parseRangeAddress("Sheet2!A1:D500");
    expect(result.sheet).toBe("Sheet2");
    expect(result.startCol).toBe(0);
    expect(result.startRow).toBe(1);
    expect(result.endCol).toBe(3);
    expect(result.endRow).toBe(500);
    expect(result.colCount).toBe(4);
  });

  it("parses single cell range", function () {
    var result = parseRangeAddress("Sheet2!A1:A1");
    expect(result.sheet).toBe("Sheet2");
    expect(result.colCount).toBe(1);
    expect(result.rowCount).toBe(1);
  });

  it("parses range with sheet name containing special chars", function () {
    var result = parseRangeAddress("My Sheet!B2:E10");
    expect(result.sheet).toBe("My Sheet");
    expect(result.startCol).toBe(1);
    expect(result.startRow).toBe(2);
    expect(result.endCol).toBe(4);
    expect(result.endRow).toBe(10);
    expect(result.colCount).toBe(4);
  });
});

describe("buildColRange", function () {
  it("builds absolute column range without sheet", function () {
    var parsed = parseRangeAddress("A1:D500");
    var result = buildColRange(parsed, 0);
    expect(result).toBe("$A$1:$A$500");
  });

  it("builds absolute column range with sheet", function () {
    var parsed = parseRangeAddress("Sheet2!A1:D500");
    var result = buildColRange(parsed, 1);
    expect(result).toBe("Sheet2!$B$1:$B$500");
  });

  it("builds range for last column", function () {
    var parsed = parseRangeAddress("Sheet2!A1:D500");
    var result = buildColRange(parsed, 3);
    expect(result).toBe("Sheet2!$D$1:$D$500");
  });
});

describe("buildIndexMatchFormula", function () {
  it("builds formula with exact match", function () {
    var result = buildIndexMatchFormula("A1", "Sheet2!$A$1:$A$500", "Sheet2!$B$1:$B$500", 0);
    expect(result).toBe("=INDEX(Sheet2!$B$1:$B$500, MATCH(A1, Sheet2!$A$1:$A$500, 0))");
  });

  it("builds formula with approximate match", function () {
    var result = buildIndexMatchFormula("A1", "$A$1:$A$500", "$C$1:$C$500", 1);
    expect(result).toBe("=INDEX($C$1:$C$500, MATCH(A1, $A$1:$A$500, 1))");
  });

  it("builds formula with different lookup cell reference", function () {
    var result = buildIndexMatchFormula("B5", "Sheet2!$A$1:$A$500", "Sheet2!$D$1:$D$500", 0);
    expect(result).toBe("=INDEX(Sheet2!$D$1:$D$500, MATCH(B5, Sheet2!$A$1:$A$500, 0))");
  });
});

describe("staticLookup", function () {
  var table = [
    ["张三", "研发部", "A001", 8000],
    ["李四", "市场部", "A002", 6000],
    ["王五", "财务部", "A003", 7000]
  ];

  it("exact match returns correct values", function () {
    var result = staticLookup(["张三", "李四"], table, 0, [1, 3], 0);
    expect(result).toEqual([
      ["研发部", 8000],
      ["市场部", 6000]
    ]);
  });

  it("returns #N/A for not found values", function () {
    var result = staticLookup(["张三", "不存在"], table, 0, [1], 0);
    expect(result).toEqual([
      ["研发部"],
      ["#N/A"]
    ]);
  });

  it("returns multiple return columns", function () {
    var result = staticLookup(["王五"], table, 0, [1, 2, 3], 0);
    expect(result).toEqual([
      ["财务部", "A003", 7000]
    ]);
  });

  it("handles null lookup value", function () {
    var result = staticLookup([null], table, 0, [1], 0);
    expect(result).toEqual([["#N/A"]]);
  });

  it("handles empty lookup array", function () {
    var result = staticLookup([], table, 0, [1], 0);
    expect(result).toEqual([]);
  });

  it("matches against non-first column", function () {
    var result = staticLookup(["A002"], table, 2, [0, 3], 0);
    expect(result).toEqual([["李四", 6000]]);
  });

  it("approximate match finds largest value <= lookup", function () {
    var numTable = [
      [100, "low"],
      [200, "mid"],
      [300, "high"]
    ];
    var result = staticLookup([150, 250, 350], numTable, 0, [1], 1);
    expect(result).toEqual([["low"], ["mid"], ["high"]]);
  });
});
```

- [ ] **Step 2: 运行测试确认失败**

```bash
npx jest src/utils/vlookup-utils.test.js
```
Expected: FAIL (vlookup-utils.js 不存在)

- [ ] **Step 3: 创建 vlookup-utils.js**（见 Task 1 Step 1 代码）

- [ ] **Step 4: 运行测试确认通过**

```bash
npx jest src/utils/vlookup-utils.test.js
```
Expected: all 19 tests PASS

- [ ] **Step 5: 提交**

```bash
git add src/utils/vlookup-utils.js src/utils/vlookup-utils.test.js
git commit -m "feat: add vlookup-utils with INDEX/MATCH formula builder and static lookup

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>"
```

---

### Task 3: vlookup-dialog.html 弹窗页面

**Files:**
- Create: `src/commands/vlookup-dialog.html`

- [ ] **Step 1: 创建弹窗页面**

```html
<!-- Enhanced VLOOKUP dialog for ribbon button -->
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    <style>
        body {
            font-family: "Segoe UI", sans-serif;
            padding: 20px;
            margin: 0;
            font-size: 14px;
        }
        h3 { margin-top: 0; }
        .row { margin-bottom: 14px; }
        label { display: block; margin-bottom: 4px; font-weight: 600; }
        input[type="text"] { width: 200px; padding: 6px; font-size: 14px; box-sizing: border-box; }
        input[type="number"] { width: 80px; padding: 6px; font-size: 14px; }
        select { padding: 6px; font-size: 14px; }
        .inline { display: inline-block; margin-right: 12px; }
        .checks { border: 1px solid #ccc; padding: 10px; max-height: 160px; overflow-y: auto; }
        .checks label { display: block; font-weight: normal; margin-bottom: 4px; }
        .checks label input { margin-right: 6px; }
        .preview { border: 1px solid #ccc; padding: 10px; max-height: 120px; overflow: auto; }
        .preview table { border-collapse: collapse; font-size: 12px; }
        .preview th, .preview td { border: 1px solid #ddd; padding: 4px 8px; text-align: left; }
        .preview th { background: #f0f0f0; }
        .buttons { margin-top: 20px; text-align: right; }
        button { padding: 6px 20px; cursor: pointer; font-size: 14px; }
        #execBtn { margin-left: 8px; }
        #execBtn:disabled { cursor: default; opacity: 0.5; }
        #status { margin-top: 12px; font-size: 13px; }
        .info { font-size: 12px; color: #666; margin-top: 2px; }
        .section-title { font-weight: 600; margin-bottom: 4px; }
    </style>
</head>
<body>
    <h3>增强查找</h3>

    <div class="row">
        <label for="lookupRange">查找表区域：</label>
        <input type="text" id="lookupRange" placeholder="例如: Sheet2!A1:D500 或 A1:D500" />
        <div class="info">输入包含表头的完整查找表区域</div>
    </div>

    <div class="row">
        <label for="headerRow">表头行号：</label>
        <input type="number" id="headerRow" value="1" min="1" style="width: 80px;" />
        <button id="readHeadersBtn" style="padding: 4px 12px; font-size: 13px;">读取表头</button>
    </div>

    <div class="row">
        <label for="matchCol">匹配列：</label>
        <select id="matchCol"></select>
        <div class="info">选择查找表中用于匹配的列（即"根据哪一列查找"）</div>
    </div>

    <div class="row">
        <div class="section-title">返回列（勾选需要的列）：</div>
        <button id="selectAllBtn" style="padding: 2px 8px; font-size: 12px;">全选</button>
        <button id="deselectAllBtn" style="padding: 2px 8px; font-size: 12px;">取消全选</button>
        <div class="checks" id="returnCols"></div>
    </div>

    <div class="row">
        <span class="inline">
            <label>匹配模式：</label>
            <label style="font-weight: normal;"><input type="radio" name="matchMode" value="0" checked /> 精确匹配</label>
            <label style="font-weight: normal; margin-left: 8px;"><input type="radio" name="matchMode" value="1" /> 近似匹配</label>
        </span>
    </div>

    <div class="row">
        <span class="inline">
            <label>输出类型：</label>
            <label style="font-weight: normal;"><input type="radio" name="outputType" value="formula" checked /> 公式</label>
            <label style="font-weight: normal; margin-left: 8px;"><input type="radio" name="outputType" value="static" /> 静态值</label>
        </span>
    </div>

    <div class="row" id="previewSection" style="display: none;">
        <div class="section-title">数据预览：</div>
        <div class="preview" id="previewTable"></div>
    </div>

    <div id="status"></div>

    <div class="buttons">
        <button id="cancelBtn">取消</button>
        <button id="execBtn" disabled>执行查找</button>
    </div>

    <script type="text/javascript">
        var MAX_ROWS = 1050000;

        // Inline utility functions (standalone dialog, no module system)
        function getColumnLetter(colIndex) {
            var letter = "";
            var remaining = colIndex;
            do {
                letter = String.fromCharCode(65 + (remaining % 26)) + letter;
                remaining = Math.floor(remaining / 26) - 1;
            } while (remaining >= 0);
            return letter;
        }

        function parseCellRef(ref) {
            var col = 0;
            var i = 0;
            while (i < ref.length) {
                var ch = ref.charCodeAt(i);
                if (ch >= 65 && ch <= 90) { col = col * 26 + (ch - 65); i++; }
                else { break; }
            }
            var row = parseInt(ref.substring(i), 10) || 1;
            return { col: col, row: row };
        }

        function parseRangeAddress(address) {
            var sheet = "";
            var rangePart = address;
            var bangIdx = address.indexOf("!");
            if (bangIdx !== -1) {
                sheet = address.substring(0, bangIdx);
                rangePart = address.substring(bangIdx + 1);
            }
            var parts = rangePart.split(":");
            var start = parseCellRef(parts[0]);
            var end = parseCellRef(parts[1] || parts[0]);
            return {
                sheet: sheet,
                startCol: start.col,
                startRow: start.row,
                endCol: end.col,
                endRow: end.row,
                colCount: end.col - start.col + 1,
                rowCount: end.row - start.row + 1
            };
        }

        function buildColRange(parsed, colIndex) {
            var colLetter = getColumnLetter(parsed.startCol + colIndex);
            var prefix = parsed.sheet ? parsed.sheet + "!" : "";
            return prefix + "$" + colLetter + "$" + parsed.startRow + ":" + "$" + colLetter + "$" + parsed.endRow;
        }

        function refreshColumns() {
            var rangeStr = document.getElementById("lookupRange").value.trim();
            var matchColEl = document.getElementById("matchCol");
            var returnColsEl = document.getElementById("returnCols");
            var execBtn = document.getElementById("execBtn");

            matchColEl.innerHTML = "";
            returnColsEl.innerHTML = "";

            if (!rangeStr) {
                execBtn.disabled = true;
                return;
            }

            var parsed;
            try {
                parsed = parseRangeAddress(rangeStr);
            } catch (e) {
                execBtn.disabled = true;
                return;
            }

            if (parsed.colCount < 1) {
                execBtn.disabled = true;
                return;
            }

            for (var i = 0; i < parsed.colCount; i++) {
                var letter = getColumnLetter(parsed.startCol + i);

                // Match column dropdown
                var opt = document.createElement("option");
                opt.value = i;
                opt.textContent = letter;
                matchColEl.appendChild(opt);

                // Return column checkboxes
                var wrapper = document.createElement("div");
                wrapper.style.display = "inline-block";
                wrapper.style.marginRight = "16px";
                wrapper.style.marginBottom = "4px";

                var cb = document.createElement("input");
                cb.type = "checkbox";
                cb.value = i;
                cb.id = "retCol_" + i;

                var lbl = document.createElement("label");
                lbl.htmlFor = "retCol_" + i;
                lbl.textContent = letter;
                lbl.style.display = "inline";
                lbl.style.marginLeft = "4px";

                wrapper.appendChild(cb);
                wrapper.appendChild(lbl);
                returnColsEl.appendChild(wrapper);
            }

            execBtn.disabled = false;
        }

        function getCheckedReturnCols() {
            var checkboxes = document.querySelectorAll("#returnCols input[type='checkbox']");
            var indices = [];
            for (var i = 0; i < checkboxes.length; i++) {
                if (checkboxes[i].checked) {
                    indices.push(parseInt(checkboxes[i].value, 10));
                }
            }
            return indices;
        }

        function readHeaders() {
            var rangeStr = document.getElementById("lookupRange").value.trim();
            var headerRow = parseInt(document.getElementById("headerRow").value, 10) || 1;

            if (!rangeStr) {
                document.getElementById("status").textContent = "请先输入查找表区域";
                document.getElementById("status").style.color = "red";
                return;
            }

            var parsed;
            try {
                parsed = parseRangeAddress(rangeStr);
            } catch (e) {
                document.getElementById("status").textContent = "区域地址格式无效";
                document.getElementById("status").style.color = "red";
                return;
            }

            var headerRowOffset = headerRow - parsed.startRow;
            if (headerRowOffset < 0 || headerRowOffset >= parsed.rowCount) {
                document.getElementById("status").textContent = "表头行号超出查找表范围";
                document.getElementById("status").style.color = "red";
                return;
            }

            document.getElementById("status").textContent = "正在读取表头...";
            document.getElementById("status").style.color = "green";

            // Try to read headers via Excel.run
            try {
                Excel.run(function (context) {
                    var sheet;
                    if (parsed.sheet) {
                        sheet = context.workbook.worksheets.getItem(parsed.sheet);
                    } else {
                        sheet = context.workbook.worksheets.getActiveWorksheet();
                    }

                    var headerRange = sheet.getRangeByIndexes(
                        headerRow - 1,
                        parsed.startCol,
                        1,
                        parsed.colCount
                    );
                    headerRange.load("values");
                    return context.sync().then(function () {
                        var headers = headerRange.values[0];
                        updateHeaderLabels(headers);

                        // Update preview: read first few rows
                        var previewRows = Math.min(5, parsed.rowCount);
                        var previewRange = sheet.getRangeByIndexes(
                            parsed.startRow - 1,
                            parsed.startCol,
                            previewRows,
                            parsed.colCount
                        );
                        previewRange.load("values");
                        return context.sync().then(function () {
                            showPreview(previewRange.values, headers);
                            document.getElementById("status").textContent = "表头读取成功，共 " + parsed.colCount + " 列";
                            document.getElementById("status").style.color = "green";
                        });
                    });
                }).catch(function (error) {
                    document.getElementById("status").textContent = "读取失败: " + error.message + "（可手动配置）";
                    document.getElementById("status").style.color = "red";
                });
            } catch (e) {
                document.getElementById("status").textContent = "无法连接 Excel（可手动配置列号）";
                document.getElementById("status").style.color = "red";
            }
        }

        function updateHeaderLabels(headers) {
            var labels = document.querySelectorAll("#returnCols label");
            for (var i = 0; i < labels.length && i < headers.length; i++) {
                var headerText = headers[i];
                if (headerText === null || headerText === undefined) headerText = "";
                headerText = String(headerText);
                var letter = getColumnLetter(parseRangeAddress(document.getElementById("lookupRange").value.trim()).startCol + i);
                labels[i].textContent = letter + " - " + (headerText || "(空)");
            }
        }

        function showPreview(values, headers) {
            var section = document.getElementById("previewSection");
            section.style.display = "block";

            var html = "<table><thead><tr>";
            for (var i = 0; i < headers.length; i++) {
                html += "<th>" + (headers[i] || "") + "</th>";
            }
            html += "</tr></thead><tbody>";

            for (var r = 0; r < values.length; r++) {
                html += "<tr>";
                for (var c = 0; c < values[r].length; c++) {
                    var v = values[r][c];
                    if (v === null || v === undefined) v = "";
                    html += "<td>" + String(v) + "</td>";
                }
                html += "</tr>";
            }
            html += "</tbody></table>";

            document.getElementById("previewTable").innerHTML = html;
        }

        Office.onReady(function () {
            document.getElementById("lookupRange").oninput = refreshColumns;
            document.getElementById("headerRow").onchange = refreshColumns;

            document.getElementById("readHeadersBtn").onclick = readHeaders;

            document.getElementById("selectAllBtn").onclick = function () {
                var cbs = document.querySelectorAll("#returnCols input[type='checkbox']");
                for (var i = 0; i < cbs.length; i++) { cbs[i].checked = true; }
            };

            document.getElementById("deselectAllBtn").onclick = function () {
                var cbs = document.querySelectorAll("#returnCols input[type='checkbox']");
                for (var i = 0; i < cbs.length; i++) { cbs[i].checked = false; }
            };

            document.getElementById("cancelBtn").onclick = function () {
                Office.context.ui.messageParent("");
            };

            document.getElementById("execBtn").onclick = function () {
                var statusEl = document.getElementById("status");
                var execBtn = document.getElementById("execBtn");

                var rangeStr = document.getElementById("lookupRange").value.trim();
                if (!rangeStr) {
                    statusEl.textContent = "请输入查找表区域";
                    statusEl.style.color = "red";
                    return;
                }

                var parsed = parseRangeAddress(rangeStr);

                var returnCols = getCheckedReturnCols();
                if (returnCols.length === 0) {
                    statusEl.textContent = "请至少勾选一列作为返回列";
                    statusEl.style.color = "red";
                    return;
                }

                var matchColIndex = parseInt(document.getElementById("matchCol").value, 10);
                var matchMode = parseInt(document.querySelector("input[name='matchMode']:checked").value, 10);
                var outputType = document.querySelector("input[name='outputType']:checked").value;

                if (isNaN(matchColIndex)) {
                    statusEl.textContent = "请选择匹配列";
                    statusEl.style.color = "red";
                    return;
                }

                if (parsed.rowCount > MAX_ROWS) {
                    statusEl.textContent = "数据量过大（" + parsed.rowCount + " 行），最多支持 " + MAX_ROWS + " 行";
                    statusEl.style.color = "red";
                    return;
                }

                execBtn.disabled = true;
                statusEl.textContent = "正在处理...";
                statusEl.style.color = "green";

                var config = {
                    type: "vlookup",
                    outputType: outputType,
                    matchMode: matchMode,
                    lookupRange: rangeStr,
                    parsed: parsed,
                    matchColIndex: matchColIndex,
                    returnColIndices: returnCols,
                    headerRow: parseInt(document.getElementById("headerRow").value, 10) || 1
                };

                if (outputType === "static") {
                    // Read lookup table data in dialog, match, then send results
                    try {
                        Excel.run(function (context) {
                            var sheet;
                            if (parsed.sheet) {
                                sheet = context.workbook.worksheets.getItem(parsed.sheet);
                            } else {
                                sheet = context.workbook.worksheets.getActiveWorksheet();
                            }
                            var dataRange = sheet.getRangeByIndexes(
                                parsed.startRow - 1,
                                parsed.startCol,
                                parsed.rowCount,
                                parsed.colCount
                            );
                            dataRange.load("values");
                            return context.sync().then(function () {
                                var data = dataRange.values;
                                // Remove header row from data
                                var headerOffset = config.headerRow - parsed.startRow;
                                if (headerOffset >= 0 && headerOffset < data.length) {
                                    data.splice(headerOffset, 1);
                                }
                                // Send config + data to parent for matching (parent has data area access)
                                config.lookupTable = data;
                                Office.context.ui.messageParent(JSON.stringify(config));
                            });
                        }).catch(function (error) {
                            statusEl.textContent = "读取查找表失败: " + error.message;
                            statusEl.style.color = "red";
                            execBtn.disabled = false;
                        });
                    } catch (e) {
                        statusEl.textContent = "无法连接 Excel";
                        statusEl.style.color = "red";
                        execBtn.disabled = false;
                    }
                } else {
                    // Formula mode: send config to parent
                    Office.context.ui.messageParent(JSON.stringify(config));
                }
            };
        });
    </script>
</body>
</html>
```

- [ ] **Step 2: 提交**

```bash
git add src/commands/vlookup-dialog.html
git commit -m "feat: add vlookup dialog HTML with configuration UI

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>"
```

---

### Task 4: 注册 enhancedVlookup ribbon 命令

**Files:**
- Modify: `src/commands/commands.html`

- [ ] **Step 1: 在 commands.html 中添加 enhancedVlookup 命令**

在 `importCSV` 的 `Office.actions.associate` 调用之后、`</script>` 之前，添加以下代码：

```javascript
Office.actions.associate("enhancedVlookup", function enhancedVlookup(event) {
    var dialogUrl = window.location.origin + "/vlookup-dialog.html";
    Office.context.ui.displayDialogAsync(
        dialogUrl,
        { height: 60, width: 35, displayInIframe: true },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                Office.context.ui.messageBox("无法打开增强查找对话框");
                event.completed();
                return;
            }
            var dialog = asyncResult.value;
            dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
                dialog.close();
                var config;
                try {
                    config = JSON.parse(arg.message);
                } catch (e) {
                    event.completed();
                    return;
                }
                if (config.type !== "vlookup") {
                    event.completed();
                    return;
                }
                performEnhancedVlookup(event, config);
            });
            dialog.addEventHandler(Office.EventType.DialogEventReceived, function () {
                event.completed();
            });
        }
    );
});
```

在 `</script>` 之前（enhancedVlookup 注册之后）添加核心执行逻辑：

```javascript
function performEnhancedVlookup(event, config) {
    Excel.run(function (context) {
        var selectedRange = context.workbook.getSelectedRange();
        selectedRange.load(["address", "columnCount", "rowCount", "columnIndex", "rowIndex"]);
        return context.sync().then(function () {
            if (selectedRange.rowCount === 0) {
                Office.context.ui.messageBox("没有数据");
                throw new Error("cancel");
            }

            if (selectedRange.rowCount > MAX_ROWS) {
                Office.context.ui.messageBox(
                    "数据量过大（" + selectedRange.rowCount + " 行），单次最多支持 " + MAX_ROWS + " 行。"
                );
                throw new Error("cancel");
            }

            var worksheet = context.workbook.worksheets.getActiveWorksheet();
            var dataColIndex = selectedRange.columnIndex;
            var dataRowIndex = selectedRange.rowIndex;
            var dataRowCount = selectedRange.rowCount;
            var dataStartCol = dataColIndex;
            var dataStartRow = dataRowIndex;

            // Build absolute range references for lookup table
            var lookupColRange = buildColRangeInline(config.parsed, config.matchColIndex);
            var returnColRanges = [];
            for (var i = 0; i < config.returnColIndices.length; i++) {
                returnColRanges.push(buildColRangeInline(config.parsed, config.returnColIndices[i]));
            }

            if (config.outputType === "formula") {
                // Write INDEX/MATCH formulas to new columns right of the data
                var dataColLetter = getColumnLetter(dataStartCol);
                var returnColCount = config.returnColIndices.length;

                // Insert new columns for each return column
                var insertPos = dataStartCol + selectedRange.columnCount;
                for (var c = 0; c < returnColCount; c++) {
                    worksheet
                        .getRange(getColumnLetter(insertPos) + ":" + getColumnLetter(insertPos))
                        .insert(Excel.InsertShiftDirection.right);
                }

                return context.sync().then(function () {
                    for (var r = 0; r < returnColCount; r++) {
                        var formulaCol = getColumnLetter(insertPos + r);
                        var lookupCellRef = dataColLetter + dataStartRow;
                        var formula = "=INDEX(" + returnColRanges[r] + ", MATCH(" + lookupCellRef + ", " + lookupColRange + ", " + config.matchMode + "))";

                        var startCell = worksheet.getRange(formulaCol + dataStartRow);
                        startCell.formulas = [[formula]];

                        if (dataRowCount > 1) {
                            var fillRange = worksheet.getRange(
                                formulaCol + dataStartRow + ":" + formulaCol + (dataStartRow + dataRowCount - 1)
                            );
                            startCell.autoFill(fillRange, Excel.AutoFillType.fillDefault);
                        }
                    }
                    return context.sync().then(function () {
                        Office.context.ui.messageBox(
                            "完成! 已在 " + returnColCount + " 列写入 " + dataRowCount + " 行 INDEX/MATCH 公式"
                        );
                    });
                });
            } else {
                // Static mode: read data area values, do local matching, write results
                var dataRange = selectedRange.getUsedRange();
                dataRange.load("values");
                return context.sync().then(function () {
                    var dataValues = dataRange.values;
                    var lookupValues = [];
                    for (var r = 0; r < dataValues.length; r++) {
                        lookupValues.push(dataValues[r][0]);
                    }

                    var results = staticLookupInline(
                        lookupValues,
                        config.lookupTable,
                        config.matchColIndex,
                        config.returnColIndices,
                        config.matchMode
                    );

                    // Write results to new columns right of data
                    var insertPos2 = dataStartCol + selectedRange.columnCount;
                    var returnColCount2 = config.returnColIndices.length;

                    for (var c2 = 0; c2 < returnColCount2; c2++) {
                        worksheet
                            .getRange(getColumnLetter(insertPos2) + ":" + getColumnLetter(insertPos2))
                            .insert(Excel.InsertShiftDirection.right);
                    }

                    return context.sync().then(function () {
                        var targetRange = worksheet.getRange(
                            getColumnLetter(insertPos2) + dataStartRow + ":" +
                            getColumnLetter(insertPos2 + returnColCount2 - 1) + (dataStartRow + results.length - 1)
                        );
                        // Transpose results from [row][col] to [col][row] for range.values
                        var valuesToWrite = [];
                        for (var rr = 0; rr < results.length; rr++) {
                            valuesToWrite.push(results[rr]);
                        }
                        targetRange.values = valuesToWrite;
                        return context.sync().then(function () {
                            Office.context.ui.messageBox(
                                "完成! 已写入 " + results.length + " 行 × " + returnColCount2 + " 列静态值"
                            );
                        });
                    });
                });
            }
        });
    }).then(function () {
        event.completed();
    }).catch(function (error) {
        if (error.message !== "cancel") {
            Office.context.ui.messageBox("操作失败: " + error.message);
        }
        event.completed();
    });
}

// Inline copies of utility functions for commands context (no module system)
function buildColRangeInline(parsed, colIndex) {
    var colLetter = getColumnLetter(parsed.startCol + colIndex);
    var prefix = parsed.sheet ? parsed.sheet + "!" : "";
    return prefix + "$" + colLetter + "$" + parsed.startRow + ":" + "$" + colLetter + "$" + parsed.endRow;
}

function staticLookupInline(lookupValues, lookupTable, matchColIndex, returnColIndices, matchMode) {
    var results = [];
    var index = {};
    if (matchMode === 0) {
        for (var r = 0; r < lookupTable.length; r++) {
            var key = lookupTable[r][matchColIndex];
            if (key === null || key === undefined) { key = ""; }
            key = String(key);
            index[key] = r;
        }
    }
    for (var i = 0; i < lookupValues.length; i++) {
        var val = lookupValues[i];
        if (val === null || val === undefined) { val = ""; }
        var valStr = String(val);
        var row = [];
        if (matchMode === 0) {
            var matchedRow = index[valStr];
            if (matchedRow !== undefined) {
                for (var j = 0; j < returnColIndices.length; j++) {
                    row.push(lookupTable[matchedRow][returnColIndices[j]]);
                }
            } else {
                for (var k = 0; k < returnColIndices.length; k++) { row.push("#N/A"); }
            }
        } else {
            var bestRow = -1;
            for (var m = 0; m < lookupTable.length; m++) {
                var tableVal = lookupTable[m][matchColIndex];
                if (tableVal === null || tableVal === undefined) continue;
                if (Number(tableVal) <= Number(val)) { bestRow = m; }
                else { break; }
            }
            if (bestRow >= 0) {
                for (var n = 0; n < returnColIndices.length; n++) {
                    row.push(lookupTable[bestRow][returnColIndices[n]]);
                }
            } else {
                for (var p = 0; p < returnColIndices.length; p++) { row.push("#N/A"); }
            }
        }
        results.push(row);
    }
    return results;
}
```

> **注意**：`enhancedVlookup` 的 `Office.actions.associate` 调用需要添加到现有 `importCSV` 的 `associate` 调用之后。`performEnhancedVlookup` 和辅助函数需要添加到 `</script>` 之前。

- [ ] **Step 2: 提交**

```bash
git add src/commands/commands.html
git commit -m "feat: register enhancedVlookup command in ribbon

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>"
```

---

### Task 5: Taskpane 增加增强查找按钮（调试用）

**Files:**
- Modify: `src/taskpane/taskpane.html`

- [ ] **Step 1: 在 taskpane.html 中添加按钮**

在 CSV 导入区域之后、`</main>` 之前，添加以下代码：

找到：
```html
        <p id="importStatus" style="margin-top: 20px; color: green;"></p>
    </main>
```

替换为：
```html
        <p id="importStatus" style="margin-top: 20px; color: green;"></p>
        <hr style="width:100%; margin: 30px 0;" />
        <h2 class="ms-font-l">增强查找</h2>
        <p class="ms-font-m">用 INDEX/MATCH 一次完成多列查找，匹配列可在任意位置</p>
        <div role="button" id="vlookupBtn" class="ms-welcome__action ms-Button ms-Button--hero ms-font-xl">
            <span class="ms-Button-label">增强查找</span>
        </div>
        <p id="vlookupStatus" style="margin-top: 20px; color: green;"></p>
    </main>
```

- [ ] **Step 2: 提交**

```bash
git add src/taskpane/taskpane.html
git commit -m "feat: add enhanced vlookup button to taskpane for debugging

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>"
```

---

### Task 6: Taskpane.js 添加增强查找处理逻辑

**Files:**
- Modify: `src/taskpane/taskpane.js`

- [ ] **Step 1: 在 taskpane.js 中添加处理逻辑**

在文件顶部的 `require` 区域，添加 vlookup-utils 引入：

找到：
```javascript
var { getColumnLetter, escapeFormulaText, buildConcatFormula } = require("../utils/concat-utils");
var { parseCSV } = require("../utils/csv-utils");
```

替换为：
```javascript
var { getColumnLetter, escapeFormulaText, buildConcatFormula } = require("../utils/concat-utils");
var { parseCSV } = require("../utils/csv-utils");
var { buildColRange, buildIndexMatchFormula, staticLookup, parseRangeAddress } = require("../utils/vlookup-utils");
```

在 `Office.onReady` 的回调中，注册 vlookup 按钮的点击事件。找到：

```javascript
    document.getElementById("csvFileInput").onchange = handleCsvFile;
  }
});
```

替换为：
```javascript
    document.getElementById("csvFileInput").onchange = handleCsvFile;
    document.getElementById("vlookupBtn").onclick = runVlookup;
  }
});
```

在 `runConcat` 函数之后（`handleCsvFile` 函数之前）添加 `runVlookup` 和 `executeVlookupFromConfig` 函数：

```javascript
async function runVlookup() {
  if (isProcessing) return;

  var statusEl = document.getElementById("vlookupStatus");
  statusEl.textContent = "正在打开配置窗口...";
  statusEl.style.color = "green";

  Office.context.ui.displayDialogAsync(
    window.location.origin + "/vlookup-dialog.html",
    { height: 60, width: 35, displayInIframe: true },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        statusEl.textContent = "错误: 无法打开配置窗口";
        statusEl.style.color = "red";
        return;
      }
      var dialog = asyncResult.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
        dialog.close();
        var config;
        try {
          config = JSON.parse(arg.message);
        } catch (e) {
          statusEl.textContent = "";
          return;
        }
        if (config.type !== "vlookup") {
          statusEl.textContent = "";
          return;
        }
        executeVlookupFromConfig(config, statusEl);
      });
      dialog.addEventHandler(Office.EventType.DialogEventReceived, function () {
        statusEl.textContent = "已取消";
        statusEl.style.color = "gray";
      });
    }
  );
}

async function executeVlookupFromConfig(config, statusEl) {
  isProcessing = true;
  statusEl.textContent = "处理中...";
  statusEl.style.color = "green";

  try {
    await Excel.run(async (context) => {
      var selectedRange = context.workbook.getSelectedRange();
      selectedRange.load(["columnCount", "rowCount", "columnIndex", "rowIndex"]);
      await context.sync();

      var worksheet = context.workbook.worksheets.getActiveWorksheet();
      var dataStartCol = selectedRange.columnIndex;
      var dataStartRow = selectedRange.rowIndex;
      var dataRowCount = selectedRange.rowCount;
      var dataColLetter = getColumnLetter(dataStartCol);

      var lookupColRange = buildColRange(config.parsed, config.matchColIndex);
      var returnColRanges = [];
      for (var i = 0; i < config.returnColIndices.length; i++) {
        returnColRanges.push(buildColRange(config.parsed, config.returnColIndices[i]));
      }

      if (config.outputType === "formula") {
        var insertPos = dataStartCol + selectedRange.columnCount;
        var returnColCount = config.returnColIndices.length;

        for (var c = 0; c < returnColCount; c++) {
          worksheet
            .getRange(getColumnLetter(insertPos) + ":" + getColumnLetter(insertPos))
            .insert(Excel.InsertShiftDirection.right);
        }
        await context.sync();

        for (var r = 0; r < returnColCount; r++) {
          var formulaCol = getColumnLetter(insertPos + r);
          var lookupCellRef = dataColLetter + dataStartRow;
          var formula = buildIndexMatchFormula(lookupCellRef, lookupColRange, returnColRanges[r], config.matchMode);

          var startCell = worksheet.getRange(formulaCol + dataStartRow);
          startCell.formulas = [[formula]];

          if (dataRowCount > 1) {
            var fillRange = worksheet.getRange(
              formulaCol + dataStartRow + ":" + formulaCol + (dataStartRow + dataRowCount - 1)
            );
            startCell.autoFill(fillRange, Excel.AutoFillType.fillDefault);
            await context.sync();
          }
        }

        statusEl.textContent =
          "完成! 已在 " + returnColCount + " 列写入 " + dataRowCount + " 行 INDEX/MATCH 公式";
      } else {
        // Static mode
        var dataValues = [];
        var dataRange = selectedRange.getUsedRange();
        dataRange.load("values");
        await context.sync();
        dataValues = dataRange.values;

        var lookupValues = [];
        for (var j = 0; j < dataValues.length; j++) {
          lookupValues.push(dataValues[j][0]);
        }

        var results = staticLookup(
          lookupValues,
          config.lookupTable,
          config.matchColIndex,
          config.returnColIndices,
          config.matchMode
        );

        var insertPos2 = dataStartCol + selectedRange.columnCount;
        var returnColCount2 = config.returnColIndices.length;

        for (var c2 = 0; c2 < returnColCount2; c2++) {
          worksheet
            .getRange(getColumnLetter(insertPos2) + ":" + getColumnLetter(insertPos2))
            .insert(Excel.InsertShiftDirection.right);
        }
        await context.sync();

        var targetRange = worksheet.getRange(
          getColumnLetter(insertPos2) + dataStartRow + ":" +
          getColumnLetter(insertPos2 + returnColCount2 - 1) + (dataStartRow + results.length - 1)
        );
        targetRange.values = results;
        await context.sync();

        statusEl.textContent =
          "完成! 已写入 " + results.length + " 行 × " + returnColCount2 + " 列静态值";
      }
    });
  } catch (error) {
    statusEl.textContent = "错误: " + error.message;
    statusEl.style.color = "red";
  }

  isProcessing = false;
}
```

- [ ] **Step 2: 提交**

```bash
git add src/taskpane/taskpane.js
git commit -m "feat: add enhanced vlookup handler to taskpane

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>"
```

---

### Task 7: manifest.xml 添加 ribbon 按钮

**Files:**
- Modify: `manifest.xml`

- [ ] **Step 1: 在 manifest.xml 中添加按钮和资源字符串**

在 `ImportCsvButton` 控件之后、`</Group>` 之前，添加新的按钮：

找到：
```xml
                <Control xsi:type="Button" id="ImportCsvButton">
                  ...
                </Control>
              </Group>
```

在 `</Group>` 之前添加：

```xml
                <Control xsi:type="Button" id="EnhancedVlookupButton">
                  <Label resid="EnhancedVlookupButton.Label"/>
                  <Supertip>
                    <Title resid="EnhancedVlookupButton.Label"/>
                    <Description resid="EnhancedVlookupButton.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16"/>
                    <bt:Image size="32" resid="Icon.32x32"/>
                    <bt:Image size="80" resid="Icon.80x80"/>
                  </Icon>
                  <Action xsi:type="ExecuteFunction">
                    <FunctionName>enhancedVlookup</FunctionName>
                  </Action>
                </Control>
```

在 ShortStrings 区域（`<bt:ShortStrings>` 内）添加：

```xml
        <bt:String id="EnhancedVlookupButton.Label" DefaultValue="增强查找"/>
```

在 LongStrings 区域（`<bt:LongStrings>` 内）添加：

```xml
        <bt:String id="EnhancedVlookupButton.Tooltip" DefaultValue="增强版查找功能：一次配置完成多列 INDEX/MATCH 查找，匹配列可在任意位置"/>
```

- [ ] **Step 2: 验证 manifest**

```bash
npm run validate
```
Expected: Validation passed

- [ ] **Step 3: 提交**

```bash
git add manifest.xml
git commit -m "feat: add enhanced vlookup ribbon button to manifest

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>"
```

---

### Task 8: webpack.config.js 添加弹窗复制

**Files:**
- Modify: `webpack.config.js`

- [ ] **Step 1: 添加 vlookup-dialog.html 到 CopyWebpackPlugin**

在 `CopyWebpackPlugin` 的 patterns 数组中，在 csv-import-dialog.html 的条目之后添加：

找到：
```javascript
          {
            from: "src/commands/csv-import-dialog.html",
            to: "csv-import-dialog.html",
          },
```

之后添加：
```javascript
          {
            from: "src/commands/vlookup-dialog.html",
            to: "vlookup-dialog.html",
          },
```

- [ ] **Step 2: 提交**

```bash
git add webpack.config.js
git commit -m "feat: add vlookup-dialog.html to webpack copy config

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>"
```

---

### Task 9: 构建验证 & 测试

**Files:**
- None (验证步骤)

- [ ] **Step 1: 运行单元测试**

```bash
npx jest
```
Expected: all tests PASS (including new vlookup-utils tests)

- [ ] **Step 2: 运行构建**

```bash
npm run build
```
Expected: build succeeds, output contains `vlookup-dialog.html`

- [ ] **Step 3: 检查构建输出**

```bash
ls dist/vlookup-dialog.html
```
Expected: file exists

- [ ] **Step 4: 提交（如有 lint/format 修复）**

```bash
git add -A
git commit -m "chore: final build verification and adjustments

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>"
```
