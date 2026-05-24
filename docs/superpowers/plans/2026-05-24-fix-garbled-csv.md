# 修复 CSV 中文乱码 — 实施计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 在"我的工具"选项卡新增"修复乱码"按钮，一键修复 Excel 误用 GBK 打开 UTF-8 CSV 导致的中文乱码。

**Architecture:** 使用 `iconv-lite` 纯 JS 库实现编码逆转（乱码文本 → GBK 编码 → UTF-8 解码）。新增 `encoding-utils.js` 工具模块，在 `commands.js` 中将工具函数暴露到全局作用域，`commands.html` 内联脚本注册命令并执行全表修复。

**Tech Stack:** iconv-lite, Office JavaScript API, Jest, webpack

---

## 文件结构

| 文件 | 操作 | 职责 |
|------|------|------|
| `package.json` | 修改 | 新增 `iconv-lite` 依赖 |
| `src/utils/encoding-utils.test.js` | 新增 | `fixGarbledText` 单元测试（TDD: 先写测试） |
| `src/utils/encoding-utils.js` | 新增 | 编码逆转工具函数，导出 `fixGarbledText` |
| `src/commands/commands.js` | 修改 | 导入 encoding-utils，挂载到 `window` |
| `src/commands/commands.html` | 修改 | 注册 `fixGarbledText` 命令，实现工作表遍历修复逻辑 |
| `manifest.xml` | 修改 | 在"我的工具"选项卡新增"修复乱码"按钮 |

---

### Task 1: 安装 iconv-lite 依赖

- [ ] **Step 1: 安装 iconv-lite**

```bash
npm install iconv-lite --save
```

Expected: `package.json` 中 `dependencies` 新增 `"iconv-lite": "^0.6.3"`

---

### Task 2: 编写 encoding-utils 单元测试（TDD — RED）

**Files:**
- Create: `src/utils/encoding-utils.test.js`

- [ ] **Step 1: 创建测试文件**

```javascript
var iconv = require("iconv-lite");
var { fixGarbledText } = require("./encoding-utils");

describe("fixGarbledText", function () {
  it("passes through plain ASCII unchanged", function () {
    expect(fixGarbledText("hello world")).toBe("hello world");
  });

  it("passes through numbers unchanged", function () {
    expect(fixGarbledText(123)).toBe(123);
    expect(fixGarbledText(0)).toBe(0);
  });

  it("passes through empty string unchanged", function () {
    expect(fixGarbledText("")).toBe("");
  });

  it("fixes garbled Chinese text (UTF-8 bytes misinterpreted as GBK)", function () {
    var original = "中文测试";
    // 模拟 Excel 的行为：UTF-8 字节 → GBK 解码 → 乱码
    var utf8Bytes = iconv.encode(original, "utf-8");
    var garbled = iconv.decode(utf8Bytes, "gbk");
    // 验证乱码确实和原文不同（排除假阳性）
    expect(garbled).not.toBe(original);
    // 我们的函数应该修复它
    expect(fixGarbledText(garbled)).toBe(original);
  });

  it("fixes garbled text with mixed Chinese and ASCII", function () {
    var original = "编号ABC_测试数据123";
    var utf8Bytes = iconv.encode(original, "utf-8");
    var garbled = iconv.decode(utf8Bytes, "gbk");
    expect(garbled).not.toBe(original);
    expect(fixGarbledText(garbled)).toBe(original);
  });

  it("does not throw on already-correct Chinese", function () {
    // 已正确的中文经 GBK→UTF-8 逆转后可能变成其他文字，
    // 但函数不应崩溃，返回字符串即可
    var result = fixGarbledText("正确中文");
    expect(typeof result).toBe("string");
    expect(result.length).toBeGreaterThan(0);
  });
});
```

- [ ] **Step 2: 运行测试，确认失败（RED）**

```bash
npm test -- --testPathPattern="encoding-utils"
```

Expected: FAIL — `Cannot find module './encoding-utils'`

- [ ] **Step 3: 创建空的 encoding-utils.js 占位文件，确认测试仍失败**

```bash
echo "module.exports = {};" > src/utils/encoding-utils.js
npm test -- --testPathPattern="encoding-utils"
```

Expected: FAIL — `fixGarbledText is not a function`

---

### Task 3: 实现 encoding-utils.js（TDD — GREEN）

**Files:**
- Modify: `src/utils/encoding-utils.js`（覆盖占位文件）

- [ ] **Step 1: 实现 fixGarbledText**

将 `src/utils/encoding-utils.js` 内容替换为：

```javascript
var iconv = require("iconv-lite");

function fixGarbledText(value) {
  if (typeof value !== "string" || value === "") {
    return value;
  }
  try {
    var bytes = iconv.encode(value, "gbk");
    var fixed = iconv.decode(bytes, "utf-8");
    return fixed;
  } catch (e) {
    return value;
  }
}

module.exports = { fixGarbledText };
```

- [ ] **Step 2: 运行测试，确认通过（GREEN）**

```bash
npm test -- --testPathPattern="encoding-utils"
```

Expected: 全部 6 个测试 PASS

- [ ] **Step 3: 提交**

```bash
git add src/utils/encoding-utils.test.js src/utils/encoding-utils.js
git commit -m "feat: add encoding-utils with fixGarbledText for GBK/UTF-8 reversal"
```

---

### Task 4: 在 commands.js 中暴露全局函数

**Files:**
- Modify: `src/commands/commands.js`

`commands.html` 的内联脚本不经过 webpack 编译，需要通过全局作用域桥接 webpack 打包的模块。

- [ ] **Step 1: 修改 commands.js**

将文件内容替换为：

```javascript
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office, Excel, window */

var { fixGarbledText } = require("../utils/encoding-utils");

// Expose to global scope for inline script in commands.html
window.fixGarbledText = fixGarbledText;
```

- [ ] **Step 2: 提交**

```bash
git add src/commands/commands.js
git commit -m "feat: expose fixGarbledText to global scope in commands.js"
```

---

### Task 5: 在 commands.html 中注册 fixGarbledText 命令

**Files:**
- Modify: `src/commands/commands.html`

- [ ] **Step 1: 添加命令注册代码**

在 `commands.html` 中，找到 `Office.actions.associate("concatenateColumns", ...)` 的闭合 `});` 行，在其之后、`</script>` 之前插入以下代码：

```javascript
        Office.actions.associate("fixGarbledText", function fixGarbledText(event) {
            Excel.run(function (context) {
                var sheet = context.workbook.worksheets.getActiveWorksheet();
                var usedRange = sheet.getUsedRange();
                usedRange.load(["rowCount", "columnCount", "values"]);
                return context.sync().then(function () {
                    var rowCount = usedRange.rowCount;
                    var colCount = usedRange.columnCount;

                    if (rowCount === 0 || colCount === 0) {
                        Office.context.ui.messageBox("没有数据");
                        throw new Error("cancel");
                    }

                    if (rowCount > 1050000) {
                        Office.context.ui.messageBox(
                            "数据量过大（" + rowCount + " 行），单次最多支持 1050000 行。"
                        );
                        throw new Error("cancel");
                    }

                    var values = usedRange.values;
                    var fixedCount = 0;
                    var totalCells = rowCount * colCount;

                    for (var r = 0; r < rowCount; r++) {
                        for (var c = 0; c < colCount; c++) {
                            var cellValue = values[r][c];
                            if (typeof cellValue === "string" && cellValue !== "") {
                                var fixed = window.fixGarbledText(cellValue);
                                if (fixed !== cellValue) {
                                    values[r][c] = fixed;
                                    fixedCount++;
                                }
                            }
                        }
                    }

                    // Write back all values at once
                    usedRange.values = values;
                    return context.sync().then(function () {
                        Office.context.ui.messageBox(
                            "修复完成，共处理 " + totalCells + " 个单元格，修复 " + fixedCount + " 个"
                        );
                    });
                });
            }).then(function () {
                event.completed();
            }).catch(function (error) {
                if (error.message !== "cancel") {
                    Office.context.ui.messageBox("操作失败: " + error.message);
                }
                event.completed();
            });
        });
```

- [ ] **Step 2: 验证构建**

```bash
npm run build:dev
```

Expected: 构建成功，无错误

- [ ] **Step 3: 提交**

```bash
git add src/commands/commands.html
git commit -m "feat: add fixGarbledText command in commands.html"
```

---

### Task 6: 在 manifest.xml 中新增"修复乱码"按钮

**Files:**
- Modify: `manifest.xml`

- [ ] **Step 1: 在 ConcatenateButton `</Control>` 之后插入新按钮**

```xml
                <Control xsi:type="Button" id="FixGarbledButton">
                  <Label resid="FixGarbledButton.Label"/>
                  <Supertip>
                    <Title resid="FixGarbledButton.Label"/>
                    <Description resid="FixGarbledButton.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16"/>
                    <bt:Image size="32" resid="Icon.32x32"/>
                    <bt:Image size="80" resid="Icon.80x80"/>
                  </Icon>
                  <Action xsi:type="ExecuteFunction">
                    <FunctionName>fixGarbledText</FunctionName>
                  </Action>
                </Control>
```

- [ ] **Step 2: 在 `<bt:ShortStrings>` 内添加**

```xml
        <bt:String id="FixGarbledButton.Label" DefaultValue="修复乱码"/>
```

- [ ] **Step 3: 在 `<bt:LongStrings>` 内添加**

```xml
        <bt:String id="FixGarbledButton.Tooltip" DefaultValue="修复因 UTF-8 编码被误读为 GBK 导致的中文乱码"/>
```

- [ ] **Step 4: 验证 manifest**

```bash
npm run validate
```

Expected: 验证通过

- [ ] **Step 5: 提交**

```bash
git add manifest.xml
git commit -m "feat: add FixGarbledButton to custom ribbon tab"
```

---

### Task 7: 端到端验证

- [ ] **Step 1: 构建并启动**

```bash
npm run build:dev && npm run start
```

Expected: Excel 启动，"我的工具"选项卡出现"修复乱码"按钮

- [ ] **Step 2: 手动测试**

1. 创建一个 UTF-8 无 BOM 的中文 CSV 文件（如 `测试数据.csv`，内容含中文列）
2. 用 Excel 直接打开该 CSV（中文显示乱码）
3. 点击"我的工具" → "修复乱码"
4. 验证：中文乱码被修复为正确文字，弹出消息框显示修复数量
5. 对纯英文/数字的工作表点击按钮，验证：消息框显示"修复 0 个"

- [ ] **Step 3: 提交（如有调整）**

```bash
git add -A
git commit -m "chore: final adjustments after e2e testing"
```

---

## 提交历史总览

```
feat: add encoding-utils with fixGarbledText for GBK/UTF-8 reversal
feat: expose fixGarbledText to global scope in commands.js
feat: add fixGarbledText command in commands.html
feat: add FixGarbledButton to custom ribbon tab
chore: final adjustments after e2e testing
```
