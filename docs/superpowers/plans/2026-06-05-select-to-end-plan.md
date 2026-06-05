# 向下选区 — 实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 在 Excel 功能区"数据处理"分组下新增"向下选区"按钮，点击后从当前选中单元格向下选取到该列最后一个有数据的单元格。

**Architecture:** 纯 ExecuteFunction 模式，按钮在 manifest.xml 配置，函数实现在 commands.html 内联脚本中。复用已有的 `getColumnLetter` 和 `MAX_ROWS` 常量，不引入新文件。

**Tech Stack:** Office JavaScript API, vanilla JavaScript (ES5 兼容)

---

### Task 1: 在 manifest.xml 添加按钮控件

**Files:**
- Modify: `manifest.xml`

- [ ] **Step 1: 在 DataProcessingGroup 末尾添加 SelectToEndButton 控件**

找到 `<Group id="DataProcessingGroup">` 内的最后一个 `</Control>`（SplitSheetButton 之后），在后面追加以下内容：

```xml
                <Control xsi:type="Button" id="SelectToEndButton">
                  <Label resid="SelectToEndButton.Label"/>
                  <Supertip>
                    <Title resid="SelectToEndButton.Label"/>
                    <Description resid="SelectToEndButton.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="SelectToEndIcon.16x16"/>
                    <bt:Image size="32" resid="SelectToEndIcon.32x32"/>
                    <bt:Image size="80" resid="SelectToEndIcon.80x80"/>
                  </Icon>
                  <Action xsi:type="ExecuteFunction">
                    <FunctionName>selectToEnd</FunctionName>
                  </Action>
                </Control>
```

> 注意：`ExecuteFunction` 模式不需要 `<TaskpaneId>` 和 `<SourceLocation>`，只需要 `<FunctionName>`。

- [ ] **Step 2: 验证 manifest.xml 完整性**

运行：`npm run validate`
期望：通过，无 XML 语法错误。

- [ ] **Step 3: 提交**

```bash
git add manifest.xml
git commit -m "feat: add SelectToEndButton control to manifest.xml"
```

---

### Task 2: 在 manifest.xml 添加资源引用

**Files:**
- Modify: `manifest.xml`

> 按钮控件通过 `resid` 引用资源，这些 resId 必须在 `<Resources>` 块中有对应定义。

- [ ] **Step 1: 在 `<bt:ShortStrings>` 中添加标签文字**

找到 `<bt:ShortStrings>` 块末尾（如 `SplitSheetButton.Label` 之后），添加：

```xml
        <bt:String id="SelectToEndButton.Label" DefaultValue="向下选区"/>
```

- [ ] **Step 2: 在 `<bt:LongStrings>` 中添加提示文字**

找到 `<bt:LongStrings>` 块末尾（如 `SplitSheetButton.Tooltip` 之后），添加：

```xml
        <bt:String id="SelectToEndButton.Tooltip" DefaultValue="从选中单元格向下选到该列最后一个有数据的单元格"/>
```

- [ ] **Step 3: 在 `<bt:Images>` 中添加图标引用（复用 tools 图标）**

找到 `<bt:Images>` 块末尾（最后一个 `</bt:Image>` 之后、`</bt:Images>` 之前），添加：

```xml
        <bt:Image id="SelectToEndIcon.16x16" DefaultValue="https://localhost:3000/assets/tools-16.png"/>
        <bt:Image id="SelectToEndIcon.32x32" DefaultValue="https://localhost:3000/assets/tools-32.png"/>
        <bt:Image id="SelectToEndIcon.80x80" DefaultValue="https://localhost:3000/assets/tools-80.png"/>
```

- [ ] **Step 4: 验证 manifest.xml**

运行：`npm run validate`
期望：通过。

- [ ] **Step 5: 提交**

```bash
git add manifest.xml
git commit -m "feat: add resource references for SelectToEndButton"
```

---

### Task 3: 在 commands.html 实现 selectToEnd 函数

**Files:**
- Modify: `src/commands/commands.html`

- [ ] **Step 1: 在 showFillSeriesPanel 的 associate 之后添加 selectToEnd 函数**

找到 `Office.actions.associate("showFillSeriesPanel", ...)` 块结束后的 `</script>` 标签之前，插入：

```javascript
        Office.actions.associate("selectToEnd", function selectToEnd(event) {
            Excel.run(function (context) {
                var range = context.workbook.getSelectedRange();
                range.load(["columnIndex", "columnCount", "rowIndex"]);
                return context.sync().then(function () {
                    var colIndex = range.columnIndex;
                    var colCount = range.columnCount;
                    var startRow = range.rowIndex + 1; // 0-based → 1-based
                    var worksheet = context.workbook.worksheets.getActiveWorksheet();

                    // 对每一列加载数据
                    var columnData = [];
                    for (var i = 0; i < colCount; i++) {
                        var colLetter = getColumnLetter(colIndex + i);
                        var dataRange = worksheet.getRange(
                            colLetter + startRow + ":" + colLetter + MAX_ROWS
                        );
                        dataRange.load("values");
                        columnData.push({ data: dataRange });
                    }

                    return context.sync().then(function () {
                        // 从下往上扫描，找每列最后一个有数据的行
                        var maxDataRow = 0;
                        for (var j = 0; j < columnData.length; j++) {
                            var values = columnData[j].data.values;
                            var lastRow = 0;
                            for (var r = values.length - 1; r >= 0; r--) {
                                var v = values[r][0];
                                if (v !== null && v !== undefined && v !== "") {
                                    lastRow = startRow + r;
                                    break;
                                }
                            }
                            if (lastRow > maxDataRow) {
                                maxDataRow = lastRow;
                            }
                        }

                        if (maxDataRow === 0) {
                            Office.context.ui.messageBox("该列没有数据");
                            return;
                        }

                        var startColLetter = getColumnLetter(colIndex);
                        var endColLetter = getColumnLetter(colIndex + colCount - 1);
                        var targetRange = worksheet.getRange(
                            startColLetter + startRow + ":" + endColLetter + maxDataRow
                        );
                        targetRange.select();
                        return context.sync();
                    });
                });
            }).catch(function (error) {
                Office.context.ui.messageBox("操作失败: " + error.message);
            }).then(function () {
                event.completed();
            });
        });
```

- [ ] **Step 2: 运行 lint 检查**

运行：`npm run lint`
期望：通过。

- [ ] **Step 3: 提交**

```bash
git add src/commands/commands.html
git commit -m "feat: implement selectToEnd function"
```

---

### Task 4: 手动验证

**不需要代码变更** — 在 Excel 中加载插件后测试。

- [ ] **Step 1: 启动开发服务器和 Excel**

运行：`npm run start`
打开 Excel，确保插件加载。

- [ ] **Step 2: 测试基础场景**

1. 在 A1 输入 "数据1"，A2 空，A3 输入 "数据2"
2. 选中 A1
3. 点击"向下选区"按钮
4. **期望**：选区变为 A1:A3

- [ ] **Step 3: 测试多列场景**

1. 在 A1:A3 和 B1:B5 中填入不同长度的数据
2. 选中 A1:B1
3. 点击"向下选区"按钮
4. **期望**：选区覆盖两列中较远的那一行（如 A1:B5）

- [ ] **Step 4: 测试空列场景**

1. 选中一个完全空白的列中的单元格
2. 点击"向下选区"按钮
3. **期望**：弹出"该列没有数据"提示

- [ ] **Step 5: 测试边界场景**

1. 选中一个有数据的单元格，该单元格已位于列的最后一行
2. 点击"向下选区"按钮
3. **期望**：选区不变（仍是这一行）
