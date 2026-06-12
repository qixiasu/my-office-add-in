# 向右选区 — 实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 在 Excel 功能区"数据处理"分组下新增"向右选区"按钮，点击后从当前选中单元格向右选取到该行最后一个有数据的单元格。

**Architecture:** 纯 ExecuteFunction 模式，完全镜像已有的 selectToEnd（向下选区）实现。按钮在 manifest.xml 配置，函数实现在 commands.html 内联脚本中。复用已有的 `getColumnLetter` 和 `MAX_ROWS` 常量，不引入新文件。

**Tech Stack:** Office JavaScript API, vanilla JavaScript (ES5 兼容)

---

### Task 0: 生成图标文件

**Files:**
- Create: `assets/select-to-right-16.png`
- Create: `assets/select-to-right-32.png`
- Create: `assets/select-to-right-80.png`
- Source: `pictures\向右选区.png`

- [ ] **Step 1: 生成 80px 图标**

直接复制源文件到 assets 目录：

```bash
cp "f:/projects/My Office Add-in/pictures/向右选区.png" "f:/projects/My Office Add-in/assets/select-to-right-80.png"
```

- [ ] **Step 2: 生成 32px 图标**

复制源文件（源文件较小，直接复用）：

```bash
cp "f:/projects/My Office Add-in/pictures/向右选区.png" "f:/projects/My Office Add-in/assets/select-to-right-32.png"
```

- [ ] **Step 3: 生成 16px 图标**

复制源文件：

```bash
cp "f:/projects/My Office Add-in/pictures/向右选区.png" "f:/projects/My Office Add-in/assets/select-to-right-16.png"
```

> 注意：如果源文件尺寸不符合具体尺寸要求，后续可以用图片工具调整大小。当前插件对图标尺寸要求不严格，Excel 会自动缩放。

- [ ] **Step 4: 提交**

```bash
git add assets/select-to-right-16.png assets/select-to-right-32.png assets/select-to-right-80.png
git commit -m "feat: add select-to-right icon files"
```

---

### Task 1: 在 manifest.xml 添加按钮控件

**Files:**
- Modify: `manifest.xml`

- [ ] **Step 1: 在 DataProcessingGroup 末尾添加 SelectToRightButton 控件**

找到 `<Group id="DataProcessingGroup">` 内的最后一个 `</Control>`（SelectToEndButton 之后），在后面追加以下内容：

```xml
                <Control xsi:type="Button" id="SelectToRightButton">
                  <Label resid="SelectToRightButton.Label"/>
                  <Supertip>
                    <Title resid="SelectToRightButton.Label"/>
                    <Description resid="SelectToRightButton.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="SelectToRightIcon.16x16"/>
                    <bt:Image size="32" resid="SelectToRightIcon.32x32"/>
                    <bt:Image size="80" resid="SelectToRightIcon.80x80"/>
                  </Icon>
                  <Action xsi:type="ExecuteFunction">
                    <FunctionName>selectToRight</FunctionName>
                  </Action>
                </Control>
```

插入位置参考：在 `SelectToEndButton` 的 `</Control>` 之后、`</Group>` 之前。

- [ ] **Step 2: 验证 manifest.xml 完整性**

运行：

```bash
npm run validate
```

期望：通过，无 XML 语法错误。

- [ ] **Step 3: 提交**

```bash
git add manifest.xml
git commit -m "feat: add SelectToRightButton control to manifest.xml"
```

---

### Task 2: 在 manifest.xml 添加资源引用

**Files:**
- Modify: `manifest.xml`

> 按钮控件通过 `resid` 引用资源，这些 resId 必须在 `<Resources>` 块中有对应定义。

- [ ] **Step 1: 在 `<bt:ShortStrings>` 中添加标签文字**

找到 `<bt:ShortStrings>` 块末尾（SelectToEndButton.Label 之后），添加：

```xml
        <bt:String id="SelectToRightButton.Label" DefaultValue="向右选区"/>
```

- [ ] **Step 2: 在 `<bt:LongStrings>` 中添加提示文字**

找到 `<bt:LongStrings>` 块末尾（SelectToEndButton.Tooltip 之后），添加：

```xml
        <bt:String id="SelectToRightButton.Tooltip" DefaultValue="从选中单元格向右选到该行最后一个有数据的单元格"/>
```

- [ ] **Step 3: 在 `<bt:Images>` 中添加图标引用**

找到 `<bt:Images>` 块末尾（最后一个 `</bt:Image>` 之后、`</bt:Images>` 之前），添加：

```xml
        <bt:Image id="SelectToRightIcon.16x16" DefaultValue="https://localhost:3000/assets/select-to-right-16.png"/>
        <bt:Image id="SelectToRightIcon.32x32" DefaultValue="https://localhost:3000/assets/select-to-right-32.png"/>
        <bt:Image id="SelectToRightIcon.80x80" DefaultValue="https://localhost:3000/assets/select-to-right-80.png"/>
```

- [ ] **Step 4: 验证 manifest.xml**

运行：

```bash
npm run validate
```

期望：通过。

- [ ] **Step 5: 提交**

```bash
git add manifest.xml
git commit -m "feat: add resource references for SelectToRightButton"
```

---

### Task 3: 在 commands.html 实现 selectToRight 函数

**Files:**
- Modify: `src/commands/commands.html`

- [ ] **Step 1: 在 selectToEnd 的 associate 块之后添加 selectToRight 函数**

找到 `Office.actions.associate("selectToEnd", ...)` 块结束后的 `</script>` 标签之前，插入以下代码：

```javascript
        Office.actions.associate("selectToRight", function selectToRight(event) {
            Excel.run(function (context) {
                var range = context.workbook.getSelectedRange();
                range.load(["columnIndex", "columnCount", "rowIndex", "rowCount"]);
                var worksheet = context.workbook.worksheets.getActiveWorksheet();
                var usedRange = worksheet.getUsedRange();
                usedRange.load(["columnIndex", "columnCount"]);
                return context.sync().then(function () {
                    var startRow = range.rowIndex;
                    var rowCount = range.rowCount;
                    var startCol = range.columnIndex + 1; // 0-based → 1-based

                    // 通过 usedRange 确定工作表的实际数据右边界
                    var sheetLastCol = usedRange.columnIndex + usedRange.columnCount;
                    var endCol = Math.min(sheetLastCol, startCol + MAX_ROWS - 1);

                    if (startCol > endCol) {
                        console.warn("selectToRight: 所选区域没有数据");
                        return;
                    }

                    // 对每一行加载数据
                    var rowData = [];
                    for (var i = 0; i < rowCount; i++) {
                        var rowNum = startRow + i + 1;
                        var dataRange = worksheet.getRange(
                            getColumnLetter(startCol) + rowNum + ":" + getColumnLetter(endCol) + rowNum
                        );
                        dataRange.load("values");
                        rowData.push({ data: dataRange });
                    }

                    return context.sync().then(function () {
                        // 从右往左扫描，找每行最后一个有数据的列
                        var maxDataCol = 0;
                        for (var j = 0; j < rowData.length; j++) {
                            var values = rowData[j].data.values;
                            if (!values || !values[0]) { continue; }
                            var lastCol = 0;
                            for (var c = values[0].length - 1; c >= 0; c--) {
                                var v = values[0][c];
                                if (v !== null && v !== undefined && v !== "") {
                                    lastCol = startCol + c;
                                    break;
                                }
                            }
                            if (lastCol > maxDataCol) {
                                maxDataCol = lastCol;
                            }
                        }

                        if (maxDataCol === 0) {
                            console.warn("selectToRight: 所选区域没有数据");
                            return;
                        }

                        var startColLetter = getColumnLetter(range.columnIndex);
                        var endColLetter = getColumnLetter(maxDataCol - 1);
                        var firstRowNum = range.rowIndex + 1;
                        var lastRowNum = range.rowIndex + range.rowCount;
                        var targetRange = worksheet.getRange(
                            startColLetter + firstRowNum + ":" + endColLetter + lastRowNum
                        );
                        targetRange.select();
                        return context.sync();
                    });
                });
            }).catch(function (error) {
                console.error("selectToRight failed:", error);
            }).then(function () {
                event.completed();
            });
        });
```

> 关键逻辑说明：
> - `startCol = range.columnIndex + 1`：从选中单元格的右边一列开始扫描（选中单元格本身已包含在选区内）
> - `endCol = Math.min(sheetLastCol, startCol + MAX_ROWS - 1)`：使用 getUsedRange 确定的工作表数据右边界，MAX_ROWS 作为安全上限
> - 扫描时每行是一个单行多列的 Range，所以用 `values[0][c]` 访问第 c 列的值
> - 最终选区范围：从选中区域的原始起始单元格到（最远数据列 × 选中区域原始行范围）

- [ ] **Step 2: 运行 lint 检查**

```bash
npm run lint
```

期望：通过。

- [ ] **Step 3: 提交**

```bash
git add src/commands/commands.html
git commit -m "feat: implement selectToRight function"
```

---

### Task 4: 手动验证

**不需要代码变更** — 在 Excel 中加载插件后测试。

- [ ] **Step 1: 启动开发服务器和 Excel**

```bash
npm run start
```

打开 Excel，确保插件加载。

- [ ] **Step 2: 测试基础场景**

1. 在 A1 输入 "数据1"，B1 空，C1 输入 "数据2"
2. 选中 A1
3. 点击"向右选区"按钮
4. **期望**：选区变为 A1:C1

- [ ] **Step 3: 测试多行场景**

1. 在 A1:C1 和 A3:D3 中填入不同长度的数据（行 2 为空）
2. 选中 A1:A3
3. 点击"向右选区"按钮
4. **期望**：选区覆盖所有行中最远的列（如 A1:D3）

- [ ] **Step 4: 测试空行场景**

1. 选中一个完全空白的行中的单元格
2. 点击"向右选区"按钮
3. **期望**：无可见错误，选区不变（console.warn 静默处理）

- [ ] **Step 5: 测试边界场景**

1. 选中一个有数据的单元格，该单元格已位于行的最后一列
2. 点击"向右选区"按钮
3. **期望**：选区不变（仍是这一列）
