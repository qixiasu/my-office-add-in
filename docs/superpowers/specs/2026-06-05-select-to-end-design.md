# 向下选区 — 设计文档

**日期**: 2026-06-05
**状态**: 已批准

## 概述

在 Excel 功能区添加"向下选区"按钮，点击后自动从用户当前选中的单元格向下，选取到该列最后一个有数据的单元格（中间空单元格包含在选区内）。

## 设计决策

| 项目 | 决定 |
|------|------|
| 按钮名称 | 向下选区 |
| 所在分组 | 数据处理（DataProcessingGroup） |
| 执行方式 | ExecuteFunction（点击即执行，无面板） |
| 函数名 | `selectToEnd` |
| 图标 | 暂时复用 tools 图标，后续可替换 |
| 提示文字 | 从选中单元格向下选到该列最后一个有数据的单元格 |

## 行为定义

### 核心逻辑

1. 获取用户当前选中的区域
2. 对于区域中的**每一列**，从选中行开始向下查找最后一个有数据的单元格
3. 最终选区为从选中区域第一个单元格到（最远数据行 × 选中区域最右列）的矩形范围
4. 调用 Excel API 的 `select()` 完成选区

### "有数据"的定义

单元格满足以下条件即为"有数据"：
- `values[r][0]` 不为 `null`、`undefined`、`""`

**已知限制**：Excel API 的 `values` 属性只返回计算后的值，无法区分空白单元格和公式返回 `""` 的单元格。如需严格区分，需额外加载 `formulas` 属性判断。当前实现仅基于 `values` 判断，公式返回 `""` 的单元格会视为空。中间空单元格/公式空单元格**仍然会包含在最终选区内**，只在确定最后一行的位置时可能少选末尾的纯公式空行。

### 多列行为

- 选中多列时，每列独立查找各自的数据末尾行
- 最终选区覆盖所有列中行号最远的那个

### 边界情况

| 情况 | 行为 |
|------|------|
| 整列无数据 | 提示"该列没有数据"，不做选区 |
| 只有一行数据 | 选区就是这一行 |
| 当前行就是最后一行 | 选区不变 |

## 实现

### 涉及文件

| 文件 | 变更内容 |
|------|----------|
| `manifest.xml` | 新增 `SelectToEndButton` 按钮控件及相关资源配置 |
| `src/commands/commands.html` | 新增 `selectToEnd` 函数并关联到 `Office.actions.associate` |

### manifest.xml 变更

在 `<Group id="DataProcessingGroup">` 下新增按钮控件：

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

新增资源引用（ShortStrings、LongStrings、Images）。

### commands.html 变更

在 `commands.html` 内联脚本中新增 `selectToEnd` 函数：

```javascript
Office.actions.associate("selectToEnd", function selectToEnd(event) {
    Excel.run(function (context) {
        var range = context.workbook.getSelectedRange();
        range.load(["address", "columnIndex", "columnCount", "rowIndex", "rowCount"]);
        return context.sync().then(function () {
            var colIndex = range.columnIndex;
            var colCount = range.columnCount;
            var startRow = range.rowIndex + 1; // 1-based
            var worksheet = context.workbook.worksheets.getActiveWorksheet();

            // 对每一列加载数据
            var columnData = [];
            for (var i = 0; i < colCount; i++) {
                var colLetter = getColumnLetter(colIndex + i);
                var dataRange = worksheet.getRange(
                    colLetter + startRow + ":" + colLetter + MAX_ROWS
                );
                dataRange.load("values");
                columnData.push({ colLetter: colLetter, data: dataRange });
            }

            return context.sync().then(function () {
                // 从下往上找每列最后一个有数据的行
                var maxDataRow = 0;
                for (var j = 0; j < columnData.length; j++) {
                    var values = columnData[j].data.values;
                    var lastRow = 0;
                    for (var r = values.length - 1; r >= 0; r--) {
                        if (values[r][0] !== null && values[r][0] !== undefined && values[r][0] !== "") {
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

复用 `commands.html` 中已有的 `getColumnLetter` 和 `MAX_ROWS`（1050000）常量。

### 错误处理

- Excel.run 异常 → messageBox 弹出错误信息
- 无数据 → messageBox 提示"该列没有数据"
- 无论成功或失败，必须调用 `event.completed()`

## 测试要点

- 单列、中间有空单元格的场景
- 多列同时选区的场景
- 全列无数据的场景
- 只有一行数据的场景
- 选中单元格已位于最后一行
- 数据量较大时的性能（如 10000 行）
