# 向右选区 — 设计文档

**日期**: 2026-06-12
**状态**: 已批准

## 概述

在 Excel 功能区"数据处理"分组下新增"向右选区"按钮，点击后从用户当前选中的单元格向右，选取到该行最后一个有数据的单元格（中间空单元格包含在选区内）。功能与已有的"向下选区"（selectToEnd）为镜像关系。

## 设计决策

| 项目 | 决定 |
|------|------|
| 按钮名称 | 向右选区 |
| 所在分组 | 数据处理（DataProcessingGroup） |
| 执行方式 | ExecuteFunction（点击即执行，无面板） |
| 函数名 | `selectToRight` |
| 图标 | `pictures\向右选区.png` → 生成 `assets/select-to-right-{16,32,80}.png` |
| 提示文字 | 从选中单元格向右选到该行最后一个有数据的单元格 |

## 行为定义

### 核心逻辑

1. 获取用户当前选中的区域
2. 用 `getUsedRange()` 确定工作表的实际数据列边界
3. 对于区域中的**每一行**，从起始列向右查找最后一个有数据的单元格
4. 最终选区为从选中区域第一个单元格到（最远数据列 × 选中区域最后一行）的矩形范围
5. 调用 Excel API 的 `select()` 完成选区

### "有数据"的定义

与"向下选区"一致：单元格 `values[r][0]` 不为 `null`、`undefined`、`""`。

### 多行行为

- 选中多行时，每行独立查找各自的数据末尾列
- 最终选区覆盖所有行中列号最远的那一个

### 边界情况

| 情况 | 行为 |
|------|------|
| 整行无数据 | 不做选区（console.warn） |
| 只有一列数据 | 选区就是这一列 |
| 当前列就是最后有数据的列 | 选区不变 |

## 实现

### 涉及文件

| 文件 | 变更内容 |
|------|----------|
| `manifest.xml` | 新增 `SelectToRightButton` 按钮控件及相关资源配置 |
| `src/commands/commands.html` | 新增 `selectToRight` 函数并关联到 `Office.actions.associate` |
| `assets/` | 新增 `select-to-right-16.png`、`select-to-right-32.png`、`select-to-right-80.png` |

### manifest.xml 变更

在 `<Group id="DataProcessingGroup">` 下 `SelectToEndButton` 之后新增按钮控件：

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

新增资源引用：

- `<bt:ShortStrings>`: `SelectToRightButton.Label` = "向右选区"
- `<bt:LongStrings>`: `SelectToRightButton.Tooltip` = "从选中单元格向右选到该行最后一个有数据的单元格"
- `<bt:Images>`: `SelectToRightIcon.{16,32,80}x{16,32,80}` → `assets/select-to-right-{16,32,80}.png`

### commands.html 变更

在 `selectToEnd` 的 `associate` 块之后，新增 `selectToRight` 函数：

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

### 与 selectToEnd 的差异

| 维度 | selectToEnd（向下） | selectToRight（向右） |
|------|---------------------|----------------------|
| 扫描方向 | 每列从下往上扫描行 | 每行从右往左扫描列 |
| 数据结构 | `values[r][0]` 取每行第一个值 | `values[0][c]` 取每列对应值 |
| 最终选区 | 起始行 → 最远列 | 起始列 → 最远行 |
| 安全上限 | `startRow + MAX_ROWS - 1` | `startCol + MAX_ROWS - 1` |

### 错误处理

- Excel.run 异常 → `console.error`（与 selectToEnd 一致）
- 无数据 → `console.warn` 静默退出（与 selectToEnd 一致）
- 无论成功或失败，必须调用 `event.completed()`

### 图标处理

1. 从 `pictures\向右选区.png` 复制并缩放为三个尺寸
2. 输出到 `assets/select-to-right-16.png`、`assets/select-to-right-32.png`、`assets/select-to-right-80.png`

## 测试要点

- 单行、中间有空单元格的场景
- 多行同时选区的场景
- 整行无数据的场景
- 只有一列数据的场景
- 选中单元格已位于最后一列
- 与"向下选区"组合使用（先向下再向右等）
