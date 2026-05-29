# 展开列工具 - 设计文档

## 概述

独立侧边栏工具「展开列」：将宽表数据转换为长表格式。将第一列作为键列，其他列的值依次展开成独立行。

## 功能需求

| 需求 | 说明 |
|------|------|
| 数据选择 | 用户在 Excel 中选中包含数据的区域 |
| 表头处理 | 第一行自动识别为表头，不参与展开 |
| 输出位置 | 在原工作表后创建 `{原名}_展开` 新工作表 |
| 展开算法 | 第一列作为键，每行保留；其他列的值依次展开，空白跳过 |

## 展开算法

```
原始数据（第0行是表头）：
省份 | 地市 | 地市 | 地市
河南 | 郑州 | 南阳 | 许昌
河北 | 石家庄 | 保定 | (空)

展开后：
省份 | 地市
河南 | 郑州
河南 | 南阳
河南 | 许昌
河北 | 石家庄
河北 | 保定
```

- 第一列（省份）作为键，每行保留
- 其他列（地市）的值依次展开
- 空白单元格跳过，不输出
- 表头行不进入结果

## 交互流程

1. 用户在 Excel 选中包含数据的区域（包含表头）
2. 侧边栏显示"当前选中：{范围}"
3. 用户点击「执行展开」按钮
4. 工具处理数据，创建新工作表
5. 显示成功状态和结果行数

## 侧边栏 UI 结构

```
┌─────────────────────────────┐
│ 📋 展开列工具               │
├─────────────────────────────┤
│ 当前选中：Sheet1!A1:D3     │
│                             │
│ [执行展开]                  │
│                             │
│ 状态：等待操作...           │
└─────────────────────────────┘
```

## 文件结构

**新建文件：**
- `src/taskpane/expand-taskpane.html` — 侧边栏 HTML 结构
- `src/taskpane/expand-taskpane.css` — 样式
- `src/taskpane/expand-taskpane.js` — 业务逻辑
- `src/utils/expand-utils.js` — 展开算法工具函数
- `src/utils/expand-utils.test.js` — 单元测试

**修改文件：**
- `manifest.xml` — 注册新 ribbon 按钮

## Ribbon 按钮

在 Home Tab 新增按钮：
- 文本：展开列
- 图标：展开图标（双向箭头）
- Action：ShowTaskpane（打开侧边栏）

## 错误处理

| 场景 | 显示 |
|------|------|
| 未选中区域 | 错误：请先在 Excel 中选中数据区域 |
| 只选中一列 | 错误：只有一列数据，无需展开 |
| 无数据 | 错误：选中区域没有数据 |
| 命名冲突 | 自动重命名：Sheet1_展开 (1) |

## 技术实现

### expand-taskpane.js 核心逻辑

```javascript
function runExpand() {
  Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    range.load(["address", "columnCount", "rowCount"]);
    return context.sync().then(function () {
      // 1. 校验数据
      if (range.columnCount < 2) {
        setStatus("错误：只有一列数据，无需展开", "error");
        return;
      }

      // 2. 读取数据（包括表头）
      var values = range.values;
      var headers = values[0];
      var keyColumn = headers[0]; // 第一列列名作为键名

      // 3. 执行展开算法
      var result = expandData(values);

      // 4. 创建新工作表
      var originalSheetName = context.workbook.worksheets.getActiveWorksheet().name;
      var newSheetName = originalSheetName + "_展开";
      // ... 插入新工作表并写入数据
    });
  });
}
```

### expand-utils.js 展开算法

```javascript
function expandData(values) {
  // values[0] 是表头，从 values[1] 开始是数据
  // 第一列是键列，其他列是待展开列
  var result = [];
  var rowCount = values.length;

  for (var i = 1; i < rowCount; i++) {
    var key = values[i][0];
    // 遍历其他列（从索引1开始）
    for (var j = 1; j < values[i].length; j++) {
      var val = values[i][j];
      if (val !== null && val !== "" && val !== undefined) {
        result.push([key, val]);
      }
    }
  }

  return result;
}
```

## 实施步骤

1. 创建 `expand-taskpane.html` — 侧边栏 HTML
2. 创建 `expand-taskpane.css` — 样式（复用 concat-taskpane 的视觉风格）
3. 创建 `expand-taskpane.js` — 核心逻辑
4. 创建 `expand-utils.js` — 展开算法函数
5. 创建 `expand-utils.test.js` — 单元测试
6. 修改 `manifest.xml` — 注册新按钮
7. 测试完整流程