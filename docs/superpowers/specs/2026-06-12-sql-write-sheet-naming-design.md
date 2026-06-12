# SQL 查询：写入新工作表的命名功能

## 问题

`writeResultToSheet()` 在写入新工作表时使用固定名称 `"查询结果"`。当该工作表已存在时，Excel 抛出错误：
> 已存在具有相同名称或标识符的资源。

## 方案

采用浏览器原生 `prompt()` 弹窗 + 自动序号兜底。

### 交互流程

1. 用户点击"📝 写入新工作表"按钮
2. 弹出 `prompt("请输入工作表名称:", "查询结果")`
3. 名称决策：
   - 用户点击取消 → 使用默认名 `"查询结果"`
   - 用户输入为空 → 使用默认名 `"查询结果"`
   - 用户输入了名称 → 使用输入的名称
4. 检查工作簿中是否已存在同名 sheet
5. 如果存在 → 自动追加序号：`"查询结果"` → `"查询结果 (1)"` → `"查询结果 (2)"` ...
6. 如果不存在 → 直接用该名称
7. 创建 sheet → 写入数据

### 实现细节

**修改文件：** `src/taskpane/sql-query-taskpane.js`

**`writeResultToSheet()` 函数重构：**

```javascript
function writeResultToSheet() {
  if (!currentQueryResult) return;

  var defaultName = "查询结果";
  var sheetName = prompt("请输入工作表名称:", defaultName);
  if (sheetName === null || sheetName.trim() === "") {
    sheetName = defaultName;
  }
  var rows = currentQueryResult.rows;
  var columns = currentQueryResult.columns;

  Excel.run(function (context) {
    var sheetCollection = context.workbook.worksheets;
    sheetCollection.load("items/name");
    return context.sync().then(function () {
      // 生成不重复的名称
      var finalName = sheetName;
      var counter = 1;
      var exists = true;
      while (exists) {
        exists = false;
        for (var i = 0; i < sheetCollection.items.length; i++) {
          if (sheetCollection.items[i].name === finalName) {
            exists = true;
            finalName = sheetName + " (" + counter + ")";
            counter++;
            break;
          }
        }
      }

      var newSheet = sheetCollection.add(finalName);
      newSheet.position = 0;

      var totalRows = rows.length + 1;
      var range = newSheet.getRangeByIndexes(0, 0, totalRows, columns.length);
      var values = [columns];
      for (var r = 0; r < rows.length; r++) {
        values.push(rows[r]);
      }
      range.values = values;
      range.format.autofitColumns();
      newSheet.activate();

      return context.sync();
    });
  }).then(function () {
    setStatusText("queryStatus", "已将 " + rows.length + " 行结果写入新工作表", "success");
  }).catch(function (error) {
    setStatusText("queryStatus", "写入失败: " + error.message, "error");
  });
}
```

### 不涉及的部分

- 不修改 HTML 结构
- 不修改样式
- 不修改其他函数
- 不需要额外的依赖
