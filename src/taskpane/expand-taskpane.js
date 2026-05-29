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
    var refreshBtn = document.getElementById("refreshBtn");

    // 按钮点击触发展开
    executeBtn.addEventListener("click", function () {
      runExpand();
    });

    // 刷新按钮点击更新选择信息
    if (refreshBtn) {
      refreshBtn.addEventListener("click", function () {
        updateSelectionInfo();
      });
    }

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
      selectionInfo.value = range.address;
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
    range.load(["address", "columnCount", "rowCount", "values"]);
    return context.sync().then(function () {
      // 校验：至少2列
      if (range.columnCount < 2) {
        executeBtn.disabled = false;
        setStatus("错误：只有一列数据，无需展开", "error");
        return;
      }

      // 读取数据
      var values = range.values;

      // 校验 values 是否有效
      if (!values || !Array.isArray(values)) {
        executeBtn.disabled = false;
        setStatus("错误：无法读取选中区域的数据", "error");
        return;
      }

      // 过滤掉完全为空的行（当选择整列时可能包含 null 行）
      var filteredRows = [];
      for (var r = 0; r < values.length; r++) {
        var row = values[r];
        if (!row || !Array.isArray(row)) {
          continue;
        }
        var hasContent = false;
        for (var c = 0; c < row.length; c++) {
          if (row[c] !== null && row[c] !== undefined && row[c] !== "") {
            hasContent = true;
            break;
          }
        }
        if (hasContent) {
          filteredRows.push(row);
        }
      }
      values = filteredRows;

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
      originalSheet.load("name");
      return context.sync().then(function () {
        var originalSheetName = originalSheet.name;
        var newSheetName = originalSheetName + "_展开";

        // 创建新工作表（在原工作表之后）
        var sheetCollection = context.workbook.worksheets;
        sheetCollection.load("items");
        return context.sync().then(function () {
          // 检查是否存在同名工作表，如果存在则添加后缀
          var existingNames = sheetCollection.items.map(function (s) {
            s.load("name");
            return s.name;
          });
          return context.sync().then(function () {

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
            headerRange.values = [[values[0][0], "展开值"]];

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
      });
    });
  }).catch(function (error) {
    setStatus("错误: " + error.message, "error");
    executeBtn.disabled = false;
  });
}