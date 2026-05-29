/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

var MAX_ROWS = 1050000;

/**
 * 将列索引（0-based）转换为 Excel 列字母
 * @param {number} colIndex - 列索引（0-based）
 * @returns {string} 列字母（如 A, B, ..., Z, AA, AB, ...）
 */
function getColumnLetter(colIndex) {
  var letter = "";
  var remaining = colIndex;
  do {
    letter = String.fromCharCode(65 + (remaining % 26)) + letter;
    remaining = Math.floor(remaining / 26) - 1;
  } while (remaining >= 0);
  return letter;
}

Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    var headerRowInput = document.getElementById("headerRow");
    var executeBtn = document.getElementById("executeBtn");

    // 自动聚焦到表头行号输入框
    headerRowInput.focus();

    // 执行按钮点击事件
    executeBtn.addEventListener("click", function () {
      runCount();
    });
  }
});

/**
 * 设置状态消息
 * @param {string} message - 状态文本
 * @param {string} type - 状态类型：idle, loading, success, error
 */
function setStatus(message, type) {
  var el = document.getElementById("status");
  el.textContent = message;
  el.className = "status-message status-" + type;
}

/**
 * 主执行函数：获取选中列，统计计数
 */
function runCount() {
  var headerRowInput = document.getElementById("headerRow");
  var executeBtn = document.getElementById("executeBtn");
  var headerRow = parseInt(headerRowInput.value) || 1;

  // 禁用按钮，显示加载状态
  executeBtn.disabled = true;
  setStatus("处理中...", "loading");

  Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    range.load(["address", "columnCount", "columnIndex", "rowCount"]);
    return context.sync().then(function () {
      if (range.columnCount < 1) {
        executeBtn.disabled = false;
        setStatus("错误: 请先选择一列", "error");
        return;
      }

      var colIndex = range.columnIndex; // 0-based
      var colCount = range.columnCount;
      var worksheet = context.workbook.worksheets.getActiveWorksheet();

      // 获取每个选中列的表头和数据范围
      var columns = [];
      for (var i = 0; i < colCount; i++) {
        var currentColIndex = colIndex + i;
        var colLetter = getColumnLetter(currentColIndex);

        // 获取表头值（加载 values 属性，返回 2D 数组）
        var headerCell = worksheet.getRange(colLetter + headerRow);
        headerCell.load("values");
        columns.push({
          colLetter: colLetter,
          colIndex: currentColIndex,
          headerCell: headerCell
        });
      }

      return context.sync().then(function () {
        // 获取每个列的表头值（Excel API 使用 .values[0][0] 获取单列元格的值）
        for (var i = 0; i < columns.length; i++) {
          var rawValue = columns[i].headerCell.values[0][0];
          var headerText = (rawValue !== null && rawValue !== undefined) ? String(rawValue) : "";
          if (headerText === "") {
            headerText = "Column" + (columns[i].colIndex + 1);
          }
          columns[i].headerText = headerText;
        }

        // 获取每个列的数据范围（表头行下方到最后一行的数据）
        var dataStartRow = headerRow + 1;

        // 获取所有列的最后一行的最大值
        var maxLastRow = 0;
        var columnDataRanges = [];

        for (var j = 0; j < columns.length; j++) {
          var col = columns[j];
          var colLetter = col.colLetter;

          // 获取特定列的UsedRange
          var singleColUsed = worksheet.getRange(colLetter + dataStartRow + ":" + colLetter + 1000000);
          singleColUsed.load("values");
          singleColUsed.load("rowCount");

          columnDataRanges.push({
            col: col,
            dataStartRow: dataStartRow,
            usedRange: singleColUsed
          });
        }

        return context.sync().then(function () {
          // 计算每个列的实际数据范围
          for (var k = 0; k < columnDataRanges.length; k++) {
            var item = columnDataRanges[k];
            var values = item.usedRange.values;
            var rowCount = item.usedRange.rowCount;

            // 找到最后有数据的行
            var lastRow = dataStartRow;
            for (var r = 0; r < rowCount; r++) {
              if (values[r] && values[r][0] !== null && values[r][0] !== "") {
                lastRow = dataStartRow + r;
              }
            }
            item.lastRow = lastRow;
            // 使用之前保存的 headerText 字符串，而不是通过 Excel proxy 对象获取
            item.headerValue = item.col.headerText;
          }

          // 执行计数操作
          return executeCount(worksheet, columnDataRanges, headerRow, executeBtn);
        });
      });
    });
  }).catch(function (error) {
    setStatus("错误: " + error.message, "error");
    executeBtn.disabled = false;
  });
}

/**
 * 执行计数操作：使用 Map 在内存中统计每个值的出现次数
 * @param {object} worksheet - Excel worksheet 对象
 * @param {Array} columnDataRanges - 列数据范围信息数组
 * @param {number} headerRow - 表头行号
 * @param {object} executeBtn - 按钮对象
 * @returns {Promise}
 */
function executeCount(worksheet, columnDataRanges, headerRow, executeBtn) {
  return Excel.run(function (context) {
    var results = [];

    // 从右到左处理，避免插入列后影响后续列的索引
    for (var i = columnDataRanges.length - 1; i >= 0; i--) {
      var item = columnDataRanges[i];
      var colLetter = item.col.colLetter;
      var colIndex = item.col.colIndex;
      var dataStartRow = item.dataStartRow;
      var lastRow = item.lastRow;
      var headerValue = item.headerValue;
      var values = item.usedRange.values;
      var rowCount = item.usedRange.rowCount;

      // Step 2: 用 Map 统计每个值的出现次数
      var countMap = new Map();
      for (var r = 0; r < rowCount; r++) {
        var key = values[r][0];
        // 跳过空值
        if (key === null || key === undefined || key === "") {
          continue;
        }
        countMap.set(key, (countMap.get(key) || 0) + 1);
      }

      // Step 3: 在当前列后插入新列
      var resultColLetter = getColumnLetter(colIndex + 1);
      worksheet.getRange(resultColLetter + ":" + resultColLetter).insert(Excel.InsertShiftDirection.right);

      // Step 4: 写入表头
      var headerCell = worksheet.getRange(resultColLetter + headerRow);
      headerCell.values = [[headerValue + "-计数"]];

      // Step 5: 生成计数结果数组（2D 数组格式）
      var dataRowCount = lastRow - dataStartRow + 1;
      var countResults = [];
      for (var dr = 0; dr < rowCount; dr++) {
        var cellValue = values[dr][0];
        var count = countMap.get(cellValue) || 0;
        countResults.push([count]);
      }

      // Step 6: 一次性写入所有计数结果
      var dataRange = worksheet.getRange(resultColLetter + dataStartRow + ":" + resultColLetter + lastRow);
      dataRange.values = countResults;

      results.push({
        colLetter: colLetter,
        resultColLetter: resultColLetter,
        rowCount: dataRowCount
      });
    }

    return context.sync().then(function () {
      var resultParts = [];
      for (var j = 0; j < results.length; j++) {
        var res = results[j];
        resultParts.push("第" + res.resultColLetter + "列(" + res.rowCount + "行)");
      }
      setStatus("完成！已在 " + resultParts.join(", ") + " 写入计数结果", "success");
      executeBtn.disabled = false;
    });
  }).catch(function (error) {
    setStatus("错误: " + error.message, "error");
    executeBtn.disabled = false;
  });
}