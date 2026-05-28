/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

var MAX_CELLS = 1050000;

Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    var startValueInput = document.getElementById("startValue");
    var stepValueInput = document.getElementById("stepValue");
    var executeBtn = document.getElementById("executeBtn");

    // Auto-focus input
    startValueInput.focus();

    // Button click triggers execution
    executeBtn.addEventListener("click", function () {
      runFillSeries();
    });
  }
});

function setStatus(message, type) {
  var el = document.getElementById("status");
  el.textContent = message;
  el.className = "status-message status-" + type;
}

function validateNumber(value) {
  return /^-?\d+(\.\d+)?$/.test(value);
}

function runFillSeries() {
  var executeBtn = document.getElementById("executeBtn");
  var startValueInput = document.getElementById("startValue");
  var stepValueInput = document.getElementById("stepValue");

  var startValue = startValueInput.value;
  var stepValue = stepValueInput.value;

  // Validate non-empty
  if (!startValue || !stepValue) {
    setStatus("错误: 起始值和步长不能为空", "error");
    return;
  }

  // Validate numeric format
  if (!validateNumber(startValue) || !validateNumber(stepValue)) {
    setStatus("错误: 请输入有效的数字", "error");
    return;
  }

  // Disable button and show loading
  executeBtn.disabled = true;
  setStatus("处理中...", "loading");

  Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    range.load(["rowCount", "columnCount"]);
    return context.sync().then(function () {
      var rowCount = range.rowCount;
      var columnCount = range.columnCount;
      var total = rowCount * columnCount;

      // Check max cells
      if (total > MAX_CELLS) {
        executeBtn.disabled = false;
        setStatus(
          "错误: 数据量过大（" + total + " 个单元格），单次最多支持 " + MAX_CELLS + " 个单元格。",
          "error"
        );
        return;
      }

      // Parse numeric values
      var start = parseFloat(startValue);
      var step = parseFloat(stepValue);

      // Generate 1D sequence array: [start, start+step, ...]
      var sequence = [];
      for (var i = 0; i < total; i++) {
        sequence.push(start + i * step);
      }

      // Convert to 2D array (row-major: index = c * rowCount + r)
      var result = [];
      for (var r = 0; r < rowCount; r++) {
        var row = [];
        for (var c = 0; c < columnCount; c++) {
          row.push(sequence[r * columnCount + c]);
        }
        result.push(row);
      }

      // Write to range
      range.values = result;

      return context.sync().then(function () {
        setStatus("完成! 已填充 " + rowCount + " 行 x " + columnCount + " 列 的序列", "success");
        executeBtn.disabled = false;
      });
    });
  }).catch(function (error) {
    setStatus("错误: " + error.message, "error");
    executeBtn.disabled = false;
  });
}