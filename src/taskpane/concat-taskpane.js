/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

var MAX_ROWS = 1050000;

function getColumnLetter(colIndex) {
  var letter = "";
  var remaining = colIndex;
  do {
    letter = String.fromCharCode(65 + (remaining % 26)) + letter;
    remaining = Math.floor(remaining / 26) - 1;
  } while (remaining >= 0);
  return letter;
}

function escapeFormulaText(text) {
  return text.replace(/"/g, '""');
}

function buildConcatFormula(firstColLetter, secondColLetter, connector) {
  var escaped = escapeFormulaText(connector);
  return (
    "=IF(" +
    firstColLetter + "1&" + secondColLetter + '1="","",' +
    firstColLetter + '1&"' + escaped + '"&' + secondColLetter +
    "1)"
  );
}

Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    var connectorInput = document.getElementById("connector");
    var executeBtn = document.getElementById("executeBtn");

    // 自动聚焦输入框
    connectorInput.focus();

    // 按钮点击触发执行
    executeBtn.addEventListener("click", function() {
      runConcat();
    });
  }
});

function runConcat() {
  var statusEl = document.getElementById("status");
  var connectorInput = document.getElementById("connector");
  var executeBtn = document.getElementById("executeBtn");
  var connector = connectorInput.value || "_";

  function setStatus(message, type) {
    var el = document.getElementById("status");
    el.textContent = message;
    el.className = "status-message status-" + type;
  }

  // 禁用按钮，显示加载状态
  executeBtn.disabled = true;
  setStatus("处理中...", "loading");

  Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    range.load(["address", "columnCount", "columnIndex"]);
    return context.sync().then(function () {
      if (range.columnCount < 2) {
        setStatus("错误: 请至少选择两列", "error");
        return;
      }

      var worksheet = context.workbook.worksheets.getActiveWorksheet();
      var colIndex = range.columnIndex;
      var firstColLetter = getColumnLetter(colIndex);
      var secondColLetter = getColumnLetter(colIndex + 1);

      var usedInSelection = range.getUsedRange();
      usedInSelection.load("rowCount");
      return context.sync().then(function () {
        var rowCount = usedInSelection.rowCount;

        if (rowCount === 0) {
          setStatus("错误: 没有数据", "error");
          return;
        }

        if (rowCount > MAX_ROWS) {
          setStatus(
            "错误: 数据量过大（" + rowCount + " 行），单次最多支持 " + MAX_ROWS + " 行。",
            "error"
          );
          return;
        }

        var targetColLetter = getColumnLetter(colIndex + 2);
        worksheet
          .getRange(targetColLetter + ":" + targetColLetter)
          .insert(Excel.InsertShiftDirection.right);
        return context.sync().then(function () {
          var formula = buildConcatFormula(firstColLetter, secondColLetter, connector);

          var startCell = worksheet.getRange(targetColLetter + "1");
          startCell.formulas = [[formula]];
          return context.sync().then(function () {
            if (rowCount > 1) {
              var fillRange = worksheet.getRange(
                targetColLetter + "1:" + targetColLetter + rowCount
              );
              startCell.autoFill(fillRange, Excel.AutoFillType.fillDefault);
              return context.sync();
            }
          }).then(function () {
            setStatus(
              "完成! 已在第 " + targetColLetter + " 列写入 " + rowCount + " 行公式",
              "success"
            );
            executeBtn.disabled = false;
          });
        });
      });
    });
  }).catch(function (error) {
    setStatus("错误: " + error.message, "error");
    executeBtn.disabled = false;
  });
}