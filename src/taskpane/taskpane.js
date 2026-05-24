/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

var { getColumnLetter, escapeFormulaText, buildConcatFormula } = require("../utils/concat-utils");

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("concatBtn").onclick = runConcat;
  }
});

var MAX_ROWS = 1050000;
var isProcessing = false;

async function runConcat() {
  if (isProcessing) return;
  isProcessing = true;

  var statusEl = document.getElementById("status");
  var connectorInput = document.getElementById("connector");
  var connector = connectorInput.value || "_";

  statusEl.textContent = "处理中...";
  statusEl.style.color = "green";

  try {
    await Excel.run(async (context) => {
      var selectedRange = context.workbook.getSelectedRange();
      selectedRange.load(["address", "columnCount", "columnIndex"]);
      await context.sync();

      if (selectedRange.columnCount < 2) {
        statusEl.textContent = "错误: 请至少选择两列";
        statusEl.style.color = "red";
        return;
      }

      var colIndex = selectedRange.columnIndex;
      var worksheet = context.workbook.worksheets.getActiveWorksheet();
      var firstColLetter = getColumnLetter(colIndex);
      var secondColLetter = getColumnLetter(colIndex + 1);

      // Get row count from selected columns only
      var usedInSelection = selectedRange.getUsedRange();
      usedInSelection.load("rowCount");
      await context.sync();
      var rowCount = usedInSelection.rowCount;

      if (rowCount === 0) {
        statusEl.textContent = "错误: 没有数据";
        statusEl.style.color = "red";
        return;
      }

      if (rowCount > MAX_ROWS) {
        statusEl.textContent =
          "错误: 数据量过大（" + rowCount + " 行），单次最多支持 " + MAX_ROWS + " 行。";
        statusEl.style.color = "red";
        return;
      }

      // Insert new column
      var targetColLetter = getColumnLetter(colIndex + 2);
      worksheet
        .getRange(targetColLetter + ":" + targetColLetter)
        .insert(Excel.InsertShiftDirection.right);
      await context.sync();

      // Write formula to first cell
      var formula = buildConcatFormula(firstColLetter, secondColLetter, connector);
      var startCell = worksheet.getRange(targetColLetter + "1");
      startCell.formulas = [[formula]];
      await context.sync();

      // Auto-fill only if more than 1 row (filling to itself would error)
      if (rowCount > 1) {
        var fillRange = worksheet.getRange(
          targetColLetter + "1:" + targetColLetter + rowCount
        );
        startCell.autoFill(fillRange, Excel.AutoFillType.fillDefault);
        await context.sync();
      }

      statusEl.textContent =
        "完成! 已在第 " + targetColLetter + " 列写入 " + rowCount + " 行公式";
    });
  } catch (error) {
    statusEl.textContent = "错误: " + error.message;
    statusEl.style.color = "red";
  }

  isProcessing = false;
}
