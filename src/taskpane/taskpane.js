/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

var { parseCSV } = require("../utils/csv-utils");
var {
  buildColRange,
  buildIndexMatchFormula,
  staticLookup,
  parseRangeAddress,
} = require("../utils/vlookup-utils");

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("importCsvBtn").onclick = function () {
      document.getElementById("csvFileInput").click();
    };
    document.getElementById("csvFileInput").onchange = handleCsvFile;
    document.getElementById("vlookupBtn").onclick = runVlookup;
  }
});

var MAX_ROWS = 1050000;
var isProcessing = false;

async function runVlookup() {
  if (isProcessing) return;
  isProcessing = true;

  var statusEl = document.getElementById("vlookupStatus");
  statusEl.textContent = "正在打开配置窗口...";
  statusEl.style.color = "green";

  Office.context.ui.displayDialogAsync(
    window.location.origin + "/vlookup-dialog.html",
    { height: 60, width: 35, displayInIframe: true },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        statusEl.textContent = "错误: 无法打开配置窗口";
        statusEl.style.color = "red";
        isProcessing = false;
        return;
      }
      var dialog = asyncResult.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
        dialog.close();
        var config;
        try {
          config = JSON.parse(arg.message);
        } catch (e) {
          statusEl.textContent = "";
          isProcessing = false;
          return;
        }
        if (config.type !== "vlookup") {
          statusEl.textContent = "";
          isProcessing = false;
          return;
        }
        executeVlookupFromConfig(config, statusEl);
      });
      dialog.addEventHandler(Office.EventType.DialogEventReceived, function () {
        statusEl.textContent = "已取消";
        statusEl.style.color = "gray";
        isProcessing = false;
      });
    }
  );
}

async function executeVlookupFromConfig(config, statusEl) {
  isProcessing = true;
  statusEl.textContent = "处理中...";
  statusEl.style.color = "green";

  function resolveDataRange(context) {
    // dataRange should always come from the user's current selection
    // (the cells that contain the values to look up)
    // config.lookupValue specifies which column to use for matching, not how many rows to process
    var range = context.workbook.getSelectedRange();
    range.load(["columnCount", "rowCount", "columnIndex", "rowIndex"]);
    return context.sync().then(function () {
      return range;
    });
  }

  try {
    await Excel.run(async (context) => {
      var dataRange = await resolveDataRange(context);

      if (dataRange.rowCount === 0) {
        statusEl.textContent = "错误: 没有数据";
        statusEl.style.color = "red";
        return;
      }

      if (dataRange.rowCount > MAX_ROWS) {
        statusEl.textContent =
          "错误: 数据量过大（" + dataRange.rowCount + " 行），单次最多支持 " + MAX_ROWS + " 行。";
        statusEl.style.color = "red";
        return;
      }

      var worksheet = context.workbook.worksheets.getActiveWorksheet();

      // Parse lookupValue to find the actual lookup value column and row
      var lookupValueStr = config.lookupValue || "";
      var lvParsed = parseRangeAddress(lookupValueStr);
      var dataStartCol = lvParsed.startCol; // 0-based column index
      var dataStartRow = lvParsed.startRow; // 1-based row number from address
      var dataRowCount = lvParsed.rowCount;
      var dataColLetter = getColumnLetter(dataStartCol);

      var lookupColRange = buildColRange(config.parsed, config.matchColIndex);
      var returnColRanges = [];
      for (var i = 0; i < config.returnColIndices.length; i++) {
        returnColRanges.push(buildColRange(config.parsed, config.returnColIndices[i]));
      }

      if (config.outputType === "formula") {
        var insertPos = dataStartCol + 1;
        var returnColCount = config.returnColIndices.length;

        for (var c = 0; c < returnColCount; c++) {
          var colLetter = getColumnLetter(insertPos + c);
          worksheet.getRange(colLetter + ":" + colLetter).insert(Excel.InsertShiftDirection.right);
        }
        await context.sync();

        // Build 2D formulas array - avoids autoFill issues by writing all formulas at once
        var formulas2D = [];
        for (var row = 0; row < dataRowCount; row++) {
          var rowFormulas = [];
          for (var col = 0; col < returnColCount; col++) {
            var lookupCellRef = dataColLetter + (dataStartRow + row);
            rowFormulas.push(
              buildIndexMatchFormula(
                lookupCellRef,
                lookupColRange,
                returnColRanges[col],
                config.matchMode
              )
            );
          }
          formulas2D.push(rowFormulas);
        }

        var outputRange = worksheet.getRange(
          getColumnLetter(insertPos) +
            dataStartRow +
            ":" +
            getColumnLetter(insertPos + returnColCount - 1) +
            (dataStartRow + dataRowCount - 1)
        );
        outputRange.formulas = formulas2D;
        await context.sync();

        statusEl.textContent =
          "完成! 已在 " + returnColCount + " 列写入 " + dataRowCount + " 行 INDEX/MATCH 公式";
      } else {
        // Static mode
        dataRange.load("values");
        await context.sync();
        var dataValues = dataRange.values;

        var lookupValues = [];
        for (var j = 0; j < dataValues.length; j++) {
          lookupValues.push(dataValues[j][0]);
        }

        var results = staticLookup(
          lookupValues,
          config.lookupTable,
          config.matchColIndex,
          config.returnColIndices,
          config.matchMode
        );

        var insertPos2 = dataStartCol + 1;
        var returnColCount2 = config.returnColIndices.length;

        for (var c2 = 0; c2 < returnColCount2; c2++) {
          var colLetter2 = getColumnLetter(insertPos2 + c2);
          worksheet
            .getRange(colLetter2 + ":" + colLetter2)
            .insert(Excel.InsertShiftDirection.right);
        }
        await context.sync();

        var targetRange = worksheet.getRange(
          getColumnLetter(insertPos2) +
            dataStartRow +
            ":" +
            getColumnLetter(insertPos2 + returnColCount2 - 1) +
            (dataStartRow + results.length - 1)
        );
        targetRange.values = results;
        await context.sync();

        statusEl.textContent =
          "完成! 已写入 " + results.length + " 行 × " + returnColCount2 + " 列静态值";
      }
    });
  } catch (error) {
    statusEl.textContent = "错误: " + error.message;
    statusEl.style.color = "red";
  }

  isProcessing = false;
}

function handleCsvFile(event) {
  var file = event.target.files[0];
  if (!file) return;

  var statusEl = document.getElementById("importStatus");
  statusEl.textContent = "正在读取文件...";
  statusEl.style.color = "green";

  var reader = new FileReader();
  reader.onload = function (e) {
    var text = e.target.result;
    importCSV(text, file.name);
  };
  reader.onerror = function () {
    statusEl.textContent = "错误: 文件读取失败";
    statusEl.style.color = "red";
  };
  reader.readAsText(file, "UTF-8");
}

async function importCSV(text, filename) {
  var statusEl = document.getElementById("importStatus");

  try {
    var delimiter = document.getElementById("csvDelimiter").value || ",";
    var rows = parseCSV(text, delimiter);

    if (rows.length === 0) {
      statusEl.textContent = "错误: CSV 文件为空";
      statusEl.style.color = "red";
      return;
    }

    if (rows.length > MAX_ROWS) {
      statusEl.textContent =
        "错误: 数据量过大（" + rows.length + " 行），单次最多支持 " + MAX_ROWS + " 行。";
      statusEl.style.color = "red";
      return;
    }

    statusEl.textContent = "正在写入 " + filename + " (" + rows.length + " 行)...";

    await Excel.run(async (context) => {
      var sheet = context.workbook.worksheets.getActiveWorksheet();
      // Start at the current selection, or A1 if nothing selected
      var startRange = context.workbook.getSelectedRange();
      startRange.load("address");
      await context.sync();

      var colCount = rows[0].length;
      var rowCount = rows.length;

      // Pad all rows to the same column count
      for (var r = 0; r < rowCount; r++) {
        while (rows[r].length < colCount) {
          rows[r].push("");
        }
      }

      // Write to worksheet starting at the selected cell
      var addr = startRange.address.split("!").pop(); // e.g., "A1" or "A1:B2"
      var startCol = getColNum(addr); // column index of start
      var startRow = getRowNum(addr); // row number of start
      var endCol = startCol + colCount - 1;
      var endRow = startRow + rowCount - 1;
      var targetRange = sheet.getRange(
        getColumnLetter(startCol) + startRow + ":" + getColumnLetter(endCol) + endRow
      );

      targetRange.values = rows;
      await context.sync();

      statusEl.textContent =
        "导入完成! " + filename + " → " + rowCount + " 行 × " + colCount + " 列";
    });
  } catch (error) {
    statusEl.textContent = "错误: " + error.message;
    statusEl.style.color = "red";
  }

  // Reset file input so the same file can be re-imported
  document.getElementById("csvFileInput").value = "";
}

// Extract column index from address like "A1" → 0
function getColNum(id) {
  var num = 0;
  for (var i = 0; i < id.length; i++) {
    var ch = id.charCodeAt(i);
    if (ch >= 65 && ch <= 90) {
      num = num * 26 + (ch - 65);
    } else {
      break;
    }
  }
  return num;
}

// Extract row number from address like "A1" → 1
function getRowNum(id) {
  return parseInt(id.replace(/[^0-9]/g, ""), 10) || 1;
}
