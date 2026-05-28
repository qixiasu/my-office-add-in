/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

var { staticLookup, parseRangeAddress } = require("../utils/vlookup-utils");

var { getColumnLetter } = require("../utils/concat-utils");

Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    initEventListeners();
    loadInitialSelection();
  }
});

var g_lookupTableData = null;
var g_lookupTableParsed = null;

function initEventListeners() {
  document.getElementById("refreshLookupValue").onclick = function () {
    refreshSelection("lookupValue");
  };

  document.getElementById("refreshLookupTable").onclick = function () {
    refreshSelection("lookupTable");
  };

  document.getElementById("selectAllColumns").onclick = function () {
    setAllCheckboxes("returnColumns", true);
    validateForm();
  };

  document.getElementById("deselectAllColumns").onclick = function () {
    setAllCheckboxes("returnColumns", false);
    validateForm();
  };

  document.getElementById("cancelBtn").onclick = function () {
    Office.context.ui.close();
  };

  document.getElementById("executeBtn").onclick = executeLookup;

  // 初始化时如果 lookupTable 已有值，自动加载表头
  setTimeout(function () {
    if (document.getElementById("lookupTable").value) {
      loadTableHeaders();
    }
  }, 0);

  var form = document.getElementById("vlookupForm");
  form.addEventListener("change", validateForm);
  form.addEventListener("input", validateForm);

  document.getElementById("headerRow").addEventListener("input", function () {
    var tableInput = document.getElementById("lookupTable").value;
    if (tableInput) {
      loadTableHeaders();
    }
  });
}

function loadInitialSelection() {
  // 仅加载选择区域到 lookupValue（查找值区域），不自动填充 lookupTable
  Excel.run(async function (context) {
    var range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();

    if (range.address) {
      document.getElementById("lookupValue").value = range.address;
    }
  }).catch(function (error) {
    setStatus("加载选择失败: " + error.message, "error");
  });
}

function refreshSelection(target) {
  Excel.run(async function (context) {
    var range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();

    document.getElementById(target).value = range.address;

    if (target === "lookupTable") {
      g_lookupTableParsed = parseRangeAddress(range.address);
      loadTableHeaders();
    }
  }).catch(function (error) {
    setStatus("刷新选择失败: " + error.message, "error");
  });
}

function refreshColumns() {
  var tableInput = document.getElementById("lookupTable").value;
  if (!tableInput) {
    return;
  }

  var parsed = parseRangeAddress(tableInput);
  g_lookupTableParsed = parsed;

  var matchColumnSelect = document.getElementById("matchColumn");
  var returnColumnsDiv = document.getElementById("returnColumns");

  matchColumnSelect.innerHTML = '<option value="">-- 选择匹配列 --</option>';
  returnColumnsDiv.innerHTML = '<span class="vlookup-label">-- 请先选择查找表并读取表头 --</span>';

  var colCount = parsed.colCount;
  for (var i = 0; i < colCount; i++) {
    var matchOption = document.createElement("option");
    matchOption.value = i;
    matchOption.textContent = "列 " + getColumnLetter(parsed.startCol + i) + " (列" + (i + 1) + ")";
    matchColumnSelect.appendChild(matchOption);

    var checkboxWrapper = document.createElement("div");
    checkboxWrapper.className = "vlookup-checkbox-item";

    var checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.name = "returnColumn";
    checkbox.value = i;
    checkbox.id = "returnCol_" + i;
    checkbox.onchange = validateForm;

    var label = document.createElement("label");
    label.htmlFor = "returnCol_" + i;
    label.textContent = "列 " + getColumnLetter(parsed.startCol + i) + " (列" + (i + 1) + ")";

    checkboxWrapper.appendChild(checkbox);
    checkboxWrapper.appendChild(label);
    returnColumnsDiv.appendChild(checkboxWrapper);
  }
}

function loadTableHeaders() {
  var tableInput = document.getElementById("lookupTable").value;
  if (!tableInput) {
    setStatus("请先选择查找表区域", "error");
    return;
  }

  var headerRow = parseInt(document.getElementById("headerRow").value, 10) || 1;

  Excel.run(async function (context) {
    var tableAddress = tableInput;
    var parsed = parseRangeAddress(tableAddress);
    g_lookupTableParsed = parsed;

    var sheet = context.workbook.worksheets.getActiveWorksheet();

    var startRow = parsed.startRow;
    var endRow = parsed.endRow;

    if (headerRow < startRow || headerRow > endRow) {
      setStatus("表头行号超出范围", "error");
      return;
    }

    var headerRange = sheet.getRange(
      getColumnLetter(parsed.startCol) +
        headerRow +
        ":" +
        getColumnLetter(parsed.endCol) +
        headerRow
    );
    headerRange.load("values");
    await context.sync();

    var headers = headerRange.values[0] || [];

    var dataRange = sheet.getRange(
      getColumnLetter(parsed.startCol) + startRow + ":" + getColumnLetter(parsed.endCol) + endRow
    );
    dataRange.load("values");
    await context.sync();

    g_lookupTableData = dataRange.values;

    refreshColumns();
    updateHeaderLabels(headers, parsed);

    setStatus("表头读取成功，共 " + parsed.colCount + " 列", "success");
    validateForm();
  }).catch(function (error) {
    setStatus("读取表头失败: " + error.message, "error");
  });
}

function updateHeaderLabels(headers, parsed) {
  var matchColumnSelect = document.getElementById("matchColumn");
  var returnColumnsDiv = document.getElementById("returnColumns");
  var options = matchColumnSelect.querySelectorAll("option");

  for (var i = 1; i < options.length; i++) {
    var headerName = headers[i - 1] || "列" + i;
    options[i].textContent = headerName + " (列" + i + ")";
  }

  var checkboxItems = returnColumnsDiv.querySelectorAll(".vlookup-checkbox-item");
  for (var j = 0; j < checkboxItems.length; j++) {
    var label = checkboxItems[j].querySelector("label");
    var headerName = headers[j] || "列" + (j + 1);
    label.textContent = headerName;
  }
}

function getCheckedReturnCols() {
  var checkboxes = document.querySelectorAll('#returnColumns input[type="checkbox"]:checked');
  var indices = [];
  for (var i = 0; i < checkboxes.length; i++) {
    indices.push(parseInt(checkboxes[i].value, 10));
  }
  return indices;
}

function extractFirstColAddress(rangeStr) {
  var parsed = parseRangeAddress(rangeStr);
  return getColumnLetter(parsed.startCol) + ":" + getColumnLetter(parsed.startCol);
}

function setStatus(message, type) {
  var statusEl = document.getElementById("statusMessage");
  statusEl.textContent = "状态：" + message;
  statusEl.className = "status-message status-" + (type || "idle");
}

function validateForm() {
  var tableInput = document.getElementById("lookupTable").value;
  var valueInput = document.getElementById("lookupValue").value;
  var matchCol = document.getElementById("matchColumn").value;
  var returnCols = getCheckedReturnCols();

  var isValid = tableInput && valueInput && matchCol !== "" && returnCols.length > 0;

  document.getElementById("executeBtn").disabled = !isValid;
}

function executeLookup() {
  var config = {
    lookupValue: document.getElementById("lookupValue").value,
    lookupTable: document.getElementById("lookupTable").value,
    headerRow: parseInt(document.getElementById("headerRow").value, 10) || 1,
    matchColIndex: parseInt(document.getElementById("matchColumn").value, 10),
    returnColIndices: getCheckedReturnCols(),
    matchMode: document.querySelector('input[name="matchMode"]:checked').value === "exact" ? 0 : 1,
    defaultValue: document.getElementById("defaultValue").value || "#N/A",
  };

  performLookup(config);
}

function performLookup(config) {
  setStatus("处理中...", "info");

  Excel.run(async function (context) {
    var BATCH_SIZE = 10000;
    var LARGE_DATA_THRESHOLD = 100000;

    var worksheet = context.workbook.worksheets.getActiveWorksheet();

    var lvParsed = parseRangeAddress(config.lookupValue);
    var dataStartRow = lvParsed.startRow;
    var dataStartCol = lvParsed.startCol;
    var dataRowCount = lvParsed.rowCount;
    var dataColLetter = getColumnLetter(dataStartCol);

    var ltParsed = parseRangeAddress(config.lookupTable);

    if (!g_lookupTableData) {
      var tableRange = worksheet.getRange(
        getColumnLetter(ltParsed.startCol) +
          ltParsed.startRow +
          ":" +
          getColumnLetter(ltParsed.endCol) +
          ltParsed.endRow
      );
      tableRange.load("values");
      await context.sync();
      g_lookupTableData = tableRange.values;
    }

    if (dataRowCount < LARGE_DATA_THRESHOLD) {
      // Small data: single read -> staticLookup -> single write
      var dataRange = worksheet.getRange(
        dataColLetter + dataStartRow + ":" + dataColLetter + (dataStartRow + dataRowCount - 1)
      );
      dataRange.load("values");
      await context.sync();

      var lookupValues = [];
      for (var j = 0; j < dataRange.values.length; j++) {
        lookupValues.push(dataRange.values[j][0]);
      }

      var results = staticLookup(
        lookupValues,
        g_lookupTableData,
        config.matchColIndex,
        config.returnColIndices,
        config.matchMode,
        config.defaultValue
      );

      var insertPos = ltParsed.endCol + 1;
      var returnColCount = config.returnColIndices.length;

      for (var c = 0; c < returnColCount; c++) {
        var colLetter = getColumnLetter(insertPos + c);
        worksheet.getRange(colLetter + ":" + colLetter).insert(Excel.InsertShiftDirection.right);
      }
      await context.sync();

      var targetRange = worksheet.getRange(
        getColumnLetter(insertPos) +
          dataStartRow +
          ":" +
          getColumnLetter(insertPos + returnColCount - 1) +
          (dataStartRow + results.length - 1)
      );
      targetRange.values = results;
      await context.sync();

      setStatus(
        "完成! 已写入 " + results.length + " 行 x " + returnColCount + " 列静态值",
        "success"
      );
    } else {
      // Large data: batch processing
      var totalRows = dataRowCount;
      var processedRows = 0;
      var insertPos = ltParsed.endCol + 1;
      var returnColCount = config.returnColIndices.length;

      // Insert all return columns first (once, before batches)
      for (var ci = 0; ci < returnColCount; ci++) {
        var colLetter = getColumnLetter(insertPos + ci);
        worksheet.getRange(colLetter + ":" + colLetter).insert(Excel.InsertShiftDirection.right);
      }
      await context.sync();

      while (processedRows < totalRows) {
        var currentBatchSize = Math.min(BATCH_SIZE, totalRows - processedRows);
        var batchStartRow = dataStartRow + processedRows;
        var batchEndRow = batchStartRow + currentBatchSize - 1;

        // Read current batch lookupValues
        var batchDataRange = worksheet.getRange(
          dataColLetter + batchStartRow + ":" + dataColLetter + batchEndRow
        );
        batchDataRange.load("values");
        await context.sync();

        var batchLookupValues = [];
        for (var bj = 0; bj < batchDataRange.values.length; bj++) {
          batchLookupValues.push(batchDataRange.values[bj][0]);
        }

        // Compute current batch
        var batchResults = staticLookup(
          batchLookupValues,
          g_lookupTableData,
          config.matchColIndex,
          config.returnColIndices,
          config.matchMode,
          config.defaultValue
        );

        // Write current batch results
        var batchTargetRange = worksheet.getRange(
          getColumnLetter(insertPos) +
            batchStartRow +
            ":" +
            getColumnLetter(insertPos + returnColCount - 1) +
            batchEndRow
        );
        batchTargetRange.values = batchResults;
        await context.sync();

        processedRows += currentBatchSize;

        // Update progress status
        setStatus("处理中... 已完成 " + processedRows + " / " + totalRows + " 行", "info");
      }

      setStatus("完成! 已写入 " + totalRows + " 行 x " + returnColCount + " 列静态值", "success");
    }
  }).catch(function (error) {
    setStatus("错误: " + error.message, "error");
  });
}

function setAllCheckboxes(containerId, checked) {
  var container = document.getElementById(containerId);
  var checkboxes = container.querySelectorAll('input[type="checkbox"]');
  for (var i = 0; i < checkboxes.length; i++) {
    checkboxes[i].checked = checked;
  }
}
