/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

var { staticLookup, parseRangeAddress, buildLookupIndex } = require("../utils/vlookup-utils");

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

    // 获取查找表所在的工作表
    var tableSheetName = parsed.sheet;
    var tableSheet;
    if (tableSheetName) {
      tableSheet = context.workbook.worksheets.getItem(tableSheetName);
    } else {
      tableSheet = context.workbook.worksheets.getActiveWorksheet();
    }
    tableSheet.load("name");
    await context.sync();
    console.log("[DEBUG] loadTableHeaders: 表头读取所在工作表:", tableSheet.name);

    var startRow = parsed.startRow;
    var endRow = parsed.endRow;

    if (headerRow < startRow || headerRow > endRow) {
      setStatus("表头行号超出范围", "error");
      return;
    }

    var headerRange = tableSheet.getRange(
      getColumnLetter(parsed.startCol) +
        headerRow +
        ":" +
        getColumnLetter(parsed.endCol) +
        headerRow
    );
    headerRange.load("values");
    await context.sync();

    var headers = headerRange.values[0] || [];

    // 读取查找表数据（整列实际数据，不是解析出来的范围）
    // 使用 headerRow 作为起始行，读取到工作表实际最后一行
    var usedRange = tableSheet.getUsedRange();
    usedRange.load("rowCount");
    await context.sync();

    var actualLastRow = usedRange.rowCount;
    console.log("[DEBUG] loadTableHeaders: usedRange rowCount =", actualLastRow);

    // 如果 endRow 小于实际最后一行，用实际最后一行
    var dataEndRow = endRow < actualLastRow ? actualLastRow : endRow;
    console.log("[DEBUG] loadTableHeaders: dataEndRow =", dataEndRow);

    // 判断查找表范围是否是整列（endRow === 1 表示可能是整列选择）
    var isFullColumn = endRow === 1 && actualLastRow > 1;
    if (isFullColumn) {
      dataEndRow = actualLastRow;
      console.log("[DEBUG] loadTableHeaders: 检测到整列选择，使用实际最后一行", dataEndRow);
    }

    // 如果 headerRow 在所选范围内，跳过 headerRow 这一行（不作为数据）
    var dataReadStartRow = startRow;
    if (headerRow >= startRow && headerRow < dataEndRow) {
      dataReadStartRow = headerRow + 1;
      console.log(
        "[DEBUG] loadTableHeaders: headerRow",
        headerRow,
        "在范围内，跳过表头行，从",
        dataReadStartRow,
        "开始读取"
      );
    }

    var dataRange = tableSheet.getRange(
      getColumnLetter(parsed.startCol) +
        dataReadStartRow +
        ":" +
        getColumnLetter(parsed.endCol) +
        dataEndRow
    );
    dataRange.load("values");
    await context.sync();

    g_lookupTableData = dataRange.values;
    console.log(
      "[DEBUG] loadTableHeaders: g_lookupTableData 行数=",
      g_lookupTableData.length,
      "列数=",
      g_lookupTableData[0] ? g_lookupTableData[0].length : 0
    );

    refreshColumns();
    updateHeaderLabels(headers, parsed);

    setStatus(
      "表头读取成功，共 " + parsed.colCount + " 列，" + g_lookupTableData.length + " 行数据",
      "success"
    );
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

function updateProgressUI(percent, completed, total) {
  var statusEl = document.getElementById("progressStatus");
  var detailEl = document.getElementById("progressDetail");
  var barEl = document.getElementById("progressBarFill");

  if (statusEl) {
    statusEl.textContent = "处理中... " + percent + "%";
    statusEl.className = "status-message status-loading";
  }

  if (barEl) {
    barEl.style.width = percent + "%";
  }

  if (detailEl && total > 0) {
    detailEl.textContent = "已完成 " + completed + " / " + total + " 行";
  }
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
  var executeBtn = document.getElementById("executeBtn");
  var progressContainer = document.getElementById("progressContainer");
  var statusEl = document.getElementById("statusMessage");
  var progressInterval = null;

  // 禁用按钮，显示进度条
  executeBtn.disabled = true;
  if (progressContainer) progressContainer.style.display = "block";
  statusEl.style.display = "none";

  Excel.run(async function (context) {
    var BATCH_SIZE = 10000;
    var LARGE_DATA_THRESHOLD = 100000;

    // 获取查找值区域所在的工作表（这是写入目标）
    var lvParsed = parseRangeAddress(config.lookupValue);
    var lvSheetName = lvParsed.sheet;
    var lvWorksheet;
    if (lvSheetName) {
      lvWorksheet = context.workbook.worksheets.getItem(lvSheetName);
    } else {
      lvWorksheet = context.workbook.worksheets.getActiveWorksheet();
    }
    lvWorksheet.load("name");
    await context.sync();
    console.log("[DEBUG] 查找值工作表:", lvWorksheet.name);

    // 获取查找表所在的工作表（用于读取查找表数据）
    var ltParsed = parseRangeAddress(config.lookupTable);
    var ltSheetName = ltParsed.sheet;
    var ltWorksheet;
    if (ltSheetName) {
      ltWorksheet = context.workbook.worksheets.getItem(ltSheetName);
    } else {
      ltWorksheet = context.workbook.worksheets.getActiveWorksheet();
    }
    ltWorksheet.load("name");
    await context.sync();
    console.log("[DEBUG] 查找表工作表:", ltWorksheet.name);

    // 使用 lvWorksheet 作为目标工作表（写入结果的位置）
    var worksheet = lvWorksheet;

    console.log("[DEBUG] lookupValue 地址:", config.lookupValue);
    console.log("[DEBUG] lvParsed:", JSON.stringify(lvParsed));
    var dataStartRow = lvParsed.startRow;
    var dataStartCol = lvParsed.startCol;
    var dataRowCount = lvParsed.rowCount;
    var dataColLetter = getColumnLetter(dataStartCol);
    console.log(
      "[DEBUG] dataStartRow:",
      dataStartRow,
      "dataStartCol:",
      dataStartCol,
      "dataRowCount:",
      dataRowCount,
      "dataColLetter:",
      dataColLetter
    );

    ltParsed = parseRangeAddress(config.lookupTable);
    console.log("[DEBUG] lookupTable 地址:", config.lookupTable);
    console.log("[DEBUG] ltParsed:", JSON.stringify(ltParsed));

    if (!g_lookupTableData) {
      console.log("[DEBUG] g_lookupTableData 为空，从工作表加载");
      // 从查找表所在工作表读取数据
      var tableRange = ltWorksheet.getRange(
        getColumnLetter(ltParsed.startCol) +
          ltParsed.startRow +
          ":" +
          getColumnLetter(ltParsed.endCol) +
          ltParsed.endRow
      );
      tableRange.load("values");
      await context.sync();
      g_lookupTableData = tableRange.values;
      console.log(
        "[DEBUG] g_lookupTableData 已加载，行数:",
        g_lookupTableData.length,
        "列数:",
        g_lookupTableData[0] ? g_lookupTableData[0].length : 0
      );
    } else {
      console.log("[DEBUG] g_lookupTableData 已存在，行数:", g_lookupTableData.length);
    }

    if (dataRowCount < LARGE_DATA_THRESHOLD) {
      // 小数据模式：启动进度条模拟
      var totalRows = dataRowCount;
      var progressStep = 0;
      progressInterval = setInterval(function() {
        progressStep = Math.min(progressStep + 10, 90);
        var percent = progressStep;
        updateProgressUI(percent, Math.round(totalRows * percent / 100), totalRows);
      }, 200);

      // Small data: single read -> staticLookup -> single write
      // 整列选择时 endRow === 1，需要用实际最后一行
      var lvIsFullColumn = lvParsed.endRow === 1;
      var lvActualEndRow = lvParsed.endRow;

      if (lvIsFullColumn) {
        var lvUsedRange = worksheet.getUsedRange();
        lvUsedRange.load("rowCount");
        await context.sync();
        lvActualEndRow = lvUsedRange.rowCount;
        console.log("[DEBUG] lookupValue 整列选择，实际最后一行:", lvActualEndRow);
      }

      var dataRange = worksheet.getRange(
        dataColLetter + lvParsed.startRow + ":" + dataColLetter + lvActualEndRow
      );
      dataRange.load("values");
      await context.sync();

      console.log("[DEBUG] dataRange.values 行数:", dataRange.values.length);

      // 跳过表头行：如果 lookupValue 从 headerRow 开始，第一行是表头不参与查找
      var lvDataStartRow = lvParsed.startRow;
      var lvFirstRowIsHeader = lvParsed.startRow === config.headerRow;
      if (lvFirstRowIsHeader) {
        lvDataStartRow = lvParsed.startRow + 1;
        console.log("[DEBUG] lookupValue 第一行是表头，跳过，从行", lvDataStartRow, "开始");
      }

      // Build lookup index once for reuse
      var lookupIndex = null;
      if (config.matchMode === 0) {
        lookupIndex = buildLookupIndex(g_lookupTableData, config.matchColIndex);
      }

      var lookupValues = [];
      var lookupValuesStartRow = lvDataStartRow;
      for (var j = lvDataStartRow - lvParsed.startRow; j < dataRange.values.length; j++) {
        lookupValues.push(dataRange.values[j][0]);
      }
      console.log("[DEBUG] lookupValues (跳过表头后):", JSON.stringify(lookupValues));

      var results = staticLookup(
        lookupValues,
        g_lookupTableData,
        config.matchColIndex,
        config.returnColIndices,
        config.matchMode,
        config.defaultValue,
        lookupIndex
      );
      console.log(
        "[DEBUG] staticLookup 结果, results.length:",
        results.length,
        "results[0]?.length:",
        results[0] ? results[0].length : 0
      );
      console.log("[DEBUG] 前3条结果:", JSON.stringify(results.slice(0, 3)));

      // 修复：写入位置应该是查找值区域的右边，而不是查找表的右边
      var insertPos = lvParsed.endCol + 1;
      console.log(
        "[DEBUG] 写入位置 insertPos:",
        insertPos,
        "= getColumnLetter(",
        getColumnLetter(insertPos),
        ")"
      );
      console.log(
        "[DEBUG] 注意：旧代码使用 ltParsed.endCol + 1 =",
        ltParsed.endCol + 1,
        ", 现改为 lvParsed.endCol + 1 =",
        insertPos
      );
      var returnColCount = config.returnColIndices.length;

      for (var c = 0; c < returnColCount; c++) {
        var colLetter = getColumnLetter(insertPos + c);
        worksheet.getRange(colLetter + ":" + colLetter).insert(Excel.InsertShiftDirection.right);
      }
      // 激活查找值区域所在的工作表，让用户能看到写入过程
      lvWorksheet.activate();
      await context.sync();

      // 写入表头（从查找表各返回列的表头复制过来）
      var headerValues = [];
      for (var hc = 0; hc < returnColCount; hc++) {
        var returnColIdx = config.returnColIndices[hc];
        // 从查找表所在工作表读取表头
        var headerCell = ltWorksheet.getRange(
          getColumnLetter(ltParsed.startCol + returnColIdx) +
            config.headerRow +
            ":" +
            getColumnLetter(ltParsed.startCol + returnColIdx) +
            config.headerRow
        );
        headerCell.load("values");
        await context.sync();
        headerValues.push(headerCell.values[0][0]);
      }
      if (lvFirstRowIsHeader) {
        // 写入表头行
        var headerRange = worksheet.getRange(
          getColumnLetter(insertPos) +
            config.headerRow +
            ":" +
            getColumnLetter(insertPos + returnColCount - 1) +
            config.headerRow
        );
        headerRange.values = [headerValues];
        await context.sync();
        console.log("[DEBUG] 已写入表头:", JSON.stringify(headerValues));
      }

      var targetRange = worksheet.getRange(
        getColumnLetter(insertPos) +
          lookupValuesStartRow +
          ":" +
          getColumnLetter(insertPos + returnColCount - 1) +
          (lookupValuesStartRow + results.length - 1)
      );
      targetRange.values = results;
      await context.sync();

      setStatus(
        "完成! 已写入 " + results.length + " 行 x " + returnColCount + " 列静态值",
        "success"
      );
      clearInterval(progressInterval);
      updateProgressUI(100, results.length, results.length);
    } else {
      // Large data: batch processing
      var totalRows = dataRowCount;
      var processedRows = 0;
      // 修复：写入位置应该是查找值区域的右边，而不是查找表的右边
      var insertPos = lvParsed.endCol + 1;
      console.log("[DEBUG] 大数据模式 insertPos:", insertPos);
      var returnColCount = config.returnColIndices.length;

      // Insert all return columns first (once, before batches)
      for (var ci = 0; ci < returnColCount; ci++) {
        var colLetter = getColumnLetter(insertPos + ci);
        worksheet.getRange(colLetter + ":" + colLetter).insert(Excel.InsertShiftDirection.right);
      }
      // 激活查找值区域所在的工作表，让用户能看到写入过程
      lvWorksheet.activate();
      await context.sync();

      // Build lookup index once for all batches
      var lookupIndex = null;
      if (config.matchMode === 0) {
        lookupIndex = buildLookupIndex(g_lookupTableData, config.matchColIndex);
      }

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
          config.defaultValue,
          lookupIndex
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
        var percent = Math.round((processedRows / totalRows) * 100);
        updateProgressUI(percent, processedRows, totalRows);
      }

      setStatus("完成! 已写入 " + totalRows + " 行 x " + returnColCount + " 列静态值", "success");
    }
  }).catch(function (error) {
    setStatus("错误: " + error.message, "error");
  }).finally(function() {
    clearInterval(progressInterval);
    executeBtn.disabled = false;
    if (progressContainer) progressContainer.style.display = "none";
    statusEl.style.display = "block";
  });
}

function setAllCheckboxes(containerId, checked) {
  var container = document.getElementById(containerId);
  var checkboxes = container.querySelectorAll('input[type="checkbox"]');
  for (var i = 0; i < checkboxes.length; i++) {
    checkboxes[i].checked = checked;
  }
}
