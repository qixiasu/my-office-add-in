/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

var { parseRangeAddress } = require("../utils/vlookup-utils");

var { getColumnLetter } = require("../utils/concat-utils");

var { groupDataByKey, truncateSheetName } = require("../utils/split-sheet-utils");

Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    initEventListeners();
    loadInitialSelection();
  }
});

function initEventListeners() {
  document.getElementById("refreshRange").onclick = function () {
    refreshSelection("dataRange");
  };

  document.getElementById("headerRow").addEventListener("input", function () {
    var dataRangeInput = document.getElementById("dataRange").value;
    if (dataRangeInput) {
      loadTableHeaders();
    }
  });

  document.getElementById("executeBtn").onclick = executeSplit;

  var form = document.getElementById("splitSheetForm");
  form.addEventListener("change", validateForm);
  form.addEventListener("input", validateForm);
}

function loadInitialSelection() {
  Excel.run(async function (context) {
    var range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();

    if (range.address) {
      document.getElementById("dataRange").value = range.address;
      refreshSelection("dataRange");
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

    if (target === "dataRange") {
      loadTableHeaders();
    }
  }).catch(function (error) {
    setStatus("刷新选择失败: " + error.message, "error");
  });
}

function loadTableHeaders() {
  var dataRangeInput = document.getElementById("dataRange").value;
  if (!dataRangeInput) {
    setStatus("请先选择数据区域", "error");
    return;
  }

  var headerRow = parseInt(document.getElementById("headerRow").value, 10) || 1;

  Excel.run(async function (context) {
    var parsed = parseRangeAddress(dataRangeInput);

    // Get the worksheet where the data range is
    var worksheet;
    if (parsed.sheet) {
      worksheet = context.workbook.worksheets.getItem(parsed.sheet);
    } else {
      worksheet = context.workbook.worksheets.getActiveWorksheet();
    }
    worksheet.load("name");
    await context.sync();

    // Read header row
    var headerRange = worksheet.getRange(
      getColumnLetter(parsed.startCol) +
        headerRow +
        ":" +
        getColumnLetter(parsed.endCol) +
        headerRow
    );
    headerRange.load("values");
    await context.sync();

    var headers = headerRange.values[0] || [];

    // Populate keyColumn dropdown
    var keyColumnSelect = document.getElementById("keyColumn");
    keyColumnSelect.innerHTML = '<option value="">-- 选择关键列 --</option>';

    for (var i = 0; i < headers.length; i++) {
      var option = document.createElement("option");
      option.value = i;
      option.textContent = (headers[i] || "列" + (i + 1)) + " (列" + (i + 1) + ")";
      keyColumnSelect.appendChild(option);
    }

    setStatus("表头读取成功，共 " + headers.length + " 列", "success");
    validateForm();
  }).catch(function (error) {
    setStatus("读取表头失败: " + error.message, "error");
  });
}

function setStatus(message, type) {
  var statusEl = document.getElementById("statusMessage");
  statusEl.textContent = "状态：" + message;
  statusEl.className = "split-sheet-status split-sheet-status--" + (type || "idle");
}

function updateProgressUI(percent, completed, total) {
  var statusEl = document.getElementById("progressStatus");
  var detailEl = document.getElementById("progressDetail");
  var barEl = document.getElementById("progressBarFill");

  if (statusEl) {
    statusEl.textContent = "处理中... " + percent + "%";
    statusEl.className = "split-sheet-status split-sheet-status--loading";
  }

  if (barEl) {
    barEl.style.width = percent + "%";
  }

  if (detailEl && total > 0) {
    detailEl.textContent = "已完成 " + completed + " / " + total + " 个工作表";
  }
}

function validateForm() {
  var dataRange = document.getElementById("dataRange").value;
  var keyColumn = document.getElementById("keyColumn").value;

  var isValid = dataRange && keyColumn !== "";

  document.getElementById("executeBtn").disabled = !isValid;
}

function executeSplit() {
  var config = {
    dataRange: document.getElementById("dataRange").value,
    headerRow: parseInt(document.getElementById("headerRow").value, 10) || 1,
    keyColumnIndex: parseInt(document.getElementById("keyColumn").value, 10),
  };

  performSplit(config);
}

function performSplit(config) {
  var executeBtn = document.getElementById("executeBtn");
  var progressContainer = document.getElementById("progressContainer");
  var statusEl = document.getElementById("statusMessage");

  // Disable button, show progress bar
  executeBtn.disabled = true;
  if (progressContainer) { progressContainer.style.display = "block"; }
  statusEl.style.display = "none";

  Excel.run(async function (context) {
    var parsed = parseRangeAddress(config.dataRange);

    // Get the worksheet where the data range is
    var worksheet;
    if (parsed.sheet) {
      worksheet = context.workbook.worksheets.getItem(parsed.sheet);
    } else {
      worksheet = context.workbook.worksheets.getActiveWorksheet();
    }
    worksheet.load("name");
    await context.sync();

    // Read entire data range
    var dataStartRow = config.headerRow;
    var dataRange = worksheet.getRange(
      getColumnLetter(parsed.startCol) +
        dataStartRow +
        ":" +
        getColumnLetter(parsed.endCol) +
        parsed.endRow
    );
    dataRange.load("values");
    await context.sync();

    var allData = dataRange.values;

    // Group data by key column
    var groupingResult = groupDataByKey(allData, config.keyColumnIndex);
    if (groupingResult.error) {
      throw new Error(groupingResult.error);
    }

    var groups = groupingResult.groups;
    var groupKeys = Object.keys(groups);
    var totalGroups = groupKeys.length;

    // Process each group
    for (var i = 0; i < groupKeys.length; i++) {
      var key = groupKeys[i];
      var group = groups[key];
      var sheetName = truncateSheetName(String(key));

      // Check if sheet already exists
      var existingSheets = context.workbook.worksheets;
      var sheetQuery = existingSheets.getItemOrNullObject(sheetName);
      sheetQuery.load("name");
      await context.sync();

      if (sheetQuery.name === sheetName) {
        throw new Error("工作表 '" + sheetName + "' 已存在，请先删除或重命名");
      }

      // Create new sheet
      var newSheet = context.workbook.worksheets.add(sheetName);
      newSheet.load("name");
      await context.sync();

      // Write header row first
      var headerTargetRange = newSheet.getRange(
        getColumnLetter(parsed.startCol) +
          "1" +
          ":" +
          getColumnLetter(parsed.startCol + group.header.length - 1) +
          "1"
      );
      headerTargetRange.values = [group.header];
      await context.sync();

      // Write data rows
      if (group.rows.length > 0) {
        var dataTargetRange = newSheet.getRange(
          getColumnLetter(parsed.startCol) +
            "2" +
            ":" +
            getColumnLetter(parsed.startCol + group.rows[0].length - 1) +
            (group.rows.length + 1)
        );
        dataTargetRange.values = group.rows;
        await context.sync();
      }

      // Update progress
      var percent = Math.round(((i + 1) / totalGroups) * 100);
      updateProgressUI(percent, i + 1, totalGroups);
    }

    setStatus("完成! 已创建 " + totalGroups + " 个工作表", "success");
  }).catch(function (error) {
    setStatus("错误: " + error.message, "error");
  }).finally(function () {
    executeBtn.disabled = false;
    if (progressContainer) { progressContainer.style.display = "none"; progressContainer.className = "split-sheet-progress"; }
    statusEl.style.display = "block";
  });
}