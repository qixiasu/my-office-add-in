/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

var {
  buildColRange,
  buildIndexMatchFormula,
  staticLookup,
  parseRangeAddress,
} = require("../utils/vlookup-utils");

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

  matchColumnSelect.innerHTML = "<option value=\"\">-- 选择匹配列 --</option>";
  returnColumnsDiv.innerHTML = "<span class=\"vlookup-label\">-- 请先选择查找表并读取表头 --</span>";

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
      getColumnLetter(parsed.startCol) +
        startRow +
        ":" +
        getColumnLetter(parsed.endCol) +
        endRow
    );
    dataRange.load("values");
    await context.sync();

    g_lookupTableData = dataRange.values;

    refreshColumns();
    updateHeaderLabels(headers, parsed);
    showPreview(headers, g_lookupTableData, headerRow);

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
    var headerName = headers[i - 1] || ("列" + i);
    options[i].textContent = headerName + " (列" + i + ")";
  }

  var checkboxItems = returnColumnsDiv.querySelectorAll(".vlookup-checkbox-item");
  for (var j = 0; j < checkboxItems.length; j++) {
    var label = checkboxItems[j].querySelector("label");
    var headerName = headers[j] || ("列" + (j + 1));
    label.textContent = headerName;
  }
}

function showPreview(headers, data, headerRow) {
  var previewSection = document.getElementById("vlookupPreview");
  var previewHead = document.getElementById("previewHead");
  var previewBody = document.getElementById("previewBody");

  previewHead.innerHTML = "";
  previewBody.innerHTML = "";

  var headerTr = document.createElement("tr");
  for (var i = 0; i < headers.length; i++) {
    var th = document.createElement("th");
    th.textContent = headers[i] || ("列" + (i + 1));
    headerTr.appendChild(th);
  }
  previewHead.appendChild(headerTr);

  var maxRows = Math.min(5, data.length);
  for (var rowIdx = 0; rowIdx < maxRows; rowIdx++) {
    var tr = document.createElement("tr");
    for (var colIdx = 0; colIdx < headers.length; colIdx++) {
      var td = document.createElement("td");
      td.textContent = data[rowIdx][colIdx] !== undefined ? data[rowIdx][colIdx] : "";
      tr.appendChild(td);
    }
    previewBody.appendChild(tr);
  }

  previewSection.classList.add("visible");
}

function getCheckedReturnCols() {
  var checkboxes = document.querySelectorAll("#returnColumns input[type=\"checkbox\"]:checked");
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

  var isValid =
    tableInput &&
    valueInput &&
    matchCol !== "" &&
    returnCols.length > 0;

  document.getElementById("executeBtn").disabled = !isValid;
}

function executeLookup() {
  var config = {
    lookupValue: document.getElementById("lookupValue").value,
    lookupTable: document.getElementById("lookupTable").value,
    headerRow: parseInt(document.getElementById("headerRow").value, 10) || 1,
    matchColIndex: parseInt(document.getElementById("matchColumn").value, 10),
    returnColIndices: getCheckedReturnCols(),
    matchMode: document.querySelector("input[name=\"matchMode\"]:checked").value === "exact" ? 0 : 1,
    outputType: document.querySelector("input[name=\"outputType\"]:checked").value,
  };

  performLookup(config);
}

function performLookup(config) {
  setStatus("处理中...", "info");

  Excel.run(async function (context) {
    var worksheet = context.workbook.worksheets.getActiveWorksheet();

    var lvParsed = parseRangeAddress(config.lookupValue);
    var dataStartRow = lvParsed.startRow;
    var dataStartCol = lvParsed.startCol;
    var dataRowCount = lvParsed.rowCount;
    var dataColLetter = getColumnLetter(dataStartCol);

    var ltParsed = parseRangeAddress(config.lookupTable);

    var lookupColRange = buildColRange(ltParsed, config.matchColIndex);
    var returnColRanges = [];
    for (var i = 0; i < config.returnColIndices.length; i++) {
      returnColRanges.push(buildColRange(ltParsed, config.returnColIndices[i]));
    }

    if (config.outputType === "formula") {
      var insertPos = ltParsed.endCol + 1;
      var returnColCount = config.returnColIndices.length;

      for (var c = 0; c < returnColCount; c++) {
        var colLetter = getColumnLetter(insertPos + c);
        worksheet.getRange(colLetter + ":" + colLetter).insert(Excel.InsertShiftDirection.right);
      }
      await context.sync();

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

      setStatus(
        "完成! 已在 " + returnColCount + " 列写入 " + dataRowCount + " 行 INDEX/MATCH 公式",
        "success"
      );
    } else {
      var dataRange = worksheet.getRange(
        dataColLetter + dataStartRow + ":" + dataColLetter + (dataStartRow + dataRowCount - 1)
      );
      dataRange.load("values");
      await context.sync();

      var lookupValues = [];
      for (var j = 0; j < dataRange.values.length; j++) {
        lookupValues.push(dataRange.values[j][0]);
      }

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

      var results = staticLookup(
        lookupValues,
        g_lookupTableData,
        config.matchColIndex,
        config.returnColIndices,
        config.matchMode
      );

      var insertPos2 = ltParsed.endCol + 1;
      var returnColCount2 = config.returnColIndices.length;

      for (var c2 = 0; c2 < returnColCount2; c2++) {
        var colLetter2 = getColumnLetter(insertPos2 + c2);
        worksheet.getRange(colLetter2 + ":" + colLetter2).insert(Excel.InsertShiftDirection.right);
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

      setStatus(
        "完成! 已写入 " + results.length + " 行 x " + returnColCount2 + " 列静态值",
        "success"
      );
    }
  }).catch(function (error) {
    setStatus("错误: " + error.message, "error");
  });
}

function setAllCheckboxes(containerId, checked) {
  var container = document.getElementById(containerId);
  var checkboxes = container.querySelectorAll("input[type=\"checkbox\"]");
  for (var i = 0; i < checkboxes.length; i++) {
    checkboxes[i].checked = checked;
  }
}
