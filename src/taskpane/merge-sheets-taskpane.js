/* global console, document, Excel, Office */

var mergeSheetsUtils = require("../utils/merge-sheets-utils");

var pendingConfirmation = null;
var allSheets = [];

function initOffice() {
  Office.onReady(function (info) {
    if (info.host === Office.HostType.Excel) {
      initEventListeners();
      loadSheetList();
    }
  });
}

if (typeof Office !== "undefined") {
  initOffice();
} else {
  // Wait for Office.js to load from CDN (handles HMR re-evaluation timing)
  var checkOffice = setInterval(function () {
    if (typeof Office !== "undefined") {
      clearInterval(checkOffice);
      initOffice();
    }
  }, 50);
  // Timeout after 10 seconds
  setTimeout(function () {
    clearInterval(checkOffice);
  }, 10000);
}

function initEventListeners() {
  document.getElementById("selectAllBtn").onclick = function () {
    setAllCheckboxes(true);
  };

  document.getElementById("deselectAllBtn").onclick = function () {
    setAllCheckboxes(false);
  };

  document.getElementById("executeBtn").onclick = executeMerge;

  document.getElementById("confirmOverwrite").onclick = function () {
    handleConfirm("overwrite");
  };

  document.getElementById("confirmRename").onclick = function () {
    handleConfirm("rename");
  };

  document.getElementById("confirmCancel").onclick = function () {
    handleConfirm("cancel");
  };
}

function setAllCheckboxes(checked) {
  var checkboxes = document.querySelectorAll(".merge-sheet-checkbox");
  checkboxes.forEach(function (cb) {
    cb.checked = checked;
  });
}

function loadSheetList() {
  Excel.run(async function (context) {
    var worksheets = context.workbook.worksheets;
    worksheets.load("items/name");
    await context.sync();

    allSheets = worksheets.items.map(function (ws) {
      return { name: ws.name, checked: false, headerRow: 1 };
    });

    renderSheetList();
  }).catch(function (error) {
    setStatus("加载工作表失败: " + error.message, "error");
  });
}

function renderSheetList() {
  var listEl = document.getElementById("sheetList");

  if (allSheets.length === 0) {
    listEl.innerHTML = '<div class="merge-empty">未找到工作表</div>';
    return;
  }

  var html = "";
  allSheets.forEach(function (sheet, index) {
    html +=
      '<div class="merge-sheet-item">' +
      '<input type="checkbox" class="merge-sheet-checkbox" data-index="' +
      index +
      '" />' +
      '<span class="merge-sheet-name">' +
      escapeHtml(sheet.name) +
      "</span>" +
      '<input type="number" class="merge-sheet-header-input" data-index="' +
      index +
      '" value="1" min="0" />' +
      "</div>";
  });

  listEl.innerHTML = html;

  // Bind checkbox and input events
  var checkboxes = listEl.querySelectorAll(".merge-sheet-checkbox");
  checkboxes.forEach(function (cb) {
    cb.onchange = function () {
      var idx = parseInt(cb.dataset.index, 10);
      allSheets[idx].checked = cb.checked;
    };
  });

  var inputs = listEl.querySelectorAll(".merge-sheet-header-input");
  inputs.forEach(function (input) {
    input.onchange = function () {
      var idx = parseInt(input.dataset.index, 10);
      allSheets[idx].headerRow = parseInt(input.value, 10) || 0;
    };
  });
}

function escapeHtml(text) {
  var div = document.createElement("div");
  div.textContent = text;
  return div.innerHTML;
}

function setStatus(message, type) {
  var el = document.getElementById("status");
  el.textContent = message;
  el.className = "merge-status merge-status--" + (type || "idle");
}

function showConfirmBox(message) {
  var confirmBox = document.getElementById("confirmBox");
  var confirmMsgEl = document.getElementById("confirmMsg");
  confirmMsgEl.textContent = message;
  confirmBox.style.display = "block";
  setStatus("状态：等待确认...", "idle");

  return new Promise(function (resolve) {
    pendingConfirmation = { resolve: resolve };
  });
}

function hideConfirmBox() {
  document.getElementById("confirmBox").style.display = "none";
  pendingConfirmation = null;
}

function handleConfirm(action) {
  hideConfirmBox();
  if (pendingConfirmation) {
    pendingConfirmation.resolve(action);
    pendingConfirmation = null;
  }
}

function getSelectedSheets() {
  return allSheets.filter(function (sheet) {
    return sheet.checked;
  });
}

function executeMerge() {
  var executeBtn = document.getElementById("executeBtn");
  executeBtn.disabled = true;
  setStatus("处理中...", "loading");

  var selectedSheets = getSelectedSheets();

  if (selectedSheets.length < 2) {
    executeBtn.disabled = false;
    setStatus("请至少选择两个 Sheet 进行合并", "error");
    return;
  }

  Excel.run(async function (context) {
    // Step 1: Read column count for each selected sheet
    var sheetsWithMeta = [];

    for (var i = 0; i < allSheets.length; i++) {
      if (allSheets[i].checked) {
        var ws = context.workbook.worksheets.getItem(allSheets[i].name);
        var usedRange = ws.getUsedRange();
        usedRange.load("columnCount");
        await context.sync();

        sheetsWithMeta.push({
          name: allSheets[i].name,
          headerRow: allSheets[i].headerRow,
          columnCount: usedRange.columnCount
        });
      }
    }

    // Step 2: Validate column consistency
    var validation = mergeSheetsUtils.validateColumnConsistency(sheetsWithMeta);
    if (!validation.valid) {
      throw new Error(validation.error);
    }

    // Step 3: Read data from each sheet
    var mergedData = [];
    var hasHeader = false;

    for (var j = 0; j < sheetsWithMeta.length; j++) {
      var sheetInfo = sheetsWithMeta[j];
      var worksheet = context.workbook.worksheets.getItem(sheetInfo.name);
      var usedRange = worksheet.getUsedRange();
      usedRange.load(["values", "rowCount", "columnCount"]);
      await context.sync();

      var data = usedRange.values;
      if (!data || data.length === 0) {
        throw new Error("工作表 '" + sheetInfo.name + "' 无可用数据");
      }

      var headerRowIndex = sheetInfo.headerRow > 0 ? sheetInfo.headerRow - 1 : -1;

      // First sheet: include header if exists
      if (j === 0) {
        if (headerRowIndex >= 0 && headerRowIndex < data.length) {
          mergedData.push(data[headerRowIndex]);
          hasHeader = true;
        }
        // Add data rows after header
        for (var r = headerRowIndex + 1; r < data.length; r++) {
          mergedData.push(data[r]);
        }
      } else {
        // Other sheets: skip header, add all data rows
        var startRow = headerRowIndex >= 0 ? headerRowIndex + 1 : 0;
        for (var r2 = startRow; r2 < data.length; r2++) {
          mergedData.push(data[r2]);
        }
      }
    }

    // Step 4: Check if "合并结果" sheet exists
    var existingSheets = context.workbook.worksheets;
    var targetSheet = existingSheets.getItemOrNullObject("合并结果");
    targetSheet.load("name");
    await context.sync();

    var targetName = "合并结果";
    var sheetExists = targetSheet.name === targetName;

    if (sheetExists) {
      var confirmAction = await showConfirmBox(
        "工作表 '" + targetName + "' 已存在，选择操作："
      );

      if (confirmAction === "cancel") {
        executeBtn.disabled = false;
        setStatus("状态：已取消", "idle");
        return;
      }

      if (confirmAction === "rename") {
        var allSheetNames = existingSheets.items.map(function (ws) {
          return ws.name;
        });
        targetName = mergeSheetsUtils.generateUniqueSheetName("合并结果", allSheetNames);
      }
      // If overwrite, use same name
    }

    // Step 5: Create or clear target sheet and write data
    var finalSheet;
    finalSheet = context.workbook.worksheets.getItemOrNullObject(targetName);
    await context.sync();
    if (finalSheet.name === targetName) {
      finalSheet.delete();
      await context.sync();
    }
    finalSheet = context.workbook.worksheets.add(targetName);
    await context.sync();

    if (mergedData.length > 0) {
      var targetRange = finalSheet.getRange(
        "A1:" +
        mergeSheetsUtils.getColumnLetter(mergedData[0].length - 1) +
        mergedData.length
      );
      targetRange.values = mergedData;
      await context.sync();
    }

    var totalRows = mergedData.length;
    var headerInfo = hasHeader ? "（含表头）" : "（无表头）";
    setStatus(
      "完成! 已合并 " + selectedSheets.length + " 个工作表，共 " + totalRows + " 行数据" + headerInfo,
      "success"
    );
  }).catch(function (error) {
    setStatus("错误: " + error.message, "error");
  }).finally(function () {
    executeBtn.disabled = false;
  });
}