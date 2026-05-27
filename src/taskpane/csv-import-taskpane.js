/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

var MAX_ROWS = 1050000;

var csvUtils = require("../utils/csv-utils");

var currentStep = 1;
var selectedFile = null;
var parsedData = null;
var rowCount = 0;
var colCount = 0;

Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    var fileInput = document.getElementById("fileInput");
    var delimiterInput = document.getElementById("delimiterInput");
    var prevBtn = document.getElementById("prevBtn");
    var nextBtn = document.getElementById("nextBtn");

    fileInput.addEventListener("change", onFileSelected);
    delimiterInput.addEventListener("input", onDelimiterChanged);
    prevBtn.addEventListener("click", onPrev);
    nextBtn.addEventListener("click", onNext);
  }
});

function setStatus(message, type) {
  var el = document.getElementById("status");
  el.textContent = message;
  el.className = "status-message status-" + (type || "idle");
}

function updateStepper() {
  for (var i = 1; i <= 3; i++) {
    var step = document.getElementById("step" + i);
    var panel = document.getElementById("panel" + i);
    step.classList.remove("active", "done");
    if (i < currentStep) {
      step.classList.add("done");
    } else if (i === currentStep) {
      step.classList.add("active");
    }
    panel.style.display = i === currentStep ? "block" : "none";
  }
  document.getElementById("prevBtn").disabled = currentStep === 1;
  updateNextButton();
}

function updateNextButton() {
  var nextBtn = document.getElementById("nextBtn");
  if (currentStep === 1) {
    nextBtn.textContent = "下一步";
    nextBtn.disabled = !selectedFile;
  } else if (currentStep === 2) {
    nextBtn.textContent = "下一步";
    nextBtn.disabled = false;
  } else if (currentStep === 3) {
    nextBtn.textContent = "导入";
    nextBtn.disabled = false;
  }
}

function showPanel(step) {
  currentStep = step;
  updateStepper();
}

function onFileSelected(e) {
  var file = e.target.files[0];
  if (!file) return;

  selectedFile = file;
  document.getElementById("fileName").textContent = file.name;
  updateNextButton();

  var reader = new FileReader();
  reader.onload = function (event) {
    var delimiter = document.getElementById("delimiterInput").value || ",";
    try {
      parsedData = csvUtils.parseCSV(event.target.result, delimiter);
      rowCount = parsedData.length;
      colCount = parsedData.length > 0 ? parsedData[0].length : 0;
    } catch (err) {
      setStatus("错误: 文件读取失败", "error");
      selectedFile = null;
      document.getElementById("fileName").textContent = "";
      updateNextButton();
      return;
    }
    if (rowCount === 0) {
      setStatus("错误: 文件为空", "error");
      selectedFile = null;
      document.getElementById("fileName").textContent = "";
      updateNextButton();
      return;
    }
    showPanel(2);
    setStatus("状态：等待操作...", "idle");
  };
  reader.onerror = function () {
    setStatus("错误: 文件读取失败", "error");
    selectedFile = null;
    document.getElementById("fileName").textContent = "";
    updateNextButton();
  };
  reader.readAsText(file);
}

function onDelimiterChanged() {
  if (currentStep === 2 && selectedFile) {
    var fileInput = document.getElementById("fileInput");
    if (fileInput.files[0]) {
      var reader = new FileReader();
      reader.onload = function (event) {
        var delimiter = document.getElementById("delimiterInput").value || ",";
        try {
          parsedData = csvUtils.parseCSV(event.target.result, delimiter);
          rowCount = parsedData.length;
          colCount = parsedData.length > 0 ? parsedData[0].length : 0;
        } catch (err) {
          // silent — user will see issue on step 3
        }
      };
      reader.readAsText(fileInput.files[0]);
    }
  }
}

function onPrev() {
  if (currentStep === 2) {
    showPanel(1);
  } else if (currentStep === 3) {
    showPanel(2);
  }
}

function onNext() {
  if (currentStep === 1) {
    showPanel(2);
  } else if (currentStep === 2) {
    updateSummary();
    showPanel(3);
  } else if (currentStep === 3) {
    doImport();
  }
}

function updateSummary() {
  var fileName = selectedFile ? selectedFile.name : "";
  var delimiter = document.getElementById("delimiterInput").value || ",";
  var delimiterLabel = delimiter === "\t" ? "Tab" : delimiter;
  var text = fileName + "，共 " + rowCount + " 行 × " + colCount + " 列，分隔符：【" + delimiterLabel + "】";
  document.getElementById("summaryText").textContent = text;
}

function doImport() {
  var nextBtn = document.getElementById("nextBtn");
  nextBtn.disabled = true;
  setStatus("处理中...", "loading");

  if (rowCount > MAX_ROWS) {
    setStatus("错误: 数据量过大（" + rowCount + " 行），单次最多支持 " + MAX_ROWS + " 行", "error");
    nextBtn.disabled = false;
    return;
  }

  Excel.run(function (context) {
    var worksheet = context.workbook.worksheets.getActiveWorksheet();
    var startCell = worksheet.getRange("A1");
    var endCell = worksheet.getRange(getColumnLetter(colCount - 1) + rowCount);
    var targetRange = worksheet.getRange(startCell.address + ":" + endCell.address);
    targetRange.values = parsedData;
    return context.sync().then(function () {
      setStatus("完成! 已写入 " + rowCount + " 行 × " + colCount + " 列", "success");
      nextBtn.disabled = false;
    });
  }).catch(function (error) {
    setStatus("错误: " + error.message, "error");
    nextBtn.disabled = false;
  });
}

function getColumnLetter(colIndex) {
  var letter = "";
  var remaining = colIndex;
  do {
    letter = String.fromCharCode(65 + (remaining % 26)) + letter;
    remaining = Math.floor(remaining / 26) - 1;
  } while (remaining >= 0);
  return letter;
}