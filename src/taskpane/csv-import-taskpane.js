/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

var MAX_ROWS = 1050000;
var BATCH_SIZE = 10000; // Rows per batch for chunked import

var csvUtils = require("../utils/csv-utils");

var currentStep = 1;
var selectedFile = null;
var selectedEncoding = "utf-8";
var rowCount = 0;
var colCount = 0;

function getEncoding() {
  var sel = document.getElementById("encodingSelect");
  return sel ? sel.value : "utf-8";
}

Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    var fileInput = document.getElementById("fileInput");
    var delimiterInput = document.getElementById("delimiterInput");
    var prevBtn = document.getElementById("prevBtn");
    var nextBtn = document.getElementById("nextBtn");
    var selectFilesBtn = document.getElementById("selectFilesBtn");

    selectFilesBtn.addEventListener("click", function() {
      fileInput.click();
    });
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

// ── Step 1: file selection ────────────────────────────────────────────

function onFileSelected(e) {
  var file = e.target.files[0];
  if (!file) return;

  selectedFile = file;
  document.getElementById("fileName").textContent = file.name;
  updateNextButton();

  setStatus("正在分析文件...", "loading");

  var delimiter = document.getElementById("delimiterInput").value || ",";
  selectedEncoding = getEncoding();

  // Stream‑scan the file in 1-MB chunks — NEVER loads full file into memory
  // Note: scanning uses default encoding; re-scan will happen in step 2 if encoding changes
  csvUtils.scanCSVFile(
    file,
    delimiter,
    "utf-8", // always scan with utf-8 first; re-scan when user confirms encoding
    function done(stats) {
      rowCount = stats.rowCount;
      colCount = stats.colCount;

      if (rowCount === 0) {
        setStatus("错误: 文件为空", "error");
        selectedFile = null;
        rowCount = 0;
        colCount = 0;
        document.getElementById("fileName").textContent = "";
        updateNextButton();
        return;
      }

      // Don't auto-advance — let user confirm encoding in step 1 first
      setStatus("已读取 " + rowCount + " 行，" + colCount + " 列，请选择编码后点击下一步", "success");
      updateNextButton();
    },
    function error(err) {
      setStatus("错误: 文件读取失败", "error");
      selectedFile = null;
      document.getElementById("fileName").textContent = "";
      updateNextButton();
    }
  );
}

// ── Delimiter change (re‑scan from file, still chunked) ───────────────

function onDelimiterChanged() {
  if (currentStep !== 2 || !selectedFile) return;

  var delimiter = document.getElementById("delimiterInput").value || ",";

  csvUtils.scanCSVFile(
    selectedFile,
    delimiter,
    selectedEncoding,
    function done(stats) {
      rowCount = stats.rowCount;
      colCount = stats.colCount;
      updateSummary();
    },
    function error() {
      // silent — user will see the issue on step 3
    }
  );
}

// ── Navigation ────────────────────────────────────────────────────────

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
  var text =
    fileName + "，共 " + rowCount + " 行 × " + colCount + " 列，分隔符：【" + delimiterLabel + "】";
  document.getElementById("summaryText").textContent = text;
}

// ── Step 3: import ────────────────────────────────────────────────────

function doImport() {
  var nextBtn = document.getElementById("nextBtn");
  nextBtn.disabled = true;
  setStatus("处理中...", "loading");

  if (rowCount > MAX_ROWS) {
    setStatus("错误: 数据量过大（" + rowCount + " 行），单次最多支持 " + MAX_ROWS + " 行", "error");
    nextBtn.disabled = false;
    return;
  }

  var delimiter = document.getElementById("delimiterInput").value || ",";
  var totalBatches = Math.ceil(rowCount / BATCH_SIZE);
  var processedBatches = 0;

  // Write a single batch to Excel
  function writeBatch(batchRows, batchIndex) {
    return Excel.run(function (context) {
      var worksheet = context.workbook.worksheets.getActiveWorksheet();
      var startRow = batchIndex * BATCH_SIZE + 1;
      var endRow = startRow + batchRows.length - 1;
      var endColLetter = getColumnLetter(colCount - 1);

      var targetRange = worksheet.getRange(
        "A" + startRow + ":" + endColLetter + endRow
      );
      targetRange.values = batchRows;

      return context.sync().then(function () {
        processedBatches++;
        setStatus(
          "写入中... " + processedBatches + "/" + totalBatches + " 批",
          "loading"
        );
      });
    });
  }

  // Stream‑parse the file and build a sequential promise‑chain for Excel writes
  var chain = Promise.resolve();

  csvUtils.parseCSVFile(
    selectedFile,
    delimiter,
    BATCH_SIZE,
    selectedEncoding,
    function onBatch(batchRows, batchIndex) {
      chain = chain.then(function () {
        return writeBatch(batchRows, batchIndex);
      });
    },
    function onDone(stats) {
      chain.then(function () {
        setStatus("完成! 已写入 " + stats.rowCount + " 行 × " + stats.colCount + " 列", "success");
        nextBtn.disabled = false;
      }).catch(function (error) {
        setStatus("错误: " + (error && error.message ? error.message : error), "error");
        nextBtn.disabled = false;
      });
    },
    function onError(err) {
      setStatus("错误: " + (err && err.message ? err.message : "文件读取失败"), "error");
      nextBtn.disabled = false;
    }
  );
}

// ── Helpers ───────────────────────────────────────────────────────────

function getColumnLetter(colIndex) {
  var letter = "";
  var remaining = colIndex;
  do {
    letter = String.fromCharCode(65 + (remaining % 26)) + letter;
    remaining = Math.floor(remaining / 26) - 1;
  } while (remaining >= 0);
  return letter;
}
