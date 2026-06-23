/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

var csvMergeUtils = require("../utils/csv-merge-utils");

// Selected files array
var selectedFiles = [];

function setStatus(message, type) {
  var el = document.getElementById("status");
  el.textContent = message;
  el.className = "status-message status-" + (type || "idle");
}

function updateFileCountText() {
  var el = document.getElementById("fileCountText");
  if (el) {
    el.textContent = "已选择 " + selectedFiles.length + " 个文件";
  }
}

function formatFileSize(bytes) {
  if (bytes < 1024) {
    return bytes + " B";
  } else if (bytes < 1024 * 1024) {
    return (bytes / 1024).toFixed(1) + " KB";
  } else {
    return (bytes / (1024 * 1024)).toFixed(1) + " MB";
  }
}

function renderFileList() {
  var fileListEl = document.getElementById("fileList");
  fileListEl.innerHTML = "";

  for (var i = 0; i < selectedFiles.length; i++) {
    var file = selectedFiles[i];
    var item = document.createElement("div");
    item.className = "cfm-file-item";

    var info = document.createElement("span");
    info.className = "cfm-file-name";
    info.textContent = file.name + " (" + formatFileSize(file.size) + ")";

    var removeBtn = document.createElement("button");
    removeBtn.className = "cfm-file-remove";
    removeBtn.textContent = "×";
    removeBtn.title = "移除文件";
    removeBtn.dataset.index = i;

    item.appendChild(info);
    item.appendChild(removeBtn);
    fileListEl.appendChild(item);
  }

  updateFileCountText();
}

function updateMergeBtn() {
  var mergeBtn = document.getElementById("mergeBtn");
  mergeBtn.disabled = selectedFiles.length === 0;
}

Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    var fileInput = document.getElementById("fileInput");
    var selectFilesBtn = document.getElementById("selectFilesBtn");
    var mergeBtn = document.getElementById("mergeBtn");
    var fileListEl = document.getElementById("fileList");

    selectFilesBtn.addEventListener("click", function() {
      fileInput.click();
    });
    fileInput.addEventListener("change", onFileInputChange);
    mergeBtn.addEventListener("click", onMergeClick);

    // Event delegation for remove buttons
    fileListEl.addEventListener("click", function (e) {
      var btn = e.target.closest(".cfm-file-remove");
      if (!btn) return;
      var index = parseInt(btn.dataset.index, 10);
      selectedFiles.splice(index, 1);
      renderFileList();
      updateMergeBtn();
    });
  }
});

function getHeaderMode() {
  var radios = document.getElementsByName("headerMode");
  for (var i = 0; i < radios.length; i++) {
    if (radios[i].checked) {
      return radios[i].value;
    }
  }
  return "keep-first";
}

function getOutputEncoding() {
  var sel = document.getElementById("outputEncoding");
  return sel ? sel.value : "utf-8";
}

function getOutputDelimiter() {
  var sel = document.getElementById("outputDelimiter");
  return sel ? sel.value : ",";
}

function onFileInputChange(e) {
  var files = e.target.files;
  for (var i = 0; i < files.length; i++) {
    selectedFiles.push(files[i]);
  }
  // Reset input so same file can be selected again if needed
  e.target.value = "";
  renderFileList();
  updateMergeBtn();
}

function onMergeClick() {
  var mergeBtn = document.getElementById("mergeBtn");
  mergeBtn.disabled = true;
  setStatus("处理中...", "loading");

  var headerMode = getHeaderMode();
  var encoding = getOutputEncoding();
  var delimiter = getOutputDelimiter();

  // Read all files sequentially
  var filesData = [];
  var emptyFileWarnings = [];

  function processFilesSequentially(index) {
    if (index >= selectedFiles.length) {
      // All files processed
      finishMerge(filesData, emptyFileWarnings, headerMode, encoding, delimiter, mergeBtn);
      return;
    }

    var file = selectedFiles[index];

    csvMergeUtils
      .readCSVFileFull(file, ",", encoding)
      .then(function (result) {
        if (result.data.length === 0) {
          emptyFileWarnings.push(file.name);
        } else {
          filesData.push(result);
        }
        processFilesSequentially(index + 1);
      })
      .catch(function () {
        emptyFileWarnings.push(file.name + " (读取失败)");
        processFilesSequentially(index + 1);
      });
  }

  processFilesSequentially(0);
}

function finishMerge(filesData, emptyFileWarnings, headerMode, encoding, delimiter, mergeBtn) {
  // Check if we have any data
  if (filesData.length === 0 || !hasDataInFiles(filesData)) {
    setStatus("错误: 所有文件为空或读取失败", "error");
    mergeBtn.disabled = false;
    return;
  }

  // Apply header mode
  var processedData = applyHeaderMode(filesData, headerMode);

  // If skip mode with no headers, take from first file with data
  if (headerMode === "skip" && (!processedData.headers || processedData.headers.length === 0)) {
    for (var i = 0; i < filesData.length; i++) {
      if (filesData[i].data.length > 0) {
        processedData.headers = filesData[i].headers;
        break;
      }
    }
  }

  // Align columns across all files
  var aligned = csvMergeUtils.alignColumns(processedData.filesData);

  // Check if we have any data to output
  if (aligned.data.length === 0) {
    setStatus("错误: 合并后无数据", "error");
    mergeBtn.disabled = false;
    return;
  }

  // Build filename
  var filename = "merged_" + getTimestamp() + ".csv";

  // Trigger download
  csvMergeUtils.buildCSVBlob(aligned.headers, aligned.data, encoding, delimiter, filename);

  // Show warnings for empty files
  var warningMsg =
    emptyFileWarnings.length > 0 ? "（已跳过空文件: " + emptyFileWarnings.join(", ") + "）" : "";

  setStatus(
    "完成! 已合并 " + filesData.length + " 个文件，共 " + aligned.data.length + " 行" + warningMsg,
    "success"
  );
  mergeBtn.disabled = false;
}

function hasDataInFiles(filesData) {
  for (var i = 0; i < filesData.length; i++) {
    if (filesData[i].data.length > 0) {
      return true;
    }
  }
  return false;
}

function applyHeaderMode(filesData, headerMode) {
  var result = {
    filesData: [],
    headers: null,
  };

  if (headerMode === "skip") {
    // Skip all headers - treat all data as regular data rows
    result.headers = []; // No header row in output
    for (var i = 0; i < filesData.length; i++) {
      result.filesData.push({
        headers: filesData[i].headers,
        data: filesData[i].data,
      });
    }
  } else if (headerMode === "keep-first") {
    // Keep first file's headers, skip others
    if (filesData.length > 0) {
      result.headers = filesData[0].headers;
      result.filesData.push({
        headers: filesData[0].headers,
        data: filesData[0].data,
      });
      for (var j = 1; j < filesData.length; j++) {
        result.filesData.push({
          headers: filesData[j].headers,
          data: filesData[j].data, // data already excludes header (readCSVFileFull separates headers)
        });
      }
    }
  } else if (headerMode === "all-data") {
    // All headers as data - no header row in output
    result.headers = [];
    for (var k = 0; k < filesData.length; k++) {
      // Add header row as first data row for each file
      var allData = [];
      if (filesData[k].headers.length > 0) {
        allData.push(filesData[k].headers);
      }
      for (var m = 0; m < filesData[k].data.length; m++) {
        allData.push(filesData[k].data[m]);
      }
      result.filesData.push({
        headers: filesData[k].headers,
        data: allData,
      });
    }
  }

  return result;
}

function getTimestamp() {
  var now = new Date();
  var year = now.getFullYear();
  var month = String(now.getMonth() + 1).padStart(2, "0");
  var day = String(now.getDate()).padStart(2, "0");
  var hours = String(now.getHours()).padStart(2, "0");
  var minutes = String(now.getMinutes()).padStart(2, "0");
  var seconds = String(now.getSeconds()).padStart(2, "0");
  return year + month + day + hours + minutes + seconds;
}
