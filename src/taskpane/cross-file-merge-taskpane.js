/* global console, document, Excel, Office, XLSX */

var crossFileMergeUtils = require("../utils/cross-file-merge-utils");

var selectedFiles = [];
var headerRowNumber = 1;

function initOffice() {
  Office.onReady(function(info) {
    if (info.host === Office.HostType.Excel) {
      initEventListeners();
    }
  });
}

if (typeof Office !== "undefined") {
  initOffice();
} else {
  var checkOffice = setInterval(function() {
    if (typeof Office !== "undefined") {
      clearInterval(checkOffice);
      initOffice();
    }
  }, 50);
  setTimeout(function() { clearInterval(checkOffice); }, 10000);
}

function initEventListeners() {
  document.getElementById("selectFilesBtn").onclick = function() {
    document.getElementById("fileInput").click();
  };

  document.getElementById("fileInput").onchange = onFileSelected;
  document.getElementById("headerRowInput").onchange = onHeaderRowChanged;
  document.getElementById("executeBtn").onclick = executeMerge;
  document.getElementById("closeModalBtn").onclick = closeSuccessModal;
}

function onFileSelected(e) {
  var files = e.target.files;
  if (!files || files.length === 0) return;

  selectedFiles = Array.from(files).map(function(file) {
    return { file: file, name: file.name };
  });

  renderFileList();
  updateExecuteButton();
}

function onHeaderRowChanged(e) {
  headerRowNumber = parseInt(e.target.value, 10) || 0;
  if (headerRowNumber < 0) headerRowNumber = 0;
  if (headerRowNumber > 100) headerRowNumber = 100;
}

function renderFileList() {
  var listEl = document.getElementById("fileList");

  if (selectedFiles.length === 0) {
    listEl.innerHTML = '<div class="cfm-empty">未选择文件</div>';
    return;
  }

  var html = "";
  selectedFiles.forEach(function(item, index) {
    html +=
      '<div class="cfm-file-item" data-index="' + index + '">' +
      '<span class="cfm-file-name" title="' + escapeHtml(item.name) + '">' +
      escapeHtml(item.name) +
      '</span>' +
      '<button class="cfm-file-remove" data-index="' + index + '">移除</button>' +
      '</div>';
  });

  listEl.innerHTML = html;

  // Bind remove buttons
  var removeButtons = listEl.querySelectorAll(".cfm-file-remove");
  removeButtons.forEach(function(btn) {
    btn.onclick = function() {
      var idx = parseInt(btn.dataset.index, 10);
      selectedFiles.splice(idx, 1);
      renderFileList();
      updateExecuteButton();
    };
  });
}

function escapeHtml(text) {
  var div = document.createElement("div");
  div.textContent = text;
  return div.innerHTML;
}

function updateExecuteButton() {
  var btn = document.getElementById("executeBtn");
  btn.disabled = selectedFiles.length < 2;
}

function setStatus(message, type) {
  var el = document.getElementById("status");
  el.textContent = message;
  el.className = "cfm-status cfm-status--" + (type || "idle");
}

function showSuccessModal(message) {
  document.getElementById("successMessage").innerHTML = message;
  document.getElementById("successModal").style.display = "flex";
}

function closeSuccessModal() {
  document.getElementById("successModal").style.display = "none";
}

function executeMerge() {
  var executeBtn = document.getElementById("executeBtn");
  executeBtn.disabled = true;
  setStatus("处理中...", "loading");

  if (selectedFiles.length < 2) {
    executeBtn.disabled = false;
    setStatus("请至少选择两个文件", "error");
    return;
  }

  headerRowNumber = parseInt(document.getElementById("headerRowInput").value, 10) || 0;

  // 步骤1：解析所有 Excel 文件
  var fileDataList = [];
  var parsePromises = selectedFiles.map(function(item) {
    return crossFileMergeUtils.parseExcelFile(item.file)
      .then(function(result) {
        fileDataList.push({
          data: result.data,
          name: item.name,
          columnCount: result.data.length > 0 ? result.data[0].length : 0
        });
        setStatus("已解析 " + fileDataList.length + "/" + selectedFiles.length + " 个文件...", "loading");
      })
      .catch(function(err) {
        throw new Error("解析文件 '" + item.name + "' 失败: " + err.message);
      });
  });

  Promise.all(parsePromises)
    .then(function() {
      // 步骤2：验证列数一致性
      var validation = crossFileMergeUtils.validateColumnConsistency(fileDataList);
      if (!validation.valid) {
        throw new Error(validation.error);
      }

      // 步骤3：合并数据
      var mergeResult = crossFileMergeUtils.mergeExcelData(fileDataList, headerRowNumber);

      if (mergeResult.mergedData.length === 0) {
        throw new Error("没有可合并的数据");
      }

      // 步骤4：写入当前工作簿
      return Excel.run(function(context) {
        // 获取所有 sheet 名称
        var worksheets = context.workbook.worksheets;
        worksheets.load("items/name");
        return context.sync().then(function() {
          var existingNames = worksheets.items.map(function(ws) { return ws.name; });
          var targetSheetName = crossFileMergeUtils.generateUniqueSheetName("合并结果", existingNames);

          // 创建新 sheet
          var newSheet = worksheets.add(targetSheetName);
          context.sync();

          // 写入数据
          var columnLetter = crossFileMergeUtils.getColumnLetter(mergeResult.columnCount - 1);
          var dataRange = newSheet.getRange("A1:" + columnLetter + mergeResult.mergedData.length);
          dataRange.values = mergeResult.mergedData;
          context.sync();

          return {
            sheetName: targetSheetName,
            fileCount: fileDataList.length,
            rowCount: mergeResult.mergedData.length,
            columnCount: mergeResult.columnCount
          };
        });
      });
    })
    .then(function(result) {
      setStatus("完成！已合并 " + result.fileCount + " 个文件，共 " + result.rowCount + " 行数据", "success");

      var message = "✅ 合并成功！<br><br>" +
        "共合并 " + result.fileCount + " 个文件，" + result.rowCount + " 行数据（" + result.columnCount + " 列）<br>" +
        "输出工作表：" + result.sheetName;
      showSuccessModal(message);
    })
    .catch(function(error) {
      setStatus("错误: " + error.message, "error");
    })
    .finally(function() {
      executeBtn.disabled = false;
    });
}
