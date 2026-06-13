/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

var sqlUtils = require("../utils/sql-utils");

var dbManager = null;
var persistenceManager = null;
var currentQueryResult = null;

// —— 初始化 ——

Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    initDB();
    bindEvents();
    updateSelectionInfo();
  }
});

function initDB() {
  dbManager = new sqlUtils.DatabaseManager();
  persistenceManager = new sqlUtils.PersistenceManager(dbManager);

  var statusEl = document.getElementById("importStatus");
  statusEl.className = "status-message status-loading";
  statusEl.textContent = "正在初始化数据库...";

  // 尝试从 IndexedDB 加载
  persistenceManager.loadFromIndexedDB().then(function (buffer) {
    return dbManager.init(buffer);
  }).then(function () {
    statusEl.textContent = "数据库就绪";
    statusEl.className = "status-message status-success";
    refreshTableList();
  }).catch(function (err) {
    // 降级：创建空数据库
    statusEl.className = "status-message status-loading";
    statusEl.textContent = "正在创建新数据库...";
    return dbManager.init(null).then(function () {
      statusEl.textContent = "新数据库已创建";
      statusEl.className = "status-message status-success";
      refreshTableList();
    });
  });
}

// —— 标签切换 ——

function bindEvents() {
  // 标签切换
  var tabBtns = document.querySelectorAll(".tab-btn");
  for (var i = 0; i < tabBtns.length; i++) {
    tabBtns[i].addEventListener("click", function () {
      switchTab(this.getAttribute("data-tab"));
    });
  }

  // 导入
  document.getElementById("importBtn").addEventListener("click", runImport);
  document.getElementById("refreshBtn").addEventListener("click", updateSelectionInfo);

  // 表管理
  document.getElementById("browseTableSelect").addEventListener("change", onBrowseTableChange);
  document.getElementById("clearTableBtn").addEventListener("click", clearSelectedTable);
  document.getElementById("dropTableBtn").addEventListener("click", dropSelectedTable);

  // SQL 查询
  document.getElementById("executeBtn").addEventListener("click", runQuery);
  document.getElementById("clearSqlBtn").addEventListener("click", clearSql);
  document.getElementById("writeSheetBtn").addEventListener("click", writeResultToSheet);
  document.getElementById("copyResultBtn").addEventListener("click", copyResult);

  // 持久化
  document.getElementById("saveDbBtn").addEventListener("click", function () {
    persistenceManager.exportToFile();
  });
  document.getElementById("loadDbBtn").addEventListener("click", function () {
    document.getElementById("dbFileInput").click();
  });
  document.getElementById("dbFileInput").addEventListener("change", function (e) {
    if (e.target.files.length > 0) {
      loadDbFile(e.target.files[0]);
    }
  });
}

function switchTab(tabName) {
  var btns = document.querySelectorAll(".tab-btn");
  for (var i = 0; i < btns.length; i++) {
    btns[i].classList.remove("active");
    if (btns[i].getAttribute("data-tab") === tabName) {
      btns[i].classList.add("active");
    }
  }

  var contents = document.querySelectorAll(".tab-content");
  for (var j = 0; j < contents.length; j++) {
    contents[j].classList.remove("active");
  }
  document.getElementById("tab-" + tabName).classList.add("active");

  if (tabName === "browse") {
    populateBrowseSelect();
  }
}

// —— 刷新当前选中区域信息 ——

function updateSelectionInfo() {
  Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    range.load(["address"]);
    return context.sync().then(function () {
      document.getElementById("selectionInfo").value = range.address;
    });
  }).catch(function (error) {
    console.error("Failed to get selection:", error);
  });
}

// —— 刷新已导入的表列表 ——

function refreshTableList() {
  var tableListEl = document.getElementById("tableList");
  if (!dbManager || !dbManager.isInitialized) {
    tableListEl.innerHTML = '<div class="table-list-empty">暂无导入的表</div>';
    return;
  }

  var tables = dbManager.getTables();
  if (tables.length === 0) {
    tableListEl.innerHTML = '<div class="table-list-empty">暂无导入的表</div>';
    return;
  }

  var html = "";
  for (var i = 0; i < tables.length; i++) {
    var schema = dbManager.getTableSchema(tables[i].name);
    var colNames = schema.map(function (col) { return col.name; }).join(", ");
    html += '<div class="table-item">' +
      '<div class="table-item-info">' +
      '<div class="table-item-name">' + tables[i].name + '</div>' +
      '<div class="table-item-columns">' + colNames + '</div>' +
      '</div>' +
      '<div class="table-item-actions">' +
      '<button onclick="deleteTable(\'' + tables[i].name + '\')">🗑</button>' +
      '</div>' +
      '</div>';
  }
  tableListEl.innerHTML = html;
}

/**
 * 自定义确认对话框（替代 window.confirm，因 Office.js 覆写后抛出异常）
 * @param {string} message - 提示信息
 * @returns {Promise<boolean>}
 */
function showConfirm(message) {
  return new Promise(function (resolve) {
    var overlay = document.createElement('div');
    overlay.className = 'confirm-overlay';

    var dialog = document.createElement('div');
    dialog.className = 'confirm-dialog';

    var msgEl = document.createElement('p');
    msgEl.className = 'confirm-message';
    msgEl.textContent = message;

    var btnGroup = document.createElement('div');
    btnGroup.className = 'confirm-buttons';

    var cancelBtn = document.createElement('button');
    cancelBtn.className = 'btn-secondary';
    cancelBtn.textContent = '取消';

    var okBtn = document.createElement('button');
    okBtn.className = 'sql-button sql-button-primary';
    okBtn.textContent = '确定';

    btnGroup.appendChild(cancelBtn);
    btnGroup.appendChild(okBtn);
    dialog.appendChild(msgEl);
    dialog.appendChild(btnGroup);
    overlay.appendChild(dialog);
    document.body.appendChild(overlay);

    function cleanup(result) {
      document.body.removeChild(overlay);
      resolve(result);
    }

    cancelBtn.addEventListener('click', function () { cleanup(false); });
    okBtn.addEventListener('click', function () { cleanup(true); });
    overlay.addEventListener('click', function (e) {
      if (e.target === overlay) cleanup(false);
    });
  });
}

// —— 删除表（全局函数，供 onclick 调用） ——

window.deleteTable = async function (tableName) {
  var confirmed = await showConfirm('确定要删除表 "' + tableName + '" 吗？');
  if (!confirmed) return;

  dbManager.exec('DROP TABLE "' + tableName + '"');
  persistenceManager.scheduleSave();
  refreshTableList();
  populateBrowseSelect();
};

// —— 导入数据 ——

function runImport() {
  var importBtn = document.getElementById("importBtn");
  var statusEl = document.getElementById("importStatus");

  importBtn.disabled = true;
  statusEl.className = "status-message status-loading";
  statusEl.textContent = "正在读取选中区域...";

  // 获取表名参数
  var tableNameInput = document.getElementById("tableNameInput");
  var tableName = tableNameInput.value.trim();
  if (!tableName) {
    var now = new Date();
    var dateStr = now.getFullYear() + "_" +
      String(now.getMonth() + 1).padStart(2, "0") + "_" +
      String(now.getDate()).padStart(2, "0");
    tableName = "import_" + dateStr;
  }
  tableName = tableName.replace(/[^a-zA-Z0-9_]/g, "_");
  if (/^[0-9]/.test(tableName)) {
    tableName = "_" + tableName;
  }
  if (tableName === "") {
    tableName = "import_data";
  }
  var firstRowIsHeader = document.getElementById("firstRowHeader").checked;

  // 分块大小：每次读取 1 万行，避免一次性加载超大数组导致内存溢出
  var CHUNK_SIZE = 10000;

  Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    range.load(["rowIndex", "columnIndex"]);
    return context.sync().then(function () {
      var usedRange = range.getUsedRange();
      usedRange.load(["rowCount", "columnCount", "rowIndex", "columnIndex"]);
      return context.sync().then(function () {
        var totalRows = usedRange.rowCount;
        var totalCols = usedRange.columnCount;

        if (totalRows === 0 || !totalCols) {
          statusEl.className = "status-message status-error";
          statusEl.textContent = "错误：选中区域没有数据";
          importBtn.disabled = false;
          return;
        }

        var worksheet = context.workbook.worksheets.getActiveWorksheet();
        var totalDataRows = totalRows - (firstRowIsHeader ? 1 : 0);
        var cumulativeInserted = 0;
        var tableCreated = false;

        /**
         * 递归读取下一块数据（通过 Promise 链避免堆栈溢出）
         * @param {number} batchStart - 本次起始行号（0-based，相对于 usedRange）
         * @returns {Promise|undefined}
         */
        function readChunk(batchStart) {
          if (batchStart >= totalRows) {
            // 全部处理完毕
            var finalMsg = "成功导入 " + cumulativeInserted + " 行数据到表 " + tableName;
            statusEl.className = "status-message status-success";
            statusEl.textContent = finalMsg;
            importBtn.disabled = false;
            persistenceManager.scheduleSave();
            refreshTableList();
            populateBrowseSelect();
            return;
          }

          var batchEnd = Math.min(batchStart + CHUNK_SIZE, totalRows);
          statusEl.textContent = "正在读取数据... (" + batchStart + "/" + totalRows + " 行)";

          var chunkRange = worksheet.getRangeByIndexes(
            usedRange.rowIndex + batchStart, usedRange.columnIndex,
            batchEnd - batchStart, totalCols
          );
          chunkRange.load("values");

          return context.sync().then(function () {
            var chunkValues = chunkRange.values;
            if (!chunkValues || chunkValues.length === 0) {
              return readChunk(batchEnd);
            }

            if (!tableCreated) {
              // 第一批：创建表并插入
              var result = sqlUtils.importData(dbManager, tableName, chunkValues, firstRowIsHeader);
              if (!result.success) {
                statusEl.className = "status-message status-error";
                statusEl.textContent = "错误: " + result.message;
                importBtn.disabled = false;
                return;
              }
              tableCreated = true;
              cumulativeInserted = result.rowsInserted;
            } else {
              // 后续批次：仅追加数据
              var insertResult = sqlUtils.insertRows(dbManager, tableName, chunkValues);
              if (!insertResult.success) {
                statusEl.className = "status-message status-error";
                statusEl.textContent = "错误: " + insertResult.message;
                importBtn.disabled = false;
                return;
              }
              cumulativeInserted += insertResult.rowsInserted;
            }

            statusEl.textContent = "已处理 " + cumulativeInserted + "/" + totalDataRows + " 行...";
            return readChunk(batchEnd);
          });
        }

        return readChunk(0);
      });
    });
  }).catch(function (error) {
    statusEl.className = "status-message status-error";
    statusEl.textContent = "错误: " + error.message;
    importBtn.disabled = false;
  });
}

// —— 表浏览器 ——

function populateBrowseSelect() {
  var select = document.getElementById("browseTableSelect");
  if (!dbManager || !dbManager.isInitialized) {
    select.innerHTML = '<option value="">-- 请先导入数据 --</option>';
    return;
  }

  var tables = dbManager.getTables();
  if (tables.length === 0) {
    select.innerHTML = '<option value="">-- 请先导入数据 --</option>';
    return;
  }

  var currentValue = select.value;
  var html = '<option value="">-- 选择表 --</option>';
  for (var i = 0; i < tables.length; i++) {
    html += '<option value="' + tables[i].name + '">' + tables[i].name + "</option>";
  }
  select.innerHTML = html;
  if (currentValue) select.value = currentValue;
}

function onBrowseTableChange() {
  var tableName = this.value;
  if (!tableName) {
    document.getElementById("schemaBody").innerHTML = "";
    document.getElementById("previewHead").innerHTML = "";
    document.getElementById("previewBody").innerHTML = "";
    document.getElementById("rowCount").textContent = "";
    return;
  }

  // 表结构
  var schema = dbManager.getTableSchema(tableName);
  var schemaHtml = "";
  for (var i = 0; i < schema.length; i++) {
    schemaHtml += "<tr><td>" + schema[i].name + "</td><td>" + schema[i].type + "</td></tr>";
  }
  document.getElementById("schemaBody").innerHTML = schemaHtml;

  // 行数
  var count = dbManager.getTableRowCount(tableName);
  document.getElementById("rowCount").textContent = "共 " + count + " 行";

  // 预览
  var preview = dbManager.previewTable(tableName, 10);
  if (preview.type === "query") {
    var headHtml = "";
    for (var c = 0; c < preview.columns.length; c++) {
      headHtml += "<th>" + preview.columns[c] + "</th>";
    }
    document.getElementById("previewHead").innerHTML = "<tr>" + headHtml + "</tr>";

    var bodyHtml = "";
    for (var r = 0; r < preview.rows.length; r++) {
      bodyHtml += "<tr>";
      for (var c2 = 0; c2 < preview.rows[r].length; c2++) {
        bodyHtml += "<td>" + (preview.rows[r][c2] !== null ? preview.rows[r][c2] : "") + "</td>";
      }
      bodyHtml += "</tr>";
    }
    document.getElementById("previewBody").innerHTML = bodyHtml;
  }
}

async function clearSelectedTable() {
  var select = document.getElementById("browseTableSelect");
  var tableName = select.value;
  if (!tableName) return;
  var confirmed = await showConfirm('确定要清空表 "' + tableName + '" 的所有数据吗？');
  if (!confirmed) return;

  dbManager.exec("DELETE FROM \"" + tableName + "\"");
  persistenceManager.scheduleSave();
  onBrowseTableChange.call(select);
  setStatusText("queryStatus", tableName + " 已清空", "success");
}

async function dropSelectedTable() {
  var select = document.getElementById("browseTableSelect");
  var tableName = select.value;
  if (!tableName) return;
  var confirmed = await showConfirm('确定要删除表 "' + tableName + '" 吗？此操作不可恢复！');
  if (!confirmed) return;

  dbManager.exec('DROP TABLE "' + tableName + '"');
  persistenceManager.scheduleSave();
  populateBrowseSelect();
  select.value = "";
  onBrowseTableChange.call(select);
  setStatusText("queryStatus", tableName + " 已删除", "success");
  refreshTableList();
}

// —— SQL 查询 ——

async function runQuery() {
  var sqlInput = document.getElementById("sqlInput");
  var sql = sqlInput.value.trim();
  if (!sql) return;

  var statusEl = document.getElementById("queryStatus");
  var executeBtn = document.getElementById("executeBtn");

  // DROP/DELETE/UPDATE 二次确认
  var upperSql = sql.toUpperCase().trim();
  if (upperSql.startsWith("DROP") || upperSql.startsWith("DELETE") || upperSql.startsWith("UPDATE")) {
    var confirmed = await showConfirm("确定要执行危险操作吗？\n\n" + sql);
    if (!confirmed) {
      return;
    }
  }

  // —— 禁用按钮、提供视觉反馈 ——
  executeBtn.disabled = true;
  executeBtn.textContent = "⏳ 执行中...";
  executeBtn.classList.add("sql-button-loading");

  statusEl.className = "status-message status-loading";
  statusEl.textContent = "执行中...";

  // 让浏览器渲染按钮的 disabled 状态后再执行同步查询
  await new Promise(function (resolve) { setTimeout(resolve, 0); });

  try {
    var result = dbManager.exec(sql);

    if (result.type === "error") {
      statusEl.className = "status-message status-error";
      statusEl.textContent = "错误: " + result.message;
      return;
    }

    if (result.type === "modification") {
      statusEl.textContent = "完成，影响 " + result.rowsAffected + " 行 (" + result.elapsed.toFixed(2) + " 秒)";
      statusEl.className = "status-message status-success";
      document.getElementById("resultActions").style.display = "none";
      document.getElementById("resultDisplay").style.display = "none";
      currentQueryResult = null;
      persistenceManager.scheduleSave();
      refreshTableList();
      populateBrowseSelect();
      addQueryHistory(sql, result.type, result.elapsed, result.rowsAffected);
      return;
    }

    // SELECT 结果
    statusEl.textContent = "查询完成，返回 " + result.rowCount + " 行 (" + result.elapsed.toFixed(2) + " 秒)";
    statusEl.className = "status-message status-success";

    currentQueryResult = result;

    // 渲染结果表格
    var headHtml = "";
    for (var c = 0; c < result.columns.length; c++) {
      headHtml += "<th>" + result.columns[c] + "</th>";
    }
    document.getElementById("resultHead").innerHTML = "<tr>" + headHtml + "</tr>";

    var MAX_DISPLAY_ROWS = 500;
    var displayRows = result.rows.slice(0, MAX_DISPLAY_ROWS);
    var bodyHtml = "";
    for (var r = 0; r < displayRows.length; r++) {
      bodyHtml += "<tr>";
      for (var c2 = 0; c2 < displayRows[r].length; c2++) {
        var val = displayRows[r][c2];
        bodyHtml += "<td>" + (val !== null ? val : "") + "</td>";
      }
      bodyHtml += "</tr>";
    }
    document.getElementById("resultBody").innerHTML = bodyHtml;
    document.getElementById("resultActions").style.display = "block";
    document.getElementById("resultDisplay").style.display = "block";

    if (result.rowCount > MAX_DISPLAY_ROWS) {
      statusEl.textContent += " (仅显示前 " + MAX_DISPLAY_ROWS + " 行)";
    }

    addQueryHistory(sql, result.type, result.elapsed, result.rowCount);
  } finally {
    executeBtn.disabled = false;
    executeBtn.textContent = "▶ 执行";
    executeBtn.classList.remove("sql-button-loading");
  }
}

function clearSql() {
  document.getElementById("sqlInput").value = "";
  document.getElementById("resultActions").style.display = "none";
  document.getElementById("resultDisplay").style.display = "none";
  document.getElementById("queryStatus").className = "status-message status-idle";
  document.getElementById("queryStatus").textContent = "";
  currentQueryResult = null;
}

// —— 结果导出 ——

var _sheetDialogActive = false;

function showSheetNameDialog(callback) {
  if (_sheetDialogActive) return;
  _sheetDialogActive = true;

  var defaultName = "查询结果";

  var overlay = document.createElement("div");
  overlay.id = "sheetNameOverlay";
  overlay.style.cssText = "position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.3);z-index:1000;display:flex;align-items:center;justify-content:center;font-family:'Segoe UI','Segoe UI Web',-apple-system,sans-serif;";

  var dialog = document.createElement("div");
  dialog.style.cssText = "background:#fff;border-radius:8px;padding:24px;min-width:260px;max-width:320px;box-shadow:0 8px 24px rgba(0,0,0,0.2);";

  var title = document.createElement("div");
  title.textContent = "请输入工作表名称";
  title.style.cssText = "font-size:14px;font-weight:600;margin-bottom:16px;color:#333;";

  var input = document.createElement("input");
  input.type = "text";
  input.value = defaultName;
  input.style.cssText = "width:100%;box-sizing:border-box;padding:8px 10px;border:1px solid #8a8886;border-radius:4px;font-size:13px;font-family:inherit;outline:none;";
  input.addEventListener("focus", function () { this.style.borderColor = "#0078d4"; });
  input.addEventListener("blur", function () { this.style.borderColor = "#8a8886"; });

  var btnContainer = document.createElement("div");
  btnContainer.style.cssText = "margin-top:20px;display:flex;gap:8px;justify-content:flex-end;";

  function cleanup(result) {
    if (overlay.parentNode) {
      overlay.parentNode.removeChild(overlay);
    }
    _sheetDialogActive = false;
    callback(result);
  }

  var cancelBtn = document.createElement("button");
  cancelBtn.textContent = "取消";
  cancelBtn.style.cssText = "padding:6px 20px;border:1px solid #8a8886;border-radius:4px;background:#fff;cursor:pointer;font-size:13px;font-family:inherit;";
  cancelBtn.addEventListener("click", function () { cleanup(defaultName); });

  var okBtn = document.createElement("button");
  okBtn.textContent = "确定";
  okBtn.style.cssText = "padding:6px 20px;border:none;border-radius:4px;background:#0078d4;color:#fff;cursor:pointer;font-size:13px;font-family:inherit;";
  okBtn.addEventListener("click", function () { cleanup(input.value.trim() || defaultName); });

  input.addEventListener("keydown", function (e) {
    if (e.key === "Enter") { cleanup(input.value.trim() || defaultName); }
    if (e.key === "Escape") { cleanup(defaultName); }
  });

  btnContainer.appendChild(cancelBtn);
  btnContainer.appendChild(okBtn);
  dialog.appendChild(title);
  dialog.appendChild(input);
  dialog.appendChild(btnContainer);
  overlay.appendChild(dialog);
  document.body.appendChild(overlay);

  setTimeout(function () { input.focus(); input.select(); }, 50);
}

/**
 * 生成不重复的工作表名
 * @param {object} sheetCollection - Office.js worksheet 集合（已加载 items/name）
 * @param {string} baseName - 基础名称
 * @returns {string}
 */
function generateUniqueSheetName(sheetCollection, baseName) {
  var finalName = baseName;
  var counter = 1;
  var exists = true;
  while (exists) {
    exists = false;
    for (var i = 0; i < sheetCollection.items.length; i++) {
      if (sheetCollection.items[i].name === finalName) {
        exists = true;
        finalName = baseName + " (" + counter + ")";
        counter++;
        break;
      }
    }
  }
  return finalName;
}

/**
 * 更新进度条和文字
 * @param {HTMLElement} fillEl - 进度条填充元素
 * @param {HTMLElement} textEl - 进度文字元素
 * @param {number} current - 当前已写入行数
 * @param {number} total - 总行数
 */
function updateProgress(fillEl, textEl, current, total) {
  var pct = Math.round((current / total) * 100);
  fillEl.style.width = pct + "%";
  fillEl.setAttribute("aria-valuenow", pct);
  textEl.textContent = "已写入 " + current + "/" + total + " 行";
}

function writeResultToSheet() {
  if (!currentQueryResult) return;

  showSheetNameDialog(function (sheetName) {
    var rows = currentQueryResult.rows;
    var columns = currentQueryResult.columns;
    if (!rows || rows.length === 0) {
      setStatusText("queryStatus", "没有数据可写入", "error");
      return;
    }

    var CHUNK_SIZE = 5000;
    var totalRows = rows.length;
    var totalCols = columns.length;

    var writeBtn = document.getElementById("writeSheetBtn");
    var progressContainer = document.getElementById("writeProgress");
    var progressFill = document.getElementById("writeProgressFill");
    var progressText = document.getElementById("writeProgressText");

    // 禁用按钮、显示 spinner、显示进度条
    writeBtn.disabled = true;
    writeBtn.textContent = "⏳ 写入中...";
    writeBtn.classList.add("sql-button-loading");
    progressContainer.style.display = "flex";
    updateProgress(progressFill, progressText, 0, totalRows);

    var batchIndex = 0;
    var finalSheetName = sheetName;
    var sheetCreated = false;

    function writeNextBatch() {
      var startRow = batchIndex * CHUNK_SIZE;
      if (startRow >= totalRows) {
        // 全部写入完成
        restoreWriteButton();
        setStatusText("queryStatus", "已将 " + totalRows + " 行结果写入新工作表", "success");
        return;
      }
      var endRow = Math.min(startRow + CHUNK_SIZE, totalRows);
      var batchSize = endRow - startRow;

      Excel.run(function (context) {
        if (!sheetCreated) {
          // 第一批：创建表 + 写表头 + 写数据
          var sheetCollection = context.workbook.worksheets;
          sheetCollection.load("items/name");
          return context.sync().then(function () {
            finalSheetName = generateUniqueSheetName(sheetCollection, sheetName);
            var newSheet = sheetCollection.add(finalSheetName);
            newSheet.position = 0;
            sheetCreated = true;

            // 写入表头 + 第一批数据行
            var rangeRows = batchSize + 1; // +1 是表头行
            var range = newSheet.getRangeByIndexes(0, 0, rangeRows, totalCols);
            var values = [columns];
            for (var r = startRow; r < endRow; r++) {
              values.push(rows[r]);
            }
            range.values = values;
            return context.sync();
          });
        } else {
          // 后续批次：通过已记录的表名获取工作表，追加数据
          var sheet = context.workbook.worksheets.getItem(finalSheetName);
          // +1 跳过表头行
          var range = sheet.getRangeByIndexes(startRow + 1, 0, batchSize, totalCols);
          var values = [];
          for (var r = startRow; r < endRow; r++) {
            values.push(rows[r]);
          }
          range.values = values;
          return context.sync();
        }
      }).then(function () {
        batchIndex++;
        updateProgress(progressFill, progressText, endRow, totalRows);
        writeNextBatch(); // 递归开始下一批
      }).catch(function (error) {
        restoreWriteButton();
        var msg = (error && error.message) ? error.message : String(error || "未知错误");
        setStatusText("queryStatus", "写入失败: " + msg, "error");
      });
    }

    function restoreWriteButton() {
      writeBtn.disabled = false;
      writeBtn.textContent = "📝 写入新工作表";
      writeBtn.classList.remove("sql-button-loading");
      progressContainer.style.display = "none";
    }

    // 启动写入
    writeNextBatch();
  });
}

function copyResult() {
  if (!currentQueryResult) return;

  var text = currentQueryResult.columns.join("\t") + "\n";
  for (var r = 0; r < currentQueryResult.rows.length; r++) {
    text += currentQueryResult.rows[r].join("\t") + "\n";
  }

  navigator.clipboard.writeText(text).then(function () {
    setStatusText("queryStatus", "已复制 " + currentQueryResult.rows.length + " 行到剪贴板", "success");
  }).catch(function () {
    // fallback
    var textarea = document.createElement("textarea");
    textarea.value = text;
    document.body.appendChild(textarea);
    textarea.select();
    document.execCommand("copy");
    document.body.removeChild(textarea);
    setStatusText("queryStatus", "已复制到剪贴板", "success");
  });
}

// —— 查询历史 ——

function addQueryHistory(sql, type, elapsed, rowInfo) {
  var MAX_HISTORY = 50;
  var storageKey = "sql-query-history";

  var history = [];
  try {
    var stored = localStorage.getItem(storageKey);
    if (stored) history = JSON.parse(stored);
  } catch (e) { /* ignore */ }

  var entry = {
    sql: sql,
    type: type,
    elapsed: elapsed,
    rowInfo: rowInfo,
    timestamp: new Date().toLocaleString(),
  };

  history.unshift(entry);
  if (history.length > MAX_HISTORY) {
    history = history.slice(0, MAX_HISTORY);
  }

  try {
    localStorage.setItem(storageKey, JSON.stringify(history));
  } catch (e) { /* ignore */ }

  renderQueryHistory();
}

function renderQueryHistory() {
  var container = document.getElementById("queryHistory");

  var history = [];
  try {
    var stored = localStorage.getItem("sql-query-history");
    if (stored) history = JSON.parse(stored);
  } catch (e) { /* ignore */ }

  if (history.length === 0) {
    container.innerHTML = '<div class="table-list-empty">暂无查询记录</div>';
    return;
  }

  var html = "";
  for (var i = 0; i < history.length; i++) {
    var entry = history[i];
    var meta = entry.timestamp + " · " + entry.elapsed.toFixed(2) + "s";
    if (entry.type === "query") meta += " · " + entry.rowInfo + " 行";
    else meta += " · 影响 " + entry.rowInfo + " 行";

    html += '<div class="history-item" data-sql="' + escapeHtml(entry.sql) + '">' +
      '<div class="history-item-sql">' + escapeHtml(truncateSql(entry.sql)) + '</div>' +
      '<div class="history-item-meta">' + meta + '</div>' +
      '</div>';
  }
  container.innerHTML = html;

  // 绑定点击回填
  var items = container.querySelectorAll(".history-item");
  for (var j = 0; j < items.length; j++) {
    items[j].addEventListener("click", function () {
      document.getElementById("sqlInput").value = this.getAttribute("data-sql");
      switchTab("query");
    });
  }
}

// —— .db 文件加载 ——

function loadDbFile(file) {
  var statusEl = document.getElementById("importStatus");
  statusEl.className = "status-message status-loading";
  statusEl.textContent = "正在加载数据库文件...";

  persistenceManager.importFromFile(file).then(function () {
    statusEl.className = "status-message status-success";
    statusEl.textContent = "数据库文件已加载";
    persistenceManager.scheduleSave();
    refreshTableList();
    populateBrowseSelect();
  }).catch(function (err) {
    statusEl.className = "status-message status-error";
    statusEl.textContent = "加载失败: " + err.message;
  });
}

// —— 工具函数 ——

function setStatusText(elId, message, type) {
  var el = document.getElementById(elId);
  el.textContent = message;
  el.className = "status-message status-" + type;
}

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function truncateSql(sql) {
  if (sql.length > 80) {
    return sql.substring(0, 80) + "...";
  }
  return sql;
}

// 初始化查询历史
renderQueryHistory();
