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

// —— 删除表（全局函数，供 onclick 调用） ——

window.deleteTable = function (tableName) {
  if (!confirm('确定要删除表 "' + tableName + '" 吗？')) return;

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

  Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    range.load(["address", "values"]);
    return context.sync().then(function () {
      var values = range.values;
      if (!values || values.length === 0) {
        statusEl.className = "status-message status-error";
        statusEl.textContent = "错误：选中区域没有数据";
        importBtn.disabled = false;
        return;
      }

      var tableNameInput = document.getElementById("tableNameInput");
      var tableName = tableNameInput.value.trim();
      if (!tableName) {
        // 自动生成表名
        var now = new Date();
        var dateStr = now.getFullYear() + "_" +
          String(now.getMonth() + 1).padStart(2, "0") + "_" +
          String(now.getDate()).padStart(2, "0");
        tableName = "import_" + dateStr;
      }
      // 清理表名
      tableName = tableName.replace(/[^a-zA-Z0-9_]/g, "_");
      if (/^[0-9]/.test(tableName)) {
        tableName = "_" + tableName;
      }
      if (tableName === "") {
        tableName = "import_data";
      }

      var firstRowIsHeader = document.getElementById("firstRowHeader").checked;

      statusEl.textContent = "正在导入数据到表 " + tableName + "...";

      var result = sqlUtils.importData(dbManager, tableName, values, firstRowIsHeader);
      if (result.success) {
        statusEl.className = "status-message status-success";
        statusEl.textContent = result.message;
        persistenceManager.scheduleSave();
        refreshTableList();
        populateBrowseSelect();
      } else {
        statusEl.className = "status-message status-error";
        statusEl.textContent = "错误: " + result.message;
      }

      importBtn.disabled = false;
    });
  }).catch(function (error) {
    statusEl.className = "status-message status-error";
    statusEl.textContent = "错误: " + error.message;
    importBtn.disabled = false;
  });
}
