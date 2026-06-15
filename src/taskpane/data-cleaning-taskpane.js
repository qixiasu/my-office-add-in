// src/taskpane/data-cleaning-taskpane.js

/* global console, document, Excel, Office */

var cleaningUtils = require("../utils/data-cleaning-utils");

var MAX_PREVIEW_ROWS = 5;
var LARGE_DATA_THRESHOLD = 10000;

// 当前状态
var state = {
  selectionAddress: null,
  rawValues: null,      // 原始数据（二维数组）— 用于撤销
  headerRow: null,      // 表头数组（如果有多行数据）
  dataRows: null,       // 数据行（不含表头）
  selectedOp: null,     // 当前选中的操作
  undoData: null,       // 撤销快照
};

// ===== 操作参数配置 =====
var OPERATION_PARAMS = {
  trimSpaces: {
    label: "修剪空格",
    template: function () {
      return (
        '<div class="param-row">' +
          '<label>清洗模式：</label>' +
          '<select id="trim-mode">' +
            '<option value="both">首尾空格</option>' +
            '<option value="all">所有多余空格</option>' +
            '<option value="leading">开头空格</option>' +
            '<option value="trailing">结尾空格</option>' +
          "</select>" +
        "</div>"
      );
    },
    getParams: function () {
      return { mode: document.getElementById("trim-mode").value };
    },
    execute: function (values, params) {
      return cleaningUtils.trimSpaces(values, params.mode);
    },
  },
  removeEmpty: {
    label: "删除空行",
    template: function () {
      return (
        '<div class="param-row">' +
          '<label>判定方式：</label>' +
          '<select id="empty-mode">' +
            '<option value="all">完全空行</option>' +
            '<option value="column">指定列为空</option>' +
            '<option value="ratio">空值率超过</option>' +
          "</select>" +
        "</div>" +
        '<div class="param-row" id="empty-column-row" style="display:none">' +
          '<label>列索引：</label>' +
          '<input type="number" id="empty-column" value="0" min="0" />' +
        "</div>" +
        '<div class="param-row" id="empty-ratio-row" style="display:none">' +
          '<label>阈值(%)：</label>' +
          '<input type="number" id="empty-ratio" value="50" min="0" max="100" />' +
        "</div>"
      );
    },
    bindEvents: function () {
      var modeSelect = document.getElementById("empty-mode");
      if (modeSelect) {
        modeSelect.addEventListener("change", function () {
          var val = this.value;
          document.getElementById("empty-column-row").style.display = val === "column" ? "flex" : "none";
          document.getElementById("empty-ratio-row").style.display = val === "ratio" ? "flex" : "none";
        });
      }
    },
    getParams: function () {
      var mode = document.getElementById("empty-mode").value;
      return {
        mode: mode,
        columnIndex: mode === "column" ? parseInt(document.getElementById("empty-column").value) || 0 : null,
        ratioThreshold: mode === "ratio" ? parseInt(document.getElementById("empty-ratio").value) || 50 : null,
      };
    },
    execute: function (values, params) {
      // removeEmptyRows 整行删除，对 dataRows 操作后合并表头
      return cleaningUtils.removeEmptyRows(values, params.mode, params.columnIndex, params.ratioThreshold);
    },
  },
  convertCase: {
    label: "大小写转换",
    template: function () {
      return (
        '<div class="param-row">' +
          '<label>转换模式：</label>' +
          '<select id="case-mode">' +
            '<option value="upper">全部大写</option>' +
            '<option value="lower">全部小写</option>' +
            '<option value="capitalize">首字母大写</option>' +
          "</select>" +
        "</div>"
      );
    },
    getParams: function () {
      return { mode: document.getElementById("case-mode").value };
    },
    execute: function (values, params) {
      return cleaningUtils.convertCase(values, params.mode);
    },
  },
  removeInvisible: {
    label: "清除不可见字符",
    template: function () {
      return (
        '<div class="param-row">' +
          '<label>清除类型：</label>' +
          '<select id="invisible-mode">' +
            '<option value="control">控制字符</option>' +
            '<option value="whitespace">空白字符(\\t\\n)</option>' +
            '<option value="zero-width">零宽字符</option>' +
            '<option value="all">全部非打印字符</option>' +
          "</select>" +
        "</div>"
      );
    },
    getParams: function () {
      return { mode: document.getElementById("invisible-mode").value };
    },
    execute: function (values, params) {
      return cleaningUtils.removeInvisible(values, params.mode);
    },
  },
  removeDuplicates: {
    label: "移除重复行",
    template: function () {
      // 列选择器在 bindEvents 时根据实际数据填充
      return (
        '<div class="param-row">' +
          '<label>依据列：</label>' +
          '<div id="dup-columns" class="checkbox-group"></div>' +
        "</div>" +
        '<div class="param-row">' +
          '<label>保留：</label>' +
          '<select id="dup-keep">' +
            '<option value="first">首行</option>' +
            '<option value="last">末行</option>' +
          "</select>" +
        "</div>"
      );
    },
    bindEvents: function () {
      var container = document.getElementById("dup-columns");
      if (!container || !state.headerRow) return;
      container.innerHTML = "";
      // 生成列选择复选框
      for (var i = 0; i < state.headerRow.length; i++) {
        var label = document.createElement("label");
        var cb = document.createElement("input");
        cb.type = "checkbox";
        cb.value = i;
        cb.checked = true;
        label.appendChild(cb);
        label.appendChild(document.createTextNode(" " + state.headerRow[i]));
        container.appendChild(label);
      }
    },
    getParams: function () {
      var checkboxes = document.querySelectorAll("#dup-columns input[type=checkbox]:checked");
      var cols = [];
      for (var i = 0; i < checkboxes.length; i++) {
        cols.push(parseInt(checkboxes[i].value));
      }
      return {
        keyColumns: cols.length > 0 ? cols : null,
        keep: document.getElementById("dup-keep").value,
      };
    },
    execute: function (values, params) {
      return cleaningUtils.removeDuplicates(values, params.keyColumns, params.keep);
    },
  },
};

// ===== Office 初始化 =====

Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("refresh-btn").addEventListener("click", loadSelection);
    document.getElementById("execute-btn").addEventListener("click", executeClean);
    document.getElementById("undo-btn").addEventListener("click", undoClean);

    // 操作按钮点击
    var opBtns = document.querySelectorAll(".op-btn");
    for (var i = 0; i < opBtns.length; i++) {
      opBtns[i].addEventListener("click", function () {
        selectOperation(this.dataset.op);
      });
    }

    // 打开面板时自动加载选区
    loadSelection();
  }
});

// ===== 选区加载 =====

function loadSelection() {
  setStatus("读取选区...", "loading");

  Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    range.load(["address", "values"]);
    return context.sync().then(function () {
      var address = range.address;
      var values = range.values;

      if (!values || !Array.isArray(values) || values.length === 0) {
        state.rawValues = null;
        state.headerRow = null;
        state.dataRows = null;
        state.selectionAddress = address || "未知";
        updateUIForEmpty();
        setStatus("请选中包含数据的目标区域", "error");
        return;
      }

      state.selectionAddress = address;
      state.rawValues = values;

      // 多行时首行为表头
      if (values.length > 1) {
        state.headerRow = values[0].map(String);
        state.dataRows = values.slice(1);
      } else {
        state.headerRow = null;
        state.dataRows = values;
      }

      document.getElementById("selection-address").textContent = "选区：" + address;
      enableExecuteIfReady();
      updatePreview();
      setStatus("已读取 " + state.dataRows.length + " 行数据", "success");
    });
  }).catch(function (error) {
    setStatus("读取选区失败: " + error.message, "error");
  });
}

// ===== 操作选择 =====

function selectOperation(opKey) {
  state.selectedOp = opKey;

  // 高亮选中按钮
  var btns = document.querySelectorAll(".op-btn");
  for (var i = 0; i < btns.length; i++) {
    btns[i].classList.remove("selected");
  }
  var activeBtn = document.querySelector('.op-btn[data-op="' + opKey + '"]');
  if (activeBtn) activeBtn.classList.add("selected");

  // 渲染参数
  var opConfig = OPERATION_PARAMS[opKey];
  if (opConfig) {
    document.getElementById("params-container").innerHTML = opConfig.template();
    if (opConfig.bindEvents) opConfig.bindEvents();
  }

  enableExecuteIfReady();
  updatePreview();
}

// ===== 预览 =====

function updatePreview() {
  var tbody = document.getElementById("preview-body");
  tbody.innerHTML = "";

  if (!state.dataRows || state.dataRows.length === 0) {
    tbody.innerHTML = '<tr><td colspan="2" class="empty-state">暂无数据</td></tr>';
    return;
  }

  if (!state.selectedOp) {
    tbody.innerHTML = '<tr><td colspan="2" class="empty-state">请先选择操作</td></tr>';
    return;
  }

  var opConfig = OPERATION_PARAMS[state.selectedOp];
  if (!opConfig) return;

  var params = opConfig.getParams();
  // Use allValues (header + dataRows) to match executeClean's input, then slice to preview limit
  var allValues = state.headerRow
    ? [state.headerRow].concat(state.dataRows)
    : state.dataRows;
  var previewData = allValues.slice(0, MAX_PREVIEW_ROWS);
  var resultData = opConfig.execute(previewData, params);

  for (var i = 0; i < previewData.length; i++) {
    var tr = document.createElement("tr");
    var originalStr = previewData[i].map(formatCell).join(" | ");
    var cleanedStr = (resultData[i] || []).map(formatCell).join(" | ");
    tr.innerHTML =
      "<td>" + escapeHtml(originalStr) + "</td>" +
      "<td>" + escapeHtml(cleanedStr) + "</td>";
    tbody.appendChild(tr);
  }
}

function formatCell(val) {
  if (val === null || val === undefined) return "";
  return String(val);
}

function escapeHtml(str) {
  if (typeof str !== "string") return String(str);
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

// ===== 执行 =====

function executeClean() {
  if (!state.selectedOp || !state.dataRows || state.dataRows.length === 0) {
    return;
  }

  if (state.dataRows.length > LARGE_DATA_THRESHOLD) {
    var confirmed = confirm(
      "数据量较大（" + state.dataRows.length + " 行），确定继续执行？"
    );
    if (!confirmed) return;
  }

  var opConfig = OPERATION_PARAMS[state.selectedOp];
  var params = opConfig.getParams();

  // 备份原始数据（撤销用）
  state.undoData = state.rawValues;

  setStatus("执行中...", "loading");
  document.getElementById("execute-btn").disabled = true;

  Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    range.load("address");
    return context.sync().then(function () {
      // 用选区数据执行清洗（含表头）
      var allValues = state.headerRow
        ? [state.headerRow].concat(state.dataRows)
        : state.dataRows;
      var cleaned = opConfig.execute(allValues, params);

      // 写回 Excel
      var writeRange = context.workbook.getSelectedRange();
      writeRange.values = cleaned;
      return context.sync();
    });
  }).then(function () {
    setStatus("清洗完成！已处理 " + state.dataRows.length + " 行", "success");
    document.getElementById("undo-btn").disabled = false;
    // 更新本地数据
    loadSelection();
  }).catch(function (error) {
    setStatus("执行失败: " + error.message, "error");
    document.getElementById("execute-btn").disabled = false;
  });
}

// ===== 撤销 =====

function undoClean() {
  if (!state.undoData) return;

  setStatus("撤销中...", "loading");

  Excel.run(function (context) {
    // Use the captured selection address so undo writes to the ORIGINAL range,
    // not whatever range happens to be selected now.
    var writeRange = context.workbook.getRange(state.selectionAddress);
    writeRange.values = state.undoData;
    return context.sync();
  }).then(function () {
    state.undoData = null;
    document.getElementById("undo-btn").disabled = true;
    setStatus("已撤销", "success");
    loadSelection();
  }).catch(function (error) {
    setStatus("撤销失败: " + error.message, "error");
  });
}

// ===== UI 辅助 =====

function updateUIForEmpty() {
  document.getElementById("selection-address").textContent = "选区：无数据";
  document.getElementById("execute-btn").disabled = true;
  document.getElementById("undo-btn").disabled = true;
  document.getElementById("preview-body").innerHTML =
    '<tr><td colspan="2" class="empty-state">请选中数据区域后点击刷新</td></tr>';
}

function enableExecuteIfReady() {
  document.getElementById("execute-btn").disabled =
    !state.selectedOp || !state.dataRows || state.dataRows.length === 0;
}

function setStatus(message, type) {
  var el = document.getElementById("status");
  el.textContent = message;
  el.className = "status-message status-" + type;
}
