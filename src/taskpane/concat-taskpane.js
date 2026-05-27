/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

var MAX_ROWS = 1050000;

var concatUtils = require("../utils/concat-utils");

function getColumnLetter(colIndex) {
  var letter = "";
  var remaining = colIndex;
  do {
    letter = String.fromCharCode(65 + (remaining % 26)) + letter;
    remaining = Math.floor(remaining / 26) - 1;
  } while (remaining >= 0);
  return letter;
}

// Confirmation state
var pendingConfirmation = null;

Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    var connectorInput = document.getElementById("connector");
    var executeBtn = document.getElementById("executeBtn");
    var confirmYes = document.getElementById("confirmYes");
    var confirmNo = document.getElementById("confirmNo");

    // 自动聚焦输入框
    connectorInput.focus();

    // 按钮点击触发执行
    executeBtn.addEventListener("click", function () {
      runConcat();
    });

    // 确认框取消按钮
    confirmNo.onclick = function () {
      document.getElementById("confirmBox").style.display = "none";
      if (pendingConfirmation) {
        pendingConfirmation.resolve(false);
        pendingConfirmation = null;
      }
      executeBtn.disabled = false;
      setStatus("状态：等待操作...", "idle");
    };

    // 确认框确定按钮
    confirmYes.onclick = function () {
      document.getElementById("confirmBox").style.display = "none";
      if (pendingConfirmation) {
        pendingConfirmation.resolve(true);
        pendingConfirmation = null;
      }
    };
  }
});

function setStatus(message, type) {
  var el = document.getElementById("status");
  el.textContent = message;
  el.className = "status-message status-" + type;
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

function runConcat() {
  var connectorInput = document.getElementById("connector");
  var executeBtn = document.getElementById("executeBtn");
  var connector = connectorInput.value || "_";

  // 禁用按钮，显示加载状态
  executeBtn.disabled = true;
  setStatus("处理中...", "loading");

  Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    range.load(["address", "columnCount", "columnIndex"]);
    return context.sync().then(function () {
      if (range.columnCount < 2) {
        executeBtn.disabled = false;
        setStatus("错误: 请至少选择两列", "error");
        return;
      }

      var colCount = range.columnCount;
      var colIndex = range.columnIndex;
      var worksheet = context.workbook.worksheets.getActiveWorksheet();

      // Build columns array: [getColumnLetter(colIndex), getColumnLetter(colIndex+1), ...]
      var columns = [];
      for (var i = 0; i < colCount; i++) {
        columns.push(getColumnLetter(colIndex + i));
      }

      var usedInSelection = range.getUsedRange();
      usedInSelection.load("rowCount");
      return context.sync().then(function () {
        var rowCount = usedInSelection.rowCount;

        if (rowCount === 0) {
          executeBtn.disabled = false;
          setStatus("错误: 没有数据", "error");
          return;
        }

        if (rowCount > MAX_ROWS) {
          executeBtn.disabled = false;
          setStatus(
            "错误: 数据量过大（" + rowCount + " 行），单次最多支持 " + MAX_ROWS + " 行。",
            "error"
          );
          return;
        }

        var targetColLetter = getColumnLetter(colIndex + colCount);

        // 超过3列显示确认框
        if (colCount > 3) {
          var startColLetter = getColumnLetter(colIndex);
          var endColLetter = getColumnLetter(colIndex + colCount - 1);
          var confirmMsg =
            "将连接第 " +
            startColLetter +
            " 列到第 " +
            endColLetter +
            " 列，使用连接符【" +
            connector +
            "】";

          // 显示确认框并等待用户响应
          return showConfirmBox(confirmMsg).then(function (confirmed) {
            if (!confirmed) {
              executeBtn.disabled = false;
              setStatus("状态：等待操作...", "idle");
              return;
            }
            // 用户确认后，重新启动 Excel.run 执行连接
            return executeConcatWithConfirm(
              columns,
              targetColLetter,
              connector,
              rowCount,
              executeBtn
            );
          });
        }

        // 直接执行
        return executeConcatDirect(columns, targetColLetter, connector, rowCount, executeBtn);
      });
    });
  }).catch(function (error) {
    setStatus("错误: " + error.message, "error");
    executeBtn.disabled = false;
  });
}

function executeConcatDirect(columns, targetColLetter, connector, rowCount, executeBtn) {
  return Excel.run(function (context) {
    var worksheet = context.workbook.worksheets.getActiveWorksheet();

    worksheet
      .getRange(targetColLetter + ":" + targetColLetter)
      .insert(Excel.InsertShiftDirection.right);
    return context.sync().then(function () {
      var formula = concatUtils.buildNConcatFormula(columns, connector);

      var startCell = worksheet.getRange(targetColLetter + "1");
      startCell.formulas = [[formula]];
      return context
        .sync()
        .then(function () {
          if (rowCount > 1) {
            var fillRange = worksheet.getRange(targetColLetter + "1:" + targetColLetter + rowCount);
            startCell.autoFill(fillRange, Excel.AutoFillType.fillDefault);
            return context.sync();
          }
        })
        .then(function () {
          setStatus(
            "完成! 已在第 " + targetColLetter + " 列写入 " + rowCount + " 行公式",
            "success"
          );
          executeBtn.disabled = false;
        });
    });
  }).catch(function (error) {
    setStatus("错误: " + error.message, "error");
    executeBtn.disabled = false;
  });
}

function executeConcatWithConfirm(columns, targetColLetter, connector, rowCount, executeBtn) {
  return executeConcatDirect(columns, targetColLetter, connector, rowCount, executeBtn);
}
