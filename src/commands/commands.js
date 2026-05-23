/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

Office.actions.associate("placeholder", async () => {
  // Placeholder for Task 2 implementation
});

/**
 * Shows a prompt dialog to get the connector string from user.
 * @returns {Promise<string>} The connector string (defaults to "_" if cancelled or empty)
 */
function promptConnector() {
  return new Promise((resolve) => {
    const connector = prompt("请输入连接符（默认 _）:", "_");
    resolve(connector === "" ? "_" : connector);
  });
}

/**
 * Concatenates two selected columns with a connector and inserts the result as a new column.
 * User must select at least two columns. The result is placed after the second column.
 * @param {*} event - The action event
 */
Office.actions.associate("concatenateColumns", async function concatenateColumns(event) {
  try {
    await Excel.run(async (context) => {
      // 1. Get selected range
      const range = context.workbook.getSelectedRange();
      range.load(["address", "columnCount", "columnIndex"]);
      await context.sync();

      // 2. Verify at least 2 columns selected
      if (range.columnCount < 2) {
        Office.context.ui.messageBox("请至少选择两列");
        event.completed();
        return;
      }

      // 3. Get connector from user
      const connector = await promptConnector();
      if (connector === null) {
        // User cancelled prompt
        event.completed();
        return;
      }

      // 4. Get first and second column data
      const firstColumn = range.getColumn(0);
      const secondColumn = range.getColumn(1);
      firstColumn.load("values");
      secondColumn.load("values");
      await context.sync();

      // 5. Insert a new column after the second column (shifting existing columns right)
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const columnIndex = range.columnIndex;
      // Create a single-column range at the insert position (right after second column)
      const insertColumnIndex = columnIndex + 2; // +2 because columnIndex is 0-based and we want after the 2nd column
      const insertAddress = `${insertColumnIndex}:${insertColumnIndex}`;
      worksheet.getRange(insertAddress).insert(Excel.InsertShiftDirection.shiftRight);
      await context.sync();

      // 6. Write concatenated data to the new column
      const rowCount = firstColumn.values.length;
      const resultValues = [];
      for (let i = 0; i < rowCount; i++) {
        const val1 = firstColumn.values[i][0] || "";
        const val2 = secondColumn.values[i][0] || "";
        resultValues.push([val1 + connector + val2]);
      }

      // Write to the inserted column (same column index, first row to last row)
      const resultAddress = `${insertColumnIndex}1:${insertColumnIndex}${rowCount}`;
      worksheet.getRange(resultAddress).values = resultValues;
      await context.sync();
    });
  } catch (error) {
    console.error(error);
    Office.context.ui.messageBox("操作失败: " + error.message);
  }
  event.completed();
});
