// src/utils/expand-utils.js

/**
 * 展开宽表数据为长表格式
 * @param {Array<Array>} values - 2D数组，第一行是表头
 * @returns {Array<Array>} 展开后的2D数组，每行 [键列值, 展开列值]
 */
function expandData(values) {
  if (!values || values.length <= 1) {
    return [];
  }

  var result = [];

  // 从第1行开始是数据（第0行是表头）
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    // 跳过空行（当选择整列时 values[i] 可能为 null）
    if (!row || !Array.isArray(row)) {
      continue;
    }
    var key = row[0];
    // 跳过键列为空的行
    if (key === null || key === "" || key === undefined) {
      continue;
    }

    // 遍历其他列（从索引1开始）
    for (var j = 1; j < row.length; j++) {
      var val = row[j];
      // 跳过空单元格
      if (val !== null && val !== "" && val !== undefined) {
        result.push([key, val]);
      }
    }
  }

  return result;
}

module.exports = {
  expandData: expandData,
};