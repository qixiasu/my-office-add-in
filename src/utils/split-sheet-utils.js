/**
 * Utility functions for splitting data by key column
 */

/**
 * Group data rows by key column value
 * @param {Array<Array>} data - 2D array (data rows including header at index 0)
 * @param {number} keyColumnIndex - 0-based index of the key column
 * @returns {{ groups: Object|null, error: string|null }}
 *   On success: { groups: { [key]: { header: [...], rows: [...] } }, error: null }
 *   On error: { groups: null, error: "错误信息" }
 */
function groupDataByKey(data, keyColumnIndex) {
  if (!data || data.length < 2) {
    return { groups: null, error: "数据不能为空且至少需要包含表头和一行数据" };
  }

  if (keyColumnIndex < 0 || keyColumnIndex >= data[0].length) {
    return { groups: null, error: "键列索引超出范围" };
  }

  var header = data[0];
  var groups = {};

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var keyValue = row[keyColumnIndex];

    // Handle null/undefined as empty string
    if (keyValue === null || keyValue === undefined) {
      keyValue = "";
    }

    // Trim the key value
    var trimmedKey = String(keyValue).trim();

    // Empty or whitespace-only keys are not allowed
    if (trimmedKey === "") {
      return { groups: null, error: "键值不能为空" };
    }

    if (!groups[trimmedKey]) {
      groups[trimmedKey] = { header: header, rows: [] };
    }
    groups[trimmedKey].rows.push(row);
  }

  return { groups: groups, error: null };
}

/**
 * Truncate sheet name to Excel's 31 character limit
 * If truncation needed, append "..." and fit as much as possible
 * @param {string} name - Original sheet name
 * @returns {string} Truncated sheet name
 */
function truncateSheetName(name) {
  var MAX_LENGTH = 31;

  if (!name || name.length <= MAX_LENGTH) {
    return name || "";
  }

  // We need to fit: truncated_name + "..."
  // So we reserve 3 characters for "..."
  var availableLength = MAX_LENGTH - 3;

  if (availableLength <= 0) {
    // Edge case: name is so long we can only fit "..." and a few chars
    return "...".substring(0, MAX_LENGTH);
  }

  return name.substring(0, availableLength) + "...";
}

module.exports = {
  groupDataByKey: groupDataByKey,
  truncateSheetName: truncateSheetName,
};