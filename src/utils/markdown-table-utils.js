/**
 * Markdown 表格转换工具
 * 将 Excel 区域数据转换为 GitHub 风格 Markdown 表格
 */

/**
 * 生成 Markdown 表格
 * @param {Array<Array>} values - 二维单元格值数组
 * @param {Object} options - 配置选项
 * @param {Array} options.mergedRanges - Excel 合并区域数组（默认 []）
 * @param {boolean} options.includeAlignment - 是否包含对齐标记（默认 true）
 * @param {boolean} options.preserveFormat - 是否保留显示格式（默认 true）
 * @returns {string} Markdown 表格字符串
 */
function generateMarkdownTable(values, options = {}) {
  const { mergedRanges = [], includeAlignment = true, preserveFormat = true } = options;

  if (!values || !values.length) return "";

  // 过滤空行空列
  const cleaned = parseRange(values);
  if (!cleaned.length) return "";

  // 检测每列对齐方式
  const alignment = detectAlignment(cleaned, includeAlignment);

  // 处理合并单元格
  const mergedInfo =
    mergedRanges.length > 0
      ? handleMergedCells(mergedRanges, cleaned.length, cleaned[0].length)
      : new Map();

  // 构建表格字符串
  return buildTableString(cleaned, alignment, mergedInfo);
}

/**
 * 处理合并单元格，标记需要 rowspan/colspan 的位置
 * @param {Array} mergedRanges - Excel merged ranges
 * @param {number} rowCount - total rows
 * @param {number} colCount - total columns
 * @returns {Map} key: "row-col", value: { rowspan, colspan, isPrimary }
 */
function handleMergedCells(mergedRanges, rowCount, colCount) {
  const mergedInfo = new Map();

  if (!mergedRanges || !mergedRanges.length) return mergedInfo;

  for (const range of mergedRanges) {
    const rangeStr = range.toString(); // e.g., "A1:C3"
    const parsed = parseMergedRangeAddress(rangeStr);
    if (!parsed) continue;

    const { startRow, startCol, endRow, endCol } = parsed;
    const rowspan = endRow - startRow + 1;
    const colspan = endCol - startCol + 1;

    // Only mark cells that span multiple rows/cols
    if (rowspan > 1 || colspan > 1) {
      const key = `${startRow}-${startCol}`;
      mergedInfo.set(key, { rowspan, colspan });

      // Mark cells that are covered by merged range but not the primary cell
      for (let r = startRow; r <= endRow; r++) {
        for (let c = startCol; c <= endCol; c++) {
          if (r !== startRow || c !== startCol) {
            mergedInfo.set(`${r}-${c}`, { covered: true });
          }
        }
      }
    }
  }

  return mergedInfo;
}

/**
 * 解析 Excel 合并范围地址
 * @param {string} address - e.g., "A1:C3"
 * @returns {Object} { startRow, startCol, endRow, endCol } (0-indexed)
 */
function parseMergedRangeAddress(address) {
  // Parse the range string to get start and end positions
  // Excel range format: "A1:C3" where A=0, B=1, C=2, etc.
  const match = address.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
  if (!match) return null;

  const colToNum = (col) => {
    let num = 0;
    for (let i = 0; i < col.length; i++) {
      num = num * 26 + (col.charCodeAt(i) - "A".charCodeAt(0) + 1);
    }
    return num - 1; // 0-indexed
  };

  return {
    startCol: colToNum(match[1]),
    startRow: parseInt(match[2]) - 1, // 0-indexed
    endCol: colToNum(match[3]),
    endRow: parseInt(match[4]) - 1, // 0-indexed
  };
}

/**
 * 过滤空行空列
 * @param {Array<Array>} values
 * @returns {Array<Array>}
 */
function parseRange(values) {
  // 过滤全空行
  const filteredRows = values.filter((row) => row.some((cell) => cell !== null && cell !== ""));

  if (!filteredRows.length) return [];

  // 过滤尾随空列
  const colCount = Math.max(...filteredRows.map((row) => row.length));
  const nonEmptyCols = new Set();

  for (let col = 0; col < colCount; col++) {
    for (const row of filteredRows) {
      if (row[col] !== null && row[col] !== "") {
        nonEmptyCols.add(col);
        break;
      }
    }
  }

  return filteredRows.map((row) =>
    Array.from(nonEmptyCols)
      .sort((a, b) => a - b)
      .map((col) => row[col] ?? "")
  );
}

/**
 * 检测列对齐方式
 * @param {Array<Array>} rows
 * @param {boolean} includeAlignment
 * @returns {Array<string>}
 */
function detectAlignment(rows, includeAlignment) {
  if (!includeAlignment || rows.length < 2) {
    return rows[0].map(() => "---");
  }

  // 第一行为表头，不用于类型检测
  const dataRows = rows.slice(1);

  return rows[0].map((_, colIndex) => {
    const values = dataRows.map((row) => row[colIndex]);
    const numericCount = values.filter(
      (v) => typeof v === "number" || (!isNaN(parseFloat(v)) && isFinite(v))
    ).length;
    const totalNonEmpty = values.filter((v) => v !== null && v !== "").length;

    // 如果超过 50% 是数字，则右对齐
    if (totalNonEmpty > 0 && numericCount / totalNonEmpty > 0.5) {
      return "---:";
    }
    return ":---";
  });
}

/**
 * 构建 Markdown 表格字符串
 * @param {Array<Array>} rows
 * @param {Array<string>} alignment
 * @param {Map} mergedInfo - 合并单元格信息
 * @returns {string}
 */
function buildTableString(rows, alignment, mergedInfo = new Map()) {
  const lines = [];
  const separator = "|" + alignment.map((a) => ` ${a} `).join("|") + "|";

  // 表头行
  const headerCells = rows[0].map((cell, colIndex) => {
    const key = `0-${colIndex}`;
    const info = mergedInfo.get(key);
    if (info && !info.covered) {
      const attrs = [];
      if (info.colspan > 1) attrs.push(`<colspan=${info.colspan}>`);
      if (info.rowspan > 1) attrs.push(`<rowspan=${info.rowspan}>`);
      return escapeCell(cell) + attrs.join("");
    }
    return escapeCell(cell);
  });
  lines.push("|" + headerCells.map((c) => ` ${c} `).join("|") + "|");
  lines.push(separator);

  // 数据行
  for (let i = 1; i < rows.length; i++) {
    const rowCells = rows[i]
      .map((cell, colIndex) => {
        const key = `${i}-${colIndex}`;
        const info = mergedInfo.get(key);
        if (info && info.covered) {
          return ""; // Skip covered cells
        }
        if (info && !info.covered) {
          const attrs = [];
          if (info.colspan > 1) attrs.push(`<colspan=${info.colspan}>`);
          if (info.rowspan > 1) attrs.push(`<rowspan=${info.rowspan}>`);
          return escapeCell(cell) + attrs.join("");
        }
        return escapeCell(cell);
      })
      .filter((c, idx, arr) => {
        // Filter out empty cells that are covered
        const key = `${i}-${idx}`;
        const info = mergedInfo.get(key);
        return !info || !info.covered;
      });

    if (rowCells.length > 0) {
      lines.push("|" + rowCells.map((c) => ` ${c} `).join("|") + "|");
    }
  }

  return lines.join("\n");
}

/**
 * 转义单元格内容中的特殊字符
 * @param {*} cell
 * @returns {string}
 */
function escapeCell(cell) {
  if (cell === null || cell === undefined) return "";
  const str = String(cell);
  return str.replace(/\\/g, "\\\\").replace(/\|/g, "\\|").replace(/\n/g, "<br>");
}

export {
  generateMarkdownTable,
  parseRange,
  detectAlignment,
  buildTableString,
  escapeCell,
  handleMergedCells,
  parseMergedRangeAddress,
};
