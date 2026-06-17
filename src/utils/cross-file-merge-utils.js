/**
 * 获取 Excel 列字母（A-Z, AA-ZZ, ...）
 * @param {number} colIndex - 0-based column index
 * @returns {string} Column letter(s)
 */
function getColumnLetter(colIndex) {
  var letter = "";
  var remaining = colIndex;
  do {
    letter = String.fromCharCode(65 + (remaining % 26)) + letter;
    remaining = Math.floor(remaining / 26) - 1;
  } while (remaining >= 0);
  return letter;
}

/**
 * 生成唯一的 Sheet 名称
 * @param {string} baseName - 基础名称
 * @param {string[]} existingNames - 已存在的 Sheet 名称数组
 * @returns {string} 唯一可用名称
 */
function generateUniqueSheetName(baseName, existingNames) {
  if (existingNames.indexOf(baseName) === -1) {
    return baseName;
  }
  for (var i = 1; i <= 100; i++) {
    var name = baseName + "_" + i;
    if (existingNames.indexOf(name) === -1) {
      return name;
    }
  }
  return baseName + "_" + Date.now();
}

/**
 * 使用 SheetJS 解析 Excel 文件的当前激活 sheet
 * @param {File} file - Browser File 对象
 * @returns {Promise<{data: Array, sheetName: string}>}
 */
function parseExcelFile(file) {
  return new Promise(function (resolve, reject) {
    var reader = new FileReader();
    reader.onload = function (e) {
      try {
        var data = new Uint8Array(e.target.result);
        var workbook = XLSX.read(data, { type: "array" });
        var sheetName = workbook.SheetNames[0]; // 当前激活 sheet
        var sheet = workbook.Sheets[sheetName];

        // 防御性检查：工作表不存在
        if (!sheet) {
          return reject(new Error("工作表 '" + sheetName + "' 不存在或为空"));
        }
        var jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 }); // 2D array
        resolve({ data: jsonData, sheetName: sheetName });
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = function (e) {
      reject(new Error("文件读取失败"));
    };
    reader.readAsArrayBuffer(file);
  });
}

/**
 * 验证所有文件的列数是否一致
 * @param {Array<{name: string, columnCount: number}>} files
 * @returns {{valid: boolean, error: string|null}}
 */
function validateColumnConsistency(files) {
  if (!files || files.length < 2) {
    return { valid: false, error: "请至少选择两个文件" };
  }
  var expected = files[0].columnCount;
  for (var i = 1; i < files.length; i++) {
    if (files[i].columnCount !== expected) {
      return {
        valid: false,
        error:
          "文件 '" +
          files[i].name +
          "' 列数(" +
          files[i].columnCount +
          ")与第一个文件(" +
          expected +
          ")不一致",
      };
    }
  }
  return { valid: true, error: null };
}

/**
 * 合并多个文件的数据
 * @param {Array<{data: Array, name: string}>} fileDataList - 文件数据列表
 * @param {number} headerRowNumber - 表头行号（0=无表头）
 * @returns {{mergedData: Array, columnCount: number}}
 */
function mergeExcelData(fileDataList, headerRowNumber) {
  var mergedData = [];
  var hasHeader = headerRowNumber > 0;
  var columnCount = 0;

  for (var i = 0; i < fileDataList.length; i++) {
    var fileData = fileDataList[i];
    var data = fileData.data;
    var fileName = fileData.name;

    if (!data || data.length === 0) {
      continue; // 跳过空文件
    }

    var headerRowIndex = hasHeader ? headerRowNumber - 1 : -1;

    // 第一文件：取表头
    if (i === 0) {
      if (hasHeader && headerRowIndex >= 0 && headerRowIndex < data.length) {
        var headerRow = data[headerRowIndex].slice();
        headerRow.unshift("来源文件");   // 第二列：来源文件
        headerRow.unshift("Sheet名");     // 第一列：Sheet名
        mergedData.push(headerRow);
        columnCount = headerRow.length;
      }
      // 添加第一文件的数据行
      var startRow = hasHeader ? headerRowIndex + 1 : 0;
      for (var r = startRow; r < data.length; r++) {
        var dataRow = data[r].slice();
        dataRow.unshift(fileName);              // 第二列：来源文件
        dataRow.unshift(fileData.sheetName);   // 第一列：Sheet名
        mergedData.push(dataRow);
      }
    } else {
      // 后续文件：跳过表头（如果存在）
      var startRow = hasHeader ? headerRowIndex + 1 : 0;
      for (var r2 = startRow; r2 < data.length; r2++) {
        var dataRow2 = data[r2].slice();
        dataRow2.unshift(fileName);              // 第二列：来源文件
        dataRow2.unshift(fileData.sheetName);    // 第一列：Sheet名
        mergedData.push(dataRow2);
      }
    }
  }

  return { mergedData: mergedData, columnCount: columnCount };
}

module.exports = {
  getColumnLetter: getColumnLetter,
  generateUniqueSheetName: generateUniqueSheetName,
  parseExcelFile: parseExcelFile,
  validateColumnConsistency: validateColumnConsistency,
  mergeExcelData: mergeExcelData,
};
