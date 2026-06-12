// src/utils/sql-utils.js

var initSqlJs = require("sql.js");

/**
 * DatabaseManager - 管理 sql.js 数据库实例
 */
function DatabaseManager() {
  this.db = null;
  this.SQL = null;
  this.isInitialized = false;
}

/**
 * 初始化 sql.js WASM 引擎
 * @param {Uint8Array} [existingBuffer] - 可选的已有数据库 buffer
 * @returns {Promise<void>}
 */
DatabaseManager.prototype.init = function (existingBuffer) {
  var self = this;
  return initSqlJs({
    locateFile: function (file) { return "assets/" + file; }
  }).then(function (SQL) {
    self.SQL = SQL;
    if (existingBuffer) {
      self.db = new SQL.Database(existingBuffer);
    } else {
      self.db = new SQL.Database();
    }
    self.isInitialized = true;
  });
};

/**
 * 执行 SQL 语句（支持多语句）
 * @param {string} sql - SQL 语句
 * @returns {{ type: string, columns?: string[], rows?: any[][], rowCount?: number, rowsAffected?: number, elapsed: number, message?: string }}
 */
DatabaseManager.prototype.exec = function (sql) {
  if (!this.isInitialized || !this.db) {
    return { type: "error", message: "数据库未初始化", elapsed: 0 };
  }

  try {
    var startTime = performance.now();
    var results = this.db.exec(sql);
    var elapsed = performance.now() - startTime;

    if (results.length === 0) {
      // 非 SELECT 语句
      var rowsAffected = this.db.getRowsModified();
      return {
        type: "modification",
        rowsAffected: rowsAffected,
        elapsed: elapsed,
      };
    }

    // SELECT 结果
    return {
      type: "query",
      columns: results[0].columns,
      rows: results[0].values,
      rowCount: results[0].values.length,
      elapsed: elapsed,
    };
  } catch (e) {
    return { type: "error", message: e.message, elapsed: 0 };
  }
};

/**
 * 导出数据库为 Uint8Array
 * @returns {Uint8Array}
 */
DatabaseManager.prototype.exportBuffer = function () {
  if (!this.db) return null;
  return this.db.export();
};

/**
 * 从 Uint8Array 加载数据库
 * @param {Uint8Array} buffer
 */
DatabaseManager.prototype.importBuffer = function (buffer) {
  if (!this.SQL) return;
  if (this.db) this.db.close();
  this.db = new this.SQL.Database(buffer);
  this.isInitialized = true;
};

/**
 * 获取表列表
 * @returns {Array<{name: string, sql: string}>}
 */
DatabaseManager.prototype.getTables = function () {
  var result = this.exec(
    "SELECT name, sql FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY name"
  );
  if (result.type === "query") {
    return result.rows.map(function (row) {
      return { name: row[0], sql: row[1] };
    });
  }
  return [];
};

/**
 * 获取表结构（列名和类型）
 * @param {string} tableName
 * @returns {Array<{name: string, type: string}>}
 */
DatabaseManager.prototype.getTableSchema = function (tableName) {
  var result = this.exec("PRAGMA table_info(" + tableName + ")");
  if (result.type === "query") {
    return result.rows.map(function (row) {
      return { name: row[1], type: row[2] };
    });
  }
  return [];
};

/**
 * 获取表行数
 * @param {string} tableName
 * @returns {number}
 */
DatabaseManager.prototype.getTableRowCount = function (tableName) {
  var result = this.exec("SELECT COUNT(*) FROM " + tableName);
  if (result.type === "query" && result.rows.length > 0) {
    return result.rows[0][0];
  }
  return 0;
};

/**
 * 预览表数据
 * @param {string} tableName
 * @param {number} [limit=10]
 * @returns {*}
 */
DatabaseManager.prototype.previewTable = function (tableName, limit) {
  if (limit === undefined || limit === null) limit = 10;
  return this.exec("SELECT * FROM " + tableName + " LIMIT " + limit);
};

/**
 * 清理列名：替换特殊字符、处理空列名、处理重复
 * @param {string} name - 原始列名
 * @param {number} index - 列索引（从0开始）
 * @param {Object} usedNames - 已使用的列名集合
 * @returns {string}
 */
function sanitizeColumnName(name, index, usedNames) {
  if (name === null || name === undefined || name === "") {
    name = "col_" + (index + 1);
  } else {
    name = String(name);
    // 替换空格和特殊字符为下划线
    name = name.replace(/[^a-zA-Z0-9_一-鿿]/g, "_");
    // 去掉开头和结尾的下划线
    name = name.replace(/^_+|_+$/g, "");
    // 如果清空后为空，使用默认名
    if (name === "") {
      name = "col_" + (index + 1);
    }
    // 不能以数字开头，加下划线前缀
    if (/^[0-9]/.test(name)) {
      name = "_" + name;
    }
  }

  // 处理重复
  if (usedNames) {
    if (usedNames[name]) {
      var counter = 2;
      while (usedNames[name + "_" + counter]) {
        counter++;
      }
      name = name + "_" + counter;
    }
    usedNames[name] = true;
  }

  return name;
}

/**
 * 推断列类型
 * @param {Array<Array>} rows - 数据行（不含表头）
 * @param {number} colIndex - 列索引
 * @returns {string} "INTEGER" | "REAL" | "TEXT"
 */
function inferColumnType(rows, colIndex) {
  var allInteger = true;
  var allNumber = true;
  var hasNonNull = false;

  for (var i = 0; i < rows.length; i++) {
    var val = rows[i][colIndex];
    if (val === null || val === undefined || val === "") {
      continue;
    }
    hasNonNull = true;

    if (typeof val === "number") {
      if (allInteger && !Number.isInteger(val)) {
        allInteger = false;
      }
    } else {
      // 尝试转为数字
      var num = Number(val);
      if (isNaN(num)) {
        allInteger = false;
        allNumber = false;
        break;
      }
      if (allInteger && !Number.isInteger(num)) {
        allInteger = false;
      }
    }
  }

  if (!hasNonNull) return "TEXT";
  if (allInteger) return "INTEGER";
  if (allNumber) return "REAL";
  return "TEXT";
}

/**
 * 将 Excel 数据导入 SQLite 表
 * @param {DatabaseManager} dbManager - 数据库管理器
 * @param {string} tableName - 表名
 * @param {Array<Array>} values - 2D 数据数组
 * @param {boolean} firstRowIsHeader - 第一行是否为列名
 * @returns {{ success: boolean, tableName: string, rowsInserted: number, message: string }}
 */
function importData(dbManager, tableName, values, firstRowIsHeader) {
  if (!values || values.length === 0) {
    return { success: false, tableName: tableName, rowsInserted: 0, message: "没有数据" };
  }

  var dataStartIndex = firstRowIsHeader ? 1 : 0;
  var columns = [];
  var usedNames = {};

  // 提取/生成列名
  if (firstRowIsHeader && values.length > 0) {
    var headerRow = values[0];
    for (var c = 0; c < headerRow.length; c++) {
      columns.push(sanitizeColumnName(headerRow[c], c, usedNames));
    }
  } else {
    for (var c2 = 0; c2 < values[0].length; c2++) {
      columns.push(sanitizeColumnName(null, c2, usedNames));
    }
  }

  // 获取数据行
  var dataRows = [];
  for (var r = dataStartIndex; r < values.length; r++) {
    var row = values[r];
    if (row !== null && Array.isArray(row) && row.length > 0) {
      dataRows.push(row);
    }
  }

  if (dataRows.length === 0) {
    return { success: false, tableName: tableName, rowsInserted: 0, message: "没有数据行" };
  }

  // 推断列类型
  var columnTypes = [];
  for (var c3 = 0; c3 < columns.length; c3++) {
    columnTypes.push(inferColumnType(dataRows, c3));
  }

  // 建表
  var colDefs = [];
  for (var c4 = 0; c4 < columns.length; c4++) {
    colDefs.push('"' + columns[c4] + '" ' + columnTypes[c4]);
  }

  var createSql = "CREATE TABLE IF NOT EXISTS \"" + tableName + "\" (" + colDefs.join(", ") + ")";
  dbManager.exec("DROP TABLE IF EXISTS \"" + tableName + "\"");
  dbManager.exec(createSql);

  // 批量插入（每 500 行一批）
  var BATCH_SIZE = 500;
  var totalInserted = 0;
  var placeholders = columns.map(function () { return "?"; }).join(", ");
  var insertSql = "INSERT INTO \"" + tableName + "\" VALUES (" + placeholders + ")";

  dbManager.exec("BEGIN TRANSACTION");
  try {
    for (var r2 = 0; r2 < dataRows.length; r2++) {
      var dataRow = dataRows[r2];
      // 补齐长度不足的行
      while (dataRow.length < columns.length) {
        dataRow.push(null);
      }
      // 只取前 columns.length 列
      var trimmedRow = dataRow.slice(0, columns.length);

      dbManager.db.run(insertSql, trimmedRow);
      totalInserted++;

      if (totalInserted % BATCH_SIZE === 0) {
        dbManager.exec("COMMIT");
        dbManager.exec("BEGIN TRANSACTION");
      }
    }
    dbManager.exec("COMMIT");
  } catch (e) {
    dbManager.exec("ROLLBACK");
    return { success: false, tableName: tableName, rowsInserted: totalInserted, message: e.message };
  }

  return {
    success: true,
    tableName: tableName,
    rowsInserted: totalInserted,
    message: "成功导入 " + totalInserted + " 行数据到表 " + tableName,
  };
}

module.exports = {
  DatabaseManager: DatabaseManager,
  sanitizeColumnName: sanitizeColumnName,
  inferColumnType: inferColumnType,
  importData: importData,
};
