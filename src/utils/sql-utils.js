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

module.exports = {
  DatabaseManager: DatabaseManager,
};
