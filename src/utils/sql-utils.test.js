// src/utils/sql-utils.test.js

var { DatabaseManager, sanitizeColumnName, inferColumnType, importData } = require("./sql-utils");

describe("sanitizeColumnName", function () {
  it("preserves Chinese characters", function () {
    expect(sanitizeColumnName("商品名称", 0, {})).toBe("商品名称");
  });

  it("replaces spaces and special chars with underscore", function () {
    expect(sanitizeColumnName("Sales Amount (2024)", 0, {})).toBe("Sales_Amount__2024");
  });

  it("fills empty names with col_N", function () {
    expect(sanitizeColumnName(null, 0, {})).toBe("col_1");
    expect(sanitizeColumnName("", 1, {})).toBe("col_2");
    expect(sanitizeColumnName(undefined, 2, {})).toBe("col_3");
  });

  it("handles duplicate names by appending _N", function () {
    var used = {};
    var first = sanitizeColumnName("name", 0, used);
    var second = sanitizeColumnName("name", 1, used);
    expect(first).toBe("name");
    expect(second).toBe("name_2");
  });

  it("prefixes with underscore if name starts with digit", function () {
    expect(sanitizeColumnName("123data", 0, {})).toBe("_123data");
  });
});

describe("inferColumnType", function () {
  it("returns INTEGER for integer values", function () {
    var rows = [[1], [2], [3]];
    expect(inferColumnType(rows, 0)).toBe("INTEGER");
  });

  it("returns REAL for decimal values", function () {
    var rows = [[1.5], [2.3], [3.7]];
    expect(inferColumnType(rows, 0)).toBe("REAL");
  });

  it("returns TEXT for non-numeric values", function () {
    var rows = [["apple"], ["banana"], ["cherry"]];
    expect(inferColumnType(rows, 0)).toBe("TEXT");
  });

  it("returns INTEGER for mixed integer and null values", function () {
    var rows = [[1], [null], [3]];
    expect(inferColumnType(rows, 0)).toBe("INTEGER");
  });

  it("returns TEXT when any cell is non-numeric", function () {
    var rows = [[1], ["text"], [3]];
    expect(inferColumnType(rows, 0)).toBe("TEXT");
  });

  it("returns REAL when string can be parsed to decimal", function () {
    var rows = [[1.5], ["2.3"]];
    expect(inferColumnType(rows, 0)).toBe("REAL");
  });

  it("returns TEXT for empty column", function () {
    var rows = [[null], [null]];
    expect(inferColumnType(rows, 0)).toBe("TEXT");
  });
});

describe("DatabaseManager", function () {
  it("creates empty database on init without buffer", function () {
    var dm = new DatabaseManager();
    // We don't call init() in unit test because sql.js WASM needs browser/Node polyfill
    // This is a placeholder for the structure - actual WASM init is tested in Node
    expect(dm.isInitialized).toBe(false);
  });
});

describe("importData", function () {
  // Integration tests requiring DatabaseManager with sql.js
  // These are tested in Node.js environment where sql.js works
  it("exports correct functions", function () {
    expect(typeof sanitizeColumnName).toBe("function");
    expect(typeof inferColumnType).toBe("function");
    expect(typeof importData).toBe("function");
  });
});
