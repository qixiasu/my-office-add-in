var {
  parseCSV,
  createParser,
} = require("./csv-utils");

// ── parseCSV (unchanged, for backward compat / small files) ──────────

describe("parseCSV", function () {
  it("parses simple comma-separated values", function () {
    var result = parseCSV("a,b,c\nd,e,f");
    expect(result).toEqual([
      ["a", "b", "c"],
      ["d", "e", "f"],
    ]);
  });

  it("parses single row", function () {
    expect(parseCSV("a,b,c")).toEqual([["a", "b", "c"]]);
  });

  it("handles empty input", function () {
    expect(parseCSV("")).toEqual([]);
  });

  it("handles quoted fields with embedded commas", function () {
    var result = parseCSV('"a,b",c\nd,"e,f"');
    expect(result).toEqual([
      ["a,b", "c"],
      ["d", "e,f"],
    ]);
  });

  it("handles doubled quotes inside quoted fields", function () {
    var result = parseCSV('"a""b",c');
    expect(result).toEqual([['a"b', "c"]]);
  });

  it("handles embedded newlines inside quoted fields", function () {
    var result = parseCSV('"multi\nline",b\nc,d');
    expect(result).toEqual([
      ["multi\nline", "b"],
      ["c", "d"],
    ]);
  });

  it("handles CRLF line endings", function () {
    var result = parseCSV("a,b\r\nc,d");
    expect(result).toEqual([
      ["a", "b"],
      ["c", "d"],
    ]);
  });

  it("strips UTF-8 BOM", function () {
    var text = "﻿" + "a,b\nc,d";
    var result = parseCSV(text);
    expect(result).toEqual([
      ["a", "b"],
      ["c", "d"],
    ]);
  });

  it("handles Chinese characters", function () {
    var result = parseCSV("姓名,年龄\n张三,25\n李四,30");
    expect(result).toEqual([
      ["姓名", "年龄"],
      ["张三", "25"],
      ["李四", "30"],
    ]);
  });

  it("handles quoted fields with Chinese characters and commas", function () {
    var result = parseCSV('"北京, 朝阳区",100\n"上海, 浦东新区",200');
    expect(result).toEqual([
      ["北京, 朝阳区", "100"],
      ["上海, 浦东新区", "200"],
    ]);
  });

  it("supports tab delimiter", function () {
    var result = parseCSV("a\tb\tc\nd\te\tf", "\t");
    expect(result).toEqual([
      ["a", "b", "c"],
      ["d", "e", "f"],
    ]);
  });

  it("supports semicolon delimiter", function () {
    var result = parseCSV("a;b;c\nd;e;f", ";");
    expect(result).toEqual([
      ["a", "b", "c"],
      ["d", "e", "f"],
    ]);
  });

  it("supports pipe delimiter", function () {
    var result = parseCSV("a|b|c\nd|e|f", "|");
    expect(result).toEqual([
      ["a", "b", "c"],
      ["d", "e", "f"],
    ]);
  });

  it("defaults to comma when no delimiter specified", function () {
    var result = parseCSV("a,b,c");
    expect(result).toEqual([["a", "b", "c"]]);
  });
});

// ── Helper: feed a full text string through the incremental parser ───

function parseText(text, delimiter, batchSize, onBatch) {
  var parser = createParser(delimiter, batchSize, onBatch, false);
  parser.feed(text);
  return parser.finish();
}

function scanText(text, delimiter) {
  var parser = createParser(delimiter, 0, null, true);
  parser.feed(text);
  return parser.finish();
}

// ── createParser (scanOnly = true) — row/col counting ─────────────────

describe("createParser (scan mode)", function () {
  it("counts rows and columns for simple CSV", function () {
    var result = scanText("a,b,c\nd,e,f\ng,h,i");
    expect(result).toEqual({ rowCount: 3, colCount: 3 });
  });

  it("handles single row", function () {
    var result = scanText("a,b,c");
    expect(result).toEqual({ rowCount: 1, colCount: 3 });
  });

  it("handles empty input", function () {
    var result = scanText("");
    expect(result).toEqual({ rowCount: 0, colCount: 0 });
  });

  it("counts rows with quoted multiline fields correctly", function () {
    var result = scanText('"multi\nline",b\nc,d');
    expect(result.rowCount).toBe(2);
    expect(result.colCount).toBe(2);
  });

  it("handles CRLF line endings", function () {
    var result = scanText("a,b\r\nc,d\r\ne,f");
    expect(result).toEqual({ rowCount: 3, colCount: 2 });
  });

  it("handles trailing newline", function () {
    var result = scanText("a,b\n");
    expect(result.rowCount).toBe(1);
  });

  it("supports tab delimiter", function () {
    var result = scanText("a\tb\tc\nd\te\tf", "\t");
    expect(result).toEqual({ rowCount: 2, colCount: 3 });
  });

  it("handles BOM in first chunk", function () {
    var text = "﻿" + "a,b\nc,d";
    var result = scanText(text);
    expect(result).toEqual({ rowCount: 2, colCount: 2 });
  });
});

// ── createParser (scanOnly = false) — chunked parsing ─────────────────

describe("createParser (parse mode)", function () {
  it("parses simple CSV in batches", function () {
    var batches = [];
    var stats = parseText("a,b,c\nd,e,f\ng,h,i", ",", 2, function (rows, idx) {
      batches.push({ idx: idx, rows: rows });
    });
    expect(stats.rowCount).toBe(3);
    expect(stats.colCount).toBe(3);
    expect(batches.length).toBe(2);
    expect(batches[0].rows).toEqual([["a", "b", "c"], ["d", "e", "f"]]);
    expect(batches[1].rows).toEqual([["g", "h", "i"]]);
  });

  it("pads rows to max column count", function () {
    var batches = [];
    parseText("a,b,c\nd,e", ",", 2, function (rows) {
      batches.push(rows);
    });
    expect(batches[0][0]).toEqual(["a", "b", "c"]);
    expect(batches[0][1]).toEqual(["d", "e", ""]);
  });

  it("handles single row with batch size 1", function () {
    var batches = [];
    var stats = parseText("a,b,c", ",", 1, function (rows) {
      batches.push(rows);
    });
    expect(stats.rowCount).toBe(1);
    expect(batches.length).toBe(1);
    expect(batches[0]).toEqual([["a", "b", "c"]]);
  });

  it("handles empty input", function () {
    var called = false;
    var stats = parseText("", ",", 100, function () { called = true; });
    expect(stats.rowCount).toBe(0);
    expect(called).toBe(false);
  });

  it("handles quoted fields with embedded delimiters", function () {
    var batches = [];
    parseText('"a,b",c\nd,"e,f"', ",", 2, function (rows) {
      batches.push(rows);
    });
    expect(batches[0]).toEqual([
      ["a,b", "c"],
      ["d", "e,f"],
    ]);
  });

  it("handles doubled quotes in quoted fields", function () {
    var batches = [];
    parseText('"a""b",c', ",", 1, function (rows) {
      batches.push(rows);
    });
    expect(batches[0][0]).toEqual(['a"b', "c"]);
  });

  it("handles embedded newlines in quoted fields", function () {
    var batches = [];
    var stats = parseText('"multi\nline",b\nc,d', ",", 2, function (rows) {
      batches.push(rows);
    });
    expect(stats.rowCount).toBe(2);
    expect(stats.colCount).toBe(2);
    expect(batches[0]).toEqual([
      ["multi\nline", "b"],
      ["c", "d"],
    ]);
  });

  it("handles Chinese characters", function () {
    var batches = [];
    parseText("姓名,年龄\n张三,25", ",", 2, function (rows) {
      batches.push(rows);
    });
    expect(batches[0]).toEqual([
      ["姓名", "年龄"],
      ["张三", "25"],
    ]);
  });

  it("supports tab delimiter", function () {
    var batches = [];
    parseText("a\tb\tc\nd\te\tf", "\t", 2, function (rows) {
      batches.push(rows);
    });
    expect(batches[0]).toEqual([
      ["a", "b", "c"],
      ["d", "e", "f"],
    ]);
  });
});

// ── createParser — chunk‑boundary safety ──────────────────────────────

describe("createParser (chunk boundary safety)", function () {
  it("handles row split across chunks", function () {
    var batches = [];
    var parser = createParser(",", 10, function (rows, idx) {
      batches.push({ idx: idx, rows: rows });
    }, false);

    // Split "a,b,c\nd,e,f" across two chunks
    parser.feed("a,b");
    parser.feed(",c\nd,e,f");

    var stats = parser.finish();
    expect(stats.rowCount).toBe(2);
    expect(stats.colCount).toBe(3);
    expect(batches[0].rows).toEqual([
      ["a", "b", "c"],
      ["d", "e", "f"],
    ]);
  });

  it("handles field split across chunks", function () {
    var batches = [];
    var parser = createParser(",", 10, function (rows) {
      batches.push(rows);
    }, false);

    parser.feed('"multi');
    parser.feed('line",b\nc,d');

    var stats = parser.finish();
    expect(stats.rowCount).toBe(2);
    expect(stats.colCount).toBe(2);
    expect(batches[0]).toEqual([
      ["multiline", "b"],
      ["c", "d"],
    ]);
  });

  it("handles \\r\\n split across chunks", function () {
    var batches = [];
    var parser = createParser(",", 10, function (rows) {
      batches.push(rows);
    }, false);

    parser.feed("a,b\r");
    parser.feed("\nc,d");

    var stats = parser.finish();
    expect(stats.rowCount).toBe(2);
    expect(stats.colCount).toBe(2);
    expect(batches[0]).toEqual([
      ["a", "b"],
      ["c", "d"],
    ]);
  });

  it("handles quoted field with escaped quote across chunks", function () {
    var batches = [];
    var parser = createParser(",", 10, function (rows) {
      batches.push(rows);
    }, false);

    // Split inside a quoted field containing ""
    parser.feed('"a""');
    parser.feed('b",c');

    var stats = parser.finish();
    expect(stats.rowCount).toBe(1);
    expect(batches[0][0]).toEqual(['a"b', "c"]);
  });
});
