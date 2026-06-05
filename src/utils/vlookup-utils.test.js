var {
  parseRangeAddress,
  parseCellRef,
  buildColRange,
  buildIndexMatchFormula,
  staticLookup,
  buildLookupIndex,
} = require("./vlookup-utils");

describe("parseCellRef", function () {
  it("parses A1 correctly", function () {
    var result = parseCellRef("A1");
    expect(result.col).toBe(0);
    expect(result.row).toBe(1);
  });

  it("parses Z1 correctly", function () {
    var result = parseCellRef("Z1");
    expect(result.col).toBe(25);
    expect(result.row).toBe(1);
  });

  it("parses AA1 correctly", function () {
    var result = parseCellRef("AA1");
    expect(result.col).toBe(26);
    expect(result.row).toBe(1);
  });

  it("parses AB10 correctly", function () {
    var result = parseCellRef("AB10");
    expect(result.col).toBe(27);
    expect(result.row).toBe(10);
  });
});

describe("parseRangeAddress", function () {
  it("parses simple range without sheet", function () {
    var result = parseRangeAddress("A1:D500");
    expect(result.sheet).toBe("");
    expect(result.startCol).toBe(0);
    expect(result.startRow).toBe(1);
    expect(result.endCol).toBe(3);
    expect(result.endRow).toBe(500);
    expect(result.colCount).toBe(4);
    expect(result.rowCount).toBe(500);
  });

  it("parses range with sheet name", function () {
    var result = parseRangeAddress("Sheet2!A1:D500");
    expect(result.sheet).toBe("Sheet2");
    expect(result.startCol).toBe(0);
    expect(result.startRow).toBe(1);
    expect(result.endCol).toBe(3);
    expect(result.endRow).toBe(500);
    expect(result.colCount).toBe(4);
  });

  it("parses single cell range", function () {
    var result = parseRangeAddress("Sheet2!A1:A1");
    expect(result.sheet).toBe("Sheet2");
    expect(result.colCount).toBe(1);
    expect(result.rowCount).toBe(1);
  });

  it("parses range with sheet name containing special chars", function () {
    var result = parseRangeAddress("My Sheet!B2:E10");
    expect(result.sheet).toBe("My Sheet");
    expect(result.startCol).toBe(1);
    expect(result.startRow).toBe(2);
    expect(result.endCol).toBe(4);
    expect(result.endRow).toBe(10);
    expect(result.colCount).toBe(4);
  });

  it("parses range with quoted sheet name (Excel format for special chars)", function () {
    var result = parseRangeAddress("'My Sheet'!B2:E10");
    expect(result.sheet).toBe("My Sheet");
    expect(result.startCol).toBe(1);
    expect(result.startRow).toBe(2);
    expect(result.endCol).toBe(4);
    expect(result.endRow).toBe(10);
    expect(result.colCount).toBe(4);
  });

  it("parses range with quoted sheet name containing Chinese characters", function () {
    var result = parseRangeAddress("'5G宏站功率及倾角'!C:C");
    expect(result.sheet).toBe("5G宏站功率及倾角");
    expect(result.startCol).toBe(2);
    expect(result.startRow).toBe(1);
    expect(result.endCol).toBe(2);
    expect(result.endRow).toBe(1);
    expect(result.colCount).toBe(1);
    expect(result.rowCount).toBe(1);
  });
});

describe("buildColRange", function () {
  it("builds absolute column range without sheet", function () {
    var parsed = parseRangeAddress("A1:D500");
    var result = buildColRange(parsed, 0);
    expect(result).toBe("$A$1:$A$500");
  });

  it("builds absolute column range with sheet", function () {
    var parsed = parseRangeAddress("Sheet2!A1:D500");
    var result = buildColRange(parsed, 1);
    expect(result).toBe("Sheet2!$B$1:$B$500");
  });

  it("builds range for last column", function () {
    var parsed = parseRangeAddress("Sheet2!A1:D500");
    var result = buildColRange(parsed, 3);
    expect(result).toBe("Sheet2!$D$1:$D$500");
  });
});

describe("buildIndexMatchFormula", function () {
  it("builds formula with exact match", function () {
    var result = buildIndexMatchFormula("A1", "Sheet2!$A$1:$A$500", "Sheet2!$B$1:$B$500", 0);
    expect(result).toBe("=INDEX(Sheet2!$B$1:$B$500, MATCH(A1, Sheet2!$A$1:$A$500, 0))");
  });

  it("builds formula with approximate match", function () {
    var result = buildIndexMatchFormula("A1", "$A$1:$A$500", "$C$1:$C$500", 1);
    expect(result).toBe("=INDEX($C$1:$C$500, MATCH(A1, $A$1:$A$500, 1))");
  });

  it("builds formula with different lookup cell reference", function () {
    var result = buildIndexMatchFormula("B5", "Sheet2!$A$1:$A$500", "Sheet2!$D$1:$D$500", 0);
    expect(result).toBe("=INDEX(Sheet2!$D$1:$D$500, MATCH(B5, Sheet2!$A$1:$A$500, 0))");
  });
});

describe("staticLookup", function () {
  var table = [
    ["张三", "研发部", "A001", 8000],
    ["李四", "市场部", "A002", 6000],
    ["王五", "财务部", "A003", 7000],
  ];

  it("exact match returns correct values", function () {
    var result = staticLookup(["张三", "李四"], table, 0, [1, 3], 0);
    expect(result).toEqual([
      ["研发部", 8000],
      ["市场部", 6000],
    ]);
  });

  it("returns #N/A for not found values", function () {
    var result = staticLookup(["张三", "不存在"], table, 0, [1], 0);
    expect(result).toEqual([["研发部"], ["#N/A"]]);
  });

  it("returns multiple return columns", function () {
    var result = staticLookup(["王五"], table, 0, [1, 2, 3], 0);
    expect(result).toEqual([["财务部", "A003", 7000]]);
  });

  it("handles null lookup value", function () {
    var result = staticLookup([null], table, 0, [1], 0);
    expect(result).toEqual([["#N/A"]]);
  });

  it("handles empty lookup array", function () {
    var result = staticLookup([], table, 0, [1], 0);
    expect(result).toEqual([]);
  });

  it("matches against non-first column", function () {
    var result = staticLookup(["A002"], table, 2, [0, 3], 0);
    expect(result).toEqual([["李四", 6000]]);
  });

  it("returns #N/A for null lookup in approximate mode", function () {
    var numTable = [
      [100, "low"],
      [200, "mid"],
    ];
    var result = staticLookup([null], numTable, 0, [1], 1);
    expect(result).toEqual([["#N/A"]]);
  });

  it("approximate match finds largest value <= lookup", function () {
    var numTable = [
      [100, "low"],
      [200, "mid"],
      [300, "high"],
    ];
    var result = staticLookup([150, 250, 350], numTable, 0, [1], 1);
    expect(result).toEqual([["low"], ["mid"], ["high"]]);
  });

  it("uses indexCache when provided, giving same results", function () {
    var table = [
      ["张三", "研发部", "A001", 8000],
      ["李四", "市场部", "A002", 6000],
      ["王五", "财务部", "A003", 7000],
    ];
    var indexCache = buildLookupIndex(table, 0);

    var resultWithCache = staticLookup(["张三", "李四", "王五"], table, 0, [1, 3], 0, "#N/A", indexCache);
    var resultWithoutCache = staticLookup(["张三", "李四", "王五"], table, 0, [1, 3], 0, "#N/A");

    expect(resultWithCache).toEqual(resultWithoutCache);
  });

  it("ignores indexCache in approximate match mode", function () {
    var numTable = [
      [100, "low"],
      [200, "mid"],
      [300, "high"],
    ];
    var indexCache = buildLookupIndex(numTable, 0);

    var result = staticLookup([150, 250, 350], numTable, 0, [1], 1, "#N/A", indexCache);

    expect(result).toEqual([["low"], ["mid"], ["high"]]);
  });

  it("works with indexCache when lookup value not found", function () {
    var table = [
      ["张三", "研发部"],
      ["李四", "市场部"],
    ];
    var indexCache = buildLookupIndex(table, 0);

    var result = staticLookup(["不存在"], table, 0, [1], 0, "#N/A", indexCache);

    expect(result).toEqual([["#N/A"]]);
  });

  it("approximate match with binary search returns correct results on large table", function () {
    var numTable = [];
    for (var n = 0; n < 1000; n++) {
      numTable.push([n * 10, "val_" + n]);
    }

    var lookupValues = [15, 25, 95, 105, 555, 9995];
    var result = staticLookup(lookupValues, numTable, 0, [1], 1);

    // 15 -> largest <= 15 is 10 -> val_1
    // 25 -> largest <= 25 is 20 -> val_2
    // 95 -> largest <= 95 is 90 -> val_9
    // 105 -> largest <= 105 is 100 -> val_10
    // 555 -> largest <= 555 is 550 -> val_55
    // 9995 -> largest <= 9995 is 9990 -> val_999
    expect(result).toEqual([
      ["val_1"],
      ["val_2"],
      ["val_9"],
      ["val_10"],
      ["val_55"],
      ["val_999"],
    ]);
  });

  it("approximate match returns defaultValue when all values > lookup", function () {
    var numTable = [
      [100, "low"],
      [200, "mid"],
    ];
    var result = staticLookup([50], numTable, 0, [1], 1);
    expect(result).toEqual([["#N/A"]]);
  });

  it("approximate match works with nulls in match column (skipped)", function () {
    var numTable = [
      [null, "skip"],
      [100, "low"],
      [200, "mid"],
    ];
    var result = staticLookup([150], numTable, 0, [1], 1);
    expect(result).toEqual([["low"]]);
  });
});

describe("buildLookupIndex", function () {
  var table = [
    ["张三", "研发部", "A001"],
    ["李四", "市场部", "A002"],
    ["王五", "财务部", "A003"],
  ];

  it("builds index keyed by match column values", function () {
    var index = buildLookupIndex(table, 0);
    expect(index["张三"]).toBe(0);
    expect(index["李四"]).toBe(1);
    expect(index["王五"]).toBe(2);
  });

  it("builds index against non-first column", function () {
    var index = buildLookupIndex(table, 2);
    expect(index["A001"]).toBe(0);
    expect(index["A002"]).toBe(1);
    expect(index["A003"]).toBe(2);
  });

  it("handles null key by converting to empty string", function () {
    var nullTable = [
      [null, "value1"],
      ["key2", "value2"],
    ];
    var index = buildLookupIndex(nullTable, 0);
    expect(index[""]).toBe(0);
    expect(index["key2"]).toBe(1);
  });

  it("later duplicate key overwrites earlier row index", function () {
    var dupTable = [
      ["dup", "first"],
      ["dup", "second"],
    ];
    var index = buildLookupIndex(dupTable, 0);
    expect(index["dup"]).toBe(1);
  });

  it("returns empty object for empty table", function () {
    var index = buildLookupIndex([], 0);
    expect(index).toEqual({});
  });
});
