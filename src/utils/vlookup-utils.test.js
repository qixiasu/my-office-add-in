var {
  parseRangeAddress,
  parseCellRef,
  buildColRange,
  buildIndexMatchFormula,
  staticLookup
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
    ["王五", "财务部", "A003", 7000]
  ];

  it("exact match returns correct values", function () {
    var result = staticLookup(["张三", "李四"], table, 0, [1, 3], 0);
    expect(result).toEqual([
      ["研发部", 8000],
      ["市场部", 6000]
    ]);
  });

  it("returns #N/A for not found values", function () {
    var result = staticLookup(["张三", "不存在"], table, 0, [1], 0);
    expect(result).toEqual([
      ["研发部"],
      ["#N/A"]
    ]);
  });

  it("returns multiple return columns", function () {
    var result = staticLookup(["王五"], table, 0, [1, 2, 3], 0);
    expect(result).toEqual([
      ["财务部", "A003", 7000]
    ]);
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
    var numTable = [[100, "low"], [200, "mid"]];
    var result = staticLookup([null], numTable, 0, [1], 1);
    expect(result).toEqual([["#N/A"]]);
  });

  it("approximate match finds largest value <= lookup", function () {
    var numTable = [
      [100, "low"],
      [200, "mid"],
      [300, "high"]
    ];
    var result = staticLookup([150, 250, 350], numTable, 0, [1], 1);
    expect(result).toEqual([["low"], ["mid"], ["high"]]);
  });
});
