var {
  getColumnLetter,
  escapeFormulaText,
  buildConcatFormula,
  buildNConcatFormula,
} = require("./concat-utils");

describe("getColumnLetter", function () {
  it("returns A for column 0", function () {
    expect(getColumnLetter(0)).toBe("A");
  });

  it("returns Z for column 25", function () {
    expect(getColumnLetter(25)).toBe("Z");
  });

  it("returns AA for column 26", function () {
    expect(getColumnLetter(26)).toBe("AA");
  });

  it("returns AZ for column 51", function () {
    expect(getColumnLetter(51)).toBe("AZ");
  });

  it("returns BA for column 52", function () {
    expect(getColumnLetter(52)).toBe("BA");
  });

  it("returns ZZ for column 701", function () {
    expect(getColumnLetter(701)).toBe("ZZ");
  });

  it("returns AAA for column 702", function () {
    expect(getColumnLetter(702)).toBe("AAA");
  });

  it("returns XFD for column 16383 (Excel max column)", function () {
    expect(getColumnLetter(16383)).toBe("XFD");
  });
});

describe("escapeFormulaText", function () {
  it("passes through text without double quotes", function () {
    expect(escapeFormulaText("_")).toBe("_");
  });

  it("doubles a single double quote", function () {
    expect(escapeFormulaText('a"b')).toBe('a""b');
  });

  it("doubles multiple double quotes", function () {
    expect(escapeFormulaText('"start"')).toBe('""start""');
  });

  it("returns empty string unchanged", function () {
    expect(escapeFormulaText("")).toBe("");
  });
});

describe("buildConcatFormula", function () {
  it("builds formula with underscore connector", function () {
    var formula = buildConcatFormula("A", "B", "_");
    expect(formula).toBe('=IF(A1&B1="","",A1&"_"&B1)');
  });

  it("builds formula with custom connector", function () {
    var formula = buildConcatFormula("C", "D", "-");
    expect(formula).toBe('=IF(C1&D1="","",C1&"-"&D1)');
  });

  it("escapes double quotes in connector", function () {
    var formula = buildConcatFormula("A", "B", 'a"b');
    expect(formula).toBe('=IF(A1&B1="","",A1&"a""b"&B1)');
  });

  it("works with multi-letter column references", function () {
    var formula = buildConcatFormula("AA", "AB", "|");
    expect(formula).toBe('=IF(AA1&AB1="","",AA1&"|"&AB1)');
  });
});

describe("buildNConcatFormula", function () {
  it("builds formula for 3 columns with underscore connector", function () {
    var formula = buildNConcatFormula(["A", "B", "C"], "_");
    expect(formula).toBe('=IF(A1&B1&C1="","",A1&"_"&B1&"_"&C1)');
  });

  it("builds formula for 4 columns with dash connector", function () {
    var formula = buildNConcatFormula(["A", "B", "C", "D"], "-");
    expect(formula).toBe('=IF(A1&B1&C1&D1="","",A1&"-"&B1&"-"&C1&"-"&D1)');
  });

  it("escapes double quotes in connector", function () {
    var formula = buildNConcatFormula(["A", "B"], 'a"b');
    expect(formula).toBe('=IF(A1&B1="","",A1&"a""b"&B1)');
  });

  it("works with multi-letter column references", function () {
    var formula = buildNConcatFormula(["AA", "AB", "AC"], "|");
    expect(formula).toBe('=IF(AA1&AB1&AC1="","",AA1&"|"&AB1&"|"&AC1)');
  });

  it("returns empty string formula for single column", function () {
    var formula = buildNConcatFormula(["A"], "_");
    expect(formula).toBe('=IF(A1="","",A1)');
  });
});
