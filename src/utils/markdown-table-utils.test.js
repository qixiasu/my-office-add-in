var {
  generateMarkdownTable,
  parseRange,
  detectAlignment,
  escapeCell,
  handleMergedCells,
  parseMergedRangeAddress,
} = require("./markdown-table-utils");

// ── parseRange ─────────────────────────────────────────────────────────

describe("parseRange", function () {
  it("should filter empty rows", function () {
    var input = [
      ["A", "B"],
      [null, null],
      ["C", "D"],
    ];
    var result = parseRange(input);
    expect(result).toEqual([
      ["A", "B"],
      ["C", "D"],
    ]);
  });

  it("should filter trailing empty columns", function () {
    var input = [
      ["A", "B", null],
      ["C", "D", null],
    ];
    var result = parseRange(input);
    expect(result).toEqual([
      ["A", "B"],
      ["C", "D"],
    ]);
  });

  it("should return empty array for all empty input", function () {
    var input = [
      [null, null],
      [null, null],
    ];
    var result = parseRange(input);
    expect(result).toEqual([]);
  });
});

// ── detectAlignment ────────────────────────────────────────────────────

describe("detectAlignment", function () {
  it("should detect numeric columns as right-aligned", function () {
    var rows = [
      ["Name", "Score"],
      ["Alice", 100],
      ["Bob", 95],
    ];
    var result = detectAlignment(rows, true);
    expect(result).toEqual([":---", "---:"]);
  });

  it("should detect text columns as left-aligned", function () {
    var rows = [
      ["Name", "City"],
      ["Alice", "NYC"],
      ["Bob", "LA"],
    ];
    var result = detectAlignment(rows, true);
    expect(result).toEqual([":---", ":---"]);
  });

  it("should return basic separator when includeAlignment is false", function () {
    var rows = [
      ["A", "B"],
      ["C", "D"],
    ];
    var result = detectAlignment(rows, false);
    expect(result).toEqual(["---", "---"]);
  });
});

// ── escapeCell ─────────────────────────────────────────────────────────

describe("escapeCell", function () {
  it("should escape pipe characters", function () {
    expect(escapeCell("A|B")).toBe("A\\|B");
  });

  it("should replace newlines with br tag", function () {
    expect(escapeCell("A\nB")).toBe("A<br>B");
  });

  it("should escape backslash", function () {
    expect(escapeCell("A\\B")).toBe("A\\\\B");
  });

  it("should handle null and undefined", function () {
    expect(escapeCell(null)).toBe("");
    expect(escapeCell(undefined)).toBe("");
  });
});

// ── generateMarkdownTable ──────────────────────────────────────────────

describe("generateMarkdownTable", function () {
  it("should generate basic markdown table", function () {
    var values = [
      ["Name", "Age"],
      ["Alice", 25],
      ["Bob", 30],
    ];
    var result = generateMarkdownTable(values);
    expect(result).toContain("| Name | Age |");
    expect(result).toContain("| :--- | ---: |");
  });

  it("should return empty string for empty input", function () {
    expect(generateMarkdownTable([])).toBe("");
    expect(generateMarkdownTable(null)).toBe("");
  });

  it("should handle merged cells with colspan and rowspan", function () {
    var values = [
      ["A", "B", "C"],
      ["D", "E", "F"],
      ["G", "H", "I"],
    ];
    // Mock merged range: B1:C2 (colspan=2, rowspan=2)
    var mergedRanges = [
      {
        toString: function () {
          return "B1:C2";
        },
      },
    ];
    var result = generateMarkdownTable(values, { mergedRanges: mergedRanges });
    // B1 should have colspan=2 and rowspan=2
    expect(result).toContain("<colspan=2>");
    expect(result).toContain("<rowspan=2>");
  });
});

// ── handleMergedCells ──────────────────────────────────────────────────

describe("handleMergedCells", function () {
  it("should return empty Map when no merged ranges", function () {
    var result = handleMergedCells([], 3, 3);
    expect(result.size).toBe(0);
  });

  it("should return empty Map when mergedRanges is null", function () {
    var result = handleMergedCells(null, 3, 3);
    expect(result.size).toBe(0);
  });

  it("should handle merged range with colspan", function () {
    // A1:B1 (colspan=2, rowspan=1)
    var mergedRanges = [
      {
        toString: function () {
          return "A1:B1";
        },
      },
    ];
    var result = handleMergedCells(mergedRanges, 3, 3);
    // Primary cell at 0-0 should have colspan=2
    expect(result.get("0-0")).toEqual({ colspan: 2, rowspan: 1 });
    // Covered cell at 0-1 should be marked as covered
    expect(result.get("0-1")).toEqual({ covered: true });
  });

  it("should handle merged range with rowspan", function () {
    // A1:A2 (colspan=1, rowspan=2)
    var mergedRanges = [
      {
        toString: function () {
          return "A1:A2";
        },
      },
    ];
    var result = handleMergedCells(mergedRanges, 3, 3);
    // Primary cell at 0-0 should have rowspan=2
    expect(result.get("0-0")).toEqual({ colspan: 1, rowspan: 2 });
    // Covered cell at 1-0 should be marked as covered
    expect(result.get("1-0")).toEqual({ covered: true });
  });

  it("should handle merged range with both colspan and rowspan", function () {
    // B2:C3 (colspan=2, rowspan=2)
    var mergedRanges = [
      {
        toString: function () {
          return "B2:C3";
        },
      },
    ];
    var result = handleMergedCells(mergedRanges, 4, 4);
    // Primary cell at 1-1 should have colspan=2, rowspan=2
    expect(result.get("1-1")).toEqual({ colspan: 2, rowspan: 2 });
    // Covered cells
    expect(result.get("1-2")).toEqual({ covered: true });
    expect(result.get("2-1")).toEqual({ covered: true });
    expect(result.get("2-2")).toEqual({ covered: true });
  });

  it("should ignore single cell ranges", function () {
    // A1:A1 (no actual merge)
    var mergedRanges = [
      {
        toString: function () {
          return "A1:A1";
        },
      },
    ];
    var result = handleMergedCells(mergedRanges, 3, 3);
    expect(result.size).toBe(0);
  });
});

// ── parseMergedRangeAddress ────────────────────────────────────────────

describe("parseMergedRangeAddress", function () {
  it("should parse simple range", function () {
    var result = parseMergedRangeAddress("A1:C3");
    expect(result).toEqual({
      startCol: 0,
      startRow: 0,
      endCol: 2,
      endRow: 2,
    });
  });

  it("should parse range with multi-letter columns", function () {
    var result = parseMergedRangeAddress("AA1:ZZ999");
    expect(result).toEqual({
      startCol: 26, // AA = 27th column, 0-indexed = 26
      startRow: 0,
      endCol: 701, // ZZ = 702nd column, 0-indexed = 701
      endRow: 998,
    });
  });

  it("should return null for invalid range format", function () {
    expect(parseMergedRangeAddress("invalid")).toBe(null);
    expect(parseMergedRangeAddress("A1")).toBe(null);
    expect(parseMergedRangeAddress("A1:B")).toBe(null);
  });
});
