var {
  generateMarkdownTable,
  parseRange,
  detectAlignment,
  escapeCell,
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
    var rows = [["A", "B"], ["C", "D"]];
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
});