// src/utils/data-cleaning-utils.test.js

var utils = require("./data-cleaning-utils");

// ===== trimSpaces =====

test("trimSpaces removes leading and trailing spaces", function () {
  var input = [["  hello  "], ["  world  "]];
  var result = utils.trimSpaces(input, "both");
  expect(result[0][0]).toBe("hello");
  expect(result[1][0]).toBe("world");
});

test("trimSpaces removes all extra spaces", function () {
  var input = [["hello   world"], ["a    b"]];
  var result = utils.trimSpaces(input, "all");
  expect(result[0][0]).toBe("hello world");
  expect(result[1][0]).toBe("a b");
});

test("trimSpaces removes all spaces", function () {
  var input = [["hello world"], [" a b "]];
  var result = utils.trimSpaces(input, "leading");
  expect(result[0][0]).toBe("hello world");
  expect(result[1][0]).toBe("a b ");
});

test("trimSpaces trailing mode", function () {
  var input = [["hello  "], ["world "]];
  var result = utils.trimSpaces(input, "trailing");
  expect(result[0][0]).toBe("hello");
  expect(result[1][0]).toBe("world");
});

test("trimSpaces handles null and empty cells", function () {
  var input = [[null], [""], ["  test  "]];
  var result = utils.trimSpaces(input, "both");
  expect(result[0][0]).toBeNull();
  expect(result[1][0]).toBe("");
  expect(result[2][0]).toBe("test");
});

test("trimSpaces handles single cell", function () {
  var input = [["  only one  "]];
  var result = utils.trimSpaces(input, "both");
  expect(result[0][0]).toBe("only one");
});

test("trimSpaces does not affect non-string values", function () {
  var input = [[123], [true], [null]];
  var result = utils.trimSpaces(input, "both");
  expect(result[0][0]).toBe(123);
  expect(result[1][0]).toBe(true);
  expect(result[2][0]).toBeNull();
});

// ===== removeEmptyRows =====

test("removeEmptyRows removes completely empty rows", function () {
  var input = [
    ["a", "b"],
    [null, null],
    ["c", "d"],
    ["", ""],
  ];
  var result = utils.removeEmptyRows(input, "all", null, null);
  expect(result.length).toBe(2);
  expect(result[0][0]).toBe("a");
  expect(result[1][0]).toBe("c");
});

test("removeEmptyRows no empty rows returns same data", function () {
  var input = [["a"], ["b"], ["c"]];
  var result = utils.removeEmptyRows(input, "all", null, null);
  expect(result.length).toBe(3);
});

test("removeEmptyRows removes rows where specific column is empty", function () {
  var input = [
    ["a", "x"],
    ["b", null],
    ["c", "y"],
  ];
  var result = utils.removeEmptyRows(input, "column", 1, null);
  expect(result.length).toBe(2);
  expect(result[0][0]).toBe("a");
  expect(result[1][0]).toBe("c");
});

test("removeEmptyRows removes rows with empty ratio above threshold", function () {
  var input = [
    ["a", "b", "c"],
    [null, "e", null],
    ["g", null, null],
    ["j", "k", null],
  ];
  var result = utils.removeEmptyRows(input, "ratio", null, 50);
  expect(result.length).toBe(2);
  expect(result[0][0]).toBe("a");
  expect(result[1][0]).toBe("j");
});

test("removeEmptyRows handles all rows empty", function () {
  var input = [[null, null], [null, null]];
  var result = utils.removeEmptyRows(input, "all", null, null);
  expect(result.length).toBe(0);
});

test("removeEmptyRows does not mutate original", function () {
  var input = [["a"], [null], ["b"]];
  var copy = [["a"], [null], ["b"]];
  utils.removeEmptyRows(input, "all", null, null);
  expect(input).toEqual(copy);
});

// ===== convertCase =====

test("convertCase converts to upper case", function () {
  var input = [["hello"], ["World"]];
  var result = utils.convertCase(input, "upper");
  expect(result[0][0]).toBe("HELLO");
  expect(result[1][0]).toBe("WORLD");
});

test("convertCase converts to lower case", function () {
  var input = [["HELLO"], ["World"]];
  var result = utils.convertCase(input, "lower");
  expect(result[0][0]).toBe("hello");
  expect(result[1][0]).toBe("world");
});

test("convertCase capitalizes first letter", function () {
  var input = [["hello world"], ["javaScript"]];
  var result = utils.convertCase(input, "capitalize");
  expect(result[0][0]).toBe("Hello world");
  expect(result[1][0]).toBe("JavaScript");
});

test("convertCase does not affect Chinese characters", function () {
  var input = [["hello 你好"], ["WORLD 世界"]];
  var result = utils.convertCase(input, "upper");
  expect(result[0][0]).toBe("HELLO 你好");
  expect(result[1][0]).toBe("WORLD 世界");
});

test("convertCase handles non-string values", function () {
  var input = [[123], [true], [null]];
  var result = utils.convertCase(input, "upper");
  expect(result[0][0]).toBe(123);
  expect(result[1][0]).toBe(true);
  expect(result[2][0]).toBeNull();
});

// ===== removeInvisible =====

test("removeInvisible removes control characters", function () {
  var input = [["hello\nworld"], ["tab\there"]];
  var result = utils.removeInvisible(input, "control");
  expect(result[0][0]).toBe("helloworld");
  expect(result[1][0]).toBe("tabhere");
});

test("removeInvisible removes zero-width characters", function () {
  var input = [["hello​world"], ["a‌b"]];
  var result = utils.removeInvisible(input, "zero-width");
  expect(result[0][0]).toBe("helloworld");
  expect(result[1][0]).toBe("ab");
});

test("removeInvisible removes all invisible characters", function () {
  var input = [["hello\nworld​!"]];
  var result = utils.removeInvisible(input, "all");
  expect(result[0][0]).toBe("helloworld!");
});

test("removeInvisible preserves normal text", function () {
  var input = [["hello world"], ["你好世界"]];
  var result = utils.removeInvisible(input, "all");
  expect(result[0][0]).toBe("hello world");
  expect(result[1][0]).toBe("你好世界");
});

test("removeInvisible handles null cells", function () {
  var input = [[null], ["hello"]];
  var result = utils.removeInvisible(input, "all");
  expect(result[0][0]).toBeNull();
  expect(result[1][0]).toBe("hello");
});

// ===== removeDuplicates =====

test("removeDuplicates removes duplicate rows based on all columns", function () {
  var input = [
    ["a", 1],
    ["b", 2],
    ["a", 1],
    ["c", 3],
  ];
  var result = utils.removeDuplicates(input, null, "first");
  expect(result.length).toBe(3);
});

test("removeDuplicates keeps first occurrence", function () {
  var input = [["a"], ["b"], ["a"], ["c"]];
  var result = utils.removeDuplicates(input, null, "first");
  expect(result[0][0]).toBe("a");
  expect(result[1][0]).toBe("b");
  expect(result[2][0]).toBe("c");
});

test("removeDuplicates keeps last occurrence", function () {
  var input = [["a", 1], ["b", 2], ["a", 99]];
  var result = utils.removeDuplicates(input, [0], "last");
  expect(result.length).toBe(2);
  expect(result[0][0]).toBe("b");
  expect(result[1][0]).toBe("a");
  expect(result[1][1]).toBe(99);
});

test("removeDuplicates with specific key columns", function () {
  var input = [
    ["a", "x", 1],
    ["a", "y", 2],
    ["b", "x", 3],
  ];
  var result = utils.removeDuplicates(input, [0], "first");
  expect(result.length).toBe(2);
});

test("removeDuplicates no duplicates returns same data", function () {
  var input = [["a"], ["b"], ["c"]];
  var result = utils.removeDuplicates(input, null, "first");
  expect(result.length).toBe(3);
});

test("removeDuplicates handles null values in key columns", function () {
  var input = [[null, "a"], [null, "b"]];
  var result = utils.removeDuplicates(input, [0], "first");
  expect(result.length).toBe(1);
});

test("removeDuplicates does not mutate original", function () {
  var input = [["a"], ["b"], ["a"]];
  var copy = [["a"], ["b"], ["a"]];
  utils.removeDuplicates(input, null, "first");
  expect(input).toEqual(copy);
});
