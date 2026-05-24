var { parseCSV } = require("./csv-utils");

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
});
