// src/utils/expand-utils.test.js
var { expandData } = require("./expand-utils");

describe("expandData", function () {
  it("expands two data rows correctly", function () {
    var values = [
      ["省份", "地市", "地市", "地市"],
      ["河南", "郑州", "南阳", "许昌"],
      ["河北", "石家庄", "保定", ""],
    ];
    var result = expandData(values);
    expect(result).toEqual([
      ["河南", "郑州"],
      ["河南", "南阳"],
      ["河南", "许昌"],
      ["河北", "石家庄"],
      ["河北", "保定"],
    ]);
  });

  it("handles single data row", function () {
    var values = [
      ["省份", "地市", "地市"],
      ["河南", "郑州", "南阳"],
    ];
    var result = expandData(values);
    expect(result).toEqual([
      ["河南", "郑州"],
      ["河南", "南阳"],
    ]);
  });

  it("skips empty cells", function () {
    var values = [
      ["省份", "地市", "地市"],
      ["河南", "郑州", ""],
    ];
    var result = expandData(values);
    expect(result).toEqual([
      ["河南", "郑州"],
    ]);
  });

  it("returns empty array for header-only data", function () {
    var values = [["省份", "地市"]];
    var result = expandData(values);
    expect(result).toEqual([]);
  });
});