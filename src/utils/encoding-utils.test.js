var iconv = require("iconv-lite");
var { fixGarbledText } = require("./encoding-utils");

describe("fixGarbledText", function () {
  it("passes through plain ASCII unchanged", function () {
    expect(fixGarbledText("hello world")).toBe("hello world");
  });

  it("passes through numbers unchanged", function () {
    expect(fixGarbledText(123)).toBe(123);
    expect(fixGarbledText(0)).toBe(0);
  });

  it("passes through empty string unchanged", function () {
    expect(fixGarbledText("")).toBe("");
  });

  it("fixes garbled Chinese text (UTF-8 bytes misinterpreted as GBK)", function () {
    var original = "中文测试";
    var utf8Bytes = iconv.encode(original, "utf-8");
    var garbled = iconv.decode(utf8Bytes, "gbk");
    expect(garbled).not.toBe(original);
    expect(fixGarbledText(garbled)).toBe(original);
  });

  it("fixes garbled text with mixed Chinese and ASCII", function () {
    var original = "编号ABC_测试数据123";
    var utf8Bytes = iconv.encode(original, "utf-8");
    var garbled = iconv.decode(utf8Bytes, "gbk");
    expect(garbled).not.toBe(original);
    expect(fixGarbledText(garbled)).toBe(original);
  });

  it("does not throw on already-correct Chinese", function () {
    var result = fixGarbledText("正确中文");
    expect(typeof result).toBe("string");
    expect(result.length).toBeGreaterThan(0);
  });
});
