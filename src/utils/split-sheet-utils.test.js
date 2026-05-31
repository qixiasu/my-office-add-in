var {
  groupDataByKey,
  truncateSheetName,
} = require("./split-sheet-utils");

describe("groupDataByKey", function () {
  it("groups 3 rows for '河南' and 2 rows for '河北'", function () {
    var data = [
      ["省份", "城市", "人口"],
      ["河南", "郑州", "1000"],
      ["河南", "洛阳", "700"],
      ["河南", "开封", "500"],
      ["河北", "石家庄", "1100"],
      ["河北", "保定", "700"],
    ];

    var result = groupDataByKey(data, 0);

    expect(result.error).toBeNull();
    expect(result.groups).not.toBeNull();

    var groups = result.groups;
    expect(groups["河南"]).not.toBeUndefined();
    expect(groups["河南"].header).toEqual(["省份", "城市", "人口"]);
    expect(groups["河南"].rows.length).toBe(3);
    expect(groups["河南"].rows[0]).toEqual(["河南", "郑州", "1000"]);
    expect(groups["河南"].rows[1]).toEqual(["河南", "洛阳", "700"]);
    expect(groups["河南"].rows[2]).toEqual(["河南", "开封", "500"]);

    expect(groups["河北"]).not.toBeUndefined();
    expect(groups["河北"].rows.length).toBe(2);
    expect(groups["河北"].rows[0]).toEqual(["河北", "石家庄", "1100"]);
    expect(groups["河北"].rows[1]).toEqual(["河北", "保定", "700"]);
  });

  it("returns error for empty key value", function () {
    var data = [
      ["省份", "城市"],
      ["", "郑州"],
      ["河南", "洛阳"],
    ];

    var result = groupDataByKey(data, 0);

    expect(result.groups).toBeNull();
    expect(result.error).toBe("键值不能为空");
  });

  it("returns error for whitespace-only key value", function () {
    var data = [
      ["省份", "城市"],
      ["   ", "郑州"],
      ["河南", "洛阳"],
    ];

    var result = groupDataByKey(data, 0);

    expect(result.groups).toBeNull();
    expect(result.error).toBe("键值不能为空");
  });

  it("handles single row per key", function () {
    var data = [
      ["省份", "城市"],
      ["河南", "郑州"],
      ["河北", "石家庄"],
      ["山东", "济南"],
    ];

    var result = groupDataByKey(data, 0);

    expect(result.error).toBeNull();
    expect(result.groups["河南"].rows.length).toBe(1);
    expect(result.groups["河北"].rows.length).toBe(1);
    expect(result.groups["山东"].rows.length).toBe(1);
  });

  it("returns error when key column index is out of bounds", function () {
    var data = [
      ["省份", "城市"],
      ["河南", "郑州"],
    ];

    var result = groupDataByKey(data, 5);

    expect(result.groups).toBeNull();
    expect(result.error).toBe("键列索引超出范围");
  });

  it("returns error when key column index is negative", function () {
    var data = [
      ["省份", "城市"],
      ["河南", "郑州"],
    ];

    var result = groupDataByKey(data, -1);

    expect(result.groups).toBeNull();
    expect(result.error).toBe("键列索引超出范围");
  });

  it("returns error when data is empty", function () {
    var result = groupDataByKey([], 0);

    expect(result.groups).toBeNull();
    expect(result.error).toBe("数据不能为空且至少需要包含表头和一行数据");
  });

  it("returns error when data has only header", function () {
    var data = [["省份", "城市"]];

    var result = groupDataByKey(data, 0);

    expect(result.groups).toBeNull();
    expect(result.error).toBe("数据不能为空且至少需要包含表头和一行数据");
  });

  it("groups by second column (keyColumnIndex=1)", function () {
    var data = [
      ["省份", "城市", "人口"],
      ["河南", "郑州", "1000"],
      ["河南", "洛阳", "700"],
      ["河北", "郑州", "800"],
    ];

    var result = groupDataByKey(data, 1);

    expect(result.error).toBeNull();
    expect(result.groups["郑州"]).not.toBeUndefined();
    expect(result.groups["郑州"].rows.length).toBe(2);
    expect(result.groups["洛阳"]).not.toBeUndefined();
    expect(result.groups["洛阳"].rows.length).toBe(1);
  });

  it("handles null key value as empty string", function () {
    var data = [
      ["省份", "城市"],
      [null, "郑州"],
      ["河南", "洛阳"],
    ];

    var result = groupDataByKey(data, 0);

    expect(result.groups).toBeNull();
    expect(result.error).toBe("键值不能为空");
  });

  it("trims key values before grouping", function () {
    var data = [
      ["省份", "城市"],
      ["  河南  ", "郑州"],
      ["河南", "洛阳"],
    ];

    var result = groupDataByKey(data, 0);

    expect(result.error).toBeNull();
    // Both rows should be grouped under "河南" after trimming
    expect(result.groups["河南"]).not.toBeUndefined();
    expect(result.groups["河南"].rows.length).toBe(2);
  });
});

describe("truncateSheetName", function () {
  it("returns original name if under 31 characters", function () {
    var name = "Sheet1";
    expect(truncateSheetName(name)).toBe("Sheet1");
  });

  it("truncates names longer than 31 characters", function () {
    var longName = "This Is A Very Long Sheet Name That Exceeds 31 Characters";
    var result = truncateSheetName(longName);

    expect(result.length).toBe(31);
    expect(result.endsWith("...")).toBe(true);
  });

  it("handles name exactly 31 characters (no truncation)", function () {
    var name = "1234567890123456789012345678901"; // 31 chars
    expect(truncateSheetName(name)).toBe(name);
  });

  it("handles name exactly 31 characters after truncation", function () {
    // Original name is 34 chars (>31), truncating gives 28 chars + "..." = 31 chars
    var name = "1234567890123456789012345678901234"; // 34 chars
    var result = truncateSheetName(name);
    // 28 chars + "..." = 31 chars total
    var expected = "1234567890123456789012345678...";

    expect(result.length).toBe(31);
    expect(result).toBe(expected);
  });

  it("handles name longer than 31 characters ending with spaces", function () {
    var name = "1234567890123456789012345678901   "; // 31 + 3 spaces
    var result = truncateSheetName(name);

    expect(result.length).toBe(31);
    expect(result.endsWith("...")).toBe(true);
  });

  it("returns empty string when name is empty", function () {
    expect(truncateSheetName("")).toBe("");
  });

  it("returns empty string when name is null/undefined", function () {
    expect(truncateSheetName(null)).toBe("");
    expect(truncateSheetName(undefined)).toBe("");
  });

  it("handles single character name", function () {
    expect(truncateSheetName("A")).toBe("A");
  });

  it("handles name exactly 32 characters", function () {
    var name = "12345678901234567890123456789012"; // 32 chars
    var result = truncateSheetName(name);

    // 29 chars + "..." = 32 chars truncated to 31
    expect(result.length).toBe(31);
    expect(result.endsWith("...")).toBe(true);
  });
});