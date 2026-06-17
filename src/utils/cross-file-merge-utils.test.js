var crossFileMergeUtils = require("../utils/cross-file-merge-utils");

describe("crossFileMergeUtils", function () {
  describe("getSheetNames", function () {
    beforeEach(function () {
      // Mock FileReader
      global.FileReader = function () {};
      global.XLSX = {
        read: function (data, opts) {
          return {
            SheetNames: ["Sheet1", "Sheet2", "数据表"]
          };
        }
      };
    });

    afterEach(function () {
      delete global.FileReader;
      delete global.XLSX;
    });

    it("returns sheet names array from workbook", function () {
      var file = { name: "test.xlsx" };
      // Mock FileReader.onload
      global.FileReader = function () {
        this.onload = null;
        var self = this;
        this.readAsArrayBuffer = function () {
          setTimeout(function () {
            self.onload({ target: { result: new ArrayBuffer(8) } });
          }, 0);
        };
      };

      return crossFileMergeUtils.getSheetNames(file).then(function (names) {
        expect(names).toEqual(["Sheet1", "Sheet2", "数据表"]);
      });
    });

    it("rejects when file read fails", function () {
      global.FileReader = function () {
        this.onerror = null;
        var self = this;
        this.readAsArrayBuffer = function () {
          setTimeout(function () {
            self.onerror({});
          }, 0);
        };
      };

      var file = { name: "test.xlsx" };
      return crossFileMergeUtils.getSheetNames(file).catch(function (err) {
        expect(err.message).toBe("文件读取失败");
      });
    });
  });

  describe("getColumnLetter", function () {
    it("returns A for column 0", function () {
      expect(crossFileMergeUtils.getColumnLetter(0)).toBe("A");
    });

    it("returns Z for column 25", function () {
      expect(crossFileMergeUtils.getColumnLetter(25)).toBe("Z");
    });

    it("returns AA for column 26", function () {
      expect(crossFileMergeUtils.getColumnLetter(26)).toBe("AA");
    });

    it("returns AZ for column 51", function () {
      expect(crossFileMergeUtils.getColumnLetter(51)).toBe("AZ");
    });

    it("returns BA for column 52", function () {
      expect(crossFileMergeUtils.getColumnLetter(52)).toBe("BA");
    });

    it("returns ZZ for column 701", function () {
      expect(crossFileMergeUtils.getColumnLetter(701)).toBe("ZZ");
    });

    it("returns AAA for column 702", function () {
      expect(crossFileMergeUtils.getColumnLetter(702)).toBe("AAA");
    });

    it("returns XFD for column 16383 (Excel max column)", function () {
      expect(crossFileMergeUtils.getColumnLetter(16383)).toBe("XFD");
    });
  });

  describe("generateUniqueSheetName", function () {
    it("returns baseName when not in existing names", function () {
      expect(
        crossFileMergeUtils.generateUniqueSheetName("合并结果", ["Sheet1"])
      ).toBe("合并结果");
    });

    it("returns baseName_1 when baseName exists", function () {
      expect(
        crossFileMergeUtils.generateUniqueSheetName("合并结果", ["合并结果"])
      ).toBe("合并结果_1");
    });

    it("returns baseName_2 when baseName and _1 exist", function () {
      expect(
        crossFileMergeUtils.generateUniqueSheetName("合并结果", [
          "合并结果",
          "合并结果_1",
        ])
      ).toBe("合并结果_2");
    });

    it("returns baseName_10 when baseName and _1 through _9 exist", function () {
      var existing = ["合并结果"];
      for (var i = 1; i <= 9; i++) {
        existing.push("合并结果_" + i);
      }
      expect(
        crossFileMergeUtils.generateUniqueSheetName("合并结果", existing)
      ).toBe("合并结果_10");
    });

    it("returns baseName with timestamp when all _N variants exhausted", function () {
      var existing = ["合并结果"];
      for (var i = 1; i <= 100; i++) {
        existing.push("合并结果_" + i);
      }
      var result = crossFileMergeUtils.generateUniqueSheetName(
        "合并结果",
        existing
      );
      expect(result.indexOf("合并结果_")).toBe(0);
      expect(result !== "合并结果_101").toBe(true);
    });
  });

  describe("validateColumnConsistency", function () {
    it("returns invalid when less than 2 files", function () {
      var result = crossFileMergeUtils.validateColumnConsistency([]);
      expect(result.valid).toBe(false);
      expect(result.error).toBe("请至少选择两个文件");
    });

    it("returns invalid for single file", function () {
      var result = crossFileMergeUtils.validateColumnConsistency([
        { name: "file1.xlsx", columnCount: 3 },
      ]);
      expect(result.valid).toBe(false);
      expect(result.error).toBe("请至少选择两个文件");
    });

    it("returns valid when all files have same column count", function () {
      var result = crossFileMergeUtils.validateColumnConsistency([
        { name: "file1.xlsx", columnCount: 3 },
        { name: "file2.xlsx", columnCount: 3 },
        { name: "file3.xlsx", columnCount: 3 },
      ]);
      expect(result.valid).toBe(true);
      expect(result.error).toBe(null);
    });

    it("returns invalid when column counts differ", function () {
      var result = crossFileMergeUtils.validateColumnConsistency([
        { name: "file1.xlsx", columnCount: 3 },
        { name: "file2.xlsx", columnCount: 5 },
      ]);
      expect(result.valid).toBe(false);
      expect(result.error).toBe(
        "文件 'file2.xlsx' 列数(5)与第一个文件(3)不一致"
      );
    });
  });

  describe("mergeExcelData", function () {
    it("returns empty data for empty file list", function () {
      var result = crossFileMergeUtils.mergeExcelData([], 0);
      expect(result.mergedData).toEqual([]);
      expect(result.columnCount).toBe(0);
    });

    it("skips empty files", function () {
      var result = crossFileMergeUtils.mergeExcelData(
        [{ data: [], name: "empty.xlsx" }],
        0
      );
      expect(result.mergedData).toEqual([]);
    });

    it("merges two files without header", function () {
      var fileDataList = [
        { data: [["A1", "B1"], ["A2", "B2"]], name: "file1.xlsx", sheetName: "Sheet1" },
        { data: [["A3", "B3"], ["A4", "B4"]], name: "file2.xlsx", sheetName: "Sheet1" },
      ];
      var result = crossFileMergeUtils.mergeExcelData(fileDataList, 0);
      expect(result.mergedData).toEqual([
        ["Sheet1", "file1.xlsx", "A1", "B1"],
        ["Sheet1", "file1.xlsx", "A2", "B2"],
        ["Sheet1", "file2.xlsx", "A3", "B3"],
        ["Sheet1", "file2.xlsx", "A4", "B4"],
      ]);
      expect(result.columnCount).toBe(0); // no header row
    });

    it("merges two files with header row 1", function () {
      var fileDataList = [
        {
          data: [
            ["Name", "Age"],
            ["Alice", 25],
            ["Bob", 30],
          ],
          name: "file1.xlsx",
          sheetName: "Sheet1",
        },
        {
          data: [
            ["Name", "Age"],
            ["Charlie", 35],
          ],
          name: "file2.xlsx",
          sheetName: "Sheet1",
        },
      ];
      var result = crossFileMergeUtils.mergeExcelData(fileDataList, 1);
      expect(result.mergedData).toEqual([
        ["Sheet名", "来源文件", "Name", "Age"],
        ["Sheet1", "file1.xlsx", "Alice", 25],
        ["Sheet1", "file1.xlsx", "Bob", 30],
        ["Sheet1", "file2.xlsx", "Charlie", 35],
      ]);
      expect(result.columnCount).toBe(4);
    });

    it("adds source column to all rows including second file", function () {
      var fileDataList = [
        {
          data: [["Name", "Age"], ["Alice", 25]],
          name: "file1.xlsx",
          sheetName: "Sheet1",
        },
        {
          data: [["Name", "Age"], ["Bob", 30]],
          name: "file2.xlsx",
          sheetName: "Sheet2",
        },
      ];
      var result = crossFileMergeUtils.mergeExcelData(fileDataList, 1);
      // Second file's header is skipped, source column added to all data rows
      expect(result.mergedData).toEqual([
        ["Sheet名", "来源文件", "Name", "Age"],
        ["Sheet1", "file1.xlsx", "Alice", 25],
        ["Sheet2", "file2.xlsx", "Bob", 30],
      ]);
    });
  });

  describe("parseExcelFile", function () {
    afterEach(function () {
      delete global.FileReader;
      delete global.XLSX;
    });

    it("parses specified sheet when sheetName parameter is provided", function () {
      var file = { name: "test.xlsx" };
      var mockSheet = { A1: { v: "test" } };
      global.FileReader = function () {
        this.onload = null;
        var self = this;
        this.readAsArrayBuffer = function () {
          setTimeout(function () {
            self.onload({ target: { result: new ArrayBuffer(8) } });
          }, 0);
        };
      };
      global.XLSX = {
        read: function (data, opts) {
          return {
            SheetNames: ["Sheet1", "Sheet2"],
            Sheets: {
              "Sheet2": mockSheet
            }
          };
        }
      };
      global.XLSX.utils = {
        sheet_to_json: function (sheet, opts) {
          return [["A1"], ["B1"]];
        }
      };

      return crossFileMergeUtils.parseExcelFile(file, "Sheet2").then(function (result) {
        expect(result.sheetName).toBe("Sheet2");
        expect(result.data).toEqual([["A1"], ["B1"]]);
      });
    });
  });
});
