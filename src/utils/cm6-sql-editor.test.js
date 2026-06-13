// src/utils/cm6-sql-editor.test.js

var { createSqlEditor } = require("./cm6-sql-editor");

describe("createSqlEditor", function () {
  it("exports createSqlEditor function", function () {
    expect(typeof createSqlEditor).toBe("function");
  });

  it("throws when container is null", function () {
    expect(function () {
      createSqlEditor(null, {});
    }).toThrow("container must be a valid HTMLElement");
  });

  it("throws when container is undefined", function () {
    expect(function () {
      createSqlEditor(undefined, {});
    }).toThrow("container must be a valid HTMLElement");
  });

  it("throws when container is not an HTMLElement", function () {
    expect(function () {
      createSqlEditor({}, {});
    }).toThrow("container must be a valid HTMLElement");
  });
});
