var iconv = require("iconv-lite");

function fixGarbledText(value) {
  if (typeof value !== "string" || value === "") {
    return value;
  }
  try {
    var bytes = iconv.encode(value, "gbk");
    var fixed = iconv.decode(bytes, "utf-8");
    return fixed;
  } catch (e) {
    return value;
  }
}

module.exports = { fixGarbledText };
