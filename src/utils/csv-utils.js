/**
 * Parse delimiter-separated text into a 2D array.
 * Follows RFC 4180: handles quoted fields, embedded delimiters,
 * embedded newlines, and doubled quotes.
 * Strips UTF-8 BOM if present.
 * @param {string} text - Raw file content
 * @param {string} delimiter - Field delimiter (default ",")
 */
function parseCSV(text, delimiter) {
  if (!text || text.length === 0) {
    return [];
  }

  var sep = delimiter || ",";

  // Strip UTF-8 BOM
  if (text.charCodeAt(0) === 0xfeff) {
    text = text.slice(1);
  }

  var rows = [];
  var row = [];
  var field = "";
  var inQuotes = false;

  for (var i = 0; i < text.length; i++) {
    var ch = text[i];

    if (inQuotes) {
      if (ch === '"') {
        var nextCh = text[i + 1];
        if (nextCh === '"') {
          field += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        field += ch;
      }
    } else {
      if (ch === '"' && field === "") {
        inQuotes = true;
      } else if (ch === sep) {
        row.push(field);
        field = "";
      } else if (ch === "\n") {
        row.push(field);
        field = "";
        if (row.length > 0 || rows.length > 0) {
          rows.push(row);
        }
        row = [];
      } else if (ch === "\r") {
        if (text[i + 1] === "\n") {
          i++;
        }
        row.push(field);
        field = "";
        if (row.length > 0 || rows.length > 0) {
          rows.push(row);
        }
        row = [];
      } else {
        field += ch;
      }
    }
  }

  // Last field
  row.push(field);
  if (row.length > 1 || row[0] !== "" || rows.length === 0) {
    rows.push(row);
  }

  return rows;
}

module.exports = { parseCSV };
