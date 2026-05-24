/**
 * Parse CSV text into a 2D array.
 * Follows RFC 4180: handles quoted fields, embedded commas,
 * embedded newlines, and doubled quotes.
 * Strips UTF-8 BOM if present.
 */
function parseCSV(text) {
  if (!text || text.length === 0) {
    return [];
  }

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
      } else if (ch === ",") {
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
        // Handle \r\n as a single newline
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
