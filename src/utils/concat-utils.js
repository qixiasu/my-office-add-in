function getColumnLetter(colIndex) {
  var letter = "";
  var remaining = colIndex;
  do {
    letter = String.fromCharCode(65 + (remaining % 26)) + letter;
    remaining = Math.floor(remaining / 26) - 1;
  } while (remaining >= 0);
  return letter;
}

function escapeFormulaText(text) {
  return text.replace(/"/g, '""');
}

function buildConcatFormula(firstColLetter, secondColLetter, connector) {
  var escaped = escapeFormulaText(connector);
  return (
    "=IF(" +
    firstColLetter +
    "1&" +
    secondColLetter +
    '1="","",' +
    firstColLetter +
    '1&"' +
    escaped +
    '"&' +
    secondColLetter +
    "1)"
  );
}

/**
 * Build concatenation formula for N columns
 * @param {string[]} columns - Array of column letters, e.g. ['A', 'B', 'C']
 * @param {string} connector - Connector string, e.g. '_'
 * @returns {string} Excel formula string
 */
function buildNConcatFormula(columns, connector) {
  var escaped = escapeFormulaText(connector);
  var rowRef = "1";

  // Build empty-check string: "A1&B1&C1"
  var emptyCheck = columns
    .map(function (col) {
      return col + rowRef;
    })
    .join("&");

  // Build concat string: 'A1&"_"&B1&"_"&C1'
  var concatParts = columns.map(function (col, index) {
    if (index === 0) {
      return col + rowRef;
    }
    return '"' + escaped + '"&' + col + rowRef;
  });
  var concatStr = concatParts.join("&");

  return "=IF(" + emptyCheck + '="","",' + concatStr + ")";
}

module.exports = { getColumnLetter, escapeFormulaText, buildConcatFormula, buildNConcatFormula };
