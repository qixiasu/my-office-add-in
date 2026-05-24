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
    firstColLetter + "1&" + secondColLetter + '1="","",' +
    firstColLetter + '1&"' + escaped + '"&' + secondColLetter +
    "1)"
  );
}

module.exports = { getColumnLetter, escapeFormulaText, buildConcatFormula };
