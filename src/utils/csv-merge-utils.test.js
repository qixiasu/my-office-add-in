/**
 * Unit tests for csv-merge-utils.js
 * Run with: node src/utils/csv-merge-utils.test.js
 */

var assert = require('assert');

// ── helpers ──────────────────────────────────────────────────────────────

function deepEqual(a, b) {
  if (a === b) return true;
  if (typeof a !== typeof b) return false;
  if (typeof a !== 'object' || a === null || b === null) return false;
  var ak = Object.keys(a);
  var bk = Object.keys(b);
  if (ak.length !== bk.length) return false;
  for (var i = 0; i < ak.length; i++) {
    var k = ak[i];
    if (!b.hasOwnProperty(k) || !deepEqual(a[k], b[k])) return false;
  }
  return true;
}

function assertEqual(actual, expected, msg) {
  if (!deepEqual(actual, expected)) {
    throw new Error(
      (msg ? msg + '\n' : '') +
      'Expected: ' + JSON.stringify(expected) + '\n' +
      'Actual:   ' + JSON.stringify(actual)
    );
  }
}

// ── tests ───────────────────────────────────────────────────────────────

var csvMerge = require('./csv-merge-utils');

// ── computeUnifiedHeaders ───────────────────────────────────────────────

console.log('Running computeUnifiedHeaders tests...');

assertEqual(
  csvMerge.computeUnifiedHeaders([['A', 'B'], ['B', 'C']]),
  ['A', 'B', 'C']
);
assertEqual(
  csvMerge.computeUnifiedHeaders([['X'], ['Y'], ['Z']]),
  ['X', 'Y', 'Z']
);
assertEqual(
  csvMerge.computeUnifiedHeaders([['name', 'age'], ['name', 'city']]),
  ['name', 'age', 'city']
);
assertEqual(
  csvMerge.computeUnifiedHeaders([[], ['B']]),
  ['B']
);
assertEqual(
  csvMerge.computeUnifiedHeaders([['A', 'B', 'A'], ['C', 'B']]),
  ['A', 'B', 'C']
);

console.log('  PASS: computeUnifiedHeaders');

// ── alignColumns ────────────────────────────────────────────────────────

console.log('Running alignColumns tests...');

function test_alignColumns_differentOrder() {
  var filesData = [
    { headers: ['A', 'B'], data: [['vA1', 'vB1'], ['vA2', 'vB2']] },
    { headers: ['B', 'A'], data: [['vB3', 'vA3']] },
  ];
  var result = csvMerge.alignColumns(filesData);
  assertEqual(result.headers, ['A', 'B']);
  assertEqual(result.data, [
    ['vA1', 'vB1'],
    ['vA2', 'vB2'],
    ['vA3', 'vB3'],
  ]);
  console.log('  PASS: test_alignColumns_differentOrder');
}

function test_alignColumns_threeFilesUnion() {
  var filesData = [
    { headers: ['name', 'age'],   data: [['Alice', '30']] },
    { headers: ['city', 'age'],   data: [['NYC', '25']] },
    { headers: ['name', 'email'], data: [['Bob', 'bob@example.com']] },
  ];
  var result = csvMerge.alignColumns(filesData);
  assertEqual(result.headers, ['name', 'age', 'city', 'email']);
  assertEqual(result.data, [
    ['Alice', '30',   '',              ''],
    ['',       '25',  'NYC',           ''],
    ['Bob',    '',    '', 'bob@example.com'],
  ]);
  console.log('  PASS: test_alignColumns_threeFilesUnion');
}

test_alignColumns_differentOrder();
test_alignColumns_threeFilesUnion();

// ── serializeRow ────────────────────────────────────────────────────────

console.log('Running serializeRow tests...');

function test_serializeRow_withQuotes() {
  // Fields with commas and double-quotes need quoting; quotes inside are doubled
  var row = ['Hello, World', 'Say "Hi"', 'Normal'];
  var result = csvMerge.serializeRow(row, ',');
  assertEqual(result, '"Hello, World","Say ""Hi""",Normal');
  console.log('  PASS: test_serializeRow_withQuotes');
}

// Additional edge cases
function test_serializeRow_noQuotes() {
  assertEqual(csvMerge.serializeRow(['a', 'b', 'c']), 'a,b,c');
}

function test_serializeRow_withNewline() {
  assertEqual(csvMerge.serializeRow(['line1\nline2', 'x']), '"line1\nline2",x');
}

function test_serializeRow_emptyFields() {
  assertEqual(csvMerge.serializeRow(['', '', '']), ',,');
}

function test_serializeRow_withCarriageReturn() {
  assertEqual(csvMerge.serializeRow(['a\rb', 'c']), '"a\rb",c');
}

test_serializeRow_withQuotes();
test_serializeRow_noQuotes();
test_serializeRow_withNewline();
test_serializeRow_emptyFields();
test_serializeRow_withCarriageReturn();

// ── computeUnifiedHeaders edge cases ────────────────────────────────────

console.log('Running computeUnifiedHeaders edge cases...');
assertEqual(csvMerge.computeUnifiedHeaders([]), []);
assertEqual(csvMerge.computeUnifiedHeaders([[]]), []);
assertEqual(csvMerge.computeUnifiedHeaders([['A'], []]), ['A']);
console.log('  PASS: edge cases');

// ── summary ─────────────────────────────────────────────────────────────

console.log('\nAll tests passed.');
