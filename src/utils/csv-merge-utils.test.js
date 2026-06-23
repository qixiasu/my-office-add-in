/**
 * Unit tests for csv-merge-utils.js
 */

var csvMerge = require('./csv-merge-utils');

describe('computeUnifiedHeaders', () => {
  test('merges headers from two arrays', () => {
    expect(csvMerge.computeUnifiedHeaders([['A', 'B'], ['B', 'C']])).toEqual(['A', 'B', 'C']);
  });

  test('preserves order from three separate arrays', () => {
    expect(csvMerge.computeUnifiedHeaders([['X'], ['Y'], ['Z']])).toEqual(['X', 'Y', 'Z']);
  });

  test('merges headers with overlapping columns', () => {
    expect(csvMerge.computeUnifiedHeaders([['name', 'age'], ['name', 'city']])).toEqual(['name', 'age', 'city']);
  });

  test('handles empty header arrays', () => {
    expect(csvMerge.computeUnifiedHeaders([[], ['B']])).toEqual(['B']);
  });

  test('removes duplicate headers', () => {
    expect(csvMerge.computeUnifiedHeaders([['A', 'B', 'A'], ['C', 'B']])).toEqual(['A', 'B', 'C']);
  });

  test('returns empty array for empty input', () => {
    expect(csvMerge.computeUnifiedHeaders([])).toEqual([]);
  });

  test('returns empty array for array with empty row', () => {
    expect(csvMerge.computeUnifiedHeaders([[]])).toEqual([]);
  });

  test('handles single row with empty headers', () => {
    expect(csvMerge.computeUnifiedHeaders([['A'], []])).toEqual(['A']);
  });
});

describe('alignColumns', () => {
  test('test_alignColumns_differentOrder', () => {
    var filesData = [
      { headers: ['A', 'B'], data: [['vA1', 'vB1'], ['vA2', 'vB2']] },
      { headers: ['B', 'A'], data: [['vB3', 'vA3']] },
    ];
    var result = csvMerge.alignColumns(filesData);
    expect(result.headers).toEqual(['A', 'B']);
    expect(result.data).toEqual([
      ['vA1', 'vB1'],
      ['vA2', 'vB2'],
      ['vA3', 'vB3'],
    ]);
  });

  test('test_alignColumns_threeFilesUnion', () => {
    var filesData = [
      { headers: ['name', 'age'],   data: [['Alice', '30']] },
      { headers: ['city', 'age'],   data: [['NYC', '25']] },
      { headers: ['name', 'email'], data: [['Bob', 'bob@example.com']] },
    ];
    var result = csvMerge.alignColumns(filesData);
    expect(result.headers).toEqual(['name', 'age', 'city', 'email']);
    expect(result.data).toEqual([
      ['Alice', '30',   '',              ''],
      ['',       '25',  'NYC',           ''],
      ['Bob',    '',    '', 'bob@example.com'],
    ]);
  });
});

describe('serializeRow', () => {
  test('test_serializeRow_withQuotes', () => {
    // Fields with commas and double-quotes need quoting; quotes inside are doubled
    var row = ['Hello, World', 'Say "Hi"', 'Normal'];
    var result = csvMerge.serializeRow(row, ',');
    expect(result).toEqual('"Hello, World","Say ""Hi""",Normal');
  });

  test('test_serializeRow_noQuotes', () => {
    expect(csvMerge.serializeRow(['a', 'b', 'c'])).toEqual('a,b,c');
  });

  test('test_serializeRow_withNewline', () => {
    expect(csvMerge.serializeRow(['line1\nline2', 'x'])).toEqual('"line1\nline2",x');
  });

  test('test_serializeRow_emptyFields', () => {
    expect(csvMerge.serializeRow(['', '', ''])).toEqual(',,');
  });

  test('test_serializeRow_withCarriageReturn', () => {
    expect(csvMerge.serializeRow(['a\rb', 'c'])).toEqual('"a\rb",c');
  });
});
