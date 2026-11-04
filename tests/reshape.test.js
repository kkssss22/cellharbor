import test from 'node:test';
import assert from 'node:assert/strict';
import { decodeDelimiter, reshapeText } from '../src/reshape.js';

test('default row fill preserves IDs, formula-like text and empty tokens', () => {
  const result = reshapeText('001,=1+1,,12,');
  assert.deepEqual(result.matrix, [['001', '=1+1', ''], ['12', '', '']]);
  assert.deepEqual(result.tokens, ['001', '=1+1', '', '12', '']);
  assert.equal(result.count, 5);
  assert.equal(result.padded, 1);
  assert.equal(result.rows, 2);
  assert.equal(result.columns, 3);
  assert.ok(result.warnings.some((warning) => warning.includes('CSV 引号')));
});

test('rows-only settings infer columns and support column-major filling', () => {
  const result = reshapeText('a,b,c,d,e', { rows: 2, fill: 'column' });
  assert.deepEqual(result.matrix, [['a', 'c', 'e'], ['b', 'd', '']]);
  assert.equal(result.columns, 3);
  assert.equal(result.padded, 1);
});

test('columns-only and explicit rectangles do not truncate', () => {
  assert.deepEqual(reshapeText('a,b,c', { columns: 2 }).matrix, [['a', 'b'], ['c', '']]);
  assert.deepEqual(reshapeText('a,b,c', { columns: 2, rows: 3 }).matrix, [['a', 'b'], ['c', ''], ['', '']]);
  assert.throws(() => reshapeText('a,b,c', { rows: 1, columns: 2 }), /只能放 2 项/);
});

test('literal multi-character and emoji delimiters are supported', () => {
  assert.deepEqual(reshapeText('甲🧩乙🧩🧩丙', { delimiter: '🧩', columns: 2 }).tokens, ['甲', '乙', '', '丙']);
  assert.deepEqual(reshapeText('a.*b.*c', { delimiter: '.*', columns: 3 }).matrix, [['a', 'b', 'c']]);
});

test('supported escape sequences and actual line endings are decoded consistently', () => {
  assert.equal(decodeDelimiter('\\t'), '\t');
  assert.equal(decodeDelimiter('\\n'), '\n');
  assert.equal(decodeDelimiter('\\r\\n'), '\n');
  assert.equal(decodeDelimiter('\r\n'), '\n');
  assert.equal(decodeDelimiter('\\r'), '\n');
  assert.equal(decodeDelimiter('\\d'), '\\d');
  assert.deepEqual(reshapeText('甲\r\n乙\r\n', { delimiter: '\\r\\n', columns: 2 }).tokens, ['甲', '乙', '']);
});

test('explicit row boundaries preserve source rows and infer the widest row', () => {
  const result = reshapeText('a\tb\r\nc\r\n', { delimiter: '\\t', rowDelimiter: '\\n' });
  assert.deepEqual(result.matrix, [['a', 'b'], ['c', ''], ['', '']]);
  assert.deepEqual(result.tokens, ['a', 'b', 'c', '']);
  assert.equal(result.rows, 3);
  assert.equal(result.columns, 2);
  assert.equal(result.padded, 2);
});

test('row boundaries reject insufficient dimensions and ignore column fill with a warning', () => {
  assert.throws(() => reshapeText('a,b\nc', { rowDelimiter: '\\n', columns: 1 }), /指定列数不足/);
  assert.throws(() => reshapeText('a,b\nc', { rowDelimiter: '\\n', rows: 1 }), /指定行数不足/);
  const result = reshapeText('a,b\nc', { rowDelimiter: '\\n', rows: 3, columns: 3, fill: 'column' });
  assert.deepEqual(result.matrix, [['a', 'b', ''], ['c', '', ''], ['', '', '']]);
  assert.ok(result.warnings.some((warning) => warning.includes('不使用按列填充')));
});

test('column boundaries keep source group order and pad uneven column heights', () => {
  const result = reshapeText('a,b|c,d,e|', { rowDelimiter: '|', boundary: 'column' });
  assert.deepEqual(result.matrix, [['a', 'c', ''], ['b', 'd', ''], ['', 'e', '']]);
  assert.deepEqual(result.tokens, ['a', 'b', 'c', 'd', 'e', '']);
  assert.equal(result.rows, 3);
  assert.equal(result.columns, 3);
  assert.equal(result.count, 6);
  assert.equal(result.padded, 3);
  assert.ok(result.warnings.some((warning) => warning.includes('每组内容向下排列')));
});

test('column boundaries respect dimensions, retain all values and ignore fill', () => {
  const options = { delimiter: ',', rowDelimiter: '|', boundary: 'column' };
  assert.deepEqual(reshapeText('a,b|c', { ...options, rows: 3, columns: 3, fill: 'column' }).matrix,
    [['a', 'c', ''], ['b', '', ''], ['', '', '']]);
  assert.throws(() => reshapeText('a,b,c|d', { ...options, rows: 2 }), /指定行数不足/);
  assert.throws(() => reshapeText('a|b', { ...options, columns: 1 }), /指定列数不足/);
  assert.equal(reshapeText('a,b|c', { ...options, rows: 4 }).columns, 2);
  assert.equal(reshapeText('a,b|c', { ...options, columns: 4 }).rows, 2);
});

test('column boundary validation applies limits and leaves flat filling unchanged', () => {
  assert.throws(() => reshapeText('a', { boundary: 'diagonal' }), /换行或换列/);
  assert.throws(() => reshapeText(Array(201).fill('a').join('|'), { rowDelimiter: '|', boundary: 'column' }), /超过 200 列/);
  assert.throws(() => reshapeText(Array(2001).fill('a').join(','), { rowDelimiter: '|', boundary: 'column' }), /超过 2000 行/);
  assert.deepEqual(reshapeText('a,b,c,d', { columns: 2, boundary: 'column' }).matrix, [['a', 'b'], ['c', 'd']]);
  assert.deepEqual(reshapeText('a\tb\r\nc', { delimiter: '\\t', rowDelimiter: '\\n', boundary: 'column' }).matrix, [['a', 'c'], ['b', '']]);
});

test('optional empty removal does not trim whitespace or delete explicit source rows', () => {
  const flat = reshapeText(',a,, ,', { keepEmpty: false, columns: 2 });
  assert.deepEqual(flat.tokens, ['a', ' ']);
  assert.equal(flat.count, 2);
  assert.ok(flat.warnings.some((warning) => warning.includes('3 个空项')));
  assert.deepEqual(reshapeText('a,,\n\nb', { rowDelimiter: '\\n', keepEmpty: false }).matrix, [['a'], [''], ['b']]);
  assert.equal(reshapeText(','.repeat(50001), { keepEmpty: false }).count, 0);
});

test('empty input has an explicit, valid rectangular result', () => {
  assert.deepEqual(reshapeText('').tokens, ['']);
  const result = reshapeText('', { keepEmpty: false, rows: 2 });
  assert.deepEqual(result.matrix, [[''], ['']]);
  assert.equal(result.count, 0);
  assert.equal(result.padded, 2);
});

test('CSV quotation marks have no special meaning and are retained', () => {
  assert.deepEqual(reshapeText('"a,b",c', { columns: 3 }).matrix, [['"a', 'b"', 'c']]);
});

test('invalid delimiters, direction, dimensions and types have guided errors', () => {
  assert.throws(() => reshapeText('x', { delimiter: '' }), /不能为空/);
  assert.throws(() => reshapeText('x', { delimiter: '\\n', rowDelimiter: '\r\n' }), /不能相同/);
  assert.throws(() => reshapeText('x', { rows: 0 }), /大于 0 的整数/);
  assert.throws(() => reshapeText('x', { columns: 1.5 }), /大于 0 的整数/);
  assert.throws(() => reshapeText('x', { columns: '2' }), /大于 0 的整数/);
  assert.throws(() => reshapeText('x', { rows: Infinity }), /大于 0 的整数/);
  assert.throws(() => reshapeText('x', { fill: 'diagonal' }), /按行或按列/);
  assert.throws(() => reshapeText('x', { keepEmpty: 1 }), /true 或 false/);
  assert.throws(() => reshapeText(null), /需要拆分/);
  assert.throws(() => reshapeText('x', null), /排版设置/);
  assert.throws(() => decodeDelimiter(1), /分隔符必须/);
});

test('dimension and total-cell limits apply to padded and inferred matrices', () => {
  assert.throws(() => reshapeText('x', { rows: 2001 }), /最多为 2000/);
  assert.throws(() => reshapeText('x', { columns: 201 }), /最多为 200/);
  assert.throws(() => reshapeText('x', { rows: 2000, columns: 26 }), /50000/);
  assert.throws(() => reshapeText(Array(2001).fill('x').join(','), { columns: 1 }), /超过 2000 行/);
  assert.throws(() => reshapeText(Array(201).fill('x').join(','), { rows: 1 }), /超过 200 列/);
  assert.throws(() => reshapeText(Array(2001).fill('x').join('\n'), { rowDelimiter: '\\n' }), /超过 2000 行/);
});

test('token limits, per-cell limits and UTF-8 byte limits are enforced', () => {
  assert.equal(reshapeText(Array(50000).fill('x').join(','), { columns: 200 }).count, 50000);
  assert.throws(() => reshapeText(Array(50001).fill('x').join(','), { columns: 200 }), /50000/);
  assert.equal(reshapeText('x'.repeat(32767)).tokens[0].length, 32767);
  assert.throws(() => reshapeText('x'.repeat(32768)), /32767/);
  assert.throws(() => reshapeText('x'.repeat(2 * 1024 * 1024 + 1)), /2 MB/);
  assert.throws(() => reshapeText('汉'.repeat(700000)), /2 MB/);
});
