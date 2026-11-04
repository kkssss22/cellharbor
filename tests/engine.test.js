import test from 'node:test';
import assert from 'node:assert/strict';
import { ACTIONS, preview } from '../src/engine.js';

function apply(values, result) {
  const copy = values.map((row) => [...row]);
  for (const { r, c, after } of result.patches) copy[r][c] = after;
  return copy;
}

test('catalog has complete unique metadata for at least 25 useful actions', () => {
  assert.ok(ACTIONS.length >= 25);
  assert.equal(new Set(ACTIONS.map((a) => a.id)).size, ACTIONS.length);
  assert.ok(ACTIONS.every((a) => a.title && a.category && a.description && ['transform', 'inspect', 'extract'].includes(a.kind)));
});

test('preview preserves the input, headers and every formula across transformations', () => {
  const values = [[' Name ', ' Value '], [' Alice ', -5], [' = A1 ', '  ']];
  const formulas = [['', ''], ['', '=2-7'], ['=A1', '=IF(TRUE,"",0)']];
  const frozen = values.map((row) => Object.freeze([...row])); Object.freeze(frozen);
  for (const action of ACTIONS.filter((a) => a.kind === 'transform')) {
    const result = preview(action.id, frozen, { formulas, find: ' ', replace: '', digits: 2 });
    assert.equal(result.summary.skippedFormulas, 3, action.id);
    assert.ok(result.patches.every((p) => p.r !== 0 && !formulas[p.r][p.c].startsWith('=')), action.id);
  }
  assert.deepEqual(frozen, values);
});

test('whitespace operations distinguish ends, all whitespace and normalized gaps', () => {
  const values = [['  A\t B \n C  ', 0, false, null]];
  assert.deepEqual(apply(values, preview('trim', values, { hasHeader: false })), [['A\t B \n C', 0, false, null]]);
  assert.equal(preview('removeSpaces', values, { hasHeader: false }).patches[0].after, 'ABC');
  assert.equal(preview('normalizeSpaces', values, { hasHeader: false }).patches[0].after, 'A B C');
  assert.equal(preview('removeInvisible', [['A\u200bB\u2060C\ufeff']], { hasHeader: false }).patches[0].after, 'ABC');
});

test('case conversion and full-width normalization operate on text only', () => {
  assert.equal(preview('titleCase', [['hELLO-world ÉCOLE 中文']], { hasHeader: false }).patches[0].after, 'Hello-World École 中文');
  assert.equal(preview('fullWidthToHalf', [['ＡＢ１２，　中文']], { hasHeader: false }).patches[0].after, 'AB12, 中文');
  assert.equal(preview('upper', [[false, 12, 'ab']], { hasHeader: false }).patches.length, 1);
});

test('number conversion protects leading-zero identifiers, long integers and precision', () => {
  const values = [['00123', '1234567890123456', '1e16', '-0012', '0.1234567890123456', '12.5', ' .5 ', '+12', '-2e2', '1,234', '', '1e-999']];
  const result = preview('textToNumber', values, { hasHeader: false });
  assert.deepEqual(result.patches.map((p) => [p.c, p.after]), [[5, 12.5], [6, 0.5], [7, 12], [8, -200]]);
  assert.equal(result.findings.length, 6);
  assert.ok(result.warnings.some((warning) => warning.includes('精度')));
  assert.equal(preview('textToNumber', [['999999999999999']], { hasHeader: false }).patches[0].after, 999999999999999);
});

test('percent text is strictly parsed as a fraction without changing numeric cells', () => {
  const values = [['12.5%', ' 05% ', '100%', '0%', '-3%', '12%%', 'abc%', '10%extra', 0.2]];
  const result = preview('percentToNumber', values, { hasHeader: false });
  assert.deepEqual(result.patches.map((p) => p.after), [0.125, 0.05, 1, 0, -0.03]);
});

test('rounding is half away from zero and handles negative decimal places', () => {
  const values = [[1.005, -1.005, 2.675, -2.675, 0, '2.675']];
  assert.deepEqual(apply(values, preview('round', values, { hasHeader: false, digits: 2 })), [[1.01, -1.01, 2.68, -2.68, 0, '2.675']]);
  assert.deepEqual(apply([[150, -150, 149]], preview('round', [[150, -150, 149]], { hasHeader: false, digits: -2 })), [[200, -200, 100]]);
  assert.throws(() => preview('round', [[1]], { digits: 1.5 }), /整数/);
  assert.equal(preview('round', [[1e308]], { hasHeader: false, digits: 15 }).patches.length, 0);
});

test('literal replacement does not treat regex characters or dollar signs as syntax', () => {
  const values = [['a.b.* a.b.*']];
  const result = preview('findReplace', values, { hasHeader: false, find: '.*', replace: '$&' });
  assert.equal(result.patches[0].after, 'a.b$& a.b$&');
  assert.throws(() => preview('findReplace', values, { find: '' }), /不能为空/);
  assert.throws(() => preview('findReplace', [['a'.repeat(32767)]], { hasHeader: false, find: 'a', replace: 'aa' }), /32,767/);
});

test('fill-down never inherits the header and treats a formula as a barrier', () => {
  const values = [['Team'], [' '], ['A'], [null], ['computed'], [''], ['B'], ['']];
  const formulas = values.map(() => ['']); formulas[4][0] = '=A1';
  const result = preview('fillDown', values, { formulas });
  assert.deepEqual(result.patches.map((p) => [p.r, p.after]), [[3, 'A'], [7, 'B']]);
  assert.equal(result.summary.skippedFormulas, 1);
});

test('blank checks and zero-fill keep false and zero distinct from missing values', () => {
  const values = [['Header'], [null], [''], ['  \t'], [false], [0]];
  assert.equal(preview('blanks', values).findings.length, 3);
  assert.equal(preview('blankToZero', values).patches.length, 3);
});

test('duplicates compare full rows and distinguish numbers, text, null and blanks', () => {
  const values = [['id', 'value'], [1, 'a'], ['1', 'a'], [1, 'a'], [null, 'a'], ['', 'a']];
  const result = preview('duplicates', values);
  assert.deepEqual(result.findings.map((f) => [f.r, f.c]), [[1, 0], [1, 1], [3, 0], [3, 1]]);
  assert.equal(result.patches.length, 0);
  assert.deepEqual(preview('uniqueRows', values).output, [values[0], values[1], values[2], values[4], values[5]]);
});

test('type audit reports minority types and all members of a tied column', () => {
  const values = [['majority', 'tie'], [1, 1], [2, 'x'], ['3', null], [null, '']];
  const result = preview('inconsistentTypes', values);
  assert.deepEqual(result.findings.map((f) => [f.r, f.c]), [[3, 0], [1, 1], [2, 1]]);
});

test('numeric-text and long-number audits leave source data unchanged', () => {
  const values = [['00123', '1234567890123456', 1e16, 123, 'word']];
  assert.equal(preview('numbersAsText', values, { hasHeader: false }).findings.length, 2);
  assert.equal(preview('longNumbers', values, { hasHeader: false }).findings.length, 2);
});

test('stats count true numeric cells, not text or booleans, and report blanks', () => {
  const values = [['Score', 'Empty'], [2, null], [4, ' '], ['6', ''], [false, null]];
  assert.deepEqual(preview('stats', values).output, [
    ['列', '数值个数', '合计', '平均值', '最小值', '最大值', '空白数'],
    ['Score', 2, 6, 3, 2, 4, 0], ['Empty', 0, 0, null, null, null, 4],
  ]);
  const overflow = preview('stats', [[1e308], [1e308]], { hasHeader: false });
  assert.equal(overflow.output[1][2], null);
  assert.equal(overflow.output[1][3], 1e308);
  assert.ok(overflow.warnings.some((warning) => warning.includes('范围')));
});

test('split-column pads shorter rows, preserves empty pieces and produces clear headers', () => {
  const values = [['Tags'], ['a,b'], ['c'], [',d,'], [null]];
  assert.deepEqual(preview('splitColumn', values).output, [
    ['Tags 1', 'Tags 2', 'Tags 3'], ['a', 'b', ''], ['c', '', ''], ['', 'd', ''], ['', '', ''],
  ]);
  assert.throws(() => preview('splitColumn', [[1, 2]]), /一列/);
  assert.throws(() => preview('splitColumn', values, { delimiter: '' }), /分隔符/);
  assert.deepEqual(preview('splitColumn', [['a.*b']], { hasHeader: false, delimiter: '.*' }).output, [['a', 'b']]);
  assert.throws(() => preview('splitColumn', Array.from({ length: 2000 }, () => [','.repeat(100)]), { hasHeader: false }), /100,000/);
  assert.throws(() => preview('splitColumn', [[','.repeat(16384)]], { hasHeader: false }), /16,384/);
});

test('extracts retain source and copy values with explicit formula limits', () => {
  const values = [['H1', 'H2'], [1, 2], ['', null], [false, 0]];
  assert.deepEqual(preview('removeEmptyRows', values).output, [values[0], values[1], values[3]]);
  assert.deepEqual(preview('transpose', values).output, [['H1', 1, '', false], ['H2', 2, null, 0]]);
  assert.deepEqual(preview('addSequence', values).output, [['序号'], [1], [2], [3]]);
  assert.deepEqual(preview('addSequence', [[1], [2]], { hasHeader: false }).output, [[1], [2]]);
  const result = preview('uniqueRows', values);
  assert.ok(result.warnings.some((warning) => warning.includes('不复制公式')));
  assert.notEqual(result.output[0], values[0]);
  result.output[1][0] = 'changed';
  assert.equal(values[1][0], 1);
});

test('empty extraction and header-only selections have a stable schema', () => {
  const empty = preview('removeEmptyRows', [[null], ['']], { hasHeader: false });
  assert.deepEqual(empty.output, []);
  assert.equal(empty.summary.outputRows, 0);
  assert.equal(empty.summary.outputColumns, 0);
  assert.equal(preview('trim', [['Header']]).summary.dataRows, 0);
  assert.deepEqual(preview('addSequence', [['Header']]).output, [['序号']]);
});

test('potential formula strings are preserved as data and flagged for safe host writing', () => {
  const result = preview('trim', [[' =SUM(A1:A2) ']], { hasHeader: false });
  assert.equal(result.patches[0].after, '=SUM(A1:A2)');
  assert.ok(result.warnings.some((warning) => warning.includes('不能执行为公式')));
  assert.equal(preview('trim', [[' =SUM(A1:A2) ']], { hasHeader: false, formulas: [['=SUM(A1:A2)']] }).patches.length, 0);
});

test('matrix validation rejects ragged, oversized or unsafe input instead of silently dropping it', () => {
  for (const values of [[], [[]], [[1], [1, 2]], [[undefined]], [[NaN]], [[Infinity]], [[{}]], [['x'.repeat(32768)]]]) {
    assert.throws(() => preview('trim', values));
  }
  assert.throws(() => preview('trim', Array.from({ length: 100001 }, () => [0])), /100,000/);
  assert.throws(() => preview('trim', [[1], [2]], { formulas: [['=1']] }), /尺寸/);
  assert.throws(() => preview('trim', [[1]], { hasHeader: 'yes' }), /布尔/);
  assert.throws(() => preview('unknown', [[1]]), /未知操作/);
  assert.throws(() => preview('transpose', Array.from({ length: 16385 }, () => [0]), { hasHeader: false }), /16,384/);
});
