import test from 'node:test';
import assert from 'node:assert/strict';
import { safeCell, validatePatches, verifyUnchanged, connectOffice, readSelection, applyPatches, undoPatches, createWorksheet } from '../src/office.js';

const snapshot = (values = [['Name', 'Code'], [' Alice ', '00123']]) => ({ values, formulas: values.map((r) => [...r]), worksheetId: 'sheet-id', rowIndex: 2, columnIndex: 3, address: "'Demo'!D3:E4" });

test('all strings are protected literals without losing existing apostrophes', () => {
  for (const value of ['=SUM(A1:A2)', '+1', '-1', '00123', "'literal", 'normal', '']) assert.equal(safeCell(value), `'${value}`);
  assert.equal(safeCell(12), 12);
  assert.equal(safeCell(false), false);
  assert.equal(safeCell(null), null);
  assert.throws(() => safeCell(Infinity));
  assert.throws(() => safeCell({}));
});

test('plans reject formulas, stale values, duplicates, bad coordinates and too many writes', () => {
  const s = snapshot();
  const good = { r: 1, c: 0, before: ' Alice ', after: 'Alice' };
  assert.equal(validatePatches(s, [good]), true);
  assert.throws(() => validatePatches(s, [good, good]), /重复/);
  assert.throws(() => validatePatches(s, [{ ...good, before: 'Bob' }]), /原始值/);
  assert.throws(() => validatePatches(s, [{ ...good, r: -1 }]), /位置/);
  assert.throws(() => validatePatches(s, Array(2001).fill(good)), /2,000/);
  s.formulas[1][0] = '=A1';
  assert.throws(() => validatePatches(s, [good]), /公式/);
});

test('snapshot checks catch value-type changes and changed formulas with same value', () => {
  const s = snapshot();
  assert.equal(verifyUnchanged(s, s.values, s.formulas), true);
  const v = structuredClone(s.values); v[1][1] = 123;
  assert.throws(() => verifyUnchanged(s, v, s.formulas), /变化/);
  const f = structuredClone(s.formulas); f[1][1] = '=123';
  assert.throws(() => verifyUnchanged(s, s.values, f), /变化/);
});

// A narrow Office object-model mock. It models the documented apostrophe text
// entry convention; real Excel host compatibility still requires a host test.
function mockExcel(input = [['Name', 'Code'], [' Alice ', '00123']]) {
  const state = { values: structuredClone(input), formulas: structuredClone(input), syncs: 0, writes: [], loaded: [], selected: true, sheetAdds: [], formatsTouched: false, fakeRows: null };
  const decode = (value) => typeof value === 'string' && value.startsWith("'") ? value.slice(1) : value;
  const range = {
    get rowCount() { return state.fakeRows ?? state.values.length; },
    get columnCount() { return state.values[0].length; }, rowIndex: 2, columnIndex: 3,
    address: "'Demo'!D3:E4",
    get values() { return state.values; },
    set values(matrix) { state.values = matrix.map((r) => r.map(decode)); state.formulas = structuredClone(state.values); state.writes.push('matrix'); },
    get formulas() { return state.formulas; },
    set numberFormat(_) { state.formatsTouched = true; throw new Error('number formats must be preserved'); },
    load(properties) { state.loaded.push(properties); },
    getCell(r, c) { return { set values(matrix) { const value = decode(matrix[0][0]); state.values[r][c] = value; state.formulas[r][c] = value; state.writes.push([r, c, matrix[0][0]]); } }; },
  };
  const sheet = { id: 'sheet-id', load() {}, getRangeByIndexes(r, c, rows, cols) { state.lastCoordinates = [r, c, rows, cols]; return range; }, activate() { state.activated = true; } };
  range.worksheet = sheet;
  const context = { workbook: { getSelectedRange() { if (!state.selected) throw new Error('selection must not be used'); return range; }, worksheets: { getItem(id) { assert.equal(id, 'sheet-id'); return sheet; }, add(name) { state.sheetAdds.push(name); return sheet; } } }, async sync() { state.syncs += 1; if (state.failSync === state.syncs) throw new Error('Host interrupted'); } };
  globalThis.Excel = { run: async (fn) => fn(context) };
  return state;
}

test('readSelection checks size before loading values and records stable worksheet coordinates', async () => {
  const state = mockExcel();
  const result = await readSelection();
  assert.equal(result.worksheetId, 'sheet-id');
  assert.equal(result.rowIndex, 2);
  assert.deepEqual(result.values, state.values);
  state.fakeRows = 50_001;
  state.loaded = [];
  await assert.rejects(readSelection(), /50,000/);
  assert.equal(state.loaded.some((s) => s.includes('values')), false);
});

test('apply and undo use captured coordinates; preserve text types, formats and unrelated edits', async () => {
  const state = mockExcel();
  const s = await readSelection();
  state.selected = false;
  const undo = await applyPatches(s, [{ r: 1, c: 0, before: ' Alice ', after: '=NOT_A_FORMULA' }]);
  assert.equal(state.values[1][0], '=NOT_A_FORMULA');
  assert.equal(state.writes[0][2], "'=NOT_A_FORMULA");
  assert.deepEqual(state.lastCoordinates, [2, 3, 2, 2]);
  state.values[0][1] = 'New heading'; state.formulas[0][1] = 'New heading';
  await undoPatches(undo);
  assert.equal(state.values[1][0], ' Alice ');
  assert.equal(state.values[0][1], 'New heading');
  assert.equal(state.formatsTouched, false);
});

test('apply refuses stale previews and undo refuses a changed target formula', async () => {
  let state = mockExcel();
  const s = await readSelection();
  state.values[0][0] = 'Changed';
  await assert.rejects(applyPatches(s, [{ r: 1, c: 0, before: ' Alice ', after: 'Alice' }]), /变化/);
  assert.equal(state.writes.length, 0);
  state = mockExcel();
  const s2 = await readSelection();
  const undo = await applyPatches(s2, [{ r: 1, c: 0, before: ' Alice ', after: 'Alice' }]);
  state.formulas[1][0] = '="Alice"';
  const writes = state.writes.length;
  await assert.rejects(undoPatches(undo), /未执行撤销/);
  assert.equal(state.writes.length, writes);
});

test('new sheets write values as literals and reject oversized or ragged matrices before mutation', async () => {
  const state = mockExcel();
  const result = await createWorksheet([['00123', '=A1'], [5, true]]);
  assert.match(result.name, /^格子结果_/);
  assert.deepEqual(state.values, [['00123', '=A1'], [5, true]]);
  assert.equal(state.activated, true);
  await assert.rejects(createWorksheet([[1], [1, 2]]), /列数/);
  await assert.rejects(createWorksheet([Array(50_001).fill('x')]), /50,000/);
  assert.equal(state.sheetAdds.length, 1);
});

test('host failures after write give explicit partial-change warning', async () => {
  const state = mockExcel();
  const s = await readSelection();
  state.failSync = state.syncs + 2;
  await assert.rejects(applyPatches(s, [{ r: 1, c: 0, before: ' Alice ', after: 'Alice' }]), /部分单元格可能已修改/);
});

test('Office connection rejects other hosts and unsupported APIs, then permits a corrected host', async () => {
  globalThis.Office = { HostType: { Excel: 'Excel' }, onReady: async () => ({ host: 'Word' }), context: { requirements: { isSetSupported: () => true } } };
  await assert.rejects(connectOffice(), /需要在Excel/);
  globalThis.Office.onReady = async () => ({ host: 'Excel', platform: 'PC' });
  globalThis.Office.context.requirements.isSetSupported = () => false;
  await assert.rejects(connectOffice(), /ExcelApi 1.4/);
  globalThis.Office.context.requirements.isSetSupported = (name, version) => name === 'ExcelApi' && version === '1.4';
  mockExcel();
  assert.equal((await connectOffice()).host, 'Excel');
});
