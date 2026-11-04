/** Office-only bridge. Importing this module never loads Office.js or reads Excel. */
const OFFICE_URL = 'https://officeapis.public.onecdn.static.microsoft/1/office.js';
const MAX_CELLS = 50_000;
const MAX_PATCHES = 2_000;
const MAX_TEXT = 32_767;
let connectionPromise;
let mutationInProgress = false;

const clone = (matrix) => matrix.map((row) => [...row]);
const equal = (a, b) => a === b; // Excel does not preserve JavaScript negative zero.
const storedValue = (value) => value === null ? '' : value;

function assertValue(value) {
  if (value === null || typeof value === 'boolean') return;
  if (typeof value === 'number' && Number.isFinite(value)) return;
  if (typeof value === 'string' && value.length <= MAX_TEXT) return;
  throw new Error('单元格只支持有限数字、布尔值、空值或不超过32,767字符的文本。');
}

/** Every string is a literal, including =, +, -, leading zeros and apostrophes. */
export function safeCell(value) {
  assertValue(value);
  return typeof value === 'string' ? `'${value}` : value;
}

function shape(matrix, label = '数据') {
  if (!Array.isArray(matrix) || !matrix.length || !Array.isArray(matrix[0]) || !matrix[0].length) {
    throw new Error(`${label}必须包含至少一个单元格。`);
  }
  const rows = matrix.length;
  const cols = matrix[0].length;
  if (rows > 1_048_576 || cols > 16_384 || rows * cols > MAX_CELLS) {
    throw new Error('Excel模式一次最多处理50,000个单元格，请缩小范围。');
  }
  for (const row of matrix) {
    if (!Array.isArray(row) || row.length !== cols) throw new Error(`${label}每行列数必须一致。`);
    for (const value of row) assertValue(value);
  }
  return { rows, cols };
}

function validateSnapshot(snapshot) {
  if (!snapshot || typeof snapshot.worksheetId !== 'string' || !snapshot.worksheetId) {
    throw new Error('缺少工作表标识，请重新读取选区。');
  }
  const size = shape(snapshot.values);
  const formulasSize = shape(snapshot.formulas, '公式');
  if (size.rows !== formulasSize.rows || size.cols !== formulasSize.cols) throw new Error('数据与公式范围不一致。');
  for (const key of ['rowIndex', 'columnIndex']) {
    if (!Number.isInteger(snapshot[key]) || snapshot[key] < 0) throw new Error('选区坐标无效，请重新读取。');
  }
  if (snapshot.rowIndex + size.rows > 1_048_576 || snapshot.columnIndex + size.cols > 16_384) throw new Error('选区超出Excel范围。');
  return size;
}

/** Reject stale plans, duplicate targets, formula writes and invalid values before Office is touched. */
export function validatePatches(snapshot, patches) {
  const { rows, cols } = validateSnapshot(snapshot);
  if (!Array.isArray(patches) || patches.length > MAX_PATCHES) throw new Error('一次最多修改2,000个单元格，请分批操作。');
  const seen = new Set();
  for (const patch of patches) {
    if (!patch || !Number.isInteger(patch.r) || !Number.isInteger(patch.c) || patch.r < 0 || patch.c < 0 || patch.r >= rows || patch.c >= cols) throw new Error('修改位置超出预览范围。');
    const key = `${patch.r}:${patch.c}`;
    if (seen.has(key)) throw new Error('修改计划中有重复位置，请重新预览。');
    seen.add(key);
    if (!equal(patch.before, snapshot.values[patch.r][patch.c])) throw new Error('修改计划与原始值不一致，请重新预览。');
    if (typeof snapshot.formulas[patch.r][patch.c] === 'string' && snapshot.formulas[patch.r][patch.c].startsWith('=')) throw new Error('不能通过原地转换覆盖公式，请重新预览。');
    assertValue(patch.after);
  }
  return true;
}

/** Check the entire previewed range, including formulas, before applying changes. */
export function verifyUnchanged(snapshot, currentValues, currentFormulas) {
  const { rows, cols } = validateSnapshot(snapshot);
  const a = shape(currentValues);
  const b = shape(currentFormulas, '公式');
  if (a.rows !== rows || a.cols !== cols || b.rows !== rows || b.cols !== cols) throw new Error('选区尺寸已经变化，请重新读取并预览。');
  for (let r = 0; r < rows; r += 1) {
    for (let c = 0; c < cols; c += 1) {
      if (!equal(snapshot.values[r][c], currentValues[r][c]) || !equal(snapshot.formulas[r][c], currentFormulas[r][c])) {
        throw new Error('预览后单元格已经变化，未应用修改。请重新读取并预览。');
      }
    }
  }
  return true;
}

export async function connectOffice() {
  if (connectionPromise) return connectionPromise;
  connectionPromise = (async () => {
    let timer;
    const timeout = new Promise((_, reject) => {
      timer = setTimeout(() => reject(new Error('连接Excel超时，请在Excel加载项中打开并检查Microsoft脚本网络连接。')), 15_000);
    });
    const connect = async () => {
      if (!globalThis.Office) {
        if (!globalThis.document) throw new Error('请在Excel加载项中打开此页面。');
        await new Promise((resolve, reject) => {
          const script = document.createElement('script');
          script.src = OFFICE_URL;
          script.async = true;
          script.onload = resolve;
          script.onerror = () => { script.remove(); reject(new Error('无法加载Microsoft Office.js，请检查网络。')); };
          document.head.append(script);
        });
      }
      const info = await globalThis.Office.onReady();
      if (info.host !== globalThis.Office.HostType.Excel) throw new Error('格子助手需要在Excel中运行。');
      if (!globalThis.Office.context?.requirements?.isSetSupported('ExcelApi', '1.4')) throw new Error('此Excel版本不支持所需的ExcelApi 1.4，请更新Excel。');
      if (!globalThis.Excel?.run) throw new Error('Excel接口尚未就绪，请重新打开加载项。');
      return info;
    };
    try { return await Promise.race([connect(), timeout]); }
    finally { clearTimeout(timer); }
  })();
  try { return await connectionPromise; }
  catch (error) { connectionPromise = undefined; throw error; }
}

function excelRun(callback) {
  if (!globalThis.Excel?.run) throw new Error('尚未连接Excel，请先打开Excel加载项。');
  return globalThis.Excel.run(callback);
}

export async function readSelection() {
  return excelRun(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load('rowCount,columnCount,address,rowIndex,columnIndex');
    range.worksheet.load('id');
    await context.sync();
    if (range.rowCount * range.columnCount > MAX_CELLS) throw new Error('选区超过50,000个单元格，请选中实际数据区域。');
    if (!range.rowCount || !range.columnCount) throw new Error('请先选择一个连续的矩形区域。');
    range.load('values,formulas');
    await context.sync();
    const snapshot = {
      values: clone(range.values), formulas: clone(range.formulas), address: range.address,
      worksheetId: range.worksheet.id, rowIndex: range.rowIndex, columnIndex: range.columnIndex,
    };
    validateSnapshot(snapshot);
    return snapshot;
  });
}

async function mutate(callback) {
  if (mutationInProgress) throw new Error('上一项操作尚未完成，请稍候。');
  mutationInProgress = true;
  try { return await callback(); }
  finally { mutationInProgress = false; }
}

function originalRange(context, snapshot) {
  const { rows, cols } = validateSnapshot(snapshot);
  const worksheet = context.workbook.worksheets.getItem(snapshot.worksheetId);
  return worksheet.getRangeByIndexes(snapshot.rowIndex, snapshot.columnIndex, rows, cols);
}

export async function applyPatches(snapshot, patches) {
  validatePatches(snapshot, patches);
  const saved = { ...snapshot, values: clone(snapshot.values), formulas: clone(snapshot.formulas) };
  const plan = patches.map((patch) => ({ ...patch }));
  return mutate(() => excelRun(async (context) => {
    const range = originalRange(context, saved);
    range.load('values,formulas');
    await context.sync();
    verifyUnchanged(saved, range.values, range.formulas);
    let writeStarted = false;
    try {
      // Office batches are not transactions. A host failure can leave partial changes.
      // No automatic rollback: it could overwrite a concurrent edit.
      for (const patch of plan) range.getCell(patch.r, patch.c).values = [[safeCell(storedValue(patch.after))]];
      writeStarted = plan.length > 0;
      await context.sync();
      range.load('values,formulas');
      await context.sync();
      for (const patch of plan) {
        if (!equal(range.values[patch.r][patch.c], storedValue(patch.after))) throw new Error('Excel返回的值或数据类型与预览不一致。');
        patch.afterFormula = range.formulas[patch.r][patch.c];
        // For a literal '=text', Range.formulas returns the value itself.
        if (typeof patch.afterFormula === 'string' && patch.afterFormula.startsWith('=') && patch.afterFormula !== range.values[patch.r][patch.c]) throw new Error('Excel把文本解释为公式。');
      }
      return { snapshot: saved, patches: plan };
    } catch (error) {
      if (writeStarted) throw new Error(`写入或校验未完成，部分单元格可能已修改。请检查原选区后再操作。${error.message}`, { cause: error });
      throw error;
    }
  }));
}

export async function undoPatches(undo) {
  if (!undo?.snapshot || !Array.isArray(undo.patches)) throw new Error('没有可以撤销的原地转换。');
  validatePatches(undo.snapshot, undo.patches);
  return mutate(() => excelRun(async (context) => {
    const range = originalRange(context, undo.snapshot);
    range.load('values,formulas');
    await context.sync();
    // Only affected cells block undo. Other cells in the old selection may change.
    for (const patch of undo.patches) {
      if (!Object.hasOwn(patch, 'afterFormula') || !equal(range.values[patch.r][patch.c], storedValue(patch.after)) || !equal(range.formulas[patch.r][patch.c], patch.afterFormula)) {
        throw new Error('目标单元格在操作后已变化，未执行撤销，以免覆盖新内容。');
      }
    }
    try {
      for (const patch of undo.patches) range.getCell(patch.r, patch.c).values = [[safeCell(storedValue(patch.before))]];
      await context.sync();
      range.load('values,formulas');
      await context.sync();
      for (const patch of undo.patches) {
        if (!equal(range.values[patch.r][patch.c], storedValue(patch.before))) throw new Error('还原值校验失败。');
      }
      return true;
    } catch (error) {
      throw new Error(`撤销未能完整确认，部分单元格可能已还原。请检查原选区。${error.message}`, { cause: error });
    }
  }));
}

export async function createWorksheet(matrix) {
  const { rows, cols } = shape(matrix);
  const data = matrix.map((row) => row.map((value) => safeCell(storedValue(value))));
  return mutate(() => excelRun(async (context) => {
    const stamp = new Date().toTimeString().slice(0, 8).replaceAll(':', '');
    const random = Math.random().toString(36).slice(2, 7);
    const name = `格子结果_${stamp}_${random}`;
    let created = false;
    try {
      const sheet = context.workbook.worksheets.add(name);
      const range = sheet.getRangeByIndexes(0, 0, rows, cols);
      range.values = data;
      sheet.activate();
      created = true;
      await context.sync();
      range.load('address,values,formulas');
      await context.sync();
      for (let r = 0; r < rows; r += 1) {
        for (let c = 0; c < cols; c += 1) {
          if (!equal(range.values[r][c], storedValue(matrix[r][c]))) throw new Error('输出数据或类型与预览不一致。');
          const formula = range.formulas[r][c];
          if (typeof formula === 'string' && formula.startsWith('=') && formula !== range.values[r][c]) throw new Error('输出文本被解释为公式。');
        }
      }
      return { name, address: range.address };
    } catch (error) {
      throw new Error(`${created ? `结果表“${name}”可能已创建，请检查它；未自动删除。` : ''}${error.message}`, { cause: error });
    }
  }));
}
