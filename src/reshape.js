const LIMITS = Object.freeze({
  bytes: 2 * 1024 * 1024,
  cells: 50000,
  rows: 2000,
  columns: 200,
  cellLength: 32767,
});

const normalizeNewlines = (value) => value.replace(/\r\n?/g, '\n');

/** Decode the supported visible delimiter escapes; all other text stays literal. */
export function decodeDelimiter(value) {
  if (typeof value !== 'string') {
    throw new TypeError('分隔符必须是文字，可输入逗号、\\t 或 \\n。');
  }
  return normalizeNewlines(value.replace(/\\r\\n|\\n|\\t|\\r/g, (escape) => ({
    '\\r\\n': '\n',
    '\\n': '\n',
    '\\t': '\t',
    '\\r': '\n',
  })[escape]));
}

function dimension(value, label, maximum) {
  if (value === undefined || value === null) return null;
  if (typeof value !== 'number' || !Number.isSafeInteger(value) || value < 1) {
    throw new RangeError(`${label}必须是大于 0 的整数，请清空该项让工具自动计算。`);
  }
  if (value > maximum) throw new RangeError(`${label}最多为 ${maximum}，请减少${label}。`);
  return value;
}

function validateSize(rows, columns) {
  if (rows > LIMITS.rows) throw new RangeError('结果超过 2000 行，请增加列数或减少输入内容。');
  if (columns > LIMITS.columns) throw new RangeError('结果超过 200 列，请增加行数或减少输入内容。');
  if (rows * columns > LIMITS.cells) {
    throw new RangeError('结果最多为 50000 个单元格，请减少行数、列数或输入内容。');
  }
}

function validateTokens(tokens) {
  if (tokens.length > LIMITS.cells) {
    throw new RangeError('拆分结果超过 50000 个单元格，请分批处理。');
  }
  const tooLong = tokens.findIndex((value) => value.length > LIMITS.cellLength);
  if (tooLong !== -1) {
    throw new RangeError(`第 ${tooLong + 1} 项超过单元格 32767 字符上限，请增加分隔符或缩短内容。`);
  }
}

/**
 * Split literal text into an Excel-shaped string matrix without coercion.
 * CRLF and CR normalize to LF in both input and delimiter settings.
 * The legacy rowDelimiter option separates source groups. boundary chooses
 * whether each group becomes a row (default) or a column. Without a group
 * delimiter, tokens fill the rectangle in row-major or column-major order.
 */
export function reshapeText(text, options = {}) {
  if (typeof text !== 'string') throw new TypeError('请输入需要拆分的文字。');
  if (!options || typeof options !== 'object' || Array.isArray(options)) {
    throw new TypeError('排版设置必须是有效的选项。');
  }
  if (text.length > LIMITS.bytes || new TextEncoder().encode(text).byteLength > LIMITS.bytes) {
    throw new RangeError('输入文字超过 2 MB，请分批处理。');
  }

  const delimiter = decodeDelimiter(options.delimiter === undefined ? ',' : options.delimiter);
  const rowDelimiter = decodeDelimiter(options.rowDelimiter === undefined ? '' : options.rowDelimiter);
  if (!delimiter) throw new RangeError('单元格分隔符不能为空，请输入逗号、空格、\\t 或其他文字。');
  if (rowDelimiter && delimiter === rowDelimiter) {
    throw new RangeError('换行/换列分隔符与单元格分隔符不能相同，请使用不同的分隔符。');
  }
  const fill = options.fill === undefined ? 'row' : options.fill;
  if (fill !== 'row' && fill !== 'column') throw new RangeError('填充方向请选择按行或按列。');
  const boundary = options.boundary === undefined ? 'row' : options.boundary;
  if (boundary !== 'row' && boundary !== 'column') throw new RangeError('分组方向请选择换行或换列。');
  const keepEmpty = options.keepEmpty === undefined ? true : options.keepEmpty;
  if (typeof keepEmpty !== 'boolean') throw new TypeError('保留空项设置必须为 true 或 false。');

  let columns = dimension(options.columns, '列数', LIMITS.columns);
  let rows = dimension(options.rows, '行数', LIMITS.rows);
  const input = normalizeNewlines(text);
  const warnings = ['按分隔符逐字拆分，不处理 CSV 引号转义'];
  let removed = 0;
  const splitCells = (source) => {
    // Walk the literal delimiter so discarded empty fields never allocate a
    // large temporary array, and only retained fields count toward the limit.
    const cells = [];
    let cursor = 0;
    for (;;) {
      const next = source.indexOf(delimiter, cursor);
      const cell = source.slice(cursor, next === -1 ? source.length : next);
      if (keepEmpty || cell !== '') cells.push(cell);
      else removed++;
      if (cells.length > LIMITS.cells) {
        throw new RangeError('拆分结果超过 50000 项，请分批处理或减少分隔符。');
      }
      if (next === -1) break;
      cursor = next + delimiter.length;
    }
    return cells;
  };

  let tokens;
  let matrix;
  if (rowDelimiter) {
    const columnGroups = boundary === 'column';
    const groupLimit = columnGroups ? LIMITS.columns : LIMITS.rows;
    const sourceGroups = input.split(rowDelimiter, groupLimit + 1);
    if (sourceGroups.length > groupLimit) {
      throw new RangeError(columnGroups ? '原文超过 200 列，请分批处理。' : '原文超过 2000 行，请分批处理。');
    }
    const sourceCells = [];
    let count = 0;
    let widest = 1;
    for (const source of sourceGroups) {
      const cells = splitCells(source);
      count += cells.length;
      if (count > LIMITS.cells) throw new RangeError('拆分结果超过 50000 个单元格，请分批处理。');
      widest = Math.max(widest, cells.length);
      sourceCells.push(cells);
    }
    const requiredRows = columnGroups ? widest : sourceGroups.length;
    const requiredColumns = columnGroups ? sourceGroups.length : widest;
    if (rows !== null && rows < requiredRows) {
      throw new RangeError(`原文需要 ${requiredRows} 行，指定行数不足；请增加行数或清空行数设置。`);
    }
    if (columns !== null && columns < requiredColumns) {
      throw new RangeError(`原文需要 ${requiredColumns} 列，指定列数不足；请增加列数或清空列数设置。`);
    }
    rows ??= requiredRows;
    columns ??= requiredColumns;
    validateSize(rows, columns);
    tokens = sourceCells.flat();
    validateTokens(tokens);
    matrix = Array.from({ length: rows }, () => Array(columns).fill(''));
    sourceCells.forEach((cells, group) => cells.forEach((cell, offset) => {
      const row = columnGroups ? offset : group;
      const column = columnGroups ? group : offset;
      matrix[row][column] = cell;
    }));
    if (columnGroups) {
      warnings.push('已按换列分隔符保留原有列，每组内容向下排列，本次不使用填充方向设置。');
    } else if (fill === 'column') {
      warnings.push('已按行分隔符保留原有行，本次不使用按列填充。');
    }
  } else {
    tokens = splitCells(input);
    validateTokens(tokens);
    if (columns === null && rows === null) columns = 3;
    if (rows === null) rows = Math.max(1, Math.ceil(tokens.length / columns));
    if (columns === null) columns = Math.max(1, Math.ceil(tokens.length / rows));
    validateSize(rows, columns);
    if (rows * columns < tokens.length) {
      throw new RangeError(`共有 ${tokens.length} 项，但 ${rows} 行 × ${columns} 列只能放 ${rows * columns} 项；请增加行列数。`);
    }
    matrix = Array.from({ length: rows }, () => Array(columns).fill(''));
    tokens.forEach((cell, index) => {
      const row = fill === 'row' ? Math.floor(index / columns) : index % rows;
      const column = fill === 'row' ? index % columns : Math.floor(index / rows);
      matrix[row][column] = cell;
    });
  }

  if (removed) warnings.push(`已去掉 ${removed} 个空项；仅含空格的文字仍会保留。`);
  const padded = rows * columns - tokens.length;
  return { matrix, tokens, warnings, count: tokens.length, rows, columns, padded };
}
