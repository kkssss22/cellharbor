/** CellHarbor: preview-only, dependency-free operations on spreadsheet values. */

const MAX_CELLS = 100_000;
const MAX_TEXT = 32_767;
const MAX_COLUMNS = 16_384;
const DECIMAL = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/;

export const ACTIONS = Object.freeze([
  { id: 'trim', title: '去除首尾空格', category: '文本整理', description: '去掉文本前后的空白，保留中间内容。', kind: 'transform' },
  { id: 'removeSpaces', title: '删除所有空白', category: '文本整理', description: '移除空格、制表符和换行。', kind: 'transform' },
  { id: 'normalizeSpaces', title: '合并多余空格', category: '文本整理', description: '把连续空白合成一个空格，并清理首尾。', kind: 'transform' },
  { id: 'removeInvisible', title: '清理零宽字符', category: '文本整理', description: '移除零宽空格、连接符和隐藏的字节序标记。', kind: 'transform' },
  { id: 'upper', title: '转为大写', category: '文本整理', description: '把文字转为大写，数字保持原样。', kind: 'transform' },
  { id: 'lower', title: '转为小写', category: '文本整理', description: '把文字转为小写，数字保持原样。', kind: 'transform' },
  { id: 'titleCase', title: '单词首字母大写', category: '文本整理', description: '每个单词首字母大写，其余字母小写。', kind: 'transform' },
  { id: 'fullWidthToHalf', title: '全角转半角', category: '文本整理', description: '转换全角英文字母、数字、标点和全角空格。', kind: 'transform' },
  { id: 'findReplace', title: '查找并替换', category: '文本整理', description: '按字面内容替换，区分大小写，不使用正则表达式。', kind: 'transform', needs: ['find', 'replace'] },
  { id: 'numberToText', title: '数字转文本', category: '数字处理', description: '把数值转换为文本；不能找回之前已丢失的数字精度。', kind: 'transform' },
  { id: 'textToNumber', title: '数字文本转数值', category: '数字处理', description: '仅转换明确的十进制数字，保留前导零编号和长数字文本。', kind: 'transform' },
  { id: 'percentToNumber', title: '百分比文本转数值', category: '数字处理', description: '例如把“12.5%”转成0.125，不更改单元格显示格式。', kind: 'transform' },
  { id: 'round', title: '四舍五入', category: '数字处理', description: '改变实际数值；负小数位可舍入到十位、百位。', kind: 'transform', needs: ['digits'] },
  { id: 'absolute', title: '转为绝对值', category: '数字处理', description: '把负数转为正数，文本保持原样。', kind: 'transform' },
  { id: 'negate', title: '正负号反转', category: '数字处理', description: '正数变负数，负数变正数。', kind: 'transform' },
  { id: 'fillDown', title: '向下填充空白', category: '表格整理', description: '用上方的普通值填充空白；表头和公式不会被覆盖。', kind: 'transform' },
  { id: 'blankToZero', title: '空白填零', category: '表格整理', description: '把空单元格和只有空白的文本改为数字0。', kind: 'transform' },
  { id: 'duplicates', title: '检查重复行', category: '数据检查', description: '比较整行值，区分数字与文本，报告所有重复行。', kind: 'inspect' },
  { id: 'blanks', title: '检查空白单元格', category: '数据检查', description: '找出空值、空字符串和只有空白的文本。', kind: 'inspect' },
  { id: 'numbersAsText', title: '检查数字文本', category: '数据检查', description: '找出看起来像数字的文本，并提示编号风险。', kind: 'inspect' },
  { id: 'longNumbers', title: '检查长数字风险', category: '数据检查', description: '提示超过15位有效数字的内容；不尝试恢复已丢失精度。', kind: 'inspect' },
  { id: 'inconsistentTypes', title: '检查列类型混用', category: '数据检查', description: '找出同列中的少数数据类型；并列时报告所有类型。', kind: 'inspect' },
  { id: 'stats', title: '生成数字摘要', category: '提取结果', description: '按列汇总数值个数、合计、均值、最小值、最大值和空白数。', kind: 'extract' },
  { id: 'uniqueRows', title: '提取不重复行', category: '提取结果', description: '保留每种行内容的第一次出现，输出一份新的值副本。', kind: 'extract' },
  { id: 'splitColumn', title: '按分隔符拆列', category: '提取结果', description: '选择一列，按指定字面分隔符拆开；不会覆盖原列。', kind: 'extract', needs: ['delimiter'] },
  { id: 'transpose', title: '行列互换副本', category: '提取结果', description: '把整个所选区域的行与列互换，包含表头。', kind: 'extract' },
  { id: 'removeEmptyRows', title: '提取非空行', category: '提取结果', description: '输出非空行的值副本，保留原来的行顺序。', kind: 'extract' },
  { id: 'addSequence', title: '生成序号列', category: '提取结果', description: '按数据行数生成1、2、3…，不改动现有单元格。', kind: 'extract' },
].map((action) => Object.freeze({ ...action, ...(action.needs ? { needs: Object.freeze(action.needs) } : {}) })));

const ACTION_MAP = new Map(ACTIONS.map((action) => [action.id, action]));

function validateMatrix(values, label = '数据') {
  if (!Array.isArray(values) || !values.length || !Array.isArray(values[0]) || !values[0].length) {
    throw new Error(`${label}必须是至少包含一个单元格的二维数组。`);
  }
  const cols = values[0].length;
  if (cols > MAX_COLUMNS) throw new Error(`${label}超过Excel的16,384列限制。`);
  if (values.length * cols > MAX_CELLS) throw new Error(`一次最多处理100,000个单元格。`);
  for (const row of values) {
    if (!Array.isArray(row) || row.length !== cols) throw new Error(`${label}每一行的列数必须相同。`);
    for (const value of row) {
      if (value === null || typeof value === 'boolean') continue;
      if (typeof value === 'number' && Number.isFinite(value)) continue;
      if (typeof value === 'string' && value.length <= MAX_TEXT) continue;
      throw new Error(`${label}只支持有限数值、布尔值、null和不超过32,767字符的文本。`);
    }
  }
  return { rows: values.length, cols };
}

const blank = (value) => value === null || (typeof value === 'string' && value.trim() === '');
const typeOf = (value) => blank(value) ? '空白' : typeof value === 'number' ? '数字' : typeof value === 'boolean' ? '逻辑值' : '文本';
const numeric = (value) => typeof value === 'number';
const isFormula = (formulas, r, c) => typeof formulas[r]?.[c] === 'string' && formulas[r][c].startsWith('=');
const noNegativeZero = (value) => Object.is(value, -0) ? 0 : value;

/** Exact type-sensitive row comparison; null and the empty string stay distinct. */
const rowKey = (row) => JSON.stringify(row);

function numericText(value, { allowLeadingZeros = false } = {}) {
  if (typeof value !== 'string') return null;
  const text = value.trim();
  if (!DECIMAL.test(text)) return null;
  const mantissa = text.replace(/^[+-]/, '').split(/[eE]/)[0];
  const integral = mantissa.split('.')[0];
  const significant = mantissa.replace('.', '').replace(/^0+/, '');
  const number = Number(text);
  let reason = null;
  if (!allowLeadingZeros && /^0\d/.test(integral)) reason = '包含前导零，可能是编号，已保留文本';
  else if (significant.length > 15 || (Number.isInteger(number) && Math.abs(number) >= 1e15)) reason = '超过15位数字精度范围，已保留文本';
  else if (!Number.isFinite(number)) reason = '数值超出可表示范围，已保留文本';
  else if (number === 0 && /[1-9]/.test(mantissa)) reason = '数值过小会丢失为零，已保留文本';
  return { value: noNegativeZero(number), reason };
}

function decimalShift(value, places) {
  const [mantissa, exponent = '0'] = String(value).split('e');
  return Number(`${mantissa}e${Number(exponent) + places}`);
}

function roundAwayFromZero(value, digits) {
  const shifted = decimalShift(Math.abs(value), digits);
  if (!Number.isFinite(shifted)) return value;
  const rounded = decimalShift(Math.round(shifted), -digits);
  return Number.isFinite(rounded) ? noNegativeZero(Math.sign(value) * rounded) : value;
}

function replaceLiteral(text, find, replacement) {
  let count = 0;
  for (let offset = 0; (offset = text.indexOf(find, offset)) !== -1; offset += find.length) count += 1;
  if (text.length + count * (replacement.length - find.length) > MAX_TEXT) {
    throw new Error('替换结果超过Excel单元格的32,767字符限制，请缩短替换内容。');
  }
  return text.split(find).join(replacement);
}

function columnName(index) {
  let text = '';
  for (let value = index + 1; value > 0; value = Math.floor((value - 1) / 26)) {
    text = String.fromCharCode(65 + ((value - 1) % 26)) + text;
  }
  return text;
}

function stableMean(numbers) {
  let sum = 0;
  let correction = 0;
  for (const number of numbers) {
    const adjusted = number / numbers.length - correction;
    const next = sum + adjusted;
    correction = (next - sum) - adjusted;
    sum = next;
  }
  return sum;
}

/**
 * Never mutates the source. Patch and finding coordinates are zero-based within
 * the selected range. hasHeader skips row 0 in transformations and inspections.
 */
export function preview(actionId, values, options = {}) {
  const action = ACTION_MAP.get(actionId);
  if (!action) throw new Error(`未知操作：${String(actionId)}。`);
  if (!options || typeof options !== 'object' || Array.isArray(options)) throw new Error('选项必须是对象。');
  const { formulas = [], hasHeader = true, find = '', replace = '', delimiter = ',', digits = 2 } = options;
  const { rows, cols } = validateMatrix(values);
  if (typeof hasHeader !== 'boolean') throw new Error('“首行是表头”必须为布尔值。');
  if (!Array.isArray(formulas)) throw new Error('公式必须是二维数组。');
  if (formulas.length) {
    const shape = validateMatrix(formulas, '公式');
    if (shape.rows !== rows || shape.cols !== cols) throw new Error('公式区域与数据区域的尺寸必须一致。');
  }
  if (actionId === 'round' && (!Number.isInteger(digits) || digits < -15 || digits > 15)) {
    throw new Error('小数位数必须是−15到15之间的整数。');
  }
  if (actionId === 'findReplace' && (typeof find !== 'string' || !find || find.length > MAX_TEXT || typeof replace !== 'string' || replace.length > MAX_TEXT)) {
    throw new Error('查找内容不能为空，查找和替换内容必须是不超过32,767字符的文本。');
  }
  if (actionId === 'splitColumn' && (cols !== 1 || typeof delimiter !== 'string' || !delimiter || delimiter.length > 100)) {
    throw new Error('拆列需要只选择一列，并提供1–100字符的分隔符。');
  }
  const first = hasHeader ? 1 : 0;
  const patches = [];
  const findings = [];
  const warnings = [];
  let output = null;
  let skippedFormulas = 0;
  let preservedNumericText = 0;
  const addFinding = (r, c, reason) => findings.push({ r, c, reason });
  const addPatch = (r, c, after) => {
    if (typeof after === 'string' && after.length > MAX_TEXT) throw new Error('处理结果超过Excel单元格的32,767字符限制。');
    if (typeof after === 'number' && !Number.isFinite(after)) throw new Error('计算结果超出可表示的数值范围。');
    if (after !== values[r][c]) patches.push({ r, c, before: values[r][c], after });
  };

  if (action.kind === 'transform') {
    if (actionId === 'fillDown') {
      for (let c = 0; c < cols; c += 1) {
        let carry;
        for (let r = first; r < rows; r += 1) {
          if (isFormula(formulas, r, c)) { skippedFormulas += 1; carry = undefined; continue; }
          const value = values[r][c];
          if (blank(value)) { if (carry !== undefined) addPatch(r, c, carry); }
          else carry = value;
        }
      }
      if (skippedFormulas) warnings.push('向下填充在公式处中断继承，避免把公式的当前结果当作固定值继续填充。');
    } else {
      for (let r = first; r < rows; r += 1) {
        for (let c = 0; c < cols; c += 1) {
          if (isFormula(formulas, r, c)) { skippedFormulas += 1; continue; }
          const before = values[r][c];
          let after = before;
          const string = typeof before === 'string';
          switch (actionId) {
            case 'trim': if (string) after = before.trim(); break;
            case 'removeSpaces': if (string) after = before.replace(/\s/g, ''); break;
            case 'normalizeSpaces': if (string) after = before.replace(/\s+/g, ' ').trim(); break;
            case 'removeInvisible': if (string) after = before.replace(/[\u200b-\u200d\u2060\ufeff]/g, ''); break;
            case 'upper': if (string) after = before.toUpperCase(); break;
            case 'lower': if (string) after = before.toLowerCase(); break;
            case 'titleCase': if (string) after = before.toLowerCase().replace(/\p{L}[\p{L}\p{M}]*/gu, (word) => { const chars = Array.from(word); return chars[0].toUpperCase() + chars.slice(1).join(''); }); break;
            case 'fullWidthToHalf': if (string) after = before.replace(/[\uff01-\uff5e]/g, (ch) => String.fromCharCode(ch.charCodeAt(0) - 0xfee0)).replace(/\u3000/g, ' '); break;
            case 'findReplace': if (string) after = replaceLiteral(before, find, replace); break;
            case 'numberToText': if (numeric(before)) after = String(before); break;
            case 'textToNumber': {
              const candidate = numericText(before);
              if (candidate?.reason) { preservedNumericText += 1; addFinding(r, c, candidate.reason); }
              else if (candidate) after = candidate.value;
              break;
            }
            case 'percentToNumber': {
              if (string && before.trim().endsWith('%')) {
                const candidate = numericText(before.trim().slice(0, -1), { allowLeadingZeros: true });
                if (candidate?.reason) { preservedNumericText += 1; addFinding(r, c, candidate.reason); }
                else if (candidate) {
                  const fraction = candidate.value / 100;
                  if (candidate.value !== 0 && fraction === 0) { preservedNumericText += 1; addFinding(r, c, '百分比转换会下溢为零，已保留文本'); }
                  else after = fraction;
                }
              }
              break;
            }
            case 'round': if (numeric(before)) after = roundAwayFromZero(before, digits); break;
            case 'absolute': if (numeric(before)) after = Math.abs(before); break;
            case 'negate': if (numeric(before)) after = noNegativeZero(-before); break;
            case 'blankToZero': if (blank(before)) after = 0; break;
            default: throw new Error('操作尚未实现。');
          }
          addPatch(r, c, after);
        }
      }
    }
    if (skippedFormulas) warnings.push(`已跳过${skippedFormulas}个公式单元格；公式不会被改成固定值。`);
    if (preservedNumericText) warnings.push(`为保护编号或数值精度，保留了${preservedNumericText}个数字文本。`);
    if (actionId === 'numberToText') warnings.push('数字转文本不能恢复Excel此前已丢失的前导零或第15位之后的精度。');
  } else if (action.kind === 'inspect') {
    if (actionId === 'duplicates') {
      const groups = new Map();
      for (let r = first; r < rows; r += 1) {
        const key = rowKey(values[r]);
        if (!groups.has(key)) groups.set(key, []);
        groups.get(key).push(r);
      }
      for (const group of groups.values()) {
        if (group.length < 2) continue;
        for (const r of group) for (let c = 0; c < cols; c += 1) addFinding(r, c, `重复行：与所选区域第${(r === group[0] ? group[1] : group[0]) + 1}行相同（区分数据类型）`);
      }
    } else if (actionId === 'inconsistentTypes') {
      for (let c = 0; c < cols; c += 1) {
        const counts = new Map();
        for (let r = first; r < rows; r += 1) {
          const type = typeOf(values[r][c]);
          if (type !== '空白') counts.set(type, (counts.get(type) ?? 0) + 1);
        }
        if (counts.size < 2) continue;
        const largest = Math.max(...counts.values());
        const leaders = [...counts].filter(([, count]) => count === largest).map(([type]) => type);
        for (let r = first; r < rows; r += 1) {
          const type = typeOf(values[r][c]);
          if (type !== '空白' && (leaders.length > 1 || type !== leaders[0])) addFinding(r, c, `该列混用${[...counts.keys()].join('、')}；当前为${type}`);
        }
      }
    } else {
      for (let r = first; r < rows; r += 1) for (let c = 0; c < cols; c += 1) {
        const value = values[r][c];
        if (actionId === 'blanks' && blank(value)) addFinding(r, c, isFormula(formulas, r, c) ? '公式当前结果显示为空白' : '空白单元格或只包含空白的文本');
        if (actionId === 'numbersAsText') {
          const candidate = numericText(value);
          if (candidate) addFinding(r, c, candidate.reason ? `数字文本：${candidate.reason}` : '以文本存储的数字');
        }
        if (actionId === 'longNumbers') {
          const candidate = numericText(typeof value === 'number' ? String(value) : value, { allowLeadingZeros: true });
          if (candidate?.reason?.includes('15位')) addFinding(r, c, typeof value === 'number' ? '长数值可能已被Excel舍入；转换成文本也无法恢复丢失的数字' : '长数字文本应保留为文本，避免Excel的15位精度限制');
        }
      }
    }
  } else {
    warnings.push('提取结果只复制当前值，不复制公式、格式或批注；原区域保持不变。');
    if (actionId === 'transpose') {
      if (rows > MAX_COLUMNS) throw new Error('输出结果超过Excel的16,384列限制。');
      output = Array.from({ length: cols }, (_, c) => values.map((row) => row[c]));
    }
    if (actionId === 'addSequence') output = [...(hasHeader ? [['序号']] : []), ...Array.from({ length: rows - first }, (_, index) => [index + 1])];
    if (actionId === 'uniqueRows') {
      const seen = new Set();
      output = hasHeader ? [[...values[0]]] : [];
      for (let r = first; r < rows; r += 1) {
        const key = rowKey(values[r]);
        if (!seen.has(key)) { output.push([...values[r]]); seen.add(key); }
      }
    }
    if (actionId === 'removeEmptyRows') output = [...(hasHeader ? [[...values[0]]] : []), ...values.slice(first).filter((row) => !row.every(blank)).map((row) => [...row])];
    if (actionId === 'splitColumn') {
      const parts = [];
      let width = 1;
      // Bound the rectangular output before allocating padding for every row.
      for (let r = first; r < rows; r += 1) {
        const row = (values[r][0] === null ? '' : String(values[r][0])).split(delimiter);
        width = Math.max(width, row.length);
        if (width > MAX_COLUMNS) throw new Error('拆列结果超过Excel的16,384列限制。');
        if (rows * width > MAX_CELLS) throw new Error('拆列结果超过100,000个单元格，请缩小所选区域。');
        parts.push(row);
      }
      const label = hasHeader && !blank(values[0][0]) ? String(values[0][0]) : '字段';
      output = [...(hasHeader ? [Array.from({ length: width }, (_, c) => width === 1 ? label : `${label} ${c + 1}`)] : []), ...parts.map((row) => [...row, ...Array(width - row.length).fill('')])];
    }
    if (actionId === 'stats') {
      output = [['列', '数值个数', '合计', '平均值', '最小值', '最大值', '空白数']];
      for (let c = 0; c < cols; c += 1) {
        const numbers = [];
        let missing = 0;
        for (let r = first; r < rows; r += 1) {
          if (numeric(values[r][c])) numbers.push(values[r][c]);
          else if (blank(values[r][c])) missing += 1;
        }
        const mean = numbers.length ? stableMean(numbers) : null;
        const total = numbers.length ? mean * numbers.length : 0;
        let min = null;
        let max = null;
        for (const number of numbers) { min = min === null ? number : Math.min(min, number); max = max === null ? number : Math.max(max, number); }
        if ((mean !== null && !Number.isFinite(mean)) || !Number.isFinite(total)) warnings.push(`第${columnName(c)}列的合计超出数值范围，无法表示的结果显示为空白。`);
        output.push([hasHeader && !blank(values[0][c]) ? String(values[0][c]) : `列 ${columnName(c)}`, numbers.length, Number.isFinite(total) ? total : null, mean !== null && Number.isFinite(mean) ? mean : null, min, max, missing]);
      }
      warnings.push('摘要仅统计真实数值，数字文本和逻辑值不计入数值个数；空白数包含只含空格的文本。');
    }
    if (output.length) validateMatrix(output, '输出结果');
  }
  const potentialFormulaText = patches.some((patch) => typeof patch.after === 'string' && /^[=+\-@]/.test(patch.after)) || output?.some((row) => row.some((value) => typeof value === 'string' && /^[=+\-@]/.test(value)));
  if (potentialFormulaText) warnings.push('结果含以=、+、-或@开头的文本，写入时必须作为文本保留，不能执行为公式。');
  return {
    patches,
    summary: { actionId, rows, columns: cols, dataRows: rows - first, changedCells: patches.length, skippedFormulas, findingCount: findings.length, outputRows: output?.length ?? 0, outputColumns: output?.[0]?.length ?? 0 },
    warnings,
    output,
    findings,
  };
}
