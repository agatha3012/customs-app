const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

// ==================== 工具函数 ====================

function dateToSerial(d) {
  const base = new Date(1899, 11, 30);
  const diff = d.getTime() - base.getTime();
  return Math.round(diff / (24 * 60 * 60 * 1000));
}

function r4(v) { return parseFloat((v || 0).toFixed(4)); }

/** 数字转中文大写（如 1131.12 → 壹仟壹佰叁拾壹圆壹角贰分） */
function toChineseRMB(amount) {
  const digits = '零壹贰叁肆伍陆柒捌玖';
  const units = ['', '拾', '佰', '仟'];
  const bigUnits = ['', '万', '亿'];
  var n = Math.round((amount || 0) * 100); // 转为分
  if (n === 0) return '零圆整';
  var jiao = Math.floor((n % 100) / 10);
  var fen = n % 10;
  var yuan = Math.floor(n / 100);
  // 整数部分
  function convertInt(num) {
    if (num === 0) return '零';
    var s = '';
    var unitPos = 0;
    var bigUnitPos = 0;
    var needZero = false;
    while (num > 0) {
      var seg = num % 10000;
      if (seg > 0) {
        var segStr = '';
        var segNum = seg;
        for (var i = 0; i < 4 && segNum > 0; i++) {
          var d = segNum % 10;
          segNum = Math.floor(segNum / 10);
          if (d > 0) {
            segStr = digits[d] + units[i] + segStr;
            needZero = false;
          } else if (segStr !== '') {
            segStr = '零' + segStr;
          }
        }
        s = segStr + bigUnits[bigUnitPos] + s;
        needZero = false;
      } else if (s !== '') {
        needZero = true;
      }
      num = Math.floor(num / 10000);
      bigUnitPos++;
    }
    return s.replace(/零+$/, '');
  }
  var yuanStr = convertInt(yuan);
  // 角分
  var jfStr = '';
  if (jiao > 0) jfStr += digits[jiao] + '角';
  if (fen > 0) jfStr += digits[fen] + '分';
  var result = yuanStr + '圆';
  if (jfStr) result += jfStr;
  else result += '整';
  return result;
}


// ==================== 核心原则：只改 cell.value，不改任何样式 ====================

/** 写任意值 — 原封不动保留单元格的字体/填充/边框/对齐/数字格式 */
function putVal(ws, row, col, value) {
  ws.getCell(row, col).value = value;
}

/** 写数值 (内部4位精度，显示2位小数) */
function putNum(ws, row, col, value) {
  const cell = ws.getCell(row, col);
  cell.value = r4(value);
  cell.numFmt = '0.00';
}

/** 写日期序列号 */
function putDate(ws, row, col, serial) {
  ws.getCell(row, col).value = serial;
}

/** 清空一行中所有单元格的值 (保留格式) */
function clearRowVals(ws, r) {
  ws.getRow(r).eachCell({ includeEmpty: false }, cell => { cell.value = null; });
}

/** 清空 [startR, endR) 行范围内所有单元格的值 */
function clearRowsVals(ws, startR, endR) {
  for (let r = startR; r < endR; r++) clearRowVals(ws, r);
}


// ==================== 统一样式 ====================

/** 统一样式常量 */
const UNIFIED = {
  font: { name: '宋体', size: 10, color: { argb: 'FF000000' } },
  fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } },
  border: {
    top:    { style: 'thin', color: { argb: 'FF000000' } },
    bottom: { style: 'thin', color: { argb: 'FF000000' } },
    left:   { style: 'thin', color: { argb: 'FF000000' } },
    right:  { style: 'thin', color: { argb: 'FF000000' } },
  },
  alignment: { horizontal: 'center', vertical: 'middle', shrinkToFit: true, wrapText: false },
};

/**
 * 统一数据区域单元格的字体/填充/边框/对齐。
 * @param {Worksheet} ws
 * @param {number} startRow - 数据起始行
 * @param {number} endRow   - 数据结束行（含）
 * @param {number} endCol   - 最右列号
 */
function unifyDataStyles(ws, startRow, endRow, endCol) {
  for (let r = startRow; r <= endRow; r++) {
    for (let c = 1; c <= endCol; c++) {
      const cell = ws.getCell(r, c);
      cell.font      = { ...UNIFIED.font };
      cell.fill      = { ...UNIFIED.fill };
      cell.border    = { ...UNIFIED.border };
      cell.alignment = { ...UNIFIED.alignment };
    }
  }
}


// ==================== 合并单元格工具 ====================
// ExcelJS 4.x 的 model.merges 是字符串数组 "A1:Y1"

/** 列字母 → 索引 (A=1, B=2, ... Z=26, AA=27) */
function colToIndex(col) {
  let n = 0;
  for (let i = 0; i < col.length; i++) n = n * 26 + (col.charCodeAt(i) - 64);
  return n;
}

/** 列索引 → 字母 */
function indexToCol(idx) {
  let s = '';
  while (idx > 0) {
    const m = (idx - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    idx = Math.floor((idx - 1) / 26);
  }
  return s;
}

/** 解析合并字符串 "A1:Y1" → {top,left,bottom,right}，按规范左<右、上<下 */
function parseMerge(str) {
  const m = str.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
  if (!m) return null;
  const left = colToIndex(m[1]), right = colToIndex(m[3]);
  return {
    top: parseInt(m[2]),
    left: Math.min(left, right),
    bottom: parseInt(m[4]),
    right: Math.max(left, right),
  };
}

/** 合并对象 → 字符串 */
function fmtMerge(top, left, bottom, right) {
  return indexToCol(left) + top + ':' + indexToCol(right) + bottom;
}

/** 合并是否水平 (同一行内) */
function isHorizontalMerge(p) { return p.top === p.bottom && p.left !== p.right; }

/** 合并是否垂直 */
function isVerticalMerge(p) { return p.left === p.right && p.top !== p.bottom; }


// ==================== 行数调整 ====================

/**
 * 确保数据区域有恰好 needRows 行可用。
 * - needRows > tplRows: 用 duplicateRow 复制最后一数据行
 * - needRows < tplRows: 清空多余行 + 移除多余合并
 * 返回汇总行位置 (1-indexed)。
 */
function ensureDataRows(ws, dataStart, tplRows, needRows) {
  if (needRows > tplRows) {
    const lastDataRow = dataStart + tplRows - 1;
    ws.duplicateRow(lastDataRow, needRows - tplRows, true);
  }
  return needRows > tplRows
    ? dataStart + needRows
    : dataStart + tplRows;
}

/**
 * 清理多余行：清除值 + 移除多余行范围内的合并单元格。
 */
function cleanExcess(ws, dataStart, needRows, tplRows) {
  if (needRows >= tplRows) return;
  const excessFirst = dataStart + needRows;
  const excessLast  = dataStart + tplRows - 1;

  // ① 清空值
  clearRowsVals(ws, excessFirst, excessLast + 1);

  // ② 移除完全落在此范围的合并
  ws.model.merges = ws.model.merges.filter(m => {
    const p = parseMerge(m);
    if (!p) return true;           // 无法解析的保留
    if (p.bottom < excessFirst) return true;   // 在上方
    if (p.top > excessLast) return true;        // 在下方（汇总行等）
    return false;                               // 与多余行有交集 → 移除
  });
}

/**
 * 重建数据行内的水平合并模式。
 * 从 refPatterns 数组（左-右列对）为每一数据行创建水平合并。
 * summaryRow: 如有传入，也为汇总行重建。
 */
function rebuildRowMerges(ws, dataStart, numRows, refPatterns, summaryRow) {
  if (numRows <= 0 || !refPatterns || refPatterns.length === 0) return;

  const cleanUpTo = summaryRow ? summaryRow : (dataStart + numRows - 1);

  // ★ 先用 API 解除所有数据区内的合并（清理旧合并 + duplicateRow 产生的坏合并）
  for (let r = dataStart; r <= cleanUpTo; r++) {
    refPatterns.forEach(ref => {
      try { ws.unMergeCells(r, ref.left, r, ref.right); } catch (_) {}
    });
  }

  // ★ 再用 API 重建合并
  for (let r = 0; r < numRows; r++) {
    const row = dataStart + r;
    refPatterns.forEach(ref => {
      try { ws.mergeCellsWithoutStyle(row, ref.left, row, ref.right); } catch (_) {}
    });
  }

  if (summaryRow) {
    refPatterns.forEach(ref => {
      try { ws.mergeCellsWithoutStyle(summaryRow, ref.left, summaryRow, ref.right); } catch (_) {}
    });
  }
}

/**
 * 保存模板第一数据行的水平合并模式，之后 restore。
 * duplicateRow 可能破坏合并模型，所以必须在任何修改之前采集。
 */
function collectMergePatterns(ws, dataStart) {
  const patterns = [];
  ws.model.merges.forEach(m => {
    const p = parseMerge(m);
    if (!p) return;
    if (p.top === dataStart && p.bottom === dataStart && p.left !== p.right) {
      patterns.push({ left: p.left, right: p.right });
    }
  });
  return patterns;
}

/**
 * 报关草单：按箱重建 Z(26) / AA(27) 列垂直合并
 */
function rebuildBoxMerges(ws, startRow, boxes) {
  // 移除数据区 Z/AA 列的旧垂直合并
  ws.model.merges = ws.model.merges.filter(m => {
    const p = parseMerge(m);
    if (!p) return true;
    if (p.top < startRow) return true;              // 数据区上方
    if ((p.left === 26 || p.left === 27) && isVerticalMerge(p)) return false;
    return true;
  });

  // 按箱重建
  let r = startRow;
  boxes.forEach(box => {
    const span = box.products.length;
    if (span > 1) {
      ws.model.merges.push(fmtMerge(r, 26, r + span - 1, 26));
      ws.model.merges.push(fmtMerge(r, 27, r + span - 1, 27));
    }
    r += span;
  });
}


// ==================== 主入口 ====================

async function generate(params) {
  const { products, boxes, destination, exchangeRate, constant, templatePath, outputPath, invoiceDate } = params;

  if (!fs.existsSync(templatePath)) throw new Error('模板文件不存在: ' + templatePath);
  fs.copyFileSync(templatePath, outputPath);

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(outputPath);

  const today = invoiceDate || new Date();
  const serialDate = dateToSerial(today);
  const dateStr = today.getFullYear().toString() +
    ('0' + (today.getMonth() + 1)).slice(-2) +
    ('0' + today.getDate()).slice(-2);
  const contractNo = 'YLW-' + dateStr + '001';

  const totalQty   = products.reduce((s, p) => s + p.quantity, 0);
  const totalBoxes = boxes.length;
  const totalGross = boxes.reduce((s, b) => s + (b.weight || 0), 0);
  const totalNet   = products.reduce((s, p) => s + (p.netWeight || 0), 0);
  const totalAmt   = products.reduce((s, p) => s + (p.unitPriceUSD * p.quantity), 0);
  const boxRows    = boxes.reduce((s, b) => s + b.products.length, 0);

  // ★ 在修改任何工作表之前，采集所有工作表的第一数据行的水平合并模式
  // duplicateRow 会破坏合并模型，必须先用模板原始状态保存参考
  const patternsDraft   = collectMergePatterns(wb.getWorksheet('报关草单'), 13);
  const patternsCont    = collectMergePatterns(wb.getWorksheet('合同'), 21);
  const patternsInv     = collectMergePatterns(wb.getWorksheet('发票'), 15);
  const patternsPack    = collectMergePatterns(wb.getWorksheet('装箱单'), 12);
  const patternsMerged  = collectMergePatterns(wb.getWorksheet('报关草单合并'), 13);

  fillCustomsDraft(wb.getWorksheet('报关草单'), products, boxes, destination, contractNo, serialDate, totalQty, totalBoxes, totalGross, totalNet, totalAmt, boxRows, patternsDraft);
  fillContract(wb.getWorksheet('合同'), products, contractNo, serialDate, totalQty, totalAmt, patternsCont);
  fillInvoice(wb.getWorksheet('发票'), products, contractNo, serialDate, totalQty, totalAmt, patternsInv);
  fillPackingList(wb.getWorksheet('装箱单'), products, boxes, contractNo, serialDate, totalBoxes, boxRows, patternsPack);
  fillCustomsMerged(wb.getWorksheet('报关草单合并'), products, destination, contractNo, serialDate, totalQty, totalBoxes, totalGross, totalNet, totalAmt, patternsMerged);

  await wb.xlsx.writeFile(outputPath);
  console.log('Generated:', outputPath);
  return outputPath;
}


// ==================== 报关草单 ====================
// 模板: 87 rows, data R13-R83 (71), summary R84
function fillCustomsDraft(ws, products, boxes, destination, contractNo, serialDate,
  totalQty, totalBoxes, totalGross, totalNet, totalAmt, boxRows, refPatterns) {

  const R = 13;
  const TPL = 71;

  // 元数据
  putVal(ws, 6, 3, contractNo);
  putVal(ws, 6, 14, destination);
  putDate(ws, 3, 14, serialDate);
  putDate(ws, 3, 17, serialDate);
  putVal(ws, 7, 7, totalBoxes + '件');
  putVal(ws, 7, 10, r4(totalGross).toFixed(2) + '千克');
  putVal(ws, 7, 12, r4(totalNet).toFixed(2) + '千克');

  const SR = ensureDataRows(ws, R, TPL, boxRows);

  // 数据
  let ri = R, sq = 1;
  boxes.forEach(box => {
    box.products.forEach((p, idx) => {
      const fi = idx === 0;
      const q  = p.quantityPerBox || p.quantity;
      const nw = p.netWeightPerBox || 0;

      putNum(ws, ri, 1, sq++);
      putNum(ws, ri, 2, p.hsCode || 0);
      putVal(ws, ri, 4, p.sku);
      putVal(ws, ri, 5, p.nameCN);
      putVal(ws, ri, 7, p.spec || p.nameCN);
      putNum(ws, ri, 11, q);
      putVal(ws, ri, 12, p.unit || '个');
      putNum(ws, ri, 13, nw);
      putNum(ws, ri, 14, p.unitPriceUSD);
      putNum(ws, ri, 15, p.unitPriceUSD * q);
      putVal(ws, ri, 16, 'USD');
      putVal(ws, ri, 17, p.originCountry || '中国');
      putVal(ws, ri, 19, destination);
      putVal(ws, ri, 22, p.domesticSource || '');
      putVal(ws, ri, 25, '照章征税');
      if (fi) { putNum(ws, ri, 26, box.weight || 0); putNum(ws, ri, 27, box.boxSeq); }

      ri++;
    });
  });

  // 汇总
  putNum(ws, SR, 11, totalQty);
  putNum(ws, SR, 13, totalNet);
  putVal(ws, SR, 14, '金额');
  putNum(ws, SR, 15, totalAmt);
  putVal(ws, SR, 16, 'USD');
  putNum(ws, SR, 26, totalGross);
  putNum(ws, SR, 27, totalBoxes);

  cleanExcess(ws, R, boxRows, TPL);
  rebuildRowMerges(ws, R, boxRows, refPatterns, SR);
  rebuildBoxMerges(ws, R, boxes);
  unifyDataStyles(ws, R, SR, 27);   // 统一颜色：数据行 + 汇总行
}


// ==================== 合同 ====================
// 模板: 56 rows, data R21-R28 (8), summary R29
function fillContract(ws, products, contractNo, serialDate, totalQty, totalAmt, refPatterns) {
  const R = 21, TPL = 8;
  putVal(ws, 5, 23, contractNo);
  putDate(ws, 7, 23, serialDate);
  const SR = ensureDataRows(ws, R, TPL, products.length);

  products.forEach((p, i) => {
    const r = R + i;
    putVal(ws, r, 1, p.spec || p.nameCN);
    putNum(ws, r, 13, p.quantity);
    putVal(ws, r, 16, p.unit || '个');
    putNum(ws, r, 18, p.unitPriceUSD);
    putNum(ws, r, 23, p.unitPriceUSD * p.quantity);
  });

  putNum(ws, SR, 13, totalQty);
  putNum(ws, SR, 23, totalAmt);

  // 总值（数字）— SR+2 行，col 23
  putNum(ws, SR + 2, 23, totalAmt);
  // 合同总值（中文大写）— SR+4 行，col 1
  putVal(ws, SR + 4, 1, '(5)合同总值：   美元' + toChineseRMB(totalAmt));

  cleanExcess(ws, R, products.length, TPL);
  rebuildRowMerges(ws, R, products.length, refPatterns, SR);
  unifyDataStyles(ws, R, SR, 28);
}


// ==================== 发票 ====================
// 模板: 32 rows, data R15-R22 (8), summary R23
function fillInvoice(ws, products, contractNo, serialDate, totalQty, totalAmt, refPatterns) {
  const R = 15, TPL = 8;
  putVal(ws, 7, 20, contractNo);
  putDate(ws, 9, 20, serialDate);
  const SR = ensureDataRows(ws, R, TPL, products.length);

  products.forEach((p, i) => {
    const r = R + i;
    putNum(ws, r, 1, i + 1);
    putVal(ws, r, 4, p.spec || p.nameCN);
    putNum(ws, r, 14, p.quantity);
    putVal(ws, r, 17, p.unit || '个');
    putNum(ws, r, 19, p.unitPriceUSD);
    putNum(ws, r, 23, p.unitPriceUSD * p.quantity);
  });

  putNum(ws, SR, 14, totalQty);
  putVal(ws, SR, 19, '总合计:');
  putNum(ws, SR, 23, totalAmt);
  cleanExcess(ws, R, products.length, TPL);
  rebuildRowMerges(ws, R, products.length, refPatterns, SR);
  unifyDataStyles(ws, R, SR, 26);
}


// ==================== 装箱单 ====================
// 模板: 89 rows, data R12-R82 (71), summary R83
function fillPackingList(ws, products, boxes, contractNo, serialDate, totalBoxes, boxRows, refPatterns) {
  const R = 12, TPL = 71;
  putDate(ws, 2, 22, serialDate);
  putVal(ws, 6, 22, contractNo);
  const SR = ensureDataRows(ws, R, TPL, boxRows);

  let ri = R, sq = 1;
  boxes.forEach(box => {
    box.products.forEach((p, idx) => {
      const fi = idx === 0;
      const q  = p.quantityPerBox || p.quantity;
      const nw = p.netWeightPerBox || 0;

      putNum(ws, ri, 1, sq++);
      putVal(ws, ri, 3, p.nameCN);
      if (fi) { putNum(ws, ri, 14, box.boxSeq); putNum(ws, ri, 21, box.weight || 0); }
      putNum(ws, ri, 17, q);
      putVal(ws, ri, 19, p.unit || '个');
      putNum(ws, ri, 24, nw);

      ri++;
    });
  });

  const qSum = boxes.reduce((s, b) => s + b.products.reduce((ss, p) => ss + (p.quantityPerBox || p.quantity || 0), 0), 0);
  const gSum = boxes.reduce((s, b) => s + (b.weight || 0), 0);
  const nSum = boxes.reduce((s, b) => s + b.products.reduce((ss, p) => ss + (p.netWeightPerBox || 0), 0), 0);

  putVal(ws, SR, 1, '合计:');
  putNum(ws, SR, 14, totalBoxes);
  putNum(ws, SR, 17, qSum);
  putNum(ws, SR, 21, gSum);
  putNum(ws, SR, 24, nSum);
  cleanExcess(ws, R, boxRows, TPL);
  rebuildRowMerges(ws, R, boxRows, refPatterns, SR);
  unifyDataStyles(ws, R, SR, 27);
}


// ==================== 报关草单合并 ====================
// 模板: 21 rows, data R13-R20 (8), summary R21
function fillCustomsMerged(ws, products, destination, contractNo, serialDate,
  totalQty, totalBoxes, totalGross, totalNet, totalAmt, refPatterns) {
  const R = 13, TPL = 8;

  putVal(ws, 6, 3, contractNo);
  putVal(ws, 6, 16, destination);
  putDate(ws, 3, 16, serialDate);
  putDate(ws, 3, 19, serialDate);
  putNum(ws, 7, 7, totalBoxes);
  putVal(ws, 7, 10, r4(totalGross).toFixed(2) + '千克');
  putNum(ws, 7, 14, totalNet);

  const SR = ensureDataRows(ws, R, TPL, products.length);

  products.forEach((p, i) => {
    const r = R + i;
    putNum(ws, r, 1, i + 1);
    putNum(ws, r, 2, p.hsCodeMerged || p.hsCode || 0);
    putVal(ws, r, 4, p.sku);
    putVal(ws, r, 5, p.nameCN);
    putVal(ws, r, 7, p.spec || p.nameCN);
    putVal(ws, r, 11, p.material || '');
    putVal(ws, r, 12, p.usage || '家居用品');
    putNum(ws, r, 13, p.quantity);
    putVal(ws, r, 14, p.unit || '个');
    putNum(ws, r, 15, p.netWeight || 0);
    putNum(ws, r, 16, p.unitPriceUSD);
    putNum(ws, r, 17, p.unitPriceUSD * p.quantity);
    putVal(ws, r, 18, 'USD');
    putVal(ws, r, 19, p.originCountry || '中国');
    putVal(ws, r, 21, destination);
    putVal(ws, r, 24, p.domesticSource || '');
    putVal(ws, r, 27, '照章征税');
  });

  putVal(ws, SR, 5, '合计');
  putNum(ws, SR, 13, totalQty);
  putVal(ws, SR, 14, '净重');
  putNum(ws, SR, 15, totalNet);
  putVal(ws, SR, 16, '金额');
  putNum(ws, SR, 17, totalAmt);
  putVal(ws, SR, 18, 'USD');
  putNum(ws, SR, 28, totalGross);
  putNum(ws, SR, 29, totalBoxes);
  cleanExcess(ws, R, products.length, TPL);
  rebuildRowMerges(ws, R, products.length, refPatterns, SR);
  unifyDataStyles(ws, R, SR, 29);
}


module.exports = { generate, dateToSerial };
