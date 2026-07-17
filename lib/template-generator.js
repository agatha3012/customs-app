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


// ==================== 核心原则：只改 cell.value，不改任何样式 ====================

/** 写任意值 — 原封不动保留单元格的字体/填充/边框/对齐/数字格式 */
function putVal(ws, row, col, value) {
  ws.getCell(row, col).value = value;
}

/** 写数值 (4 位精度) */
function putNum(ws, row, col, value) {
  ws.getCell(row, col).value = r4(value);
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


// ==================== 行数调整 (唯一涉及结构变更的地方) ====================

/**
 * 确保数据区域有恰好 needRows 行可用。
 * - needRows > tplRows: 用 duplicateRow 在数据末尾插入行
 * - needRows < tplRows: 只清空多余行值，不删行
 * 返回汇总行位置 (1-indexed)。
 *
 * duplicateRow(src, count, true): 在 src 位置插入 count 份副本，原 src 行及以下全部下移。
 */
function ensureDataRows(ws, dataStart, tplRows, needRows) {
  if (needRows > tplRows) {
    // 在最后一个模板数据行位置插入副本 (数据行格式一致)
    const lastDataRow = dataStart + tplRows - 1;
    ws.duplicateRow(lastDataRow, needRows - tplRows, true);
  }
  // 新汇总行总是紧随最后一个数据行
  return needRows > tplRows
    ? dataStart + needRows
    : dataStart + tplRows;
}

/**
 * 清理"幻影行"——needRows < tplRows 时残留在数据区域下方的旧模板数据
 */
function cleanExcess(ws, dataStart, needRows, tplRows) {
  if (needRows < tplRows) {
    clearRowsVals(ws, dataStart + needRows, dataStart + tplRows);
  }
}

/**
 * 仅在报关草单使用：重建 Z(26) / AA(27) 列按箱垂直合并
 */
function rebuildBoxMerges(ws, startRow, boxes) {
  // 先删除数据区内的垂直合并 (Z/AA 列)
  if (ws.model.merges) {
    ws.model.merges = ws.model.merges.filter(m => {
      if (m.top < startRow) return true;
      if (m.top === m.bottom) return true;        // 水平合并保留
      if (m.left !== m.right) return true;         // 水平合并保留
      const isZ  = m.left === 26;
      const isAA = m.left === 27;
      return !(isZ || isAA);
    });
  }
  // 按箱重建
  let r = startRow;
  boxes.forEach(box => {
    const span = box.products.length;
    if (span > 1) {
      try { ws.mergeCells(r, 26, r + span - 1, 26); } catch (e) {}
      try { ws.mergeCells(r, 27, r + span - 1, 27); } catch (e) {}
    }
    r += span;
  });
}


// ==================== 主入口 ====================

async function generate(params) {
  const { products, boxes, destination, exchangeRate, constant, templatePath, outputPath } = params;

  if (!fs.existsSync(templatePath)) throw new Error('模板文件不存在: ' + templatePath);
  fs.copyFileSync(templatePath, outputPath);

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(outputPath);

  const today = new Date();
  const serialDate = dateToSerial(today);
  const dateStr = today.toISOString().slice(0, 10).replace(/-/g, '');
  const contractNo = 'YLW-' + dateStr + '001';

  const totalQty   = products.reduce((s, p) => s + p.quantity, 0);
  const totalBoxes = boxes.length;
  const totalGross = boxes.reduce((s, b) => s + (b.weight || 0), 0);
  const totalNet   = products.reduce((s, p) => s + (p.netWeight || 0), 0);
  const totalAmt   = products.reduce((s, p) => s + (p.unitPriceUSD * p.quantity), 0);
  const boxRows    = boxes.reduce((s, b) => s + b.products.length, 0);

  fillCustomsDraft(wb.getWorksheet('报关草单'), products, boxes, destination, contractNo, serialDate, totalQty, totalBoxes, totalGross, totalNet, totalAmt, boxRows);
  fillContract(wb.getWorksheet('合同'), products, contractNo, serialDate, totalQty, totalAmt);
  fillInvoice(wb.getWorksheet('发票'), products, contractNo, serialDate, totalQty, totalAmt);
  fillPackingList(wb.getWorksheet('装箱单'), products, boxes, contractNo, serialDate, totalBoxes, boxRows);
  fillCustomsMerged(wb.getWorksheet('报关草单合并'), products, destination, contractNo, serialDate, totalQty, totalBoxes, totalGross, totalNet, totalAmt);

  await wb.xlsx.writeFile(outputPath);
  console.log('Generated:', outputPath);
  return outputPath;
}


// ==================== 报关草单 ====================
// 模板: 87 rows, data R13-R83 (71), summary R84
function fillCustomsDraft(ws, products, boxes, destination, contractNo, serialDate,
  totalQty, totalBoxes, totalGross, totalNet, totalAmt, boxRows) {

  const R = 13;       // data start
  const TPL = 71;       // template data row count

  // 元数据
  putVal(ws, 6, 3, contractNo);                                  // C6
  putVal(ws, 6, 14, destination);                                // N6
  putDate(ws, 3, 14, serialDate);                                 // N3
  putDate(ws, 3, 17, serialDate);                                 // Q3
  putVal(ws, 7, 7, totalBoxes + '件');                           // G7
  putVal(ws, 7, 10, r4(totalGross).toFixed(2) + '千克');         // J7
  putVal(ws, 7, 12, r4(totalNet).toFixed(2) + '千克');           // L7

  // 行数
  const SR = ensureDataRows(ws, R, TPL, boxRows);               // summary row

  // Z/AA 按箱合并
  rebuildBoxMerges(ws, R, boxes);

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
}


// ==================== 合同 ====================
// 模板: 56 rows, data R21-R28 (8), summary R29
function fillContract(ws, products, contractNo, serialDate, totalQty, totalAmt) {
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
  cleanExcess(ws, R, products.length, TPL);
}


// ==================== 发票 ====================
// 模板: 32 rows, data R15-R22 (8), summary R23
function fillInvoice(ws, products, contractNo, serialDate, totalQty, totalAmt) {
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
}


// ==================== 装箱单 ====================
// 模板: 89 rows, data R12-R82 (71), summary R83
function fillPackingList(ws, products, boxes, contractNo, serialDate, totalBoxes, boxRows) {
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
}


// ==================== 报关草单合并 ====================
// 模板: 21 rows, data R13-R20 (8), summary R21
function fillCustomsMerged(ws, products, destination, contractNo, serialDate,
  totalQty, totalBoxes, totalGross, totalNet, totalAmt) {
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
}


module.exports = { generate, dateToSerial };
