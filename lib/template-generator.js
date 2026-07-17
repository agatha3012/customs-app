const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

// ==================== 工具函数 ====================

function colLetter(idx) {
  if (idx < 26) return String.fromCharCode(65 + idx);
  return String.fromCharCode(64 + Math.floor(idx / 26)) + String.fromCharCode(65 + (idx % 26));
}
function cellRef(r, c) { return colLetter(c) + (r + 1); }

function dateToSerial(d) {
  const base = new Date(1899, 11, 30);
  const diff = d.getTime() - base.getTime();
  return Math.round(diff / (24 * 60 * 60 * 1000));
}

function r4(v) { return parseFloat((v || 0).toFixed(4)); }

/** 写入数值单元格，内部4位精度，Excel显示2位 */
function setNum(ws, r, c, value) {
  const ref = cellRef(r, c);
  ws[ref] = { t: 'n', v: r4(value), z: '0.00' };
}
function setStr(ws, r, c, value) {
  const ref = cellRef(r, c);
  ws[ref] = { t: 's', v: String(value) };
}
function setNumRaw(ws, r, c, value) {
  const ref = cellRef(r, c);
  ws[ref] = { t: 'n', v: value };
}

function clearRow(ws, r) {
  const toDelete = [];
  Object.keys(ws).forEach(ref => {
    if (ref.startsWith('!')) return;
    const cell = XLSX.utils.decode_cell(ref);
    if (cell.r === r) toDelete.push(ref);
  });
  toDelete.forEach(ref => delete ws[ref]);
}
function clearRows(ws, startR, endR) {
  for (let r = startR; r < endR; r++) clearRow(ws, r);
}
function clearMergesInRange(ws, startR, endR) {
  if (!ws['!merges']) return;
  ws['!merges'] = ws['!merges'].filter(m => m.s.r < startR || m.s.r > endR);
}

/** 插入空白行：在 atRow 处插入 count 行 */
function insertRows(ws, atRow, count) {
  const cellEntries = [];
  const metaKeys = ['!ref', '!merges', '!cols', '!rows', '!autofilter', '!protect'];
  Object.keys(ws).forEach(ref => {
    if (metaKeys.includes(ref)) return;
    const cell = XLSX.utils.decode_cell(ref);
    cellEntries.push({ cell, value: ws[ref] });
    delete ws[ref];
  });
  cellEntries.forEach(({ cell, value }) => {
    if (cell.r >= atRow) cell.r += count;
    ws[XLSX.utils.encode_cell(cell)] = value;
  });
  if (ws['!merges']) {
    ws['!merges'] = ws['!merges'].map(m => {
      if (m.s.r >= atRow) return { s: { r: m.s.r + count, c: m.s.c }, e: { r: m.e.r + count, c: m.e.c } };
      if (m.e.r >= atRow && m.s.r < atRow) return { s: { r: m.s.r, c: m.s.c }, e: { r: m.e.r + count, c: m.e.c } };
      return m;
    });
  }
  if (ws['!ref']) {
    const range = XLSX.utils.decode_range(ws['!ref']);
    range.e.r += count;
    ws['!ref'] = XLSX.utils.encode_range(range);
  }
}

// ==================== 主入口 ====================

function generate(params) {
  const { products, boxes, destination, exchangeRate, constant, templatePath, outputPath } = params;

  if (!fs.existsSync(templatePath)) throw new Error('模板文件不存在: ' + templatePath);
  fs.copyFileSync(templatePath, outputPath);

  const wb = XLSX.readFile(outputPath);

  const today = new Date();
  const serialDate = dateToSerial(today);
  const dateStr = today.toISOString().slice(0, 10).replace(/-/g, '');
  const contractNo = 'YLW-' + dateStr + '001';

  // 汇总数据
  const totalQuantity = products.reduce((s, p) => s + p.quantity, 0);
  const totalBoxes = boxes.length;
  const totalGrossWeight = boxes.reduce((s, b) => s + (b.weight || 0), 0);
  const totalNetWeight = products.reduce((s, p) => s + (p.netWeight || 0), 0);
  const totalAmount = products.reduce((s, p) => s + (p.unitPriceUSD * p.quantity), 0);

  // 箱行数
  const boxRowCount = boxes.reduce((s, b) => s + b.products.length, 0);

  processCustomsDraft(wb.Sheets['报关草单'], products, boxes, destination, contractNo, serialDate, totalQuantity, totalBoxes, totalGrossWeight, totalNetWeight, totalAmount, boxRowCount);
  processContract(wb.Sheets['合同'], products, destination, contractNo, serialDate, totalQuantity, totalAmount);
  processInvoice(wb.Sheets['发票'], products, destination, contractNo, serialDate, totalQuantity, totalAmount);
  processPackingList(wb.Sheets['装箱单'], products, boxes, contractNo, serialDate, totalBoxes, boxRowCount);
  processCustomsMerged(wb.Sheets['报关草单合并'], products, destination, contractNo, serialDate, totalQuantity, totalBoxes, totalGrossWeight, totalNetWeight, totalAmount);

  XLSX.writeFile(wb, outputPath);
  console.log('Generated:', outputPath);
  return outputPath;
}

// ==================== 报关草单 ====================
// 模板: 87 rows, header R1-R12, data R13-R83 (71行), summary R84
// 列: A=序号 B=商品编码 D=商检编码 E=商品名称 G=规格型号
//     K=数量 L=单位 M=净重 N=单价 O=总价 P=币制 Q=原产国
//     S=最终目的地 V=境内货源地 Y=照章征税 Z=毛重 AA=件数
function processCustomsDraft(ws, products, boxes, destination, contractNo, serialDate, totalQuantity, totalBoxes, totalGrossWeight, totalNetWeight, totalAmount, boxRowCount) {
  // --- 元数据 ---
  // C6 (row5,col2): 合同协议号
  setStr(ws, 5, 2, contractNo);
  // N6 (row5,col13): 运抵国
  setStr(ws, 5, 13, destination);
  // N3 (row2,col13): 出口日期
  setNumRaw(ws, 2, 13, serialDate);
  // Q3 (row2,col16): 申报日期 → 模板是 Q3 申报日期，值放在后面
  setNumRaw(ws, 2, 16, serialDate);
  // G7 (row6,col6): 件数
  setStr(ws, 6, 6, totalBoxes + '件');
  // J7 (row6,col9): 毛重
  setStr(ws, 6, 9, r4(totalGrossWeight).toFixed(2) + '千克');
  // L7 (row6,col11): 净重
  setStr(ws, 6, 11, r4(totalNetWeight).toFixed(2) + '千克');

  // --- 数据行 ---
  const DATA_START = 12;    // R13 (0-indexed 12)
  const TEMPLATE_ROWS = 71; // R13-R83
  const TEMPLATE_END = DATA_START + TEMPLATE_ROWS; // 83

  clearMergesInRange(ws, DATA_START, TEMPLATE_END - 1);
  clearRows(ws, DATA_START, TEMPLATE_END);

  if (boxRowCount > TEMPLATE_ROWS) {
    insertRows(ws, DATA_START, boxRowCount - TEMPLATE_ROWS);
  }

  let rowIdx = DATA_START;
  let seq = 1;

  boxes.forEach(box => {
    box.products.forEach(prod => {
      const isFirstInBox = box.products.indexOf(prod) === 0;
      const qty = prod.quantityPerBox || prod.quantity;
      const netW = prod.netWeightPerBox || 0;

      // A: 序号
      setNumRaw(ws, rowIdx, 0, seq++);
      // B-C: 商品编码
      setNumRaw(ws, rowIdx, 1, prod.hsCode || '');
      // D: 商检编码
      setStr(ws, rowIdx, 3, prod.sku);
      // E-F: 商品名称
      setStr(ws, rowIdx, 4, prod.nameCN);
      // G-J: 规格型号
      setStr(ws, rowIdx, 6, prod.spec || prod.nameCN);
      // K: 数量
      setNumRaw(ws, rowIdx, 10, qty);
      // L: 单位
      setStr(ws, rowIdx, 11, prod.unit || '个');
      // M: 净重
      setNum(ws, rowIdx, 12, netW);
      // N: 单价
      setNum(ws, rowIdx, 13, prod.unitPriceUSD);
      // O: 总价
      setNum(ws, rowIdx, 14, prod.unitPriceUSD * qty);
      // P: 币制
      setStr(ws, rowIdx, 15, 'USD');
      // Q: 原产国
      setStr(ws, rowIdx, 16, prod.originCountry || '中国');
      // S: 最终目的地 (col 18)
      setStr(ws, rowIdx, 18, destination);
      // V: 境内货源地 (col 21)
      setStr(ws, rowIdx, 21, prod.domesticSource || '');
      // Y: 照章征税 (col 24)
      setStr(ws, rowIdx, 24, '照章征税');
      // Z: 毛重 - only first product per box
      if (isFirstInBox) {
        setNum(ws, rowIdx, 25, box.weight || 0);
      }
      // AA: 件数 (boxSeq) - only first product per box
      if (isFirstInBox) {
        setNumRaw(ws, rowIdx, 26, box.boxSeq);
      }

      rowIdx++;
    });
  });

  // --- 汇总行 (rowIdx) ---
  setNumRaw(ws, rowIdx, 10, totalQuantity);
  setNum(ws, rowIdx, 12, totalNetWeight);
  setStr(ws, rowIdx, 13, '金额');
  setNum(ws, rowIdx, 14, totalAmount);
  setStr(ws, rowIdx, 15, 'USD');
  setNum(ws, rowIdx, 25, totalGrossWeight);
  setNumRaw(ws, rowIdx, 26, totalBoxes);
}

// ==================== 合同 ====================
// 模板: 56 rows, W5=合同号, W7=日期, 数据 R21-R28 (8行), 汇总 R29
// 列: A-L=名称 M-O=数量 P-Q=单位 R-V=单价 W-AB=金额
function processContract(ws, products, destination, contractNo, serialDate, totalQuantity, totalAmount) {
  // W5 (row4,col22): 合同号
  setStr(ws, 4, 22, contractNo);
  // W7 (row6,col22): 日期
  setNumRaw(ws, 6, 22, serialDate);

  const DATA_START = 20;    // R21 (0-indexed 20)
  const TEMPLATE_ROWS = 8;  // R21-R28

  clearMergesInRange(ws, DATA_START, DATA_START + TEMPLATE_ROWS - 1);
  clearRows(ws, DATA_START, DATA_START + TEMPLATE_ROWS);

  if (products.length > TEMPLATE_ROWS) {
    insertRows(ws, DATA_START, products.length - TEMPLATE_ROWS);
  }

  products.forEach((prod, i) => {
    const r = DATA_START + i;
    // A-L: 名称及规格 (col 0, merged A-L)
    setStr(ws, r, 0, prod.spec || prod.nameCN);
    // M-O: 数量 (col 12)
    setNumRaw(ws, r, 12, prod.quantity);
    // P-Q: 单位 (col 15)
    setStr(ws, r, 15, prod.unit || '个');
    // R-V: 单价 (col 17)
    setNum(ws, r, 17, prod.unitPriceUSD);
    // W-AB: 金额 (col 22)
    setNum(ws, r, 22, prod.unitPriceUSD * prod.quantity);
  });

  // 汇总行
  const sumR = DATA_START + Math.max(products.length, TEMPLATE_ROWS);
  setNumRaw(ws, sumR, 12, totalQuantity);
  setNum(ws, sumR, 22, totalAmount);
}

// ==================== 发票 ====================
// 模板: 32 rows, T7=合同号, T9=日期, 数据 R15-R22 (8行), 汇总 R23
// 列: A=序号 D=名称 N=数量 Q=单位 S=单价 W=金额
function processInvoice(ws, products, destination, contractNo, serialDate, totalQuantity, totalAmount) {
  // T7 (row6,col19): 合同号
  setStr(ws, 6, 19, contractNo);
  // T9 (row8,col19): 日期
  setNumRaw(ws, 8, 19, serialDate);

  const DATA_START = 14;    // R15 (0-indexed 14)
  const TEMPLATE_ROWS = 8;  // R15-R22

  clearMergesInRange(ws, DATA_START, DATA_START + TEMPLATE_ROWS - 1);
  clearRows(ws, DATA_START, DATA_START + TEMPLATE_ROWS);

  if (products.length > TEMPLATE_ROWS) {
    insertRows(ws, DATA_START, products.length - TEMPLATE_ROWS);
  }

  products.forEach((prod, i) => {
    const r = DATA_START + i;
    // A: 序号 (col 0, merged A-C)
    setNumRaw(ws, r, 0, i + 1);
    // D: 货物名称 (col 3, merged D-M)
    setStr(ws, r, 3, prod.spec || prod.nameCN);
    // N: 数量 (col 13, merged N-P)
    setNumRaw(ws, r, 13, prod.quantity);
    // Q: 单位 (col 16, merged Q-R)
    setStr(ws, r, 16, prod.unit || '个');
    // S: 单价 (col 18, merged S-V)
    setNum(ws, r, 18, prod.unitPriceUSD);
    // W: 金额 (col 22, merged W-Z)
    setNum(ws, r, 22, prod.unitPriceUSD * prod.quantity);
  });

  // 汇总行
  const sumR = DATA_START + Math.max(products.length, TEMPLATE_ROWS);
  setNumRaw(ws, sumR, 13, totalQuantity);
  setStr(ws, sumR, 18, '总合计:');
  setNum(ws, sumR, 22, totalAmount);
}

// ==================== 装箱单 ====================
// 模板: 89 rows, V2=日期, V6=合同号, 数据 R12-R82 (71行), 汇总 R83
// 列: A=箱号 C=名称 N=总箱数 Q=总数量 S=单位 U=总毛重 X=总净重
function processPackingList(ws, products, boxes, contractNo, serialDate, totalBoxes, boxRowCount) {
  // V2 (row1,col21): 日期
  setNumRaw(ws, 1, 21, serialDate);
  // V6 (row5,col21): 合同号
  setStr(ws, 5, 21, contractNo);

  const DATA_START = 11;    // R12 (0-indexed 11)
  const TEMPLATE_ROWS = 71; // R12-R82

  clearMergesInRange(ws, DATA_START, DATA_START + TEMPLATE_ROWS - 1);
  clearRows(ws, DATA_START, DATA_START + TEMPLATE_ROWS);

  if (boxRowCount > TEMPLATE_ROWS) {
    insertRows(ws, DATA_START, boxRowCount - TEMPLATE_ROWS);
  }

  let rowIdx = DATA_START;
  let seq = 1;

  boxes.forEach(box => {
    box.products.forEach(prod => {
      const isFirstInBox = box.products.indexOf(prod) === 0;
      const qty = prod.quantityPerBox || prod.quantity;
      const netW = prod.netWeightPerBox || 0;

      // A: 箱号 (col 0)
      setNumRaw(ws, rowIdx, 0, seq++);
      // C: 货物名称 (col 2, merged C-M)
      setStr(ws, rowIdx, 2, prod.nameCN);
      // N: 总箱数 (col 13) - only first product per box
      if (isFirstInBox) {
        setNumRaw(ws, rowIdx, 13, box.boxSeq);
      }
      // Q: 总数量 (col 16, merged Q-R)
      setNumRaw(ws, rowIdx, 16, qty);
      // S: 单位 (col 18, merged S-T)
      setStr(ws, rowIdx, 18, prod.unit || '个');
      // U: 总毛重 (col 20) - only first product per box
      if (isFirstInBox) {
        setNum(ws, rowIdx, 20, box.weight || 0);
      }
      // X: 总净重 (col 23)
      setNum(ws, rowIdx, 23, netW);

      rowIdx++;
    });
  });

  // 汇总行 (rowIdx)
  const totalQtyAll = boxes.reduce((s, b) => s + b.products.reduce((ss, p) => ss + (p.quantityPerBox || 0), 0), 0);
  const totalGrossAll = boxes.reduce((s, b) => s + (b.weight || 0), 0);
  const totalNetAll = boxes.reduce((s, b) => s + b.products.reduce((ss, p) => ss + (p.netWeightPerBox || 0), 0), 0);
  setStr(ws, rowIdx, 0, '合计:');
  setNumRaw(ws, rowIdx, 13, totalBoxes);
  setNumRaw(ws, rowIdx, 16, totalQtyAll);
  setNum(ws, rowIdx, 20, totalGrossAll);
  setNum(ws, rowIdx, 23, totalNetAll);
}

// ==================== 报关草单合并 ====================
// 模板: 21 rows, R12 header, data R13-R20 (8行), summary R21
// 列: A=序号 B=商品编码 D=商检编码 E=商品名称 G=规格型号
//     K=材质 L=用途 M=数量 N=单位 O=净重 P=单价 Q=总价
//     R=币制 S=原产国 U=最终目的地 X=境内货源地 AA=照章征税
//     AB=毛重(summary) AC=件数(summary)
function processCustomsMerged(ws, products, destination, contractNo, serialDate, totalQuantity, totalBoxes, totalGrossWeight, totalNetWeight, totalAmount) {
  // C6 (row5,col2): 合同号
  setStr(ws, 5, 2, contractNo);
  // P6 (row5,col15): 运抵国
  setStr(ws, 5, 15, destination);
  // P3 (row2,col15): 出口日期
  setNumRaw(ws, 2, 15, serialDate);
  // S3 (row2,col18): 申报日期
  setNumRaw(ws, 2, 18, serialDate);
  // G7 (row6,col6): 件数
  setNumRaw(ws, 6, 6, totalBoxes);
  // J7 (row6,col9): 毛重
  setStr(ws, 6, 9, r4(totalGrossWeight).toFixed(2) + '千克');
  // N7 (row6,col13): 净重
  setNum(ws, 6, 13, totalNetWeight);

  const DATA_START = 12;    // R13 (0-indexed 12)
  const TEMPLATE_ROWS = 8;  // R13-R20

  clearMergesInRange(ws, DATA_START, DATA_START + TEMPLATE_ROWS - 1);
  clearRows(ws, DATA_START, DATA_START + TEMPLATE_ROWS);

  if (products.length > TEMPLATE_ROWS) {
    insertRows(ws, DATA_START, products.length - TEMPLATE_ROWS);
  }

  products.forEach((prod, i) => {
    const r = DATA_START + i;

    // A: 序号
    setNumRaw(ws, r, 0, i + 1);
    // B-C: 商品编码
    setNumRaw(ws, r, 1, prod.hsCodeMerged || prod.hsCode || '');
    // D: 商检编码
    setStr(ws, r, 3, prod.sku);
    // E-F: 商品名称
    setStr(ws, r, 4, prod.nameCN);
    // G-J: 规格型号
    setStr(ws, r, 6, prod.spec || prod.nameCN);
    // K: 材质
    setStr(ws, r, 10, prod.material || '');
    // L: 用途
    setStr(ws, r, 11, prod.usage || '家居用品');
    // M: 数量
    setNumRaw(ws, r, 12, prod.quantity);
    // N: 单位
    setStr(ws, r, 13, prod.unit || '个');
    // O: 净重
    setNum(ws, r, 14, prod.netWeight || 0);
    // P: 单价
    setNum(ws, r, 15, prod.unitPriceUSD);
    // Q: 总价
    setNum(ws, r, 16, prod.unitPriceUSD * prod.quantity);
    // R: 币制
    setStr(ws, r, 17, 'USD');
    // S: 原产国
    setStr(ws, r, 18, prod.originCountry || '中国');
    // U: 最终目的地 (col 20)
    setStr(ws, r, 20, destination);
    // X: 境内货源地 (col 23)
    setStr(ws, r, 23, prod.domesticSource || '');
    // AA: 照章征税 (col 26)
    setStr(ws, r, 26, '照章征税');
  });

  // 汇总行
  const sumR = DATA_START + Math.max(products.length, TEMPLATE_ROWS);
  setStr(ws, sumR, 4, '合计');
  setNumRaw(ws, sumR, 12, totalQuantity);
  setStr(ws, sumR, 13, '净重');
  setNum(ws, sumR, 14, totalNetWeight);
  setStr(ws, sumR, 15, '金额');
  setNum(ws, sumR, 16, totalAmount);
  setStr(ws, sumR, 17, 'USD');
  setNum(ws, sumR, 27, totalGrossWeight);
  setNumRaw(ws, sumR, 28, totalBoxes);
}

module.exports = { generate, dateToSerial };
