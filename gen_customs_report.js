const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// ====== SETTINGS ======
const exchangeRate = 6.7989;
const constant = 1.25;
const destination = '德国';
const today = new Date();
const dateStr = today.toISOString().slice(0, 10);
const contractNo = 'YLW-' + dateStr.replace(/-/g, '') + '001';

// ====== PRODUCT DB ======
const db = JSON.parse(fs.readFileSync('data/product_db_merged.json', 'utf-8'));
const manualData = JSON.parse(fs.readFileSync('data/manual_product_data.json', 'utf-8'));

function lookupProduct(sku) {
  let records = db[sku] || [];
  if (!records || records.length === 0) {
    const skuSuffix = sku.replace(/^[A-Z]+-/, '');
    for (const [key, val] of Object.entries(db)) {
      const keySuffix = key.replace(/^[A-Z]+-/, '');
      if (keySuffix === skuSuffix && val.length > 0) { records = val; break; }
    }
  }
  let maxPrice = 0, supplier = '', city = '';
  records.forEach(r => {
    if (r.price > maxPrice || (r.price === maxPrice && r.supplier && !supplier)) {
      maxPrice = r.price; supplier = r.supplier;
    }
  });
  const cityPatterns = [
    [/^东莞/, '东莞'], [/^深圳/, '深圳'], [/^广州/, '广州'],
    [/^中山/, '中山'], [/^惠州/, '惠州'], [/^佛山/, '佛山'],
    [/^珠海/, '珠海'], [/^江门/, '江门'], [/^义乌/, '义乌'],
    [/^杭州/, '杭州'], [/^宁波/, '宁波'], [/^温州/, '温州'],
    [/^厦门/, '厦门'], [/^福州/, '福州'], [/^泉州/, '泉州'],
    [/^上海/, '上海'], [/^北京/, '北京'], [/^苏州/, '苏州'],
    [/^沧州/, '沧州'], [/^任丘/, '任丘'], [/^河间/, '沧州'],
  ];
  for (const [re, c] of cityPatterns) { if (re.test(supplier)) { city = c; break; } }
  // 供应商名称不含城市前缀的，通过供应商全名匹配（工商查询结果）
  if (!city) {
    const supplierCityMap = {
      '伟能(广东)新材料有限公司': '广州',
      '伟能（广东）新材料有限公司': '广州',
    };
    if (supplierCityMap[supplier]) city = supplierCityMap[supplier];
  }
  // 人工补充数据优先覆盖
  const md = manualData[sku];
  if (md) {
    if (md.maxPrice > 0) maxPrice = md.maxPrice;
    if (md.supplier) supplier = md.supplier;
    if (md.city) city = md.city;
  }
  // 仍无货源地 → 标记为待确认
  if (!city) city = '待确认';
  const netWPerUnit = (md && md.netWeightPerUnit > 0) ? md.netWeightPerUnit : 0.5;
  const unitPriceUSD = (maxPrice * constant) / exchangeRate;
  return { maxPrice, supplier, city, netWPerUnit, unitPriceUSD };
}

// ====== PARSE INVOICE ======
const invPath = 'D:/桌面/7-14凡洋德国陆运包税FBA不带电20箱发票.xlsx';
const wb = XLSX.readFile(invPath);
const ws = wb.Sheets['CIPI'];
const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true });

const boxes = [];
let currentBox = null, prevBoxId = '';

for (let i = 9; i < raw.length; i++) {
  const row = raw[i];
  if (!row || !row[0]) continue;
  const sku = String(row[0]).trim();
  if (!sku) continue;
  const boxId = row[1] ? String(row[1]).trim() : '';
  const isNewBox = boxId !== '' && boxId !== prevBoxId;
  if (boxId !== '') prevBoxId = boxId;
  const cnName = row[7] ? String(row[7]).trim() : '';
  let hsCode = 0, hsIdx = -1;
  for (let j = 8; j < Math.min(row.length, 14); j++) {
    if (typeof row[j] === 'number' && row[j] >= 1000000 && row[j] <= 9999999999) { hsCode = row[j]; hsIdx = j; break; }
  }
  let material = '', usage = '家居用品';
  if (hsIdx > 0) {
    for (let j = hsIdx + 1; j < Math.min(row.length, hsIdx + 3); j++) {
      if (typeof row[j] === 'string') {
        if (!material) material = row[j].trim();
        else usage = row[j].trim();
      }
    }
  }
  const tailNums = [];
  for (let j = Math.max(hsIdx + 1, 10); j < row.length; j++) {
    if (typeof row[j] === 'number' && row[j] > 0) tailNums.push(row[j]);
  }
  let qtyPerBox = 0;
  if (isNewBox) { qtyPerBox = 1; }
  else if (tailNums.length === 1) { qtyPerBox = tailNums[0]; }
  else if (tailNums.length >= 2) {
    for (let k = tailNums.length - 1; k >= 0; k--) {
      if ((tailNums[k] !== Math.floor(tailNums[k]) || tailNums[k] < 50) && tailNums[k] < 1000) {
        qtyPerBox = tailNums.slice(0, k).reduce((a, b) => a + b, 0); break;
      }
    }
    if (qtyPerBox === 0) qtyPerBox = tailNums[0];
  }
  if (isNewBox) {
    if (currentBox && currentBox.products.length > 0) boxes.push(currentBox);
    currentBox = { id: boxId, boxSeq: boxes.length + 1, weight: parseFloat(row[5]) || 0, products: [] };
  }
  if (currentBox) currentBox.products.push({ sku, nameCN: cnName, hsCode, material, usage, unit: '个', qtyPerBox });
}
if (currentBox && currentBox.products.length > 0) boxes.push(currentBox);

// ====== BUILD PRODUCT AGGREGATE ======
const prodMap = new Map();
boxes.forEach(box => {
  box.products.forEach(p => {
    if (!prodMap.has(p.sku)) {
      const info = lookupProduct(p.sku);
      prodMap.set(p.sku, {
        sku: p.sku, nameCN: p.nameCN, hsCode: p.hsCode,
        material: p.material, usage: p.usage, unit: p.unit,
        totalQty: 0, netWPerUnit: info.netWPerUnit,
        unitPriceUSD: info.unitPriceUSD, city: info.city,
        supplier: info.supplier,
      });
    }
    prodMap.get(p.sku).totalQty += p.qtyPerBox;
  });
});

const products = Array.from(prodMap.values());
const totalQty = products.reduce((s,p) => s + p.totalQty, 0);
const totalNetWt = products.reduce((s,p) => s + p.totalQty * p.netWPerUnit, 0);
const totalAmount = products.reduce((s,p) => s + p.unitPriceUSD * p.totalQty, 0);
const totalGrossWt = boxes.reduce((s,b) => s + b.weight, 0);
const boxRowCount = boxes.reduce((s,b) => s + b.products.length, 0);
const r4 = v => parseFloat(v.toFixed(4));

// ====== PRODUCE OUTPUT ======
const lines = [];

lines.push('# 中华人民共和国海关出口货物报关单');
lines.push('');
lines.push('**预录入编号：** 　　　　　　　　　　　　**申报口岸：** 　　　　　　　　　　　　**海关编号：**');
lines.push('');
lines.push('| 字段 | 内容 | | 字段 | 内容 | | 字段 | 内容 |');
lines.push('|------|------|---|------|------|---|------|------|');
lines.push('| **境内发货人** | 深圳市亿莱沃实业有限公司 | | **出境关别** | | | **出口日期** | ' + dateStr + ' |');
lines.push('| **境外收货人(AEO)** | HK Kuntaikang Industrial Co., Limited | | **运输方式** | 铁路运输 | | **申报日期** | ' + dateStr + ' |');
lines.push('| **生产销售单位** | 深圳市亿莱沃实业有限公司 | | **监管方式** | 一般贸易 | | **备案号** | |');
lines.push('| **合同协议号** | ' + contractNo + ' | | **征免性质** | 一般征税 | | **许可证号** | |');
lines.push('| **贸易国(地区)** | 中国香港 | | **运抵国(地区)** | ' + destination + ' | | **指运港** | |');
lines.push('| **包装种类** | 纸箱 | | **件数** | ' + boxes.length + '件 | | **离境口岸** | |');
lines.push('| **毛重** | ' + totalGrossWt.toFixed(2) + '千克 | | **净重** | ' + totalNetWt.toFixed(2) + '千克 | | **成交方式** | FOB |');
lines.push('| **随附单证及编号** | | | **运费** | | | **保费** | |');
lines.push('| **特殊关系确认** | 否 | | **价格影响确认** | 否 | | **支付特许权使用费** | 否 |');
lines.push('');
lines.push('> 唛头及备注：');
lines.push('> 原产国/目的国/货源地未注明的项默认与已提供项相同');
lines.push('');

// Product summary first
lines.push('## 商品汇总（按HS编码归并）');
lines.push('');
lines.push('| 序号 | 商品编码 | 商检编码 | 商品名称 | 规格型号 | 材质 | 用途 | 总数量 | 单位 | 总净重(kg) | 单价(USD) | 总价(USD) | 币制 | 原产国 | 最终目的地 | 境内货源地 | 征免方式 |');
lines.push('|------|----------|----------|----------|----------|------|------|--------|------|------------|-----------|-----------|------|--------|------------|------------|----------|');

products.forEach((p, i) => {
  const netW = r4(p.totalQty * p.netWPerUnit);
  const totalUSD = r4(p.unitPriceUSD * p.totalQty);
  lines.push('| ' + [
    i + 1,
    String(p.hsCode),
    p.sku,
    p.nameCN,
    p.nameCN,
    p.material || '',
    p.usage || '家居用品',
    p.totalQty,
    p.unit,
    netW.toFixed(2),
    p.unitPriceUSD.toFixed(4),
    totalUSD.toFixed(2),
    'USD',
    '中国',
    destination,
    p.city || '',
    '照章征税',
  ].join(' | ') + ' |');
});

// Total row
lines.push('| **合计** | | | | | | | **' + totalQty + '** | | **' + totalNetWt.toFixed(2) + '** | | **' + totalAmount.toFixed(2) + '** | USD | | | | |');

lines.push('');
lines.push('## 装箱明细（按箱号逐行列出）');
lines.push('');
lines.push('| 序号 | 商品编码 | 商检编码 | 商品名称 | 规格型号 | 材质 | 用途 | 单箱数量 | 单位 | 单箱净重(kg) | 单价(USD) | 单箱总价(USD) | 币制 | 原产国 | 最终目的地 | 境内货源地 | 征免方式 | 箱毛重(kg) | 箱序号 |');
lines.push('|------|----------|----------|----------|----------|------|------|----------|------|-------------|-----------|--------------|------|--------|------------|------------|----------|-----------|--------|');

let seq = 1;
boxes.forEach(box => {
  box.products.forEach((p, idx) => {
    const info = prodMap.get(p.sku);
    const isFirst = idx === 0;
    const netW = r4(p.qtyPerBox * info.netWPerUnit);
    const totalUSD = r4(info.unitPriceUSD * p.qtyPerBox);

    lines.push('| ' + [
      seq++,
      String(p.hsCode),
      p.sku,
      p.nameCN,
      p.nameCN,
      p.material || '',
      p.usage || '家居用品',
      p.qtyPerBox,
      p.unit,
      netW.toFixed(2),
      info.unitPriceUSD.toFixed(4),
      totalUSD.toFixed(2),
      'USD',
      '中国',
      destination,
      info.city || '',
      '照章征税',
      isFirst ? box.weight.toFixed(1) : '',
      isFirst ? box.boxSeq : '',
    ].join(' | ') + ' |');
  });
});

// Summary row
lines.push('| **合计** | | | | | | | **' + totalQty + '** | | **' + totalNetWt.toFixed(2) + '** | | **' + totalAmount.toFixed(2) + '** | USD | | | | | | **' + totalGrossWt.toFixed(2) + '** | **' + boxes.length + '** |');

// Print stats
lines.push('');
lines.push('---');
lines.push('');
lines.push('**制单信息：**');
lines.push('- 申报日期：' + dateStr);
lines.push('- 合同协议号：' + contractNo);
lines.push('- 成交方式：FOB');
lines.push('- 运输方式：铁路运输');
lines.push('- 运抵国（地区）：' + destination);
lines.push('- 境内发货人：深圳市亿莱沃实业有限公司');
lines.push('- 境外收货人：HK Kuntaikang Industrial Co., Limited');
lines.push('- 总件数：' + boxes.length + ' 箱');
lines.push('- 总毛重：' + totalGrossWt.toFixed(2) + ' 千克');
lines.push('- 总净重：' + totalNetWt.toFixed(2) + ' 千克');
lines.push('- 总金额：USD ' + totalAmount.toFixed(2));
lines.push('- 总数量：' + totalQty + ' 个');
lines.push('- SKU 品名数：' + products.length);
lines.push('- 装箱明细行数：' + boxRowCount);

console.log(lines.join('\n'));
