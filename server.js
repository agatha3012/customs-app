const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const productLookup = require('./lib/product-lookup');
const templateGen = require('./lib/template-generator');
const supplierLookup = require('./lib/supplier-lookup');

const app = express();
const PORT = 3000;

app.use(express.json({ limit: '200mb' }));
app.use(express.urlencoded({ limit: '200mb', extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

const SETTINGS_FILE = path.join(__dirname, 'data', 'settings.json');
const MANUAL_DATA_FILE = path.join(__dirname, 'data', 'manual_product_data.json');
const CONFIRMATIONS_FILE = path.join(__dirname, 'data', 'confirmations.json');
const REVIEW_OVERRIDES_FILE = path.join(__dirname, 'data', 'review_overrides.json');
const UPLOADS_DIR = path.join(__dirname, 'uploads');
const OUTPUT_DIR = path.join(__dirname, 'output');
const TEMPLATE_PATH = path.join(__dirname, 'templates', '报关资料模板(7.9).xlsx');

// 确保 data 目录存在
const DATA_DIR = path.join(__dirname, 'data');
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });

[UPLOADS_DIR, OUTPUT_DIR].forEach(d => {
  if (!fs.existsSync(d)) fs.mkdirSync(d, { recursive: true });
});

// ==================== 设置 ====================
function readSettings() {
  try {
    return JSON.parse(fs.readFileSync(SETTINGS_FILE, 'utf-8'));
  } catch (e) {
    return { exchangeRate: 7.25, constant: 1.25, destination: '' };
  }
}
function writeSettings(data) {
  if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
  fs.writeFileSync(SETTINGS_FILE, JSON.stringify(data, null, 2), 'utf-8');
}

// ==================== 人工补充数据管理 ====================
/**
 * 人工补充的产品数据格式：
 * { "SKU": { sku, maxPrice, supplier, city, netWeightPerUnit, description, updatedAt } }
 */
function readManualData() {
  try {
    if (!fs.existsSync(MANUAL_DATA_FILE)) return {};
    return JSON.parse(fs.readFileSync(MANUAL_DATA_FILE, 'utf-8'));
  } catch (e) {
    return {};
  }
}
function writeManualData(data) {
  if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
  fs.writeFileSync(MANUAL_DATA_FILE, JSON.stringify(data, null, 2), 'utf-8');
}

app.get('/api/settings', (req, res) => {
  res.json(readSettings());
});

app.post('/api/settings', (req, res) => {
  try {
    const current = readSettings();
    const u = req.body;
    if (u.exchangeRate !== undefined) current.exchangeRate = parseFloat(u.exchangeRate) || current.exchangeRate;
    if (u.constant !== undefined) current.constant = parseFloat(u.constant) || current.constant;
    if (u.destination !== undefined) current.destination = String(u.destination);
    writeSettings(current);
    res.json({ success: true, data: current });
  } catch (err) {
    res.status(500).json({ success: false, message: '保存失败: ' + err.message });
  }
});

// ==================== 上传发票并解析 ====================
/** 生成安全的 ASCII 文件名，避免中文在 Windows 文件系统上编码乱码 */
function safeFilename(originalName) {
  const ext = path.extname(originalName);
  const ts = Date.now();
  const rnd = Math.random().toString(36).slice(2, 8);
  return ts + '-' + rnd + ext;
}

const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, UPLOADS_DIR),
  filename: (req, file, cb) => {
    cb(null, safeFilename(file.originalname));
  },
});
const upload = multer({
  storage,
  limits: { fileSize: 200 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (ext === '.xlsx' || ext === '.xls' || ext === '.et') cb(null, true);
    else cb(new Error('仅支持 .xlsx / .xls / .et 格式'));
  },
});

app.post('/api/upload-invoice', (req, res) => {
  upload.single('file')(req, res, (err) => {
    if (err) return res.status(400).json({ success: false, message: err.message });
    if (!req.file) return res.status(400).json({ success: false, message: '请选择文件' });

    try {
      const result = parseInvoice(req.file.path);
      const sessionFile = path.join(UPLOADS_DIR, 'last_parsed.json');
      // 修复文件名编码：浏览器上传的中文名可能被按 Latin-1 误解
      // 不同浏览器/系统可能用 UTF-8 或 GBK 编码
      let decodedName = req.file.originalname;
      try {
        const buf = Buffer.from(decodedName, 'latin1');
        const utf8 = buf.toString('utf8');
        if (utf8.indexOf('�') < 0) {
          decodedName = utf8;
        } else {
          // UTF-8 失败，尝试 GBK（中文 Windows 常见）
          try {
            const iconv = require('iconv-lite');
            const gbk = iconv.decode(buf, 'gbk');
            if (gbk && gbk.indexOf('�') < 0) decodedName = gbk;
          } catch (_) {}
        }
      } catch (_) {}
      // 兜底：百分号解码
      try { decodedName = decodeURIComponent(decodedName); } catch (_) {}
      const sessionData = { ...result, _originalName: decodedName };
      fs.writeFileSync(sessionFile, JSON.stringify(sessionData, null, 2), "utf-8");
      res.json({ success: true, ...result });
    } catch (e) {
      res.status(500).json({ success: false, message: '解析失败: ' + e.message });
    }
  });
});

/**
 * 解析发票 Excel (CIPI sheet)
 * 返回 { products: ProductItem[], boxes: BoxItem[] }
 */
function parseInvoice(filePath) {
  const wb = XLSX.readFile(filePath);

  // ★ 检测文件格式
  if (wb.SheetNames.includes('箱单') && wb.SheetNames.includes('发票')) {
    // 智谷供应链格式：箱单 sheet 含箱级产品明细
    return parseZhiguFormat(wb);
  }

  // ★ 凡洋英国 / 顺沃格式：单 sheet「箱单发票」含箱级明细+HS编码
  if (wb.SheetNames.includes('箱单发票')) {
    return parseDanSheetFormat(wb);
  }

  const ws = wb.Sheets['CIPI'];
  if (!ws) throw new Error('未找到 CIPI 工作表，也不支持此文件格式');

  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true });
  if (raw.length < 9) throw new Error('发票数据行数不足');

  // ★ 数据起始行检测：支持两种发票格式 ★
  // 格式A (标准SKU格式): col 0 = SKU 编码 (如 FD-506-TCM)
  // 格式B (Amazon Ref格式): col 0 = Amazon Reference ID (如 5T8NNJWM)
  let dataStart = 0;
  let usesAmazonRef = false;

  // 1. 先尝试标准 SKU 模式匹配
  for (let i = 0; i < raw.length; i++) {
    const firstVal = String(raw[i][0] || '').trim();
    if (/^[A-Z]{2,4}\d*-/i.test(firstVal) || /^[A-Z]{2,4}-\d/i.test(firstVal)) {
      dataStart = i;
      break;
    }
  }

  // 2. SKU 未匹配，尝试通过表头行定位（Amazon Ref 格式）
  if (dataStart === 0) {
    for (let i = 0; i < Math.min(raw.length, 20); i++) {
      const col0 = String(raw[i][0] || '').trim();
      const col1 = String(raw[i][1] || '').trim();
      // 检测表头行：col 0 含 "Amazon reference"，col 1 含 "货箱编号"
      if ((col0.includes('Amazon reference') || col0.includes('亚马逊')) &&
          (col1.includes('货箱编号') || col1.includes('Amazon Reference'))) {
        dataStart = i + 1;
        usesAmazonRef = true;
        console.log('[解析] 检测到 Amazon Ref 格式发票，从第 ' + dataStart + ' 行开始解析');
        break;
      }
    }
  }

  if (dataStart === 0) {
    throw new Error('未找到数据起始行。请确认：\n' +
      '1) 文件是否为 CIPI 格式的发票模板\n' +
      '2) 第一列是否包含 SKU 编码（如 FD-xxx）或 Amazon Reference ID\n' +
      '3) 如果是 Amazon Ref 格式，请确保表头行包含"Amazon reference ID"和"货箱编号"');
  }

  const boxes = [];
  const skuPrices = new Map(); // SKU → [prices]
  let currentBox = null;
  let prevBoxId = '';

  for (let i = dataStart; i < raw.length; i++) {
    const row = raw[i];
    if (!row) continue;

    // ★ 格式B (Amazon Ref): 只有每箱首行有 Ref ID，其余行为空 → 不能跳过
    // ★ 格式A (SKU):       每行都有 SKU，col 0 为空说明是无效行
    const col0Val = String(row[0] || '').trim();

    // 两格式通用：完全空行（无 col1 箱号）则跳过
    const boxId = row[1] ? String(row[1]).trim() : '';
    const hasProductData = (row[6] && String(row[6]).trim()) || (row[7] && String(row[7]).trim());

    if (!usesAmazonRef && !col0Val) continue;           // 格式A: col 0 为空 → 跳过
    if (usesAmazonRef && !boxId && !hasProductData) continue; // 格式B: 无箱号且无品名 → 跳过

    // ---- 提取 SKU (格式A) 或 Amazon Ref ID (格式B) ----
    let sku = '';
    if (usesAmazonRef) {
      // 格式B: 用 Amazon Ref ID + 英文名生成唯一标识
      const enName = row[6] ? String(row[6]).trim() : '';
      if (col0Val) {
        sku = col0Val; // Amazon Reference ID 作为临时 SKU
      } else {
        // 同行无 Ref ID，沿用上一个箱子的同产品 SKU（通过英文名匹配）
        // 在汇总阶段会按英文名+HS编码归并
        sku = 'AMZN-' + (enName || 'UNKNOWN');
      }
    } else {
      sku = col0Val;
    }

    // 因为合并单元格，每一行都有 boxId；通过比较前一行来判断是否新箱
    const isNewBox = boxId !== '' && boxId !== prevBoxId;
    if (boxId !== '') prevBoxId = boxId;

    // 提取中文名和英文名 (cols 6-7)
    const enName = row[6] ? String(row[6]).trim() : '';
    const cnName = row[7] ? String(row[7]).trim() : '';

    // 智能查找HS编码: 在 row[8]~row[10] 中找4-10位纯数字
    let hsCode = 0;
    let hsIdx = -1;
    for (let j = 8; j < Math.min(row.length, 12); j++) {
      const val = row[j];
      if (typeof val === 'number' && val >= 1000000 && val <= 9999999999) {
        hsCode = val;
        hsIdx = j;
        break;
      }
      // ★ 兼容 HS 编码含空格格式（如 "382499 9999"）
      if (typeof val === 'string' && /^\d{4,10}\s*\d{0,6}$/.test(val.trim())) {
        const cleaned = parseInt(val.replace(/\s/g, ''));
        if (cleaned >= 1000000) { hsCode = cleaned; hsIdx = j; break; }
      }
    }

    // 材质和用途在 HS 编码之后
    let material = '';
    let usage = '家居用品';
    if (hsIdx > 0) {
      for (let j = hsIdx + 1; j < Math.min(row.length, hsIdx + 3); j++) {
        const val = row[j];
        if (typeof val === 'string') {
          if (!material) material = val.trim();
          else usage = val.trim();
        }
      }
    }

    // ★ 直接按发票模板列位读取数量和单价 ★
    // 模板 R9 表头: col 13=产品数量(单箱), col 15=申报单价(EUR)
    // 不再使用"猜尾数"的启发式逻辑
    const qtyPerBox = (typeof row[13] === 'number' && row[13] > 0) ? row[13] : 1;
    const unitPriceEUR = (typeof row[15] === 'number' && row[15] > 0) ? row[15] : 0;

    // 收集SKU价格（格式B使用英文名+HS编码归并）
    const priceKey = usesAmazonRef ? (enName + '|' + hsCode) : sku;
    if (unitPriceEUR > 0) {
      if (!skuPrices.has(priceKey)) skuPrices.set(priceKey, []);
      skuPrices.get(priceKey).push(unitPriceEUR);
    }

    // 产品条目
    const productEntry = {
      sku: sku,
      nameCN: cnName,
      nameEN: enName,
      hsCode: hsCode,
      material: material,
      usage: usage,
      unit: '个',
      qtyPerBox: qtyPerBox,
      unitPriceEUR: unitPriceEUR,
      _priceKey: priceKey,  // 用于汇总归并
    };

    // 新建或切换箱子
    if (isNewBox || !currentBox) {
      // 保存上一个箱子
      if (currentBox && currentBox.products.length > 0) {
        boxes.push(currentBox);
      }
      currentBox = {
        id: boxId,
        boxSeq: boxes.length + 1,
        length: parseFloat(row[2]) || 56,
        width: parseFloat(row[3]) || 44,
        height: parseFloat(row[4]) || 42,
        weight: parseFloat(row[5]) || 0,
        products: [],
      };
    }

    // 将产品加入当前箱
    if (currentBox) {
      currentBox.products.push(productEntry);
    }
  }

  // 保存最后一个箱子
  if (currentBox && currentBox.products.length > 0) {
    boxes.push(currentBox);
  }

  // 按SKU（或格式B的品名+HS编码）汇总产品
  const productMap = new Map();
  boxes.forEach(box => {
    box.products.forEach(prod => {
      // 格式B 使用英文名+HS编码归并（不同Amazon Ref可能是同一产品）
      const key = usesAmazonRef ? (prod.nameEN + '|' + prod.hsCode) : prod.sku;
      if (!productMap.has(key)) {
        productMap.set(key, {
          sku: usesAmazonRef ? ('AMZN-' + (prod.nameEN || 'ITEM')) : prod.sku,
          nameCN: prod.nameCN,
          nameEN: prod.nameEN,
          hsCode: prod.hsCode,
          material: prod.material,
          usage: prod.usage,
          unit: prod.unit,
          totalQty: 0,
          netWeightEstimate: 0,
          bestPriceEUR: 0,
          netWeight: 0,
          _mergeKey: key,
        });
      }
      const agg = productMap.get(key);
      agg.totalQty += prod.qtyPerBox;
      // 如果已存 SKU 是 AMZN- 前缀，尝试用新的非 AMZN Ref ID 替换
      if (usesAmazonRef && prod.sku && !prod.sku.startsWith('AMZN-')) {
        agg.sku = prod.sku;
      }
    });
  });

  // 为每个产品解析最佳价格
  const products = [];
  productMap.forEach((agg) => {
    const priceKey = usesAmazonRef ? agg._mergeKey : agg.sku;
    const prices = skuPrices.get(priceKey) || [];
    let bestPriceEUR = 0;
    if (prices.length > 0) {
      bestPriceEUR = Math.min(...prices);
    }
    agg.bestPriceEUR = bestPriceEUR;
    // 估算净重 = 总数量 * 0.5kg/个
    agg.netWeightEstimate = agg.totalQty * 0.5;
    agg.netWeight = agg.netWeightEstimate;
    delete agg._mergeKey; // 清理内部字段
    products.push(agg);
  });

  if (usesAmazonRef) {
    console.log('[解析] Amazon Ref 格式归并: ' + boxes.length + ' 箱, ' + products.length + ' 种产品');
    console.log('[解析] ⚠ 注意: Amazon Ref 格式不含供应商SKU，产品数据库匹配将失效');
    console.log('[解析] ⚠ 请在人工审查步骤中手动补充进价和货源地信息');
  }

  return {
    products: products,
    boxes: boxes,
    totalBoxes: boxes.length,
  };
}

/**
 * 解析智谷供应链格式的发票
 * 工作表「箱单」含箱级产品明细。
 *
 * ── V1 原始格式 ──
 *   col 0: SKU(Amazon Ref)  col 2: 材积(长*宽*高)   col 3: 箱号(数字)
 *   col 4: 单箱毛重          col 5: HS编码            col 6: 中文品名
 *   col 7: 英文品名          col 8: 单箱数量           col 10: 单价(GBP/USD)
 *   表头: col 0 = "FBA编号/纸箱单号"
 *   箱号列为数字时表示新箱；为空表示属于上一箱。
 *
 * ── V2 变体 ──
 *   col 0: SKU(Amazon Ref)  col 1: 箱号(字符串)      col 2: 材积
 *   col 4: 单箱毛重          col 5: HS编码            col 6: 中文品名
 *   col 7: 英文品名          col 8: 单箱数量           col 10: 单价(GBP)
 *   表头: col 0 = "Amazon reference ID" + col 1 = "FBA编号/纸箱单号"
 *   箱号列(col 1)值变化时表示新箱。
 *
 * ── V3 变体 ──
 *   col 0: 箱号(字符串)      col 1: SKU(Reference ID)  col 2: 材积
 *   col 4: 单箱毛重          col 5: HS编码            col 6: 中文品名
 *   col 7: 英文品名          col 8: 单箱数量           col 10: 单价(GBP)
 *   表头: col 0 = "FBA编号/纸箱单号" + col 1 = "Reference ID（FBA卡派必填）"
 *   箱号列(col 0)值变化时表示新箱。
 */
function parseZhiguFormat(wb) {
  const XLSX = require('xlsx');
  const ws = wb.Sheets['箱单'];
  if (!ws) throw new Error('未找到箱单工作表');

  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true });
  if (raw.length < 12) throw new Error('箱单数据行数不足');

  // ★ 检测 V1 / V2 / V3 变体
  let headerRow = -1;
  let variant = 1; // 1=V1(col3箱号), 2=V2(col1箱号), 3=V3(col0箱号)

  for (let i = 0; i < Math.min(raw.length, 20); i++) {
    const col0 = String(raw[i][0] || '').trim();
    const col1 = String(raw[i][1] || '').trim();
    const col5 = String(raw[i][5] || '').trim();

    // ★ 检测顺序重要：V2/V3 比 V1 更具体，必须先匹配
    // V2: col 0 = "Amazon reference ID", col 1 = "FBA编号/纸箱单号"
    if ((col0.includes('Amazon reference') || col0.includes('亚马逊')) &&
        (col1.includes('FBA编号') || col1.includes('纸箱单号'))) {
      headerRow = i;
      variant = 2;
      break;
    }
    // V3: col 0 = "FBA编号/纸箱单号", col 1 = "Reference ID" (非FBA前缀)
    if ((col0 === 'FBA编号/纸箱单号' || col0.includes('FBA编号/纸箱单号')) &&
        (col1.includes('Reference') || col1.includes('ID'))) {
      headerRow = i;
      variant = 3;
      break;
    }
    // V1: col 0 = "FBA编号/纸箱单号"（只有这一列是表头关键词，col1为空或非Reference）
    if ((col0 === 'FBA编号/纸箱单号' || col0.includes('FBA编号')) && col5.includes('HS')) {
      headerRow = i;
      variant = 1;
      break;
    }
  }
  if (headerRow < 0) throw new Error('未找到箱单表头行（支持V1/V2/V3智谷格式）');

  const dataStart = headerRow + 1;
  const boxes = [];
  const skuPrices = new Map();

  let currentBox = null;
  let prevBoxId = '';
  let boxSeq = 0;

  for (let i = dataStart; i < raw.length; i++) {
    const row = raw[i];
    if (!row) continue;

    // ★ 从正确列提取 SKU
    let sku = '';
    let boxId = '';
    let isNewBox = false;

    if (variant === 3) {
      // V3: col 0 = 箱号, col 1 = SKU
      boxId = String(row[0] || '').trim();
      sku = String(row[1] || '').trim();
      isNewBox = (boxId !== '' && boxId !== prevBoxId);
    } else if (variant === 2) {
      // V2: col 0 = SKU, col 1 = 箱号
      sku = String(row[0] || '').trim();
      boxId = String(row[1] || '').trim();
      isNewBox = (boxId !== '' && boxId !== prevBoxId);
    } else {
      // V1: col 0 = SKU, col 3 = 箱号(数字)
      sku = String(row[0] || '').trim();
      const boxNum = row[3];
      isNewBox = (typeof boxNum === 'number' && boxNum > 0);
      if (isNewBox) boxId = 'BOX' + (boxes.length + 1);
    }

    if (!sku) continue;

    // 跳过非数据行
    const metaKeywords = ['Shipper', 'Company', 'Address', 'City', 'Postal', '额外', '备注', 'Consignee'];
    if (metaKeywords.some(kw => sku.startsWith(kw))) continue;

    if (isNewBox) {
      if (boxId !== '') prevBoxId = boxId;

      if (currentBox && currentBox.products.length > 0) {
        boxes.push(currentBox);
      }
      boxSeq++;

      // 解析材积 (所有变体都在 col 2)
      let l = 56, w = 44, h = 42;
      const dimStr = String(row[2] || '56*44*42').trim();
      const dimMatch = dimStr.match(/(\d+)\s*\*\s*(\d+)\s*\*\s*(\d+)/);
      if (dimMatch) {
        l = parseInt(dimMatch[1]) || 56;
        w = parseInt(dimMatch[2]) || 44;
        h = parseInt(dimMatch[3]) || 42;
      }

      const weight = parseFloat(row[4]) || 0;

      currentBox = {
        id: (variant === 1) ? ('BOX' + boxSeq) : boxId,
        boxSeq: boxSeq,
        length: l, width: w, height: h,
        weight: weight,
        products: [],
      };
    }

    if (!currentBox) {
      boxSeq++;
      const boxDefault = variant === 1 ? ('BOX' + boxSeq) : (boxId || 'BOX1');
      currentBox = {
        id: boxDefault, boxSeq: boxSeq,
        length: 56, width: 44, height: 42, weight: parseFloat(row[4]) || 0,
        products: [],
      };
      if (variant !== 1 && boxId) prevBoxId = boxId;
    }

    // —— 产品数据 (所有变体 col 5-10 列位相同) ——
    const enName = String(row[7] || '').trim();
    const cnName = String(row[6] || '').trim();

    // HS编码 (col 5), 兼容 "382499 9999" 格式
    const hsRaw = row[5];
    let hsCode = 0;
    if (typeof hsRaw === 'number') {
      hsCode = Math.floor(hsRaw);
    } else if (hsRaw) {
      const hsClean = String(hsRaw).replace(/[^\d]/g, '');
      hsCode = parseInt(hsClean) || 0;
    }

    const qtyPerBox = (typeof row[8] === 'number' && row[8] > 0) ? row[8] : 1;
    const unitPrice = (typeof row[10] === 'number' && row[10] > 0) ? row[10] : 0;

    if (unitPrice > 0) {
      if (!skuPrices.has(sku)) skuPrices.set(sku, []);
      skuPrices.get(sku).push(unitPrice);
    }

    currentBox.products.push({
      sku, nameCN: cnName, nameEN: enName,
      hsCode, material: '', usage: '家居用品',
      unit: '个', qtyPerBox, unitPriceEUR: unitPrice,
    });
  }

  // 最后一个箱子
  if (currentBox && currentBox.products.length > 0) {
    boxes.push(currentBox);
  }

  // 按SKU汇总
  const productMap = new Map();
  boxes.forEach(box => {
    box.products.forEach(prod => {
      const key = prod.sku;
      if (!productMap.has(key)) {
        productMap.set(key, {
          sku: prod.sku,
          nameCN: prod.nameCN, nameEN: prod.nameEN,
          hsCode: prod.hsCode,
          material: prod.material, usage: prod.usage,
          unit: prod.unit,
          totalQty: 0,
          netWeightEstimate: 0,
          bestPriceEUR: 0,
          netWeight: 0,
        });
      }
      const agg = productMap.get(key);
      agg.totalQty += prod.qtyPerBox;
    });
  });

  const products = [];
  productMap.forEach(agg => {
    const prices = skuPrices.get(agg.sku) || [];
    agg.bestPriceEUR = prices.length > 0 ? Math.min(...prices) : 0;
    agg.netWeightEstimate = agg.totalQty * 0.5;
    agg.netWeight = agg.netWeightEstimate;
    products.push(agg);
  });

  console.log('[智谷格式V' + variant + '] 解析完成: ' + products.length + ' 种产品, ' + boxes.length + ' 箱');
  return { products, boxes, totalBoxes: boxes.length };
}

/**
 * 单 sheet「箱单发票」格式自动检测 + 分发
 * 支持两种子格式：
 *  - 凡洋格式: col 0 = 箱号（非空）
 *  - 顺沃格式: col 0 = 永远为空，按 SKU 重复识别箱边界
 */
function parseDanSheetFormat(wb) {
  const XLSX = require('xlsx');
  const ws = wb.Sheets['箱单发票'];
  if (!ws) throw new Error('未找到箱单发票工作表');

  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true });

  // 查找表头行
  let headerRow = -1;
  for (let i = 0; i < Math.min(raw.length, 15); i++) {
    const col0 = String(raw[i][0] || '').trim();
    if (col0.includes('唛头') || col0.includes('MARKS')) { headerRow = i; break; }
  }
  if (headerRow < 0) throw new Error('未找到箱单发票表头行');

  // ★ 检测子格式：检查前 30 行数据中 col 0 是否全为空
  let isShunwo = true;
  let checked = 0;
  for (let i = headerRow + 1; i < Math.min(raw.length, headerRow + 31); i++) {
    const row = raw[i];
    if (!row) continue;
    const c0 = String(row[0] || '').trim();
    const c1 = String(row[1] || '').trim();
    if (!c1) continue; // 没有 SKU 的行跳过
    checked++;
    if (c0 !== '') { isShunwo = false; break; }
  }

  if (isShunwo && checked > 0) {
    console.log('[检测] 箱单发票 → 顺沃格式 (col 0 为空)');
    return parseShunwoFormat(raw, headerRow);
  }
  console.log('[检测] 箱单发票 → 凡洋格式 (col 0 有箱号)');
  return parseFanyangUKFormat(wb);
}

/**
 * 解析凡洋英国格式的发票（单 sheet「箱单发票」）
 *
 * 表头 (Row 7):
 *   col 0: 唛头(箱号)   col 1: Amazon Ref    col 2: 英文品名      col 3: 中文品名
 *   col 4: 品牌         col 5: 品牌类型       col 6: 型号          col 7: 材质
 *   col 8: 用途         col 9: 单箱数量       col 10: 单位         col 11: 净重(首行)
 *   col 12: 毛重(首行)  col 13: 箱规           col 15: 件数CTN     col 16: 单价
 *   col 18: HS编码
 *
 * 特点：无 SKU 编码，按英文品名+HS编码归并产品
 */
function parseFanyangUKFormat(wb) {
  const XLSX = require('xlsx');
  const ws = wb.Sheets['箱单发票'];
  if (!ws) throw new Error('未找到箱单发票工作表');

  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true });
  if (raw.length < 9) throw new Error('箱单发票数据行数不足');

  // 查找表头行：col 0 含 "唛头" 且 col 18 含 "海关编码"
  let headerRow = -1;
  for (let i = 0; i < Math.min(raw.length, 15); i++) {
    const col0 = String(raw[i][0] || '').trim();
    const col18 = String(raw[i][18] || '').trim();
    if (col0.includes('唛头') || (col0.includes('MARKS') && col18.includes('海关'))) {
      headerRow = i; break;
    }
  }
  if (headerRow < 0) throw new Error('未找到箱单发票表头行（需要 col 0 含"唛头"）');

  const dataStart = headerRow + 1;
  const boxes = [];
  const skuPrices = new Map(); // enName|hsCode → [prices]
  let currentBox = null;
  let prevBoxId = '';

  for (let i = dataStart; i < raw.length; i++) {
    const row = raw[i];
    if (!row) continue;

    // col 0 = 箱号（合并单元格，仅首行有值）
    const boxId = String(row[0] || '').trim();
    // col 2/3 = 品名
    const enName = String(row[2] || '').trim();
    const cnName = String(row[3] || '').trim();

    // 空行判断：无箱号且无产品名 → 跳过
    const isNewBox = boxId !== '' && boxId !== prevBoxId;
    if (boxId !== '') prevBoxId = boxId;
    if (!isNewBox && !boxId && !enName) continue;

    // 新建或切换箱子
    if (isNewBox) {
      if (currentBox && currentBox.products.length > 0) boxes.push(currentBox);

      // 解析箱规 "56*44*42"
      let l = 56, w = 44, h = 42;
      const dimStr = String(row[13] || '56*44*42').trim();
      const dimMatch = dimStr.match(/(\d+)\s*\*\s*(\d+)\s*\*\s*(\d+)/);
      if (dimMatch) {
        l = parseInt(dimMatch[1]) || 56;
        w = parseInt(dimMatch[2]) || 44;
        h = parseInt(dimMatch[3]) || 42;
      }

      // 箱重 = 净重 (col 11, 仅首行) + 毛重 (col 12, 仅首行)
      // 报关系统用毛重
      const grossWeight = parseFloat(row[12]) || parseFloat(row[11]) || 0;

      currentBox = {
        id: boxId, boxSeq: boxes.length + 1,
        length: l, width: w, height: h,
        weight: grossWeight,
        products: [],
      };
    }

    if (!currentBox) {
      // 第一行没有新箱号时创建默认箱
      currentBox = {
        id: 'BOX1', boxSeq: 1,
        length: 56, width: 44, height: 42,
        weight: parseFloat(row[12]) || parseFloat(row[11]) || 0,
        products: [],
      };
    }

    // —— 产品数据提取 ——
    const material = String(row[7] || '').trim();
    const usage = String(row[8] || '').trim() || '家居用品';

    // HS编码 (col 18)，兼容 "382499 9999" 带空格格式
    const hsRaw = row[18];
    let hsCode = 0;
    if (typeof hsRaw === 'number') {
      hsCode = Math.floor(hsRaw);
    } else if (hsRaw) {
      const hsClean = String(hsRaw).replace(/[^\d]/g, '');
      hsCode = parseInt(hsClean) || 0;
    }

    // 数量 (col 9 / 15)：箱首行的 col 15 通常有箱数
    // col 9 = 单箱产品数量, 非首行不重复
    const qtyPerBox = (typeof row[9] === 'number' && row[9] > 0) ? row[9]
      : (typeof row[15] === 'number' && row[15] > 0) ? row[15] : 1;

    // 单价 (col 16)
    const unitPrice = (typeof row[16] === 'number' && row[16] > 0) ? row[16] : 0;

    // 产品归并 key = 英文名|HS编码（无 SKU）
    const mergeKey = enName + '|' + hsCode;
    if (unitPrice > 0) {
      if (!skuPrices.has(mergeKey)) skuPrices.set(mergeKey, []);
      skuPrices.get(mergeKey).push(unitPrice);
    }

    currentBox.products.push({
      sku: 'FY-' + (enName || 'ITEM'), // 凡洋格式无 SKU → 用英文名生成
      nameCN: cnName, nameEN: enName,
      hsCode, material, usage,
      unit: String(row[10] || '').trim() || '个',
      qtyPerBox, unitPriceEUR: unitPrice,
      _priceKey: mergeKey,
    });
  }

  // 最后一个箱子
  if (currentBox && currentBox.products.length > 0) boxes.push(currentBox);

  // 按英文名+HS编码归并产品
  const productMap = new Map();
  boxes.forEach(box => {
    box.products.forEach(prod => {
      const key = prod._priceKey || (prod.nameEN + '|' + prod.hsCode);
      if (!productMap.has(key)) {
        productMap.set(key, {
          sku: prod.sku, nameCN: prod.nameCN, nameEN: prod.nameEN,
          hsCode: prod.hsCode, material: prod.material, usage: prod.usage,
          unit: prod.unit, totalQty: 0,
          netWeightEstimate: 0, bestPriceEUR: 0, netWeight: 0,
          _mergeKey: key,
        });
      }
      productMap.get(key).totalQty += prod.qtyPerBox;
    });
  });

  const products = [];
  productMap.forEach(agg => {
    const prices = skuPrices.get(agg._mergeKey) || [];
    agg.bestPriceEUR = prices.length > 0 ? Math.min(...prices) : 0;
    agg.netWeightEstimate = agg.totalQty * 0.5;
    agg.netWeight = agg.netWeightEstimate;
    delete agg._mergeKey;
    products.push(agg);
  });

  console.log('[凡洋英国格式] 解析完成: ' + products.length + ' 种产品, ' + boxes.length + ' 箱');
  console.log('[凡洋英国格式] ⚠ 此格式不含SKU编码，产品数据库匹配将失效');
  console.log('[凡洋英国格式] ⚠ 请在人工审查步骤手动补充进价和货源地');

  return { products, boxes, totalBoxes: boxes.length };
}

/**
 * 解析顺沃格式的发票（单 sheet「箱单发票」）
 *
 * 与凡洋格式的区别：
 *  - col 0 永远为空（无箱号标识）
 *  - col 1 = SKU 编码（如 SWK603-3004153J）
 *  - col 2 = 英文品名  col 3 = 中文品名
 *  - col 7 = 材质       col 8 = 用途
 *  - col 9 = 单箱数量   col 11 = 净重(仅每箱首行有值)
 *  - col 13 = 箱规      col 16 = 单价
 *  - col 18 = HS编码
 *
 * 箱边界检测：col 11（净重）仅每箱首行有值 → 遇到非空净重即开启新箱。
 *
 * @param {Array} raw - 已解析的原始行数据
 * @param {number} headerRow - 表头行号
 */
function parseShunwoFormat(raw, headerRow) {
  const XLSX = require('xlsx');
  const dataStart = headerRow + 1;
  const boxes = [];
  const skuPrices = new Map();
  let currentBox = null;
  let boxSeq = 0;

  for (let i = dataStart; i < raw.length; i++) {
    const row = raw[i];
    if (!row) continue;

    // col 1 = SKU (顺沃格式)
    const sku = String(row[1] || '').trim();
    if (!sku) continue;

    const enName = String(row[2] || '').trim();
    const cnName = String(row[3] || '').trim();

    // ★ 箱边界检测：col 11（净重）非空 → 新箱首行
    const rowNetWeight = parseFloat(row[11]);
    const isNewBox = (!isNaN(rowNetWeight) && rowNetWeight > 0);

    if (isNewBox) {
      // 保存上一个箱子
      if (currentBox && currentBox.products.length > 0) {
        boxes.push(currentBox);
      }
      boxSeq++;
      currentBox = null; // 下面会创建新的
    }

    if (!currentBox) {
      // 解析箱规 col 13
      let l = 56, w = 44, h = 42;
      const dimStr = String(row[13] || '56*44*42').trim();
      const dimMatch = dimStr.match(/(\d+)\s*\*\s*(\d+)\s*\*\s*(\d+)/);
      if (dimMatch) {
        l = parseInt(dimMatch[1]) || 56;
        w = parseInt(dimMatch[2]) || 44;
        h = parseInt(dimMatch[3]) || 42;
      }

      currentBox = {
        id: 'BOX' + boxSeq,
        boxSeq: boxSeq,
        length: l, width: w, height: h,
        weight: isNaN(rowNetWeight) ? 0 : rowNetWeight,
        products: [],
      };
    }

    // —— 产品数据 ——
    const material = String(row[7] || '').trim();
    const usage = String(row[8] || '').trim() || '家居用品';

    // HS编码 col 18
    const hsRaw = row[18];
    let hsCode = 0;
    if (typeof hsRaw === 'number') {
      hsCode = Math.floor(hsRaw);
    } else if (hsRaw) {
      const hsClean = String(hsRaw).replace(/[^\d]/g, '');
      hsCode = parseInt(hsClean) || 0;
    }

    // 单箱数量 col 9
    const qtyPerBox = (typeof row[9] === 'number' && row[9] > 0) ? row[9] : 1;

    // 单价 col 16
    const unitPrice = (typeof row[16] === 'number' && row[16] > 0) ? row[16] : 0;

    // 收集价格
    if (unitPrice > 0) {
      if (!skuPrices.has(sku)) skuPrices.set(sku, []);
      skuPrices.get(sku).push(unitPrice);
    }

    currentBox.products.push({
      sku, nameCN: cnName, nameEN: enName,
      hsCode, material, usage,
      unit: String(row[10] || '').trim() || '个',
      qtyPerBox, unitPriceEUR: unitPrice,
    });
  }

  // 最后一个箱子
  if (currentBox && currentBox.products.length > 0) {
    boxes.push(currentBox);
  }

  // 按SKU汇总
  const productMap = new Map();
  boxes.forEach(box => {
    box.products.forEach(prod => {
      if (!productMap.has(prod.sku)) {
        productMap.set(prod.sku, {
          sku: prod.sku, nameCN: prod.nameCN, nameEN: prod.nameEN,
          hsCode: prod.hsCode, material: prod.material, usage: prod.usage,
          unit: prod.unit, totalQty: 0,
          netWeightEstimate: 0, bestPriceEUR: 0, netWeight: 0,
        });
      }
      productMap.get(prod.sku).totalQty += prod.qtyPerBox;
    });
  });

  const products = [];
  productMap.forEach(agg => {
    const prices = skuPrices.get(agg.sku) || [];
    agg.bestPriceEUR = prices.length > 0 ? Math.min(...prices) : 0;
    agg.netWeightEstimate = agg.totalQty * 0.5;
    agg.netWeight = agg.netWeightEstimate;
    products.push(agg);
  });

  console.log('[顺沃格式] 解析完成: ' + products.length + ' 种产品, ' + boxes.length + ' 箱');
  return { products, boxes, totalBoxes: boxes.length };
}

// ==================== 数据一致性核验 ====================
/**
 * 核验报关草单、装箱单、报关草单合并之间的毛重/净重一致性。
 * 以箱级数据为基准，修正产品级汇总数据。
 * 最多尝试 10 轮修正；若仍未通过则标出可疑数据。
 *
 * @param {Array} products - genProducts
 * @param {Array} boxes - genBoxes
 * @returns {{ products, boxes, report }}
 */
function validateConsistency(products, boxes) {
  const report = { passed: true, rounds: 0, fixes: [], warnings: [] };
  let currentProducts = products;
  let currentBoxes = boxes;
  const MAX_ROUNDS = 10;

  for (let round = 1; round <= MAX_ROUNDS; round++) {
    report.rounds = round;
    let fixedThisRound = false;

    // ---- 1. 核验净重：箱级汇总 vs 产品级汇总 ----
    // 从箱数据按 SKU 汇总净重
    const netFromBoxes = {};
    currentBoxes.forEach(box => {
      box.products.forEach(bp => {
        if (!netFromBoxes[bp.sku]) netFromBoxes[bp.sku] = 0;
        netFromBoxes[bp.sku] += (bp.netWeightPerBox || bp.netWeight || 0);
      });
    });

    // 逐产品比对
    currentProducts.forEach(prod => {
      const boxNet = netFromBoxes[prod.sku];
      if (boxNet === undefined) {
        report.warnings.push('SKU ' + prod.sku + ' 在箱数据中未找到');
        return;
      }
      const diff = Math.abs(boxNet - (prod.netWeight || 0));
      if (diff > 0.005) {
        const oldVal = prod.netWeight;
        prod.netWeight = parseFloat(boxNet.toFixed(3));
        report.fixes.push(
          '净重修正 [' + prod.sku + ']: ' + oldVal + ' → ' + prod.netWeight + ' (来源: 箱级汇总)'
        );
        fixedThisRound = true;
      }
    });

    // ---- 2. 核验每箱内 netWeightPerBox 与 netWeight 一致性 ----
    currentBoxes.forEach(box => {
      box.products.forEach(bp => {
        if (Math.abs((bp.netWeightPerBox || 0) - (bp.netWeight || 0)) > 0.005) {
          bp.netWeight = bp.netWeightPerBox;
          fixedThisRound = true;
        }
      });
    });

    // ---- 3. 核验总净重：箱行级求和 vs 产品级求和 ----
    const totalNetFromBoxRows = currentBoxes.reduce(
      (s, b) => s + b.products.reduce((ss, bp) => ss + (bp.netWeightPerBox || 0), 0), 0
    );
    const totalNetFromProducts = currentProducts.reduce((s, p) => s + (p.netWeight || 0), 0);
    if (Math.abs(totalNetFromBoxRows - totalNetFromProducts) > 0.01) {
      // 强制以箱数据为准
      const fixedProducts = currentProducts.map(p => {
        const bNet = netFromBoxes[p.sku];
        if (bNet !== undefined) p.netWeight = parseFloat(bNet.toFixed(3));
        return p;
      });
      currentProducts = fixedProducts;
      report.fixes.push(
        '总净重修正: 产品级=' + totalNetFromProducts.toFixed(2) +
        ' → 箱级汇总=' + totalNetFromBoxRows.toFixed(2)
      );
      fixedThisRound = true;
    }

    // ---- 4. 核验毛重（箱重）：确保 totalGrossWeight 源于 boxes ----
    const totalGrossFromBoxes = currentBoxes.reduce((s, b) => s + (b.weight || 0), 0);
    // 毛重不存在独立的产品级计算，仅记录供调试
    // 但检查是否有异常值（单箱毛重为 0 或超大）
    currentBoxes.forEach(box => {
      if (!box.weight || box.weight <= 0) {
        report.warnings.push('箱 ' + box.id + ' 毛重异常: ' + box.weight + 'kg');
      }
      if (box.weight > 500) {
        report.warnings.push('箱 ' + box.id + ' 毛重异常偏大: ' + box.weight + 'kg');
      }
    });

    if (!fixedThisRound) {
      break; // 数据已一致，无需继续
    }
  }

  // 最终判定
  if (report.rounds >= MAX_ROUNDS && report.fixes.length > 0) {
    report.passed = false;
    report.warnings.push('⚠ 经 ' + MAX_ROUNDS + ' 轮核验仍未完全一致，请人工核查上述可疑数据');
  }

  if (report.fixes.length > 0) {
    console.log('[核验] ' + report.rounds + ' 轮, 修正 ' + report.fixes.length + ' 项:');
    report.fixes.forEach(f => console.log('  - ' + f));
  } else {
    console.log('[核验] 数据一致，无需修正');
  }

  return { products: currentProducts, boxes: currentBoxes, report };
}

// ==================== 产品数据库查询 ====================
app.post('/api/lookup-products', async (req, res) => {
  try {
    const { skus } = req.body;
    if (!skus || !Array.isArray(skus) || skus.length === 0) {
      return res.status(400).json({ success: false, message: '请提供SKU列表' });
    }

    const settings = readSettings();
    const exchangeRate = settings.exchangeRate || 7.25;
    const constant = settings.constant || 1.25;

    // 调用产品查询模块
    const result = productLookup.lookupProducts(skus);

    // ★ 人工补充数据优先：覆盖已找到SKU的空值/更低价格
    const manualData = readManualData();
    Object.keys(result.found).forEach(sku => {
      if (manualData[sku]) {
        const md = manualData[sku];
        const item = result.found[sku];
        if (md.maxPrice > item.maxPrice) item.maxPrice = md.maxPrice;
        if (!item.supplier && md.supplier) item.supplier = md.supplier;
        if (!item.city && md.city) { item.city = md.city; item.cityUncertain = false; }
        if (!item.description && md.description) item.description = md.description;
        item.source = item.source || 'manual-override';
      }
    });
    // 对未找到的SKU，尝试从人工补充数据中匹配
    const stillNotFound = [];
    result.notFound.forEach(sku => {
      if (manualData[sku]) {
        const md = manualData[sku];
        result.found[sku] = {
          sku: sku,
          maxPrice: md.maxPrice,
          allPrices: [md.maxPrice],
          supplier: md.supplier || '',
          description: md.description || '',
          recordCount: 1,
          city: md.city || '',
          cityUncertain: false,
          rawSupplier: md.supplier || '',
          source: 'manual',
        };
      } else {
        stillNotFound.push(sku);
      }
    });
    result.notFound = stillNotFound;

    // ★ 自动从供应商名称推断城市（同步：缓存/内置映射/名称模式）
    const uncertainItems = [];
    Object.values(result.found).forEach(item => {
      if ((!item.city || item.cityUncertain) && item.supplier && item.supplier.length >= 3) {
        uncertainItems.push(item);
      }
    });

    if (uncertainItems.length > 0) {
      const suppliers = [...new Set(uncertainItems.map(it => it.supplier))];
      console.log('[产品查询] 自动推断 ' + uncertainItems.length + ' 个不确定货源地, ' + suppliers.length + ' 个供应商');

      // 逐个查询（优先缓存/内置映射，Web搜索设短超时）
      const cityResults = {};
      for (const supplier of suppliers) {
        try {
          const info = await supplierLookup.lookupSupplierCity(supplier);
          if (info) cityResults[supplier] = info;
        } catch (_) {}
      }

      // 将结果应用到产品
      if (Object.keys(cityResults).length > 0) {
        uncertainItems.forEach(item => {
          const info = cityResults[item.supplier];
          if (info) {
            item.city = info.city;
            item.cityUncertain = false;
            item._citySource = info.source;
          }
        });
        console.log('[产品查询] 自动推断结果: ' + Object.keys(cityResults).length + ' 个供应商→城市已填充');
      }
    }

    // 计算单价(USD) = 进价 × 常数 ÷ 汇率
    Object.values(result.found).forEach(item => {
      item.unitPriceUSD = (item.maxPrice * constant) / exchangeRate;
    });

    res.json(result);
  } catch (err) {
    res.status(500).json({ success: false, message: '产品查询失败: ' + err.message });
  }
});

// ==================== 人工补充数据 API ====================
// 获取所有已保存的人工数据
app.get('/api/manual-data', (req, res) => {
  try {
    res.json(readManualData());
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

// ==================== 供应商在线查询 ====================
// POST: 根据供应商名称查询公司所在地（城市）
app.post('/api/lookup-supplier-city', async (req, res) => {
  try {
    const { suppliers } = req.body;
    if (!suppliers || !Array.isArray(suppliers) || suppliers.length === 0) {
      return res.status(400).json({ success: false, message: '请提供供应商名称列表' });
    }

    // 过滤空值
    const names = suppliers.filter(s => s && s.length >= 3);
    if (names.length === 0) {
      return res.json({ success: true, results: {}, notFound: suppliers, errors: [] });
    }

    const { results, notFound } = await supplierLookup.batchLookup(names);
    console.log('[API] 供应商查询: 找到 ' + Object.keys(results).length + '/' + names.length);
    res.json({ success: true, results, notFound });
  } catch (err) {
    res.status(500).json({ success: false, message: '查询失败: ' + err.message });
  }
});

// 保存/更新人工数据（支持批量）
app.post('/api/manual-data', (req, res) => {
  try {
    const entries = req.body;
    if (!entries || typeof entries !== 'object') {
      return res.status(400).json({ success: false, message: '请提供数据' });
    }

    const current = readManualData();
    const now = new Date().toISOString();
    let saved = 0;

    Object.entries(entries).forEach(([sku, info]) => {
      if (!sku) return;
      current[sku] = {
        sku: sku,
        maxPrice: parseFloat(info.maxPrice) || 0,
        supplier: String(info.supplier || ''),
        city: String(info.city || ''),
        netWeightPerUnit: parseFloat(info.netWeightPerUnit) || 0.5,
        description: String(info.description || ''),
        updatedAt: now,
      };
      saved++;
    });

    writeManualData(current);
    console.log('Manual data saved: ' + saved + ' SKUs, total: ' + Object.keys(current).length);
    res.json({ success: true, saved: saved, total: Object.keys(current).length });
  } catch (err) {
    res.status(500).json({ success: false, message: '保存失败: ' + err.message });
  }
});

// 删除单条人工数据
app.delete('/api/manual-data/:sku', (req, res) => {
  try {
    const sku = req.params.sku;
    const current = readManualData();
    if (current[sku]) {
      delete current[sku];
      writeManualData(current);
      res.json({ success: true, deleted: sku });
    } else {
      res.status(404).json({ success: false, message: '未找到: ' + sku });
    }
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

// ==================== 产品数据库管理 ====================

// 上传新的备货单/出货记录，合并到现有数据库
const dbStorage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, UPLOADS_DIR),
  filename: (req, file, cb) => {
    cb(null, 'db_' + safeFilename(file.originalname));
  },
});
const dbUpload = multer({
  storage: dbStorage,
  limits: { fileSize: 200 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (ext === '.xlsx' || ext === '.xls' || ext === '.et') cb(null, true);
    else cb(new Error('仅支持 .xlsx / .xls / .et 格式'));
  },
});

app.post('/api/upload-product-db', (req, res) => {
  dbUpload.single('file')(req, res, (err) => {
    if (err) return res.status(400).json({ success: false, message: err.message });
    if (!req.file) return res.status(400).json({ success: false, message: '请选择文件' });

    try {
      const result = productLookup.mergeFromFile(req.file.path);
      res.json({
        success: true,
        message: '数据库已更新',
        ...result,
      });
    } catch (e) {
      res.status(500).json({ success: false, message: '数据库更新失败: ' + e.message });
    }
  });
});

// 获取当前数据库状态
app.get('/api/product-db-status', (req, res) => {
  try {
    const status = productLookup.getStatus();
    res.json(status);
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

// 手动重载数据库（指定文件路径）
app.post('/api/reload-db', (req, res) => {
  try {
    const { filePath } = req.body || {};
    productLookup.setDBPath(filePath);
    res.json({ success: true, message: '数据库已重新加载' });
  } catch (err) {
    res.status(500).json({ success: false, message: '重载失败: ' + err.message });
  }
});

// ==================== 确认数据持久化 ====================
function readConfirmations() {
  try {
    if (!fs.existsSync(CONFIRMATIONS_FILE)) return {};
    return JSON.parse(fs.readFileSync(CONFIRMATIONS_FILE, 'utf-8'));
  } catch (e) { return {}; }
}
function writeConfirmations(data) {
  if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
  fs.writeFileSync(CONFIRMATIONS_FILE, JSON.stringify(data, null, 2), 'utf-8');
}

// ==================== 人工审查覆盖数据持久化 ====================
function readReviewOverrides() {
  try {
    if (!fs.existsSync(REVIEW_OVERRIDES_FILE)) return null;
    return JSON.parse(fs.readFileSync(REVIEW_OVERRIDES_FILE, 'utf-8'));
  } catch (e) { return null; }
}
function writeReviewOverrides(data) {
  if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
  fs.writeFileSync(REVIEW_OVERRIDES_FILE, JSON.stringify(data, null, 2), 'utf-8');
}

// GET: 获取上一次人工审查的覆盖数据（用于"是否延用"提示）
app.get('/api/review-overrides', (req, res) => {
  try {
    const data = readReviewOverrides();
    if (data && data.overrides && Object.keys(data.overrides).length > 0) {
      res.json({ success: true, hasOverrides: true, overrides: data.overrides, savedAt: data.savedAt });
    } else {
      res.json({ success: true, hasOverrides: false });
    }
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

// POST: 保存人工审查覆盖数据
app.post('/api/review-overrides', (req, res) => {
  try {
    const { overrides } = req.body;
    if (!overrides || typeof overrides !== 'object') {
      return res.status(400).json({ success: false, message: '请提供覆盖数据' });
    }
    writeReviewOverrides({ overrides, savedAt: new Date().toISOString() });
    res.json({ success: true, saved: Object.keys(overrides).length });
  } catch (err) {
    res.status(500).json({ success: false, message: '保存失败: ' + err.message });
  }
});

// ==================== 生成前数据核验 ====================
app.post('/api/preflight-check', async (req, res) => {
  try {
    const { skus } = req.body;
    if (!skus || !Array.isArray(skus) || skus.length === 0) {
      return res.status(400).json({ success: false, message: '请提供SKU列表' });
    }

    const settings = readSettings();
    const exchangeRate = settings.exchangeRate || 7.25;
    const constant = settings.constant || 1.25;

    // 查询产品数据库
    const lookupResult = productLookup.lookupProducts(skus);

    // 人工补充数据
    const manualData = readManualData();
    lookupResult.notFound.forEach(sku => {
      if (manualData[sku]) {
        const md = manualData[sku];
        lookupResult.found[sku] = {
          sku, maxPrice: md.maxPrice, allPrices: [md.maxPrice],
          supplier: md.supplier || '', description: md.description || '',
          recordCount: 1, city: md.city || '', cityUncertain: false,
          rawSupplier: md.supplier || '', source: 'manual',
        };
      }
    });

    // ★ 自动从供应商名称推断城市（同步：缓存/内置映射/名称模式）
    const fuzzyItems = [];
    Object.values(lookupResult.found).forEach(item => {
      if ((!item.city || item.cityUncertain) && item.supplier && item.supplier.length >= 3) {
        fuzzyItems.push(item);
      }
    });
    if (fuzzyItems.length > 0) {
      const suppliers = [...new Set(fuzzyItems.map(it => it.supplier))];
      for (const supplier of suppliers) {
        try {
          const info = await supplierLookup.lookupSupplierCity(supplier);
          if (info) {
            fuzzyItems.filter(it => it.supplier === supplier).forEach(it => {
              it.city = info.city;
              it.cityUncertain = false;
            });
          }
        } catch (_) {}
      }
    }

    // 历史确认数据
    const confirmations = readConfirmations();

    // 读取发票数据
    const sessionFile = path.join(UPLOADS_DIR, 'last_parsed.json');
    const invoiceData = fs.existsSync(sessionFile)
      ? JSON.parse(fs.readFileSync(sessionFile, 'utf-8'))
      : { products: [] };

    // 构建SKU→中文名映射
    const skuNames = {};
    (invoiceData.products || []).forEach(p => { skuNames[p.sku] = p.nameCN; });

    const uncertainItems = [];

    skus.forEach(sku => {
      const found = lookupResult.found[sku];
      const nameCN = skuNames[sku] || sku;
      const saved = confirmations[sku] || null;

      // 情况1：数据库和人工数据都找不到 → 需要补全进价+供应商+货源地
      if (!found) {
        uncertainItems.push({
          sku, nameCN,
          field: 'all',
          label: '全部信息（进价/供应商/货源地）',
          currentValue: '',
          savedValue: saved ? JSON.stringify({ supplier: saved.supplier, domesticSource: saved.domesticSource, maxPrice: saved.maxPrice }) : '',
          savedData: saved,
          reason: '产品数据库中未找到此SKU',
          supplier: '',
        });
        return;
      }

      // 情况2：进价为0
      if (!found.maxPrice || found.maxPrice <= 0) {
        uncertainItems.push({
          sku, nameCN,
          field: 'maxPrice',
          label: '最高进价(¥)',
          currentValue: '',
          savedValue: saved ? String(saved.maxPrice || '') : '',
          savedData: saved,
          reason: '未找到有效进价记录',
          supplier: found.supplier || '',
        });
      }

      // 情况3：境内货源地缺失或不明确
      const cityMissing = !found.city || found.city === '';
      const cityUncertain = found.cityUncertain === true;
      if (cityMissing || cityUncertain) {
        uncertainItems.push({
          sku, nameCN,
          field: 'domesticSource',
          label: '境内货源地',
          currentValue: found.city || '',
          savedValue: saved ? String(saved.domesticSource || '') : '',
          savedData: saved,
          reason: cityMissing ? '供应商信息缺失，无法推断货源地' : '货源地不确定（供应商: ' + (found.rawSupplier || found.supplier) + '）',
          supplier: found.supplier || found.rawSupplier || '',
        });
      }

      // 情况4：供应商为空
      if (!found.supplier || found.supplier === '') {
        // 如果已经因为货源地缺失加了，就合并到同一条
        const alreadyListed = uncertainItems.find(
          it => it.sku === sku && it.field === 'domesticSource'
        );
        if (alreadyListed) {
          alreadyListed.alsoMissingSupplier = true;
        } else {
          uncertainItems.push({
            sku, nameCN,
            field: 'supplier',
            label: '供应商',
            currentValue: '',
            savedValue: saved ? String(saved.supplier || '') : '',
            savedData: saved,
            reason: '供应商信息缺失',
            supplier: '',
          });
        }
      }

      // 记录当前计算值供前端展示
      if (found.maxPrice > 0) {
        const unitPriceUSD = (found.maxPrice * constant) / exchangeRate;
        uncertainItems.forEach(it => {
          if (it.sku === sku) {
            it.maxPrice = found.maxPrice;
            it.unitPriceUSD = unitPriceUSD;
          }
        });
      }
    });

    res.json({
      success: true,
      uncertainItems,
      hasUncertain: uncertainItems.length > 0,
    });
  } catch (err) {
    res.status(500).json({ success: false, message: '核验失败: ' + err.message });
  }
});

// ==================== 保存用户确认数据 ====================
app.post('/api/save-confirmations', (req, res) => {
  try {
    const { confirmations: newData } = req.body;
    if (!newData || typeof newData !== 'object') {
      return res.status(400).json({ success: false, message: '请提供确认数据' });
    }

    const current = readConfirmations();
    const now = new Date().toISOString();
    let saved = 0;

    Object.entries(newData).forEach(([sku, info]) => {
      if (!sku) return;
      current[sku] = {
        ...(current[sku] || {}),
        ...info,
        confirmedAt: now,
      };
      saved++;
    });

    writeConfirmations(current);
    console.log('Confirmations saved: ' + saved + ' SKUs');
    res.json({ success: true, saved });
  } catch (err) {
    res.status(500).json({ success: false, message: '保存失败: ' + err.message });
  }
});

// ==================== 产品重量缓存 ====================
const WEIGHTS_FILE = path.join(__dirname, 'data', 'product_weights.json');
let _weightsCache = null;
function loadWeights() {
  if (_weightsCache) return _weightsCache;
  try {
    if (fs.existsSync(WEIGHTS_FILE)) {
      _weightsCache = JSON.parse(fs.readFileSync(WEIGHTS_FILE, 'utf-8'));
      console.log('Product weights loaded: ' + Object.keys(_weightsCache).length + ' SKUs');
    } else {
      _weightsCache = {};
    }
  } catch (e) {
    _weightsCache = {};
  }
  return _weightsCache;
}

// ==================== 智能净重推断 ====================
/**
 * 获取单件产品净重 (kg)，按优先级：
 * 1. 人工补充数据 (manual_data.json)
 * 2. 亚马逊出货产品信息表 (product_weights.json)
 * 3. 默认 0.5kg
 */
function getNetWeightPerUnit(sku, description) {
  const manualData = readManualData();
  // 1. 人工数据
  if (manualData[sku] && manualData[sku].netWeightPerUnit > 0) {
    return manualData[sku].netWeightPerUnit;
  }

  // 2. 出货产品信息表中的重量
  const weights = loadWeights();
  if (weights[sku] && weights[sku] > 0) {
    return weights[sku];
  }

  // 3. 默认
  return 0.5;
}

// ==================== 预览待生成数据 ====================
app.post('/api/preview-data', async (req, res) => {
  try {
    const settings = readSettings();
    const exchangeRate = parseFloat(settings.exchangeRate) || 7.25;
    const constant = parseFloat(settings.constant) || 1.25;
    const destination = settings.destination || '';

    const sessionFile = path.join(UPLOADS_DIR, 'last_parsed.json');
    if (!fs.existsSync(sessionFile)) {
      return res.status(400).json({ success: false, message: '请先上传发票文件' });
    }

    const invoiceData = JSON.parse(fs.readFileSync(sessionFile, 'utf-8'));
    const { products: invoiceProducts, boxes: invoiceBoxes } = invoiceData;

    if (!invoiceProducts || invoiceProducts.length === 0) {
      return res.status(400).json({ success: false, message: '发票数据为空' });
    }

    const { genProducts, genBoxes } = await computeGenData({
      invoiceProducts, invoiceBoxes, exchangeRate, constant, destination,
      confirmedLocations: req.body.confirmedLocations || {},
    });

    res.json({
      success: true,
      products: genProducts,
      boxes: genBoxes,
      exchangeRate, constant, destination,
    });
  } catch (err) {
    console.error('Preview error:', err);
    res.status(500).json({ success: false, message: '预览失败: ' + err.message });
  }
});

// ==================== 核心计算：发票数据 + 数据库 → 生成数据 ====================
async function computeGenData(params) {
  const { invoiceProducts, invoiceBoxes, exchangeRate, constant, destination, confirmedLocations, overrides } = params;

  // 查询产品数据库获取进价
  const skus = invoiceProducts.map(p => p.sku);
  const lookupResult = productLookup.lookupProducts(skus);

  // ★ 人工补充数据优先：覆盖已找到SKU的空值/更低价格
  const manualData = readManualData();
  Object.keys(lookupResult.found).forEach(sku => {
    if (manualData[sku]) {
      const md = manualData[sku];
      const item = lookupResult.found[sku];
      if (md.maxPrice > item.maxPrice) item.maxPrice = md.maxPrice;
      if (!item.supplier && md.supplier) item.supplier = md.supplier;
      if (!item.city && md.city) { item.city = md.city; item.cityUncertain = false; }
      if (!item.description && md.description) item.description = md.description;
      item.source = item.source || 'manual-override';
    }
  });
  // 对未找到的SKU，从人工补充数据中补全
  const stillNotFound = [];
  lookupResult.notFound.forEach(sku => {
    if (manualData[sku]) {
      const md = manualData[sku];
      lookupResult.found[sku] = {
        sku: sku, maxPrice: md.maxPrice, allPrices: [md.maxPrice],
        supplier: md.supplier || '', description: md.description || '',
        recordCount: 1, city: md.city || '', cityUncertain: false,
        rawSupplier: md.supplier || '', source: 'manual',
      };
    } else { stillNotFound.push(sku); }
  });
  lookupResult.notFound = stillNotFound;

  // ★ 自动从供应商名称推断城市（同步：缓存/内置映射/名称模式）
  const uncertainItems = [];
  Object.values(lookupResult.found).forEach(item => {
    if ((!item.city || item.cityUncertain) && item.supplier && item.supplier.length >= 3) {
      uncertainItems.push(item);
    }
  });
  if (uncertainItems.length > 0) {
    const suppliers = [...new Set(uncertainItems.map(it => it.supplier))];
    for (const supplier of suppliers) {
      try {
        const info = await supplierLookup.lookupSupplierCity(supplier);
        if (info) {
          uncertainItems.filter(it => it.supplier === supplier).forEach(it => {
            it.city = info.city;
            it.cityUncertain = false;
            it._citySource = info.source;
          });
        }
      } catch (_) {}
    }
    console.log('[computeGenData] 自动推断: ' + Object.keys(suppliers).length + ' 供应商 → 货源地');
  }

  // 用户确认的数据
  const savedConfirmations = readConfirmations();
  const mergedConfirmations = {};
  Object.entries(savedConfirmations).forEach(([sku, info]) => {
    mergedConfirmations[sku] = { ...info };
  });
  Object.entries(confirmedLocations).forEach(([sku, city]) => {
    if (!mergedConfirmations[sku]) mergedConfirmations[sku] = {};
    mergedConfirmations[sku].domesticSource = city;
  });

  // 构建产品列表
  const genProducts = invoiceProducts.map(invProd => {
    const found = lookupResult.found[invProd.sku];
    const conf = mergedConfirmations[invProd.sku] || {};

    // 进价：用户确认 > 数据库
    let maxPrice = conf.maxPrice ? parseFloat(conf.maxPrice) : (found ? found.maxPrice : 0);
    let unitPriceUSD = (maxPrice * constant) / exchangeRate;

    // 货源地
    let domesticSource = conf.domesticSource || '';
    if (!domesticSource && found && found.city) domesticSource = found.city;

    // 供应商
    const supplier = conf.supplier || (found ? found.supplier : '') || '';

    // 净重
    const netWPerUnit = getNetWeightPerUnit(invProd.sku, found ? found.description : '');
    let calcNetWeight = invProd.totalQty * netWPerUnit;
    let quantity = invProd.totalQty;
    let unit = invProd.unit || '个';
    let hsCodeMerged = invProd.hsCode;
    let nameCN = invProd.nameCN;
    let spec = (found && found.description) ? found.description : invProd.nameCN;
    let material = invProd.material;
    let usage = invProd.usage;
    let inspectionCode = '';  // 商检编码，默认空

    // ★ 应用用户人工审查的覆盖
    if (overrides && overrides[invProd.sku]) {
      const ov = overrides[invProd.sku];
      if (ov.unitPriceUSD !== undefined) {
        unitPriceUSD = parseFloat(ov.unitPriceUSD);
        maxPrice = (unitPriceUSD * exchangeRate) / constant;
      }
      if (ov.unitPriceCNY !== undefined) {
        maxPrice = parseFloat(ov.unitPriceCNY);
        unitPriceUSD = (maxPrice * constant) / exchangeRate;
      }
      if (ov.maxPrice !== undefined) {
        maxPrice = parseFloat(ov.maxPrice);
        unitPriceUSD = (maxPrice * constant) / exchangeRate;
      }
      if (ov.domesticSource !== undefined) domesticSource = String(ov.domesticSource);
      if (ov.netWeight !== undefined) calcNetWeight = parseFloat(ov.netWeight);
      if (ov.quantity !== undefined) quantity = parseInt(ov.quantity) || quantity;
      if (ov.unit !== undefined) unit = String(ov.unit);
      if (ov.hsCode !== undefined) hsCodeMerged = parseInt(ov.hsCode) || hsCodeMerged;
      if (ov.nameCN !== undefined) nameCN = String(ov.nameCN);
      if (ov.spec !== undefined) spec = String(ov.spec);
      if (ov.material !== undefined) material = String(ov.material);
      if (ov.usage !== undefined) usage = String(ov.usage);
      if (ov.supplier !== undefined) ov._supplier = String(ov.supplier);
      if (ov.inspectionCode !== undefined) inspectionCode = String(ov.inspectionCode);
    }

    return {
      sku: invProd.sku,
      nameCN, nameEN: invProd.nameEN,
      spec, hsCode: invProd.hsCode, hsCodeMerged,
      material, usage,
      quantity, unit,
      unitPriceCNY: maxPrice,
      unitPriceUSD,
      netWeight: calcNetWeight,
      originCountry: '中国',
      destination,
      domesticSource,
      supplier,
      inspectionCode,
    };
  });

  // 构建箱级产品明细
  const genBoxes = invoiceBoxes.map(box => ({
    id: box.id, boxSeq: box.boxSeq,
    length: box.length, width: box.width, height: box.height, weight: box.weight,
    products: box.products.map(bp => {
      const found = lookupResult.found[bp.sku];
      const conf = mergedConfirmations[bp.sku] || {};

      let maxPrice = conf.maxPrice ? parseFloat(conf.maxPrice) : (found ? found.maxPrice : 0);
      let unitPriceUSD = (maxPrice * constant) / exchangeRate;

      let domesticSource = conf.domesticSource || '';
      if (!domesticSource && found && found.city) domesticSource = found.city;

      const bNetWPerUnit = getNetWeightPerUnit(bp.sku, found ? found.description : '');
      let netWBox = bp.qtyPerBox * bNetWPerUnit;
      let qtyPerBox = bp.qtyPerBox;

      // ★ 应用覆盖到箱级
      if (overrides && overrides[bp.sku]) {
        const ov = overrides[bp.sku];
        if (ov.unitPriceUSD !== undefined) {
          unitPriceUSD = parseFloat(ov.unitPriceUSD);
          maxPrice = (unitPriceUSD * exchangeRate) / constant;
        }
        if (ov.maxPrice !== undefined) {
          maxPrice = parseFloat(ov.maxPrice);
          unitPriceUSD = (maxPrice * constant) / exchangeRate;
        }
        if (ov.domesticSource !== undefined) domesticSource = String(ov.domesticSource);
        if (ov.netWeightPerUnit !== undefined) {
          netWBox = qtyPerBox * parseFloat(ov.netWeightPerUnit);
        }
      }

      return {
        sku: bp.sku, nameCN: bp.nameCN, nameEN: bp.nameEN,
        spec: (found && found.description) ? found.description : bp.nameCN,
        hsCode: bp.hsCode, material: bp.material, usage: bp.usage,
        quantityPerBox: qtyPerBox, quantity: qtyPerBox,
        unit: bp.unit || '个',
        unitPriceCNY: maxPrice, unitPriceUSD,
        netWeightPerBox: netWBox, netWeight: netWBox,
        originCountry: '中国', destination,
        domesticSource,
      };
    }),
  }));

  return { genProducts, genBoxes, lookupResult };
}

// ==================== 生成报关文件 ====================
app.post('/api/generate', async (req, res) => {
  try {
    const settings = readSettings();
    const exchangeRate = parseFloat(settings.exchangeRate) || 7.25;
    const constant = parseFloat(settings.constant) || 1.25;
    const destination = settings.destination || '';

    if (!destination) {
      return res.status(400).json({ success: false, message: '请先在设置中填写最终目的地' });
    }

    // 读取发票数据
    const sessionFile = path.join(UPLOADS_DIR, 'last_parsed.json');
    if (!fs.existsSync(sessionFile)) {
      return res.status(400).json({ success: false, message: '请先上传发票文件' });
    }

    const invoiceData = JSON.parse(fs.readFileSync(sessionFile, 'utf-8'));
    const { products: invoiceProducts, boxes: invoiceBoxes } = invoiceData;

    if (!invoiceProducts || invoiceProducts.length === 0) {
      return res.status(400).json({ success: false, message: '发票数据为空' });
    }

    // 使用统一计算函数
    const { genProducts, genBoxes } = await computeGenData({
      invoiceProducts, invoiceBoxes, exchangeRate, constant, destination,
      confirmedLocations: req.body.confirmedLocations || {},
      overrides: req.body.overrides || {},
    });

    // 从发票原始文件名提取日期和生成文件名
    const origName = invoiceData._originalName || '';
    const dateMatch = origName.match(/^(\d+)[-.](\d+)/);
    let invoiceDate = new Date(); // fallback: today
    let fileDateLabel = '未知日期';
    if (dateMatch) {
      const m = parseInt(dateMatch[1]), d = parseInt(dateMatch[2]);
      invoiceDate = new Date(2026, m - 1, d); // 月-日 → Date
      fileDateLabel = dateMatch[0].replace('.', '-');
    }
    // 输出文件名 = 发票日期 + 待审核.xlsx
    const fileName = fileDateLabel + '待审核.xlsx';
    const outputPath = path.join(OUTPUT_DIR, fileName);

    // === 数据一致性核验（净重、毛重交叉校验） ===
    const validation = validateConsistency(genProducts, genBoxes);

    // 调用模板生成模块
    await templateGen.generate({
      products: validation.products,
      boxes: validation.boxes,
      destination: destination,
      exchangeRate: exchangeRate,
      constant: constant,
      templatePath: TEMPLATE_PATH,
      outputPath: outputPath,
      invoiceDate: invoiceDate,
    });

    res.json({
      success: true,
      fileName: fileName,
      downloadUrl: '/api/download/' + encodeURIComponent(fileName),
      validation: validation.report,
    });
  } catch (err) {
    console.error('Generate error:', err);
    res.status(500).json({ success: false, message: '生成失败: ' + err.message });
  }
});

// ==================== 文件下载 ====================
app.get('/api/download/:filename', (req, res) => {
  try {
    const filename = decodeURIComponent(req.params.filename);
    const filePath = path.join(OUTPUT_DIR, filename);

    if (!fs.existsSync(filePath)) {
      return res.status(404).json({ success: false, message: '文件不存在' });
    }

    res.download(filePath, filename);
  } catch (err) {
    res.status(500).json({ success: false, message: '下载失败: ' + err.message });
  }
});

// ==================== 全局错误处理 ====================
app.use((err, req, res, next) => {
  console.error('Unhandled error:', err);
  if (err.code === 'LIMIT_FILE_SIZE') {
    return res.status(400).json({ success: false, message: '文件过大' });
  }
  res.status(500).json({ success: false, message: '服务器内部错误: ' + err.message });
});

// ==================== 启动服务 ====================
function startServer(port) {
  const server = app.listen(port, () => {
    console.log('='.repeat(50));
    console.log('  报关辅助系统 Customs App v2.0');
    console.log('  Server running at: http://localhost:' + port);
    console.log('='.repeat(50));

    // 服务器就绪后自动打开浏览器
    const url = 'http://localhost:' + port;
    const platform = process.platform;
    const cmd = platform === 'win32'
      ? 'start'
      : (platform === 'darwin' ? 'open' : 'xdg-open');
    require('child_process').exec(cmd + ' ' + url, (err) => {
      if (err) console.log('请手动打开浏览器访问: ' + url);
    });
  });

  server.on('error', (err) => {
    if (err.code === 'EADDRINUSE') {
      console.log('端口 ' + port + ' 已被占用，尝试自动处理...');

      // Windows: 查找并终止占用端口的进程
      if (process.platform === 'win32') {
        const { execSync } = require('child_process');
        try {
          const output = execSync('netstat -ano | findstr :' + port, { encoding: 'utf-8' });
          const lines = output.trim().split(/\r?\n/);
          const pids = new Set();
          lines.forEach(line => {
            const parts = line.trim().split(/\s+/);
            const pid = parts[parts.length - 1];
            if (pid && pid !== '0' && pid !== 'LISTENING') pids.add(pid);
          });
          pids.forEach(pid => {
            try {
              console.log('正在终止进程 PID: ' + pid);
              execSync('taskkill /F /PID ' + pid, { encoding: 'utf-8' });
            } catch (e) { /* 进程可能已结束 */ }
          });
          // 等待端口释放后重试
          console.log('等待端口释放...');
          setTimeout(() => startServer(port), 1000);
          return;
        } catch (e) {
          console.log('无法自动释放端口，请手动关闭占用程序后重试');
        }
      }

      console.log('请尝试: 1) 关闭之前的命令行窗口  2) 重启电脑');
      console.log('或在命令行执行: netstat -ano | findstr :' + port);
    } else {
      console.error('服务器启动失败:', err.message);
    }
    process.exit(1);
  });
}

startServer(PORT);
