const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const productLookup = require('./lib/product-lookup');
const templateGen = require('./lib/template-generator');

const app = express();
const PORT = 3000;

app.use(express.json({ limit: '50mb' }));
app.use(express.static(path.join(__dirname, 'public')));

const SETTINGS_FILE = path.join(__dirname, 'data', 'settings.json');
const MANUAL_DATA_FILE = path.join(__dirname, 'data', 'manual_product_data.json');
const CONFIRMATIONS_FILE = path.join(__dirname, 'data', 'confirmations.json');
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
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (ext === '.xlsx' || ext === '.xls') cb(null, true);
    else cb(new Error('仅支持 .xlsx / .xls 格式'));
  },
});

app.post('/api/upload-invoice', (req, res) => {
  upload.single('file')(req, res, (err) => {
    if (err) return res.status(400).json({ success: false, message: err.message });
    if (!req.file) return res.status(400).json({ success: false, message: '请选择文件' });

    try {
      const result = parseInvoice(req.file.path);
      // 保存解析结果到会话文件
      const sessionFile = path.join(UPLOADS_DIR, 'last_parsed.json');
      fs.writeFileSync(sessionFile, JSON.stringify(result, null, 2), 'utf-8');
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
  const ws = wb.Sheets['CIPI'];
  if (!ws) throw new Error('未找到 CIPI 工作表');

  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true });
  if (raw.length < 9) throw new Error('发票数据行数不足');

  // 数据起始行：找到第一个 SKU 行（以 FD-/FUK-/SWK- 等开头）
  let dataStart = 0;
  for (let i = 0; i < raw.length; i++) {
    const firstVal = String(raw[i][0] || '').trim();
    if (/^[A-Z]{2,4}\d*-/i.test(firstVal) || /^[A-Z]{2,4}-\d/i.test(firstVal)) {
      dataStart = i;
      break;
    }
  }
  if (dataStart === 0) throw new Error('未找到数据起始行');

  const boxes = [];
  const skuPrices = new Map(); // SKU → [prices]
  let currentBox = null;
  let prevBoxId = '';

  for (let i = dataStart; i < raw.length; i++) {
    const row = raw[i];
    if (!row || !row[0]) continue;

    const sku = String(row[0]).trim();
    if (!sku || sku === '') continue;

    const boxId = row[1] ? String(row[1]).trim() : '';
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

    // 收集SKU价格
    if (unitPriceEUR > 0) {
      if (!skuPrices.has(sku)) skuPrices.set(sku, []);
      skuPrices.get(sku).push(unitPriceEUR);
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
    };

    // 新建或切换箱子
    if (isNewBox) {
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

  // 按SKU汇总产品
  const productMap = new Map();
  boxes.forEach(box => {
    box.products.forEach(prod => {
      const key = prod.sku;
      if (!productMap.has(key)) {
        productMap.set(key, {
          sku: prod.sku,
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
        });
      }
      const agg = productMap.get(key);
      agg.totalQty += prod.qtyPerBox;
    });
  });

  // 为每个产品解析最佳价格
  const products = [];
  productMap.forEach((agg) => {
    const prices = skuPrices.get(agg.sku) || [];
    let bestPriceEUR = 0;
    if (prices.length > 0) {
      // 取最低价作为最优进价
      bestPriceEUR = Math.min(...prices);
    }
    agg.bestPriceEUR = bestPriceEUR;
    // 估算净重 = 总数量 * 0.5kg/个
    agg.netWeightEstimate = agg.totalQty * 0.5;
    agg.netWeight = agg.netWeightEstimate;
    products.push(agg);
  });

  return {
    products: products,
    boxes: boxes,
    totalBoxes: boxes.length,
  };
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
app.post('/api/lookup-products', (req, res) => {
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

    // 对未找到的SKU，尝试从人工补充数据中匹配
    const manualData = readManualData();
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
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (ext === '.xlsx' || ext === '.xls') cb(null, true);
    else cb(new Error('仅支持 .xlsx / .xls 格式'));
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

// ==================== 生成前数据核验 ====================
app.post('/api/preflight-check', (req, res) => {
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

// ==================== 智能净重推断 ====================
/**
 * 获取单件产品净重 (kg)，按优先级：
 * 1. 人工补充数据 (manual_data.json)
 * 2. 产品描述中的重量（如 "1KG"、"170g"）
 * 3. SKU 名称中的重量（如 "C2KG" → 2kg, "C1KG" → 1kg）
 * 4. 默认 0.5kg
 */
function getNetWeightPerUnit(sku, description) {
  const manualData = readManualData();
  // 1. 人工数据
  if (manualData[sku] && manualData[sku].netWeightPerUnit > 0) {
    return manualData[sku].netWeightPerUnit;
  }

  // 2. 产品描述中的重量
  if (description) {
    const m = description.match(/(\d+\.?\d*)\s*(KG|kg|Kg|g|G)/);
    if (m) {
      let w = parseFloat(m[1]);
      if (m[2] === 'g' || m[2] === 'G') w /= 1000;
      if (w > 0) return w;
    }
  }

  // 3. SKU 名称中的重量 (如 C2KG→2, C1KG→1)
  const skuMatch = sku.match(/C?(\d+)\s*KG/i);
  if (skuMatch) {
    const w = parseFloat(skuMatch[1]);
    if (w > 0 && w < 500) return w;
  }

  // 4. 默认
  return 0.5;
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

    // 读取上次解析的发票数据
    const sessionFile = path.join(UPLOADS_DIR, 'last_parsed.json');
    if (!fs.existsSync(sessionFile)) {
      return res.status(400).json({ success: false, message: '请先上传发票文件' });
    }

    const invoiceData = JSON.parse(fs.readFileSync(sessionFile, 'utf-8'));
    const { products: invoiceProducts, boxes: invoiceBoxes } = invoiceData;

    if (!invoiceProducts || invoiceProducts.length === 0) {
      return res.status(400).json({ success: false, message: '发票数据为空' });
    }

    // 查询产品数据库获取进价
    const skus = invoiceProducts.map(p => p.sku);
    const lookupResult = productLookup.lookupProducts(skus);

    // 对未找到的SKU，从人工补充数据中补全
    const manualData = readManualData();
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
      } else {
        stillNotFound.push(sku);
      }
    });
    lookupResult.notFound = stillNotFound;

    // 用户确认的货源地（前端传来的 + 历史保存的）
    const confirmedLocations = req.body.confirmedLocations || {};
    const savedConfirmations = readConfirmations();
    // 合并：前端传来的优先，历史保存的作为后备
    Object.entries(savedConfirmations).forEach(([sku, info]) => {
      if (!confirmedLocations[sku] && info.domesticSource) {
        confirmedLocations[sku] = info.domesticSource;
      }
    });

    // 构建用于生成的产品列表
    const genProducts = invoiceProducts.map(invProd => {
      const found = lookupResult.found[invProd.sku];
      const maxPrice = found ? found.maxPrice : 0;
      const unitPriceUSD = (maxPrice * constant) / exchangeRate;

      // 确定境内货源地
      let domesticSource = '';
      if (confirmedLocations[invProd.sku]) {
        domesticSource = confirmedLocations[invProd.sku];
      } else if (found && found.city) {
        domesticSource = found.city;
      }

      // 从多来源智能推断单件净重
      const netWPerUnit = getNetWeightPerUnit(invProd.sku, found ? found.description : '');
      const calcNetWeight = invProd.totalQty * netWPerUnit;

      return {
        sku: invProd.sku,
        nameCN: invProd.nameCN,
        nameEN: invProd.nameEN,
        spec: invProd.nameCN,
        hsCode: invProd.hsCode,
        hsCodeMerged: invProd.hsCode,
        material: invProd.material,
        usage: invProd.usage,
        quantity: invProd.totalQty,
        unit: invProd.unit || '个',
        unitPriceCNY: maxPrice,
        unitPriceUSD: unitPriceUSD,
        netWeight: calcNetWeight,
        originCountry: '中国',
        destination: destination,
        domesticSource: domesticSource,
      };
    });

    // 构建箱级产品明细（带单价和货源地信息）
    const genBoxes = invoiceBoxes.map(box => ({
      id: box.id,
      boxSeq: box.boxSeq,
      length: box.length,
      width: box.width,
      height: box.height,
      weight: box.weight,
      products: box.products.map(bp => {
        const found = lookupResult.found[bp.sku];
        const maxPrice = found ? found.maxPrice : 0;
        const unitPriceUSD = (maxPrice * constant) / exchangeRate;

        let domesticSource = '';
        if (confirmedLocations[bp.sku]) {
          domesticSource = confirmedLocations[bp.sku];
        } else if (found && found.city) {
          domesticSource = found.city;
        }

        const bNetWPerUnit = getNetWeightPerUnit(bp.sku, found ? found.description : '');

        return {
          sku: bp.sku,
          nameCN: bp.nameCN,
          nameEN: bp.nameEN,
          spec: bp.nameCN,
          hsCode: bp.hsCode,
          material: bp.material,
          usage: bp.usage,
          quantityPerBox: bp.qtyPerBox,
          quantity: bp.qtyPerBox,
          unit: bp.unit || '个',
          unitPriceCNY: maxPrice,
          unitPriceUSD: unitPriceUSD,
          netWeightPerBox: bp.qtyPerBox * bNetWPerUnit,
          netWeight: bp.qtyPerBox * bNetWPerUnit,
          originCountry: '中国',
          destination: destination,
          domesticSource: domesticSource,
        };
      }),
    }));

    // 生成输出文件名
    const now = new Date();
    const dateStr = now.toISOString().slice(0, 10).replace(/-/g, '');
    const timeStr = now.toTimeString().slice(0, 8).replace(/:/g, '');
    const fileName = '报关资料_' + dateStr + '_' + timeStr + '.xlsx';
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

    // 安全检查：防止路径穿越
    const normalizedPath = path.normalize(filePath);
    if (!normalizedPath.startsWith(path.normalize(OUTPUT_DIR))) {
      return res.status(403).json({ success: false, message: '禁止访问' });
    }

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
app.listen(PORT, () => {
  console.log('='.repeat(50));
  console.log('  报关辅助系统 Customs App v2.0');
  console.log('  Server running at: http://localhost:' + PORT);
  console.log('='.repeat(50));
});
