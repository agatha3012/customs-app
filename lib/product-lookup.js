const XLSX = require('xlsx');
const path = require('path');

// === 缓存 ===
let _db = null; // Map: SKU → [{price, supplier, description, date}]

// === 城市提取规则 ===
// 供应商名称通常以城市/地区开头
const CITY_PATTERNS = [
  { regex: /^(东莞市|东莞)/, city: '东莞' },
  { regex: /^(深圳市|深圳)/, city: '深圳' },
  { regex: /^(广州市|广州)/, city: '广州' },
  { regex: /^(中山市|中山)/, city: '中山' },
  { regex: /^(惠州市|惠州)/, city: '惠州' },
  { regex: /^(佛山市|佛山)/, city: '佛山' },
  { regex: /^(珠海市|珠海)/, city: '珠海' },
  { regex: /^(江门市|江门)/, city: '江门' },
  { regex: /^(肇庆市|肇庆)/, city: '肇庆' },
  { regex: /^(汕头市|汕头)/, city: '汕头' },
  { regex: /^(潮州市|潮州)/, city: '潮州' },
  { regex: /^(揭阳市|揭阳)/, city: '揭阳' },
  { regex: /^(汕尾市|汕尾)/, city: '汕尾' },
  { regex: /^(湛江市|湛江)/, city: '湛江' },
  { regex: /^(茂名市|茂名)/, city: '茂名' },
  { regex: /^(阳江市|阳江)/, city: '阳江' },
  { regex: /^(云浮市|云浮)/, city: '云浮' },
  { regex: /^(清远市|清远)/, city: '清远' },
  { regex: /^(韶关市|韶关)/, city: '韶关' },
  { regex: /^(河源市|河源)/, city: '河源' },
  { regex: /^(梅州市|梅州)/, city: '梅州' },
  { regex: /^(义乌市|义乌)/, city: '义乌' },
  { regex: /^(杭州市|杭州)/, city: '杭州' },
  { regex: /^(宁波市|宁波)/, city: '宁波' },
  { regex: /^(温州市|温州)/, city: '温州' },
  { regex: /^(嘉兴市|嘉兴)/, city: '嘉兴' },
  { regex: /^(湖州市|湖州)/, city: '湖州' },
  { regex: /^(绍兴市|绍兴)/, city: '绍兴' },
  { regex: /^(金华市|金华)/, city: '金华' },
  { regex: /^(衢州市|衢州)/, city: '衢州' },
  { regex: /^(台州市|台州)/, city: '台州' },
  { regex: /^(丽水市|丽水)/, city: '丽水' },
  { regex: /^(舟山市|舟山)/, city: '舟山' },
  { regex: /^(北京市|北京)/, city: '北京' },
  { regex: /^(上海市|上海)/, city: '上海' },
  { regex: /^(天津市|天津)/, city: '天津' },
  { regex: /^(重庆市|重庆)/, city: '重庆' },
  { regex: /^(南京市|南京)/, city: '南京' },
  { regex: /^(苏州市|苏州)/, city: '苏州' },
  { regex: /^(无锡市|无锡)/, city: '无锡' },
  { regex: /^(常州市|常州)/, city: '常州' },
  { regex: /^(南通市|南通)/, city: '南通' },
  { regex: /^(扬州市|扬州)/, city: '扬州' },
  { regex: /^(镇江市|镇江)/, city: '镇江' },
  { regex: /^(泰州市|泰州)/, city: '泰州' },
  { regex: /^(盐城市|盐城)/, city: '盐城' },
  { regex: /^(淮安市|淮安)/, city: '淮安' },
  { regex: /^(连云港市|连云港)/, city: '连云港' },
  { regex: /^(徐州市|徐州)/, city: '徐州' },
  { regex: /^(宿迁市|宿迁)/, city: '宿迁' },
  { regex: /^(合肥市|合肥)/, city: '合肥' },
  { regex: /^(芜湖市|芜湖)/, city: '芜湖' },
  { regex: /^(任丘市|任丘)/, city: '任丘' },
  { regex: /^(石家庄市|石家庄)/, city: '石家庄' },
  { regex: /^(唐山市|唐山)/, city: '唐山' },
  { regex: /^(保定市|保定)/, city: '保定' },
  { regex: /^(沧州市|沧州)/, city: '沧州' },
  { regex: /^(河间市|河间)/, city: '沧州' },
  { regex: /^(清河县|清河)/, city: '清河' },
  { regex: /^(邢台市|邢台)/, city: '邢台' },
  { regex: /^(衡水市|衡水)/, city: '衡水' },
  { regex: /^(廊坊市|廊坊)/, city: '廊坊' },
  { regex: /^(邯郸市|邯郸)/, city: '邯郸' },
  { regex: /^(秦皇岛市|秦皇岛)/, city: '秦皇岛' },
  { regex: /^(承德市|承德)/, city: '承德' },
  { regex: /^(张家口市|张家口)/, city: '张家口' },
  { regex: /^(成都市|成都)/, city: '成都' },
  { regex: /^(绵阳市|绵阳)/, city: '绵阳' },
  { regex: /^(德阳市|德阳)/, city: '德阳' },
  { regex: /^(宜宾市|宜宾)/, city: '宜宾' },
  { regex: /^(武汉市|武汉)/, city: '武汉' },
  { regex: /^(宜昌市|宜昌)/, city: '宜昌' },
  { regex: /^(襄阳市|襄阳)/, city: '襄阳' },
  { regex: /^(长沙市|长沙)/, city: '长沙' },
  { regex: /^(株洲市|株洲)/, city: '株洲' },
  { regex: /^(湘潭市|湘潭)/, city: '湘潭' },
  { regex: /^(郑州市|郑州)/, city: '郑州' },
  { regex: /^(洛阳市|洛阳)/, city: '洛阳' },
  { regex: /^(新乡市|新乡)/, city: '新乡' },
  { regex: /^(济南市|济南)/, city: '济南' },
  { regex: /^(青岛市|青岛)/, city: '青岛' },
  { regex: /^(烟台市|烟台)/, city: '烟台' },
  { regex: /^(潍坊市|潍坊)/, city: '潍坊' },
  { regex: /^(威海市|威海)/, city: '威海' },
  { regex: /^(淄博市|淄博)/, city: '淄博' },
  { regex: /^(厦门市|厦门)/, city: '厦门' },
  { regex: /^(福州市|福州)/, city: '福州' },
  { regex: /^(泉州市|泉州)/, city: '泉州' },
  { regex: /^(漳州市|漳州)/, city: '漳州' },
  { regex: /^(西安市|西安)/, city: '西安' },
  { regex: /^(咸阳市|咸阳)/, city: '咸阳' },
  { regex: /^(兰州市|兰州)/, city: '兰州' },
  { regex: /^(贵阳市|贵阳)/, city: '贵阳' },
  { regex: /^(昆明市|昆明)/, city: '昆明' },
  { regex: /^(南宁市|南宁)/, city: '南宁' },
  { regex: /^(桂林市|桂林)/, city: '桂林' },
  { regex: /^(南昌市|南昌)/, city: '南昌' },
  { regex: /^(赣州市|赣州)/, city: '赣州' },
  { regex: /^(太原市|太原)/, city: '太原' },
  { regex: /^(大同市|大同)/, city: '大同' },
  { regex: /^(大连市|大连)/, city: '大连' },
  { regex: /^(沈阳市|沈阳)/, city: '沈阳' },
  { regex: /^(长春市|长春)/, city: '长春' },
  { regex: /^(吉林市|吉林)/, city: '吉林' },
  { regex: /^(哈尔滨市|哈尔滨)/, city: '哈尔滨' },
  { regex: /^(呼和浩特市|呼和浩特)/, city: '呼和浩特' },
  { regex: /^(包头市|包头)/, city: '包头' },
  { regex: /^(银川市|银川)/, city: '银川' },
  { regex: /^(西宁市|西宁)/, city: '西宁' },
  { regex: /^(拉萨市|拉萨)/, city: '拉萨' },
  { regex: /^(乌鲁木齐市|乌鲁木齐)/, city: '乌鲁木齐' },
  { regex: /^(海口市|海口)/, city: '海口' },
  { regex: /^(三亚市|三亚)/, city: '三亚' },
];

/**
 * 从供应商名称中提取城市
 * @param {string} supplier - 供应商全名
 * @returns {{city: string|null, uncertain: boolean}}
 */
function extractCity(supplier) {
  if (!supplier || typeof supplier !== 'string') {
    return { city: null, uncertain: false };
  }

  // 尝试匹配已知城市模式
  for (const pattern of CITY_PATTERNS) {
    if (pattern.regex.test(supplier)) {
      return { city: pattern.city, uncertain: false };
    }
  }

  // 供应商全名匹配（供应商名称不以城市开头但可通过工商信息确定）
  const SUPPLIER_CITY_MAP = {
    '伟能(广东)新材料有限公司': '广州',
    '伟能（广东）新材料有限公司': '广州',
  };
  if (SUPPLIER_CITY_MAP[supplier]) {
    return { city: SUPPLIER_CITY_MAP[supplier], uncertain: false };
  }

  // 尝试通用模式: 以"XX市"或"XX省"开头
  const genericMatch = supplier.match(/^([^\s省]+(?:市|县|区|镇|州))/);
  if (genericMatch) {
    const rawCity = genericMatch[1].replace(/[市县区镇州]$/, '');
    return { city: rawCity, uncertain: false };
  }

  // 无法确定
  return { city: null, uncertain: true, rawSupplier: supplier };
}

/**
 * 加载/重新加载产品数据库
 * @param {string} filePath - 亚马逊备记录xlsx路径
 * @returns {Map} SKU → [{price, supplier, description, date}]
 */
function loadProductDB(filePath) {
  // 优先尝试从合并 JSON 加载（更快）
  if (!filePath && loadFromMerged()) {
    return _db;
  }

  const defaultPath = path.join(__dirname, '..', 'data', 'product_db_source.txt');

  // 尝试使用提供的路径，否则回退到默认
  let actualPath = filePath;
  if (!actualPath) {
    // 尝试读取默认路径文件获取最新位置
    try {
      const fs = require('fs');
      actualPath = fs.readFileSync(defaultPath, 'utf-8').trim();
    } catch (e) {
      // 使用桌面默认路径
      actualPath = 'D:/桌面/2026亚马逊备记录 (22).xlsx';
    }
  }

  console.log('Loading product DB from:', actualPath);

  const fs = require('fs');
  if (!fs.existsSync(actualPath)) {
    console.error('Product DB file not found:', actualPath);
    _db = new Map();
    return _db;
  }

  const wb = XLSX.readFile(actualPath);
  const ws = wb.Sheets['Sheet1'];
  const data = XLSX.utils.sheet_to_json(ws, { defval: '' });

  _db = new Map();

  data.forEach(row => {
    const sku = (row['SKU'] || '').toString().trim();
    if (!sku) return;

    const price = parseFloat(row['单价']) || 0;
    const supplier = (row['供应商'] || '').toString().trim();
    const description = (row['产品描述'] || '').toString().trim();
    const dateVal = row['具体日期'];

    if (!_db.has(sku)) {
      _db.set(sku, []);
    }
    _db.get(sku).push({ price, supplier, description, date: dateVal });
  });

  console.log(`Product DB loaded: ${_db.size} unique SKUs`);
  return _db;
}

/**
 * 根据SKU列表查询进价信息
 * @param {string[]} skuList - 要查询的SKU列表
 * @param {string} dbPath - 数据库文件路径（可选）
 * @returns {Object} { found: {...}, notFound: [...] }
 */
function lookupProducts(skuList, dbPath) {
  if (!_db) {
    loadProductDB(dbPath);
  }

  const result = { found: {}, notFound: [] };

  skuList.forEach(sku => {
    const skuNorm = sku.toString().trim();
    let records = _db.get(skuNorm);

    // 尝试模糊匹配 - 有些SKU前缀不同（如 FD- vs FUK-）
    if (!records) {
      // 尝试匹配最后一部分
      for (const [key, val] of _db.entries()) {
        // e.g. FD-602-GY505 vs FUK-602-GY505
        const keySuffix = key.replace(/^[A-Z]+-/, '');
        const skuSuffix = skuNorm.replace(/^[A-Z]+-/, '');
        if (keySuffix === skuSuffix) {
          records = val;
          console.log(`Fuzzy match: ${skuNorm} → ${key}`);
          break;
        }
      }
    }

    if (!records || records.length === 0) {
      result.notFound.push(skuNorm);
      return;
    }

    // 取最高进价（同价时优先取有供应商信息的记录）
    let maxPrice = 0;
    let bestRecord = null;
    records.forEach(r => {
      if (r.price > maxPrice || (r.price === maxPrice && r.supplier && (!bestRecord || !bestRecord.supplier))) {
        maxPrice = r.price;
        bestRecord = r;
      }
    });

    const locationInfo = extractCity(bestRecord.supplier);

    result.found[skuNorm] = {
      sku: skuNorm,
      maxPrice: maxPrice,           // 最高进价 (人民币)
      allPrices: records.map(r => r.price).sort((a, b) => b - a),
      supplier: bestRecord.supplier,
      description: bestRecord.description,
      recordCount: records.length,
      city: locationInfo.city,
      cityUncertain: locationInfo.uncertain,
      rawSupplier: locationInfo.rawSupplier,
    };
  });

  return result;
}

/**
 * 设置产品数据库路径并加载
 */
function setDBPath(filePath) {
  _db = null;
  return loadProductDB(filePath);
}

// === 合并后的持久化数据 ===
const MERGED_DB_FILE = path.join(__dirname, '..', 'data', 'product_db_merged.json');
const MERGED_META_FILE = path.join(__dirname, '..', 'data', 'product_db_meta.json');

/**
 * 尝试从合并 JSON 加载（优先于 Excel 解析，更快）
 */
function loadFromMerged() {
  try {
    const fs = require('fs');
    if (!fs.existsSync(MERGED_DB_FILE)) return false;
    const raw = JSON.parse(fs.readFileSync(MERGED_DB_FILE, 'utf-8'));
    _db = new Map();
    Object.entries(raw).forEach(([sku, records]) => {
      _db.set(sku, records);
    });
    console.log('Product DB loaded from merged JSON: ' + _db.size + ' SKUs');
    return true;
  } catch (e) {
    console.error('Failed to load merged DB:', e.message);
    return false;
  }
}

/**
 * 获取当前 DB 状态信息
 * @returns {{ totalSKUs: number, totalRecords: number, lastMerge: string|null }}
 */
function getStatus() {
  const totalSKUs = _db ? _db.size : 0;
  let totalRecords = 0;
  if (_db) { _db.forEach(v => { totalRecords += v.length; }); }

  let lastMerge = null;
  try {
    const fs = require('fs');
    if (fs.existsSync(MERGED_META_FILE)) {
      const meta = JSON.parse(fs.readFileSync(MERGED_META_FILE, 'utf-8'));
      lastMerge = meta.lastMerge || null;
    }
  } catch (e) { /* ignore */ }

  return { totalSKUs, totalRecords, lastMerge };
}

/**
 * 上传新的备货单/出货记录，与现有 DB 合并
 * - 新文件中的 SKU 覆盖旧数据
 * - 旧文件中独有的 SKU 保留
 *
 * @param {string} filePath - 上传的 Excel 文件路径
 * @returns {{ totalSKUs: number, newAdded: number, updated: number, kept: number }}
 */
function mergeFromFile(filePath) {
  const fs = require('fs');
  const XLSX = require('xlsx');

  // 1. 解析新文件（复用 loadProductDB 的逻辑但返回独立 Map）
  const wb = XLSX.readFile(filePath);
  const ws = wb.Sheets['Sheet1'];
  if (!ws) throw new Error('未找到 Sheet1 工作表');
  const data = XLSX.utils.sheet_to_json(ws, { defval: '' });

  const newDB = new Map();
  data.forEach(row => {
    const sku = (row['SKU'] || '').toString().trim();
    if (!sku) return;
    const price = parseFloat(row['单价']) || 0;
    const supplier = (row['供应商'] || '').toString().trim();
    const description = (row['产品描述'] || '').toString().trim();
    const dateVal = row['具体日期'];

    if (!newDB.has(sku)) newDB.set(sku, []);
    newDB.get(sku).push({ price, supplier, description, date: dateVal });
  });
  console.log('New file parsed: ' + newDB.size + ' SKUs');

  // 2. 确保旧 DB 已加载
  if (!_db) {
    if (!loadFromMerged()) {
      loadProductDB();
    }
  }

  // 3. 合并: 新覆盖旧, 旧保留独有
  let newAdded = 0, updated = 0, kept = 0;

  if (!_db || _db.size === 0) {
    // 没有旧数据，直接用新的
    _db = newDB;
    newAdded = newDB.size;
  } else {
    const merged = new Map(_db); // 先复制旧的全部
    kept = merged.size;

    newDB.forEach((records, sku) => {
      if (merged.has(sku)) {
        // SKU 已存在 → 新数据覆盖
        kept--;
        updated++;
      } else {
        // SKU 不存在 → 新增
        newAdded++;
      }
      merged.set(sku, records);
    });
    _db = merged;
  }

  // 4. 持久化合并结果
  const dir = path.dirname(MERGED_DB_FILE);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });

  const plain = {};
  _db.forEach((records, sku) => { plain[sku] = records; });
  fs.writeFileSync(MERGED_DB_FILE, JSON.stringify(plain, null, 2), 'utf-8');

  const now = new Date().toISOString();
  fs.writeFileSync(MERGED_META_FILE, JSON.stringify({
    lastMerge: now,
    sourceFile: path.basename(filePath),
    totalSKUs: _db.size,
    totalRecords: Array.from(_db.values()).reduce((s, r) => s + r.length, 0),
  }, null, 2), 'utf-8');

  console.log('Merge complete: ' + newAdded + ' new, ' + updated + ' updated, ' + kept + ' kept → ' + _db.size + ' total SKUs');
  return { totalSKUs: _db.size, newAdded, updated, kept };
}

module.exports = {
  loadProductDB,
  loadFromMerged,
  lookupProducts,
  extractCity,
  setDBPath,
  mergeFromFile,
  getStatus,
};
