/**
 * 供应商 → 城市 在线查询模块
 *
 * 策略：
 *   1. 本地缓存 (data/supplier_city_cache.json) — 已查过的供应商直接返回
 *   2. 供应商名称模式匹配 — 从公司名中提取城市（如"东莞XX公司"→东莞）
 *   3. 内置已知映射 — 城市不在名称中的知名供应商
 *   4. 在线搜索（尽力而为，可能因反爬机制失败）
 */

const https = require('https');
const fs = require('fs');
const path = require('path');

const CACHE_FILE = path.join(__dirname, '..', 'data', 'supplier_city_cache.json');

// ==================== 内置供应商→城市映射 ====================
// 供应商名称不含城市时，手动维护的映射表
const BUILTIN_MAP = {
  '伟能(广东)新材料有限公司': '广州',
  '伟能（广东）新材料有限公司': '广州',
  '河北鼎诚麻纺有限公司': '邢台',
};

// ==================== 城市列表 ====================
const KNOWN_CITIES = [
  '北京', '上海', '天津', '重庆',
  '广州', '深圳', '东莞', '中山', '惠州', '佛山', '珠海', '江门',
  '肇庆', '汕头', '潮州', '揭阳', '汕尾', '湛江', '茂名', '阳江',
  '云浮', '清远', '韶关', '河源', '梅州',
  '杭州', '宁波', '温州', '嘉兴', '湖州', '绍兴', '金华', '衢州',
  '台州', '丽水', '舟山', '义乌',
  '南京', '苏州', '无锡', '常州', '南通', '扬州', '镇江', '泰州',
  '盐城', '淮安', '连云港', '徐州', '宿迁',
  '石家庄', '唐山', '保定', '沧州', '邢台', '衡水', '廊坊', '邯郸',
  '秦皇岛', '承德', '张家口', '任丘', '河间', '清河',
  '合肥', '芜湖', '蚌埠', '淮南', '马鞍山', '淮北', '铜陵', '安庆',
  '黄山', '滁州', '阜阳', '宿州', '六安', '亳州', '池州', '宣城',
  '福州', '厦门', '泉州', '漳州', '莆田', '三明', '南平', '龙岩', '宁德',
  '济南', '青岛', '淄博', '枣庄', '东营', '烟台', '潍坊', '济宁',
  '泰安', '威海', '日照', '临沂', '德州', '聊城', '滨州', '菏泽',
  '郑州', '开封', '洛阳', '平顶山', '安阳', '鹤壁', '新乡', '焦作',
  '濮阳', '许昌', '漯河', '三门峡', '南阳', '商丘', '信阳', '周口', '驻马店',
  '武汉', '黄石', '十堰', '宜昌', '襄阳', '鄂州', '荆门', '孝感',
  '荆州', '黄冈', '咸宁', '随州', '恩施',
  '长沙', '株洲', '湘潭', '衡阳', '邵阳', '岳阳', '常德', '张家界',
  '益阳', '郴州', '永州', '怀化', '娄底',
  '成都', '自贡', '攀枝花', '泸州', '德阳', '绵阳', '广元', '遂宁',
  '内江', '乐山', '南充', '眉山', '宜宾', '广安', '达州', '雅安', '巴中', '资阳',
  '南昌', '景德镇', '萍乡', '九江', '新余', '鹰潭', '赣州', '吉安',
  '宜春', '抚州', '上饶',
  '南宁', '柳州', '桂林', '梧州', '北海', '防城港', '钦州', '贵港', '玉林', '贺州',
  '昆明', '曲靖', '玉溪', '保山', '昭通', '丽江', '普洱', '临沧',
  '贵阳', '六盘水', '遵义', '安顺', '毕节', '铜仁',
  '西安', '铜川', '宝鸡', '咸阳', '渭南', '延安', '汉中', '榆林', '安康', '商洛',
  '兰州', '嘉峪关', '金昌', '白银', '天水', '武威', '张掖', '平凉', '酒泉', '庆阳',
  '西宁', '海东',
  '呼和浩特', '包头', '乌海', '赤峰', '通辽', '鄂尔多斯', '呼伦贝尔', '巴彦淖尔', '乌兰察布',
  '银川', '石嘴山', '吴忠', '固原', '中卫',
  '拉萨', '日喀则', '昌都', '林芝', '山南', '那曲',
  '乌鲁木齐', '克拉玛依', '吐鲁番', '哈密',
  '长春', '吉林', '四平', '辽源', '通化', '白山', '松原', '白城', '延边',
  '哈尔滨', '齐齐哈尔', '鸡西', '鹤岗', '双鸭山', '大庆', '伊春', '佳木斯', '七台河', '牡丹江', '黑河', '绥化',
  '沈阳', '大连', '鞍山', '抚顺', '本溪', '丹东', '锦州', '营口', '阜新', '辽阳', '盘锦', '铁岭', '朝阳', '葫芦岛',
  '海口', '三亚', '三沙', '儋州',
  '太原', '大同', '阳泉', '长治', '晋城', '朔州', '晋中', '运城', '忻州', '临汾', '吕梁',
];

// ==================== 缓存管理 ====================

function loadCache() {
  try {
    if (fs.existsSync(CACHE_FILE)) {
      return JSON.parse(fs.readFileSync(CACHE_FILE, 'utf-8'));
    }
  } catch (e) {
    console.log('[供应商查询] 缓存读取失败:', e.message);
  }
  return {};
}

function saveCache(cache) {
  try {
    const dir = path.dirname(CACHE_FILE);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    fs.writeFileSync(CACHE_FILE, JSON.stringify(cache, null, 2), 'utf-8');
  } catch (e) {
    console.log('[供应商查询] 缓存写入失败:', e.message);
  }
}

// ==================== 城市提取 ====================

/**
 * 从供应商名称推断城市（提取名称中的城市前缀）
 * 与 product-lookup.js 中的 extractCity 逻辑一致
 */
function extractCityFromName(supplierName) {
  if (!supplierName) return null;

  // 匹配城市名称前缀（如"东莞市"、"东莞"、"深圳"等）
  for (const city of KNOWN_CITIES) {
    if (supplierName.startsWith(city + '市') || supplierName.startsWith(city)) {
      return { city, address: supplierName, source: 'name-prefix' };
    }
  }

  // 匹配 "XX省XX市" 模式
  const provinceCityMatch = supplierName.match(/^([一-鿿]{2,4})(?:省|自治区)([一-鿿]{2,4})(?:市|地区|州)/);
  if (provinceCityMatch && KNOWN_CITIES.includes(provinceCityMatch[2])) {
    return { city: provinceCityMatch[2], address: supplierName, source: 'name-province' };
  }

  // 匹配含 "XX市" 模式（可能不在开头）
  const cityMatch = supplierName.match(/([一-鿿]{2,4})市/);
  if (cityMatch && KNOWN_CITIES.includes(cityMatch[1])) {
    return { city: cityMatch[1], address: supplierName, source: 'name-contains' };
  }

  return null;
}

/**
 * 从任意文本中提取城市名
 */
function extractCityFromText(text) {
  if (!text) return null;
  for (const city of KNOWN_CITIES) {
    if (text.includes(city)) {
      const idx = text.indexOf(city);
      const start = Math.max(0, idx - 30);
      const end = Math.min(text.length, idx + city.length + 80);
      return { city, address: text.substring(start, end).replace(/\s+/g, '').trim() };
    }
  }
  return null;
}

// ==================== 在线搜索（尽力而为） ====================

/**
 * 搜索单个源
 */
function searchSource(hostname, path, timeout = 6000) {
  return new Promise((resolve) => {
    const req = https.get({
      hostname, path, method: 'GET', timeout,
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
        'Accept': 'text/html,application/xhtml+xml',
        'Accept-Language': 'zh-CN,zh;q=0.9',
      },
    }, (res) => {
      if (res.statusCode >= 300) { req.destroy(); resolve(null); return; }
      let data = '';
      res.on('data', chunk => { data += chunk.toString('utf8'); if (data.length > 200000) { req.destroy(); resolve(null); } });
      res.on('end', () => {
        // 验证码检测
        if (data.includes('验证') || data.includes('captcha') || data.includes('请输入')) { resolve(null); return; }
        const text = data.replace(/<[^>]+>/g, ' ').replace(/&[a-z]+;/g, ' ').replace(/\s+/g, ' ').trim();
        resolve(text);
      });
    });
    req.on('error', () => resolve(null));
    req.on('timeout', () => { req.destroy(); resolve(null); });
  });
}

/**
 * 在线查询（多源并发，第一个成功即返回）
 */
async function lookupOnline(name) {
  const query = encodeURIComponent('"' + name + '" 注册地址');
  const encodedCn = encodeURIComponent(name);

  // 多源并发
  const sources = await Promise.allSettled([
    // 搜狗搜索
    searchSource('www.sogou.com', '/web?query=' + encodeURIComponent(name + ' 注册地址')),
    // 360 搜索
    searchSource('www.so.com', '/s?q=' + query),
    // 必应国际版
    searchSource('www.bing.com', '/search?q=' + query + '&setlang=zh-hans&cc=sg'),
  ]);

  for (const s of sources) {
    if (s.status === 'fulfilled' && s.value) {
      const result = extractCityFromText(s.value);
      if (result) return { ...result, source: 'web-search' };
    }
  }

  return null;
}

// ==================== 公共 API ====================

/**
 * 查询单个供应商的公司所在地
 * @param {string} supplierName
 * @returns {Promise<{ city: string, source: string, address: string } | null>}
 */
async function lookupSupplierCity(supplierName) {
  if (!supplierName || supplierName.length < 3) return null;

  // 1. 本地缓存
  const cache = loadCache();
  if (cache[supplierName]) {
    return cache[supplierName];
  }

  // 2. 内置映射
  if (BUILTIN_MAP[supplierName]) {
    const city = BUILTIN_MAP[supplierName];
    const result = { city, source: 'builtin', address: supplierName, updatedAt: new Date().toISOString() };
    cache[supplierName] = result;
    saveCache(cache);
    return result;
  }

  // 3. 供应商名称中提取城市
  const nameResult = extractCityFromName(supplierName);
  if (nameResult) {
    const result = { ...nameResult, updatedAt: new Date().toISOString() };
    cache[supplierName] = result;
    saveCache(cache);
    return result;
  }

  // 4. 在线搜索
  const onlineResult = await lookupOnline(supplierName);
  if (onlineResult) {
    const result = { ...onlineResult, updatedAt: new Date().toISOString() };
    cache[supplierName] = result;
    saveCache(cache);
    return result;
  }

  return null;
}

/**
 * 批量查询
 * @param {string[]} suppliers
 * @returns {Promise<{ results: object, notFound: string[] }>}
 */
async function batchLookup(suppliers) {
  const results = {};
  const notFound = [];

  for (const name of suppliers) {
    try {
      const result = await lookupSupplierCity(name);
      if (result) {
        results[name] = result;
      } else {
        notFound.push(name);
      }
    } catch (err) {
      notFound.push(name);
    }
    // 避免请求过于密集
    await new Promise(r => setTimeout(r, 200));
  }

  return { results, notFound };
}

module.exports = { lookupSupplierCity, batchLookup, loadCache, saveCache, extractCityFromName, BUILTIN_MAP };
