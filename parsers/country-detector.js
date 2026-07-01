/**
 * 通用国家/地区检测器
 *
 * 根据 Sheet 名称自动识别线路国家，所有解析器共用。
 * 优先级：Sheet名 > 文件名
 */

const COUNTRY_PATTERNS = [
  { keywords: ["美国", "美线", "美西", "美中", "美东", "usa", "united states", "美森", "洛杉矶", "休斯顿", "芝加哥", "萨凡纳", "纽约"], country: "美国" },
  { keywords: ["英国", "英线", "uk", "united kingdom", "伦敦", "曼城", "费力克斯托", "felixstowe", "南安普顿", "southampton"], country: "英国" },
  { keywords: ["欧洲", "欧线", "europe", "eu", "德国", "法国", "意大利", "西班牙", "荷兰", "比利时", "波兰", "捷克", "匈牙利", "罗马尼亚"], country: "欧线" },
  { keywords: ["加拿大", "加线", "canada", "ca", "温哥华", "多伦多", "蒙特利尔", "vancouver", "toronto", "montreal"], country: "加拿大" },
  { keywords: ["墨西哥", "墨线", "mexico", "mx"], country: "墨西哥" },
  { keywords: ["巴西", "brazil", "br"], country: "巴西" },
  { keywords: ["澳大利亚", "澳洲", "澳线", "australia", "au", "悉尼", "墨尔本", "布里斯班", "sydney", "melbourne", "brisbane"], country: "澳大利亚" },
  { keywords: ["日本", "日线", "japan", "jp", "东京", "大阪"], country: "日本" },
];

/**
 * 从 Sheet 名称检测国家
 * @param {string} sheetName Excel Sheet 名称
 * @returns {string|null} 国家名，未识别返回 null
 */
function detectCountry(sheetName) {
  const n = sheetName.toLowerCase();
  for (const entry of COUNTRY_PATTERNS) {
    for (const kw of entry.keywords) {
      if (n.includes(kw.toLowerCase())) {
        return entry.country;
      }
    }
  }
  return null;
}

/**
 * 从文件名检测国家（兜底）
 * @param {string} fileName 文件名
 * @returns {string|null} 国家名，未识别返回 null
 */
function detectCountryFromFileName(fileName) {
  const n = fileName.toLowerCase();
  // 天图英国文件特殊处理：含"英国"且不含"美"
  if (n.includes("英国") && !n.includes("美")) return "英国";
  // 通用匹配
  for (const entry of COUNTRY_PATTERNS) {
    for (const kw of entry.keywords) {
      if (n.includes(kw.toLowerCase())) {
        return entry.country;
      }
    }
  }
  return null;
}

/**
 * 扫描 Excel 所有 Sheet，返回去重后的国家列表
 * @param {object} workbook XLSX workbook
 * @returns {string[]} 检测到的国家列表
 */
function detectAllCountries(workbook) {
  const countries = new Set();
  for (const name of workbook.SheetNames) {
    const c = detectCountry(name);
    if (c) countries.add(c);
  }
  return Array.from(countries);
}

module.exports = { detectCountry, detectCountryFromFileName, detectAllCountries, COUNTRY_PATTERNS };
