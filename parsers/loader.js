/**
 * 解析器加载器 —— 放在 parsers/ 目录避开 webpack 打包
 * 所有 require("xlsx") 调用均通过原生 Node.js require 解析
 *
 * 使用方式（从 upload route）:
 *   const { parseFile } = require(path.join(parsersDir, "loader.js"));
 *   const records = parseFile(filePath, supplierHint);
 */

const path = require("path");
const fs = require("fs");
const { createRequire } = require("module");

const parsersDir = __dirname;
const parsersRequire = createRequire(path.join(parsersDir, "_index.js"));

// ── 供应商识别（与 build_db.js 保持一致） ──
function identifySupplier(fileName) {
  const n = fileName.toLowerCase();
  // 跳过非价格表文件
  if (n.includes("出运计划") || n.includes("船期") || n.includes("schedule")) return "skip";
  // 天图英国（必须不含"美"）
  if ((n.includes("天图") || n.includes("tiantu")) && n.includes("英国") && !n.includes("美")) return "tiantu_uk";
  // 天图空运（必须在普通 tiantu 之前）
  if ((n.includes("天图") || n.includes("tiantu")) && (n.includes("空运") || n.includes("air"))) return "tiantu_air";
  if (n.includes("皓辉") || n.includes("haohui")) return "haohui";
  if (n.includes("皓鹏") || n.includes("haopeng")) return "haopeng";
  if (n.includes("星链") || n.includes("xinglian")) return "xinglian";
  if (n.includes("心一") || n.includes("xinyi")) return "xinyi";
  if (n.includes("航乐") || n.includes("hangle") || n.includes("yue")) return "hangle";
  if (n.includes("etton") || n.includes("易通")) return "etton";
  if (n.includes("天图") || n.includes("tiantu")) return "tiantu";
  if (n.includes("英美") || n.includes("yingmei")) return "yingmei";
  if (n.includes("丰运") || n.includes("fengyun")) return "fengyun";
  if (n.includes("华威尔") || n.includes("huaweier")) return "huaweier";
  if (n.includes("凯鑫") || n.includes("kaixin")) return "kaixin";
  if (n.includes("新胜") || n.includes("xinsheng")) return "xinsheng";
  if (n.includes("美琦") || n.includes("meiqi")) return "meiqi";
  if (n.includes("劲港") || n.includes("jingang")) return "jingang";
  if (n.includes("瑞秋") || n.includes("ruiqiu")) return "ruiqiu";
  if (n.includes("纽酷") || n.includes("niuku")) return "niuku";
  if (n.includes("博创兴") || n.includes("bochuangxing")) return "bochuangxing";
  if (n.includes("环洋") || n.includes("hye")) return "hye";
  return null;
}

// ── 解析器注册表 ──
const PARSER_REGISTRY = {
  etton:      { file: "etton_us.js",    exportName: "parseETTON" },
  tiantu:     { file: "tiantu_us.js",   exportName: "parseTiantu" },
  tiantu_uk:  { file: "tiantu_uk.js",   exportName: "parseTiantuUK" },
  tiantu_air: { file: "tiantu_air.js",  exportName: "parseTiantuAir" },
  yingmei:    { file: "yingmei_us.js",  exportName: "parseYingmei" },
  haohui:     { file: "haohui_us.js",   exportName: "parseHaohui" },
  haopeng:    { file: "haopeng_us.js",  exportName: "parseHaopeng" },
  xinglian:   { file: "xinglian_us.js", exportName: "parseXinglian" },
  xinyi:      { file: "xinyi_eu.js",    exportName: "parseXinyi" },
  hangle:     { file: "hangle.js",      exportName: "parseHangle" },
  fengyun:    { file: "fengyun.js",     exportName: "parseXinyun" },
  huaweier:   { file: "huaweier.js",    exportName: "parseHuaweier" },
  kaixin:     { file: "kaixin.js",      exportName: "parseKaixin" },
  xinsheng:   { file: "xinsheng.js",    exportName: "parseXinsheng" },
  meiqi:      { file: "meiqi_us.js",    exportName: "parseMeiQi" },
  jingang:    { file: "jingang.js",     exportName: "parseJingang" },
  ruiqiu:     { file: "ruiqiu.js",      exportName: "parseRuiqiu" },
  niuku:      { file: "niuku.js",       exportName: "parseNiuku" },
  bochuangxing: { file: "bochuangxing.js", exportName: "parseBochuangxing" },
  hye:        { file: "hye.js",         exportName: "parseHYE" },
};

// ── 多线路供应商分组 ──
const SUPPLIER_PARSER_GROUP = {
  tiantu: ["tiantu", "tiantu_uk", "tiantu_air"],
};

function getParserKeys(supplier) {
  const base = supplier.replace(/_(uk|air|us)$/, "");
  return SUPPLIER_PARSER_GROUP[base] || [supplier];
}

/**
 * 加载并执行指定解析器
 * @param {string} fileName - 解析器文件名 (如 "tiantu_us.js")
 * @param {string} exportName - 导出函数名 (如 "parseTiantu")
 * @param {string} filePath - Excel 文件路径
 * @returns {Array} 解析结果
 */
function loadAndParse(fileName, exportName, filePath) {
  const mod = parsersRequire("./" + fileName);
  if (typeof mod[exportName] !== "function") {
    throw new Error(`${fileName} 未导出 ${exportName}`);
  }
  return mod[exportName](filePath);
}

/**
 * 主解析入口 —— 自动识别供应商并解析 Excel 文件
 * @param {string} filePath - Excel 文件路径
 * @param {string} supplierHint - 可选供应商提示
 * @returns {Array} 价格记录数组
 */
function parseFile(filePath, supplierHint) {
  const fileName = path.basename(filePath);
  let supplier = supplierHint || "";

  // 自动识别供应商
  if (!supplier) {
    const identified = identifySupplier(fileName);
    if (!identified) {
      throw new Error(
        `无法识别供应商，文件名需包含供应商标识（如 ETTON/易通、天图/tiantu、英美/yingmei、皓辉/haohui、皓鹏/haopeng、星链/xinglian 等）`
      );
    }
    if (identified === "skip") {
      throw new Error(`非价格表文件（船期/出运计划），已跳过: ${fileName}`);
    }
    supplier = identified;
  }

  // 提取生效日期
  let effectiveDate = "";
  const dm1 = fileName.match(/(\d{4})[年.-]?(\d{1,2})[月.-]?(\d{1,2})/);
  if (dm1) {
    effectiveDate = `${dm1[1]}-${String(dm1[2]).padStart(2, "0")}-${String(dm1[3]).padStart(2, "0")}`;
  } else {
    const dm2 = fileName.match(/(\d{1,2})[.·](\d{1,2})/);
    if (dm2) {
      effectiveDate = `${new Date().getFullYear()}-${String(parseInt(dm2[1])).padStart(2, "0")}-${String(parseInt(dm2[2])).padStart(2, "0")}`;
    }
  }

  // 多线路供应商：依次尝试所有子解析器
  const parserKeys = getParserKeys(supplier);
  const allResults = [];
  const parsedLines = [];
  const parserErrors = [];

  for (const key of parserKeys) {
    const entry = PARSER_REGISTRY[key];
    if (!entry) { parserErrors.push(`${key}: 无注册信息`); continue; }

    try {
      const mod = parsersRequire("./" + entry.file);
      const parseFn = mod[entry.exportName];
      if (typeof parseFn !== "function") {
        const msg = `${key}: ${entry.file} 未导出 ${entry.exportName}`;
        console.warn(`[loader] ⚠ ${msg}`);
        parserErrors.push(msg);
        continue;
      }

      const results = parseFn(filePath);
      if (results.length > 0) {
        for (const r of results) {
          r.source_file = fileName;
          r.effective_date = effectiveDate;
        }
        allResults.push(...results);
        parsedLines.push(`${key}(${results.length}条)`);
        console.log(`[loader]   ✅ ${key}: ${results.length} 条`);
      } else {
        parserErrors.push(`${key}: 解析完成但返回0条数据`);
      }
    } catch (err) {
      const msg = `${key}: ${err.message}`;
      console.error(`[loader]   ❌ ${msg}`);
      if (err.stack) console.error(`[loader]   📚 ${err.stack.split('\n').slice(0, 4).join('\n')}`);
      parserErrors.push(msg);
    }
  }

  if (allResults.length === 0) {
    const detail = parserErrors.length > 0 ? ` 详情: ${parserErrors.join("; ")}` : "";
    throw new Error(`文件 "${fileName}" 未能解析出任何价格数据。${detail}`);
  }

  console.log(`[loader] 📊 总计: ${allResults.length} 条 (线路: ${parsedLines.join(" + ")})`);
  return allResults;
}

module.exports = { loadAndParse, parsersRequire, parseFile, identifySupplier, PARSER_REGISTRY };
