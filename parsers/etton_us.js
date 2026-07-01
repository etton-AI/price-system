/**
 * ETTON 易通科技 — 美线海运价格解析器
 *
 * 处理 ETTON Excel 中 12 个美线价格 Sheet，每个 Sheet 含 1-2 个渠道并排排列。
 * 仓库代码在 A 列，含 "/" 的需拆分为多条记录。
 * 国内仓固定 4 区域：东莞/中山/广州 | 嘉兴/义乌 | 汕头/厦门/泉州 | 武汉/长沙
 */

const XLSX = require("xlsx");
const { detectCountry } = require("./country-detector");

// ── 常量 ──────────────────────────────────────────
const SUPPLIER = "易通ETTON";
const COUNTRY = "美国";

const REGIONS = [
  "东莞/中山/广州",
  "嘉兴/义乌",
  "汕头/厦门/泉州",
  "武汉/长沙",
];

// 船配置关键词标签（小写，用于搜索匹配）
const VESSEL_TAG_MAP = {
  CLX: ["matson", "clx"],
  MAX: ["matson", "max"],
  EXX: ["exx", "zim", "合德"],
  以星: ["zim", "合德"],
  合德: ["合德"],
  COSCO: ["cosco", "oa"],
  OA: ["oa", "msc", "whl"],
  MSC: ["msc", "oa", "whl"],
  ONE: ["one"],
  WHL: ["whl"],
  HMM: ["hmm"],
  TSL: ["tsl"],
  ZIM: ["zim"],
  OOCL: ["oocl"],
  海领: ["海领"],
};

function extractVesselTags(vesselConfig) {
  if (!vesselConfig) return [];
  const lower = vesselConfig.toLowerCase();
  const tags = new Set();
  for (const [key, vals] of Object.entries(VESSEL_TAG_MAP)) {
    if (lower.includes(key.toLowerCase())) {
      vals.forEach((v) => tags.add(v));
    }
  }
  return [...tags];
}

// ── 渠道 → 统一维度映射 ──────────────────────────
function inferChannelMeta(channelName) {
  const meta = { speed_tier: "", vessel_config: "", delivery_method: "" };

  // 送仓方式
  if (channelName.includes("海派")) {
    meta.delivery_method = "海派(快递派)";
  } else if (
    channelName.includes("整柜直送") ||
    channelName.includes("直送")
  ) {
    meta.delivery_method = "整柜直送";
  } else {
    meta.delivery_method = "卡派(拆派)";
  }

  // 时效层级 + 船配置
  if (channelName.includes("12日达")) {
    meta.speed_tier = "12-15日达";
    meta.vessel_config = "美森正班CLX";
  } else if (channelName.includes("15日达")) {
    meta.speed_tier = "12-15日达";
    meta.vessel_config = "美森正班CLX";
  } else if (channelName.includes("17日达")) {
    meta.speed_tier = "17-20日达";
    meta.vessel_config = "EXX/以星带托架/合德快提统配";
  } else if (channelName.includes("20日达")) {
    meta.speed_tier = "17-20日达";
    meta.vessel_config = "以星不带托架/合德普提/SEA3/CEN";
  } else if (channelName.includes("22日达")) {
    meta.speed_tier = "22-28日达";
    meta.vessel_config = "COSCO/OA";
  } else if (channelName.includes("25日达")) {
    meta.speed_tier = "22-28日达";
    meta.vessel_config = "OA/MSC/ONE/WHL/HMM/TSL普船统配";
  } else if (channelName.includes("26日达") || channelName.includes("FBA全美")) {
    meta.speed_tier = "22-28日达";
    meta.vessel_config = "OA/ZIM(靠洛杉矶港)";
  } else if (channelName.includes("28日达")) {
    meta.speed_tier = "22-28日达";
    meta.vessel_config = "OA/MSC/WHL统配(LA)";
  } else if (channelName.includes("美西北")) {
    meta.speed_tier = "22-28日达";
    meta.vessel_config = "OA/ONE/MSC/HMM统配";
  } else if (channelName.includes("芝加哥")) {
    meta.speed_tier = "22-28日达";
    meta.vessel_config = "OA/ZIM/MSC统配";
  } else if (channelName.includes("纽约")) {
    meta.speed_tier = "40-45日达";
    meta.vessel_config = "OA/ZIM/MSC统配(NY)";
  } else if (channelName.includes("萨凡纳")) {
    meta.speed_tier = "40-45日达";
    meta.vessel_config = "OA/ZIM/MSC/WHL统配(SAV)";
  } else if (channelName.includes("休斯顿")) {
    meta.speed_tier = "40-45日达";
    meta.vessel_config = "OA/ZIM/MSC统配(HOU)";
  } else if (channelName.includes("东部稳速达")) {
    meta.speed_tier = "22-28日达";
    meta.vessel_config = "OA/MSC/WHL统配";
  } else if (channelName.includes("DG海卡")) {
    meta.speed_tier = "22-28日达";
    meta.vessel_config = "OA/MSC/WHL统配";
  }

  return meta;
}

// ── 目的地区域推断 ──────────────────────────────
function inferRegion(destCode, sheetName) {
  const sheet = sheetName.toLowerCase();
  if (sheet.includes("美西")) return "美西";
  if (sheet.includes("美西北")) return "美西北";
  if (sheet.includes("芝加哥") || sheet.includes("美中")) return "美中";
  if (sheet.includes("休斯顿")) return "美东南";
  if (sheet.includes("纽约") || sheet.includes("美东") || sheet.includes("东部")) return "美东北";
  if (sheet.includes("萨凡纳")) return "美东南";
  if (sheet.includes("fba")) return "美西"; // FBA渠道默认美西
  if (sheet.includes("整柜直送")) return "美西";
  return "美西";
}

// ── 解析时效文本 ─────────────────────────────────
function parseTransitTime(text) {
  if (!text) return { min: null, max: null, desc: "" };
  const str = String(text).trim();
  const desc = str;
  // 匹配 "12-13自然日" / "12-13天" / "26-28自然日入仓"
  const match = str.match(/(\d+)\s*[-–]\s*(\d+)/);
  if (match) {
    return { min: parseInt(match[1]), max: parseInt(match[2]), desc };
  }
  // 单天数
  const single = str.match(/(\d+)/);
  if (single) {
    return { min: parseInt(single[1]), max: parseInt(single[1]), desc };
  }
  return { min: null, max: null, desc };
}

// ── 解析理赔规则 ─────────────────────────────────
function parseClaimRule(text) {
  if (!text) return "";
  const str = String(text).trim();
  if (str.includes("无") || str === "-" || str === "") return "";
  return str.replace(/\n/g, "");
}

// ── Sheet 配置: 美西12-15日达, 17-20日达, 22-25日达 ─
// 这三个 Sheet 为双渠道并排布局，每个渠道 4 区域
// 12-15日达: 12日达(KG only) | 15日达(KG+CBM)
// 17-20日达: 17日达(KG+CBM) | 20日达(KG+CBM)
// 22-25日达: 22日达(KG+CBM) | 25日达(KG+CBM)

function parseDualChannelSheet(ws, sheetName, channels) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 6) return [];

  // 从 row[4] 表头获取各列的计费方式
  const headerRow = data[4]; // row index 4 (0-based) = row 5 in Excel

  const results = [];

  for (const ch of channels) {
    for (let ri = 5; ri < data.length; ri++) {
      const row = data[ri];
      const wareCell = String(row[0] || "").trim();
      if (!wareCell || wareCell.includes("海外仓自提")) continue;

      const transitText = String(row[ch.transitCol] || "").trim();
      const claimText = String(row[ch.claimCol] || "").trim();
      const transit = parseTransitTime(transitText);
      const claim = parseClaimRule(claimText);

      for (const pg of ch.priceGroups) {
        let billingType = "含税KG";
        let priceUnit = "元/KG";
        let minQty = "12KG+";
        let minQtyVal = 12;

        // 从表头判断计费方式
        const hdrText = String(headerRow[pg.col] || "");
        if (hdrText.includes("不含税") || hdrText.includes("CBM")) {
          billingType = "不含税CBM";
          priceUnit = "元/CBM";
          minQty = "0.5CBM+";
          minQtyVal = 0.5;
        } else if (hdrText.includes("含税")) {
          billingType = "含税KG";
          priceUnit = "元/KG";
        }

        const price = parseFloat(row[pg.col]);
        if (isNaN(price) || price <= 0) continue;

        const meta = inferChannelMeta(ch.name);
        const destCodes = wareCell.split("/").map((s) => s.trim()).filter(Boolean);

        for (const dc of destCodes) {
          results.push({
            supplier: SUPPLIER,
            country: COUNTRY,
            channel_name: ch.name,
            speed_tier: meta.speed_tier,
            vessel_config: meta.vessel_config,
            vessel_tags: extractVesselTags(meta.vessel_config),
            delivery_method: meta.delivery_method,
            destination_type: "warehouse",
            destination_code: dc,
            destination_region: inferRegion(dc, sheetName),
            origin_region: pg.region,
            origin_cities: pg.region.split("/").map((s) => s.trim()),
            billing_type: billingType,
            min_quantity: minQty,
            min_quantity_value: minQtyVal,
            unit_price: price,
            price_unit: priceUnit,
            transit_time_min: transit.min,
            transit_time_max: transit.max,
            transit_time_desc: transit.desc,
            claim_rule: claim,
            effective_date: "",
            source_sheet: sheetName,
          });
        }
      }
    }
  }
  return results;
}

// ── FBA渠道, DG海卡, 美西北专线 等单渠道 Sheet ──
function parseSingleChannelSheet(ws, sheetName, config) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 5) return [];

  const results = [];
  // 行4通常是表头
  const headerRow = data[4] || [];

  for (let ri = config.dataStartRow || 5; ri < data.length; ri++) {
    const row = data[ri];
    const wareCell = String(row[0] || "").trim();
    if (!wareCell || wareCell.includes("海外仓自提") || wareCell.includes("全美所有")) continue;

    if (wareCell.startsWith("全美所有")) {
      // 这是一个汇总行，跳过（已在其他Sheet中覆盖各仓）
      continue;
    }

    const transitText = String(row[config.transitCol] || "").trim();
    const claimText = String(row[config.claimCol] || "").trim();
    const transit = parseTransitTime(transitText);
    const claim = parseClaimRule(claimText);

    const meta = inferChannelMeta(config.channelName);

    for (const pg of config.priceGroups) {
      const price = parseFloat(row[pg.col]);
      if (isNaN(price) || price <= 0) continue;

      const destCodes = wareCell.split("/").map((s) => s.trim()).filter(Boolean);

      for (const dc of destCodes) {
        results.push({
          supplier: SUPPLIER,
          country: COUNTRY,
          channel_name: config.channelName,
          speed_tier: meta.speed_tier,
          vessel_config: meta.vessel_config,
          vessel_tags: extractVesselTags(meta.vessel_config),
          delivery_method: meta.delivery_method,
          destination_type: "warehouse",
          destination_code: dc,
          destination_region: inferRegion(dc, sheetName),
          origin_region: pg.region,
          origin_cities: pg.region.split("/").map((s) => s.trim()),
          billing_type: pg.billingType,
          min_quantity: pg.minQuantity,
          min_quantity_value: pg.minQuantityValue,
          unit_price: price,
          price_unit: pg.priceUnit,
          transit_time_min: transit.min,
          transit_time_max: transit.max,
          transit_time_desc: transit.desc,
          claim_rule: claim,
          effective_date: "",
          source_sheet: sheetName,
        });
      }
    }
  }
  return results;
}

// ── 美国海派 (ZIP code zone based) ──
function parseSeaParcelSheet(ws, sheetName) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 5) return [];

  const results = [];

  // 海派结构特殊: 行3起，每行一个邮编分区 × 4区域 × 2重量段(12-50KG / 50KG+)
  // 渠道名在行2中
  let currentChannel = null;
  let currentVessel = null;

  for (let ri = 2; ri < data.length; ri++) {
    const row = data[ri];
    const col0 = String(row[0] || "").trim();
    const col1 = String(row[1] || "").trim();

    // 检测渠道名称行 (如 "美西12-15日达海派" 或 "易·12日达海派")
    if (col0.includes("渠道") || col0.includes("使用船司")) {
      if (col0.includes("使用船司") && col1) {
        currentVessel = col1.trim();
      }
      continue;
    }

    // 区域表头行
    if (col0.includes("交货仓库") || col0.includes("国内仓") || col0.includes("东莞")) {
      continue;
    }

    // 邮编分区行
    if (col0.includes("邮编") || col0.includes("开头") || col0.match(/^\d/)) {
      const zipZone = col0;
      // 行结构: col0=邮编分区, col1-2=区域A(12-50KG,50KG+), col3-4=区域B, ...
      const regionCols = [
        { region: REGIONS[0], kg12Col: 1, kg50Col: 2 },
        { region: REGIONS[1], kg12Col: 3, kg50Col: 4 },
        { region: REGIONS[2], kg12Col: 5, kg50Col: 6 },
        { region: REGIONS[3], kg12Col: 7, kg50Col: 8 },
      ];

      // 从附近行提取渠道信息
      // 简化处理：从sheet名推断渠道系列
      const channelSeries = sheetName; // "美国海派"

      // 需要从sheet上下文中找到具体的渠道名
      // 这里我们只解析数据，渠道名从上下文推断
      for (const rc of regionCols) {
        // 12-50KG
        const price12 = parseFloat(row[rc.kg12Col]);
        if (!isNaN(price12) && price12 > 0) {
          results.push({
            supplier: SUPPLIER,
            country: COUNTRY,
            channel_name: channelSeries,
            speed_tier: "",
            vessel_config: "",
            vessel_tags: [],
            delivery_method: "海派(快递派)",
            destination_type: "zip_zone",
            destination_code: zipZone,
            destination_region: "",
            origin_region: rc.region,
            origin_cities: rc.region.split("/").map((s) => s.trim()),
            billing_type: "海派KG",
            min_quantity: "12-50KG",
            min_quantity_value: 12,
            unit_price: price12,
            price_unit: "元/KG",
            transit_time_min: null,
            transit_time_max: null,
            transit_time_desc: "",
            claim_rule: "",
            effective_date: "",
            source_sheet: sheetName,
          });
        }

        // 50KG+
        const price50 = parseFloat(row[rc.kg50Col]);
        if (!isNaN(price50) && price50 > 0) {
          results.push({
            supplier: SUPPLIER,
            country: COUNTRY,
            channel_name: channelSeries,
            speed_tier: "",
            vessel_config: "",
            vessel_tags: [],
            delivery_method: "海派(快递派)",
            destination_type: "zip_zone",
            destination_code: zipZone,
            destination_region: "",
            origin_region: rc.region,
            origin_cities: rc.region.split("/").map((s) => s.trim()),
            billing_type: "海派KG",
            min_quantity: "50KG+",
            min_quantity_value: 50,
            unit_price: price50,
            price_unit: "元/KG",
            transit_time_min: null,
            transit_time_max: null,
            transit_time_desc: "",
            claim_rule: "",
            effective_date: "",
            source_sheet: sheetName,
          });
        }
      }
    }
  }

  return results;
}

// ── 美西整柜直送 ──────────────────────────────
function parseContainerDirectSheet(ws, sheetName) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 5) return [];

  const results = [];
  // 结构: col0=仓库, col1-2=区域A(KG,CBM), col3-4=区域B, ...

  for (let ri = 4; ri < data.length; ri++) {
    const row = data[ri];
    const wareCell = String(row[0] || "").trim();
    if (!wareCell || wareCell === "仓库代码") continue;

    const priceGroups = [
      { region: REGIONS[0], kgCol: 1, cbmCol: 2 },
      { region: REGIONS[1], kgCol: 3, cbmCol: 4 },
      { region: REGIONS[2], kgCol: 5, cbmCol: 6 },
      { region: REGIONS[3], kgCol: 7, cbmCol: 8 },
    ];

    const destCodes = wareCell.split("/").map((s) => s.trim()).filter(Boolean);

    for (const pg of priceGroups) {
      const kgPrice = parseFloat(row[pg.kgCol]);
      if (!isNaN(kgPrice) && kgPrice > 0) {
        for (const dc of destCodes) {
          results.push({
            supplier: SUPPLIER,
            country: COUNTRY,
            channel_name: "美西整柜直送",
            speed_tier: "22-28日达",
            vessel_config: "普船统配",
            vessel_tags: [],
            delivery_method: "整柜直送",
            destination_type: "warehouse",
            destination_code: dc,
            destination_region: "美西",
            origin_region: pg.region,
            origin_cities: pg.region.split("/").map((s) => s.trim()),
            billing_type: "含税KG",
            min_quantity: "1CBM+",
            min_quantity_value: 1,
            unit_price: kgPrice,
            price_unit: "元/KG",
            transit_time_min: 28,
            transit_time_max: 30,
            transit_time_desc: "28-30自然日入仓",
            claim_rule: "",
            effective_date: "",
            source_sheet: sheetName,
          });
        }
      }
      const cbmPrice = parseFloat(row[pg.cbmCol]);
      if (!isNaN(cbmPrice) && cbmPrice > 0) {
        for (const dc of destCodes) {
          results.push({
            supplier: SUPPLIER,
            country: COUNTRY,
            channel_name: "美西整柜直送",
            speed_tier: "22-28日达",
            vessel_config: "普船统配",
            vessel_tags: [],
            delivery_method: "整柜直送",
            destination_type: "warehouse",
            destination_code: dc,
            destination_region: "美西",
            origin_region: pg.region,
            origin_cities: pg.region.split("/").map((s) => s.trim()),
            billing_type: "不含税CBM",
            min_quantity: "0.5CBM+",
            min_quantity_value: 0.5,
            unit_price: cbmPrice,
            price_unit: "元/CBM",
            transit_time_min: 28,
            transit_time_max: 30,
            transit_time_desc: "28-30自然日入仓",
            claim_rule: "",
            effective_date: "",
            source_sheet: sheetName,
          });
        }
      }
    }
  }
  return results;
}

// ── 主解析入口 ──────────────────────────────────
function parseETTON(filePath) {
  console.log("[ETTON] 开始解析:", filePath);
  const wb = XLSX.readFile(filePath);
  const allResults = [];

  // ── 美西12-15日达 (双渠道: 12日达 + 15日达) ──
  if (wb.SheetNames.includes("美西12-15日达")) {
    const results = parseDualChannelSheet(wb.Sheets["美西12-15日达"], "美西12-15日达", [
      {
        name: "易·12日达卡派",
        priceGroups: [
          { col: 1, region: REGIONS[0] },
          { col: 2, region: REGIONS[1] },
          { col: 3, region: REGIONS[2] },
          { col: 4, region: REGIONS[3] },
        ],
        transitCol: 5,
        claimCol: 6,
      },
      {
        name: "易·15日达卡派",
        priceGroups: [
          { col: 7, region: REGIONS[0] },
          { col: 8, region: REGIONS[0] }, // 0.5CBM+不含税
          { col: 9, region: REGIONS[1] },
          { col: 10, region: REGIONS[1] }, // 0.5CBM+不含税
          { col: 11, region: REGIONS[2] },
          { col: 12, region: REGIONS[2] }, // 0.5CBM+不含税
          { col: 13, region: REGIONS[3] },
          { col: 14, region: REGIONS[3] }, // 0.5CBM+不含税
        ],
        transitCol: 15,
        claimCol: 16,
      },
    ]);
    console.log(`  [美西12-15日达] ${results.length} 条`);
    allResults.push(...results);
  }

  // ── 美西17-20日达 (双渠道: 17日达 + 20日达) ──
  if (wb.SheetNames.includes("美西17-20日达")) {
    const results = parseDualChannelSheet(wb.Sheets["美西17-20日达"], "美西17-20日达", [
      {
        name: "易·17日达卡派",
        priceGroups: [
          { col: 1, region: REGIONS[0] }, // 12KG+含税
          { col: 2, region: REGIONS[0] }, // 0.5CBM+不含税
          { col: 3, region: REGIONS[1] },
          { col: 4, region: REGIONS[1] },
          { col: 5, region: REGIONS[2] },
          { col: 6, region: REGIONS[2] },
          { col: 7, region: REGIONS[3] },
          { col: 8, region: REGIONS[3] },
        ],
        transitCol: 9,
        claimCol: 10,
      },
      {
        name: "易·20日达卡派",
        priceGroups: [
          { col: 11, region: REGIONS[0] },
          { col: 12, region: REGIONS[0] },
          { col: 13, region: REGIONS[1] },
          { col: 14, region: REGIONS[1] },
          { col: 15, region: REGIONS[2] },
          { col: 16, region: REGIONS[2] },
          { col: 17, region: REGIONS[3] },
          { col: 18, region: REGIONS[3] },
        ],
        transitCol: 19,
        claimCol: 20,
      },
    ]);
    console.log(`  [美西17-20日达] ${results.length} 条`);
    allResults.push(...results);
  }

  // ── 美西22-25日达 (双渠道: 22日达 + 25日达) ──
  if (wb.SheetNames.includes("美西22-25日达")) {
    const results = parseDualChannelSheet(wb.Sheets["美西22-25日达"], "美西22-25日达", [
      {
        name: "易·22日达卡派",
        priceGroups: [
          { col: 1, region: REGIONS[0] },
          { col: 2, region: REGIONS[0] },
          { col: 3, region: REGIONS[1] },
          { col: 4, region: REGIONS[1] },
          { col: 5, region: REGIONS[2] },
          { col: 6, region: REGIONS[2] },
          { col: 7, region: REGIONS[3] },
          { col: 8, region: REGIONS[3] },
        ],
        transitCol: 9,
        claimCol: 10,
      },
      {
        name: "易·25日达卡派",
        priceGroups: [
          { col: 11, region: REGIONS[0] },
          { col: 12, region: REGIONS[0] },
          { col: 13, region: REGIONS[1] },
          { col: 14, region: REGIONS[1] },
          { col: 15, region: REGIONS[2] },
          { col: 16, region: REGIONS[2] },
          { col: 17, region: REGIONS[3] },
          { col: 18, region: REGIONS[3] },
        ],
        transitCol: 19,
        claimCol: 20,
      },
    ]);
    console.log(`  [美西22-25日达] ${results.length} 条`);
    allResults.push(...results);
  }

  // ── FBA渠道 ──
  if (wb.SheetNames.includes("FBA渠道")) {
    const results = parseSingleChannelSheet(wb.Sheets["FBA渠道"], "FBA渠道", {
      channelName: "FBA美西28日达",
      dataStartRow: 5,
      transitCol: 9,
      claimCol: 9,
      priceGroups: [
        { col: 1, region: REGIONS[0], billingType: "含税KG", minQuantity: "12KG+", minQuantityValue: 12, priceUnit: "元/KG" },
        { col: 2, region: REGIONS[0], billingType: "不含税CBM", minQuantity: "1CBM+", minQuantityValue: 1, priceUnit: "元/CBM" },
        { col: 3, region: REGIONS[1], billingType: "含税KG", minQuantity: "12KG+", minQuantityValue: 12, priceUnit: "元/KG" },
        { col: 4, region: REGIONS[1], billingType: "不含税CBM", minQuantity: "1CBM+", minQuantityValue: 1, priceUnit: "元/CBM" },
        { col: 5, region: REGIONS[2], billingType: "含税KG", minQuantity: "12KG+", minQuantityValue: 12, priceUnit: "元/KG" },
        { col: 6, region: REGIONS[2], billingType: "不含税CBM", minQuantity: "1CBM+", minQuantityValue: 1, priceUnit: "元/CBM" },
        { col: 7, region: REGIONS[3], billingType: "含税KG", minQuantity: "12KG+", minQuantityValue: 12, priceUnit: "元/KG" },
        { col: 8, region: REGIONS[3], billingType: "不含税CBM", minQuantity: "1CBM+", minQuantityValue: 1, priceUnit: "元/CBM" },
      ],
    });
    console.log(`  [FBA渠道] ${results.length} 条`);
    allResults.push(...results);
  }

  // ── 美西整柜直送 ──
  if (wb.SheetNames.includes("美西整柜直送")) {
    const results = parseContainerDirectSheet(wb.Sheets["美西整柜直送"], "美西整柜直送");
    console.log(`  [美西整柜直送] ${results.length} 条`);
    allResults.push(...results);
  }

  // ── 专线 Sheets (单渠道, KG only) ──
  const singleSheets = [
    { name: "美西北专线", channel: "美西北专线卡派" },
    { name: "芝加哥专线", channel: "芝加哥专线卡派" },
    { name: "休斯顿专线", channel: "休斯顿专线卡派" },
    { name: "纽约专线", channel: "纽约专线卡派" },
    { name: "萨凡纳专线", channel: "萨凡纳专线卡派" },
    { name: "DG海卡", channel: "DG海卡" },
  ];

  for (const ss of singleSheets) {
    if (wb.SheetNames.includes(ss.name)) {
      // 专线 sheet 结构: 4 区域 × 含税KG，各一列
      const results = parseSingleChannelSheet(wb.Sheets[ss.name], ss.name, {
        channelName: ss.channel,
        dataStartRow: 5,
        transitCol: 5,
        claimCol: 6,
        priceGroups: [
          { col: 1, region: REGIONS[0], billingType: "含税KG", minQuantity: "12KG+", minQuantityValue: 12, priceUnit: "元/KG" },
          { col: 2, region: REGIONS[1], billingType: "含税KG", minQuantity: "12KG+", minQuantityValue: 12, priceUnit: "元/KG" },
          { col: 3, region: REGIONS[2], billingType: "含税KG", minQuantity: "12KG+", minQuantityValue: 12, priceUnit: "元/KG" },
          { col: 4, region: REGIONS[3], billingType: "含税KG", minQuantity: "12KG+", minQuantityValue: 12, priceUnit: "元/KG" },
        ],
      });
      console.log(`  [${ss.name}] ${results.length} 条`);
      allResults.push(...results);
    }
  }

  // ── 加拿大 + 澳洲 Sheets（自动检测）──
  const caAuSheets = wb.SheetNames.filter((sn) => {
    const c = detectCountry(sn);
    return c === "加拿大" || c === "澳大利亚" || sn.includes("澳洲");
  });
  // 也检查文件名的国家
  const hasCA = caAuSheets.length > 0;
  // 如果没有按国家检测到，检查Sheet名关键词
  const extraSheets = wb.SheetNames.filter((sn) =>
    sn.includes("美转加") || sn.includes("澳洲") || sn.includes("加拿大")
  );
  const allExtraSheets = [...new Set([...caAuSheets, ...extraSheets])];

  for (const sn of allExtraSheets) {
    // Skip metadata sheets
    if (sn.includes("须知") || sn.includes("仓库列表") || sn.includes("偏远") || sn.includes("增值服务")
      || sn.includes("分区表") || sn.includes("区域表") || sn.includes("DG货物") || sn.includes("船期")
      || sn.includes("查询") || sn.includes("目录")) continue;

    try {
      const country = detectCountry(sn) === "加拿大" || sn.includes("加拿大") || sn.includes("美转加") ? "加拿大"
        : detectCountry(sn) === "澳大利亚" || sn.includes("澳洲") ? "澳大利亚" : "加拿大";
      // Skip 香港UPS红单 for now
      if (sn.includes("香港UPS")) continue;

      const results = parseETTON_CA_AU(wb.Sheets[sn], sn, country, {});
      console.log(`  [${sn}] ${results.length} 条 → ${country}`);
      allResults.push(...results);
    } catch (err) {
      console.error(`  [${sn}] 解析失败: ${err.message}`);
    }
  }

  console.log(`[ETTON] 总计解析 ${allResults.length} 条价格记录`);
  return allResults;
}

// ── 加拿大+澳洲解析器 ──────────────────────────────
// ETTON 加拿大澳洲价格表格式: R1-R2=渠道名+船司, R3-R4=城市+重量段表头, R5+=数据行
function parseETTON_CA_AU(ws, sheetName, countryOverride, defaults) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 6) return [];
  const results = [];

  // Extract channel name from R1
  let channelName = "";
  const r1 = data[1] || [];
  const r2 = data[2] || [];
  for (let i = 1; i < r1.length; i++) {
    const c = String(r1[i] || "").trim();
    if (c && c.length > 2 && !c.includes("DISPIMG")) { channelName = c; break; }
  }
  if (!channelName) {
    // Try from R2
    for (let i = 1; i < r2.length; i++) {
      const c = String(r2[i] || "").trim();
      if (c && c.length > 3 && !c.includes("DISPIMG")) { channelName = c; break; }
    }
  }
  if (!channelName) channelName = sheetName;

  // Extract vessel info from R2
  let vesselConfig = "";
  for (let i = 1; i < r2.length; i++) {
    const c = String(r2[i] || "").trim();
    if (c && (c.includes("美森") || c.includes("OA") || c.includes("COSCO") || c.includes("船司") || c.includes("ERS") || c.includes("统配"))) {
      vesselConfig = c.replace("使用船司", "").trim();
      break;
    }
  }

  // Find city groups from R3-R5
  // R3 typically has city names, R4/R5 has weight tier labels
  const r3 = data[3] || [];
  const r4 = data[4] || [];
  const r5 = data[5] || [];

  // Detect city columns: look for Chinese city names or region labels in R3
  const cityGroups = [];
  let currentCity = null;
  for (let col = 1; col < Math.min(r3.length, 20); col++) {
    const cell = String(r3[col] || "").trim();
    if (cell && (cell.includes("/") || cell.includes("东莞") || cell.includes("嘉兴") || cell.includes("汕头") || cell.includes("武汉") || cell.includes("青岛") || cell.includes("华东") || cell.includes("华南"))) {
      if (currentCity) {
        // Finish previous city group
        currentCity.endCol = col - 1;
        cityGroups.push(currentCity);
      }
      currentCity = { name: cell, startCol: col, endCol: col };
    }
  }
  if (currentCity) {
    currentCity.endCol = Math.min(r3.length, r4.length) - 1;
    cityGroups.push(currentCity);
  }

  // If no cities detected, use default single group
  if (cityGroups.length === 0) {
    cityGroups.push({ name: "华南", startCol: 2, endCol: Math.min(r4.length - 1, 10) });
  }

  // For each city group, detect weight tiers from R4/R5
  for (const cg of cityGroups) {
    cg.tiers = [];
    for (let col = cg.startCol; col <= cg.endCol; col++) {
      const r4cell = String(r4[col] || "").trim();
      const r5cell = col < r5.length ? String(r5[col] || "").trim() : "";

      // Determine billing type and weight/volume tier
      let tierLabel = r4cell || r5cell;
      let billingType = "含税KG";
      let priceUnit = "元/KG";

      if (r4cell.includes("不含税") || r4cell.includes("不包税") || r5cell.includes("CBM") || r5cell.includes("cbm")) {
        billingType = "不含税CBM";
        priceUnit = "元/CBM";
        tierLabel = r5cell || r4cell;
      } else if (r4cell.includes("含税") || r5cell.includes("KG+") || r5cell.includes("kg+")) {
        billingType = "含税KG";
        tierLabel = r5cell || r4cell;
      }

      // Extract weight/volume value
      let minQtyVal = 21;
      const qtyMatch = String(tierLabel).match(/(\d+)\s*(?:KG\+|kg\+|CBM\+|cbm\+)/i);
      if (qtyMatch) minQtyVal = parseInt(qtyMatch[1]);

      if (tierLabel && tierLabel.length > 0 && tierLabel !== "仓库代码" && tierLabel !== "含税" && tierLabel !== "不含税") {
        cg.tiers.push({ col, label: tierLabel, billingType, priceUnit, minQtyVal,
          minQty: tierLabel.includes("CBM") || tierLabel.includes("cbm") ? tierLabel : tierLabel.includes("KG") ? tierLabel : minQtyVal + "KG+" });
      }
    }
    // If no tiers found, use default
    if (cg.tiers.length === 0) {
      for (let col = cg.startCol; col <= cg.endCol; col++) {
        cg.tiers.push({ col, label: "21KG+", billingType: "含税KG", priceUnit: "元/KG", minQtyVal: 21, minQty: "21KG+" });
      }
    }
  }

  // Determine country
  let country = countryOverride || "美国";
  if (sheetName.includes("加拿大") || sheetName.includes("美转加")) country = "加拿大";
  if (sheetName.includes("澳洲") || sheetName.includes("澳大利亚")) country = "澳大利亚";

  // Transport mode
  let transportMode = defaults.tm || "海运";
  if (channelName.includes("空运") || sheetName.includes("空运")) transportMode = "空运";
  if (channelName.includes("快递") || channelName.includes("UPS") || channelName.includes("红单")) transportMode = "快递";

  // Delivery method
  let deliveryMethod = "卡派";
  if (channelName.includes("海派") || channelName.includes("UPS派") || channelName.includes("快递")) deliveryMethod = "快递派";
  if (channelName.includes("DG")) deliveryMethod = "卡派";

  // Find transit/claim columns (usually at the end)
  let transitCol = -1, claimCol = -1;
  for (let col = cityGroups[cityGroups.length-1]?.endCol + 1 || 8; col < Math.min(r3.length, 15); col++) {
    const c = String(r3[col] || "").trim();
    if (c.includes("时效") || c.includes("签收")) transitCol = col;
    if (c.includes("赔付") || c.includes("赔偿")) claimCol = col;
  }

  // Parse data rows
  for (let ri = (r5[0] && String(r5[0]).includes("仓库")) ? 6 : 5; ri < data.length; ri++) {
    const row = data[ri];
    const whCell = String(row[0] || row[1] || "").trim();
    if (!whCell || whCell.length < 3) continue;
    if (whCell.includes("仓库类型") || whCell.includes("仓库代码") || whCell.includes("DISPIMG")) continue;

    // Extract warehouse codes
    let warehouses = [];
    const whMatch = whCell.match(/[A-Z]{2,}\d/g);
    if (whMatch) {
      warehouses = whMatch;
    } else if (whCell.includes("(") || whCell.includes("（")) {
      // "多伦多(YYZ/YOO)" → extract codes from parentheses
      const pm = whCell.match(/[（(]([^)）]+)[)）]/);
      if (pm) warehouses = pm[1].split("/").map(s => s.trim()).filter(s => s.match(/[A-Z]+\d/));
      if (warehouses.length === 0) warehouses = [whCell.split(/[（(]/)[0].trim().slice(0, 40)];
    } else {
      warehouses = [whCell.slice(0, 40)];
    }

    const transitText = transitCol >= 0 ? String(row[transitCol] || "").trim() : "";
    const claimText = claimCol >= 0 ? String(row[claimCol] || "").trim() : "";
    const transitMin = transitText.match(/(\d+)/) ? parseInt(transitText.match(/(\d+)/)[1]) : null;
    const transitMax = transitText.match(/(\d+)[-–](\d+)/) ? parseInt(transitText.match(/(\d+)[-–](\d+)/)[2]) : transitMin;

    for (const cg of cityGroups) {
      for (const tier of cg.tiers) {
        const price = parseFloat(row[tier.col]);
        if (isNaN(price) || price <= 0 || price > 99999) continue;

        for (const wh of warehouses) {
          results.push({
            supplier: SUPPLIER,
            country,
            channel_name: channelName,
            speed_tier: "",
            vessel_config: vesselConfig,
            vessel_tags: vesselConfig ? extractVesselTags(vesselConfig) : [],
            delivery_method: deliveryMethod,
            destination_type: "warehouse",
            destination_code: wh,
            destination_region: country,
            origin_region: cg.name,
            origin_cities: cg.name.split("/").map((s) => s.trim()),
            billing_type: tier.billingType,
            min_quantity: tier.minQty,
            min_quantity_value: tier.minQtyVal,
            unit_price: price,
            price_unit: tier.priceUnit,
            transit_time_min: transitMin,
            transit_time_max: transitMax,
            transit_time_desc: transitText,
            claim_rule: claimText,
            effective_date: "",
            source_sheet: sheetName,
          });
        }
      }
    }
  }
  return results;
}

module.exports = { parseETTON };
