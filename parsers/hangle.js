/**
 * 航乐国际 — 英国线 + 欧洲大陆线价格解析器
 *
 * Sheet结构:
 *   "英国卡航，海运" — 英国线: 卡航(不包税/包税) + 海运(不包税/包税)
 *   "欧洲卡航-快递派+专仓卡派" — 欧洲大陆: 卡航包税(专仓卡派 + 快递派)
 *   "欧洲海运-快递派+专仓卡派" — 欧洲大陆: 海运包税(专仓卡派 + 快递派)
 *   "欧洲铁路-快递派+专仓卡派" — 欧洲大陆: 铁路包税(专仓卡派 + 快递派)
 *
 * 特点:
 * - 仓行: 仓库代码以"+"分隔需拆分 (如 WRO5+DTM2+HAJ1 拆为3条)
 * - 国家行: 国家组以逗号分隔需拆分
 * - 含 10CBM+ 按方行
 */

const XLSX = require("xlsx");

const SUPPLIER = "航乐国际";

// ── 默认发货城市 ──
const DEFAULT_CITIES = ["深圳", "东莞", "广州", "中山", "惠州", "义乌", "上海", "宁波", "杭州"];

// ── 仓库拆分 ──
function parseWarehouses(cell) {
  const text = String(cell).replace(/\r?\n/g, " ").trim();
  return text.split(/[+＋]/).map(s => s.trim()).filter(s => s.match(/[A-Z]{2,}\d/) || s === "万邑通" || s === "谷仓" || s === "4PX");
}

// ── 国家拆分 ──
const EU_COUNTRIES = ["德国", "法国", "荷兰", "波兰", "比利时", "奥地利", "卢森堡", "捷克", "意大利", "西班牙",
  "英国", "丹麦", "瑞典", "芬兰", "匈牙利", "希腊", "葡萄牙", "爱尔兰", "罗马尼亚", "保加利亚",
  "克罗地亚", "斯洛文尼亚", "斯洛伐克", "立陶宛", "拉脱维亚", "爱沙尼亚"];

function parseCountries(cell) {
  const text = String(cell).replace(/\r?\n/g, " ").trim();
  if (text === "德国FBA") return ["德国"];
  for (const cn of EU_COUNTRIES) { if (text === cn) return [cn]; }
  const parts = text.split(/[,，、]/).map(s => s.trim()).filter(Boolean);
  const result = [];
  for (const p of parts) {
    for (const cn of EU_COUNTRIES) {
      if (p.includes(cn)) { result.push(cn); break; }
    }
  }
  return result.length > 0 ? result : [text];
}

// ── 时效解析 ──
function parseTransit(text) {
  const cleaned = String(text || "").replace(/\r?\n/g, " ").trim();
  const match = cleaned.match(/(\d+)\s*[-–~约]*\s*(\d+)\s*(?:个)?(?:自然日|天)/);
  if (match) return { min: parseInt(match[1]), max: parseInt(match[2]), desc: cleaned };
  const single = cleaned.match(/(\d+)\s*(?:个)?(?:自然日|天)/);
  if (single) return { min: parseInt(single[1]), max: parseInt(single[1]), desc: cleaned };
  return { min: null, max: null, desc: cleaned };
}

function makeRecord(opts) {
  return {
    supplier: SUPPLIER,
    country: opts.country || "欧洲",
    channel_name: opts.channelName || "",
    transport_mode: opts.transportMode || "海运",
    vessel_config: opts.vesselConfig || "",
    vessel_tags: opts.vesselTags || [],
    delivery_method: opts.deliveryMethod || "卡派",
    destination_type: opts.destType || "warehouse",
    destination_code: opts.destCode || "",
    destination_region: opts.destRegion || "",
    origin_region: "华南",
    origin_cities: DEFAULT_CITIES,
    billing_type: opts.billingType || "包税",
    tax_mode: opts.taxMode || opts.billingType || "包税",
    min_quantity: opts.minQty || "",
    min_quantity_value: opts.minQtyValue || 0,
    unit_price: opts.price || 0,
    price_unit: "元/KG",
    transit_time_min: opts.transitMin || null,
    transit_time_max: opts.transitMax || null,
    transit_time_desc: opts.transitDesc || "",
    claim_rule: opts.claimRule || "",
    effective_date: "",
    source_sheet: opts.sourceSheet || "",
  };
}

// ═══════════════════════════════════════════
// 英国线
// ═══════════════════════════════════════════
function parseUK(ws) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 3) return [];

  const results = [];
  let currentMode = ""; // "卡航" or "海运"

  for (let ri = 1; ri < data.length; ri++) {
    const row = data[ri];
    const col0 = String(row[0] || "").replace(/\r?\n/g, " ").trim();

    // Detect section
    if (col0.includes("卡航不包税") || col0.includes("卡航包税")) { currentMode = "卡航"; }
    if (col0.includes("海运不包税") || col0.includes("海运包税")) { currentMode = "海运"; }
    if (col0.includes("英国卡航") && col0.includes("海运")) continue; // section header
    if (col0.includes("英国海运")) continue;

    // Parse data rows
    if (col0.includes("不包税") || col0.includes("包税")) {
      const taxMode = col0.includes("不包税") ? "不包税" : "包税";
      const transportMode = currentMode === "卡航" ? "卡航" : "海运";
      const channelName = `英国${currentMode}-${taxMode}`;

      const tiers = [
        { col: 1, qty: "21KG+", val: 21 }, { col: 2, qty: "101KG+", val: 101 },
        { col: 3, qty: currentMode === "卡航" ? "501KG+" : "501KG+", val: 501 },
      ];

      for (const tier of tiers) {
        const price = parseFloat(row[tier.col]);
        if (!isNaN(price) && price > 0) {
          results.push(makeRecord({
            country: "英国", channelName, transportMode,
            vesselConfig: transportMode, vesselTags: [transportMode],
            deliveryMethod: "DPD派", destCode: "英国", destType: "country", destRegion: "英国",
            billingType: taxMode, taxMode, minQty: tier.qty, minQtyValue: tier.val, price,
            transitDesc: currentMode === "卡航" ? "约22-25天提取" : "35-40天提取",
            sourceSheet: "英国卡航海运",
          }));
        }
      }
    }
  }
  return results;
}

// ═══════════════════════════════════════════
// 欧洲大陆线 — 通用解析
// ═══════════════════════════════════════════

function parseEUSheet(ws, sheetName, transportMode) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 3) return [];

  const results = [];
  let currentSection = ""; // "专仓卡派" or "快递派"
  let currentTaxMode = "包税"; // 航乐欧洲线全是包税

  for (let ri = 1; ri < data.length; ri++) {
    const row = data[ri];
    const col0 = String(row[0] || "").replace(/\r?\n/g, " ").trim();
    const col1 = String(row[1] || "").replace(/\r?\n/g, " ").trim();

    // Detect section
    if (col0.includes("专仓卡派") || col1.includes("专仓卡派")) { currentSection = "专仓卡派"; continue; }
    if (col0.includes("快递派送") || col1.includes("快递派送")) { currentSection = "快递派"; continue; }
    if (col0.includes("包税") && !col0.includes("派送")) {
      if (col0.includes("卡派")) currentSection = "专仓卡派";
      else if (col0.includes("快递")) currentSection = "快递派";
      continue;
    }

    // Skip headers
    if (col1 === "派送方式" || col0.includes("清关费") || col0.includes("派送方式")) continue;
    if (col0.includes("注意事项") || col0.includes("针对包税")) continue;
    if (!col0 && !col1) continue;

    // CBM rows (10CBM+)
    if (col0.includes("10CBM") || col0.includes("CBM")) {
      const warehouses = parseWarehouses(col0);
      const price = parseFloat(row[1]);
      if (!isNaN(price) && price > 0 && warehouses.length > 0) {
        for (const wh of warehouses) {
          results.push(makeRecord({
            country: "欧洲", channelName: `${sheetName}-${currentSection}`, transportMode,
            vesselConfig: transportMode, vesselTags: [transportMode],
            deliveryMethod: currentSection === "专仓卡派" ? "卡派" : "快递派",
            destCode: wh, destType: "warehouse", destRegion: wh,
            billingType: "包税", taxMode: "包税",
            minQty: "10CBM+", minQtyValue: 10, price,
            transitDesc: "", sourceSheet: sheetName,
          }));
        }
      }
      continue;
    }

    // Data rows
    const isWarehouseRow = col0.match(/[A-Z]{2,}\d/) || col0.match(/万邑通|谷仓|4PX/);
    let destinations;
    let destType;

    if (isWarehouseRow) {
      destinations = parseWarehouses(col0);
      destType = "warehouse";
    } else {
      destinations = parseCountries(col0);
      destType = destinations[0] && destinations[0].length >= 2 ? "country" : "warehouse";
    }

    if (destinations.length === 0) continue;

    const weightTiers = [
      { col: 1, qty: "21KG+", val: 21 }, { col: 2, qty: "101KG+", val: 101 }, { col: 3, qty: "501KG+", val: 501 },
    ];
    const transit = parseTransit(String(row[5] || row[6] || ""));

    for (const dest of destinations) {
      for (const tier of weightTiers) {
        const price = parseFloat(row[tier.col]);
        if (!isNaN(price) && price > 0) {
          results.push(makeRecord({
            country: "欧洲", channelName: `${sheetName}-${currentSection}`, transportMode,
            vesselConfig: transportMode, vesselTags: [transportMode],
            deliveryMethod: currentSection === "专仓卡派" ? "卡派" : "快递派",
            destCode: dest, destType, destRegion: dest,
            billingType: "包税", taxMode: "包税",
            minQty: tier.qty, minQtyValue: tier.val, price,
            transitMin: transit.min, transitMax: transit.max,
            transitDesc: transit.desc, sourceSheet: sheetName,
          }));
        }
      }
    }
  }
  return results;
}

// ═══════════════════════════════════════════
// 主入口
// ═══════════════════════════════════════════

function parseHangle(filePath) {
  console.log("[航乐] 开始解析:", filePath);
  const wb = XLSX.readFile(filePath);
  const allResults = [];

  // UK
  if (wb.SheetNames.includes("英国卡航，海运")) {
    try {
      const results = parseUK(wb.Sheets["英国卡航，海运"]);
      console.log(`  [英国卡航海运] ${results.length} 条`);
      allResults.push(...results);
    } catch (err) {
      console.error(`  [英国卡航海运] 解析失败: ${err.message}`);
    }
  }

  // EU
  const euSheets = {
    "欧洲卡航-快递派+专仓卡派": "卡航",
    "欧洲海运-快递派+专仓卡派": "海运",
    "欧洲铁路-快递派+专仓卡派": "铁路",
  };

  for (const [sheetName, mode] of Object.entries(euSheets)) {
    if (wb.SheetNames.includes(sheetName)) {
      try {
        const results = parseEUSheet(wb.Sheets[sheetName], sheetName, mode);
        console.log(`  [${sheetName}] ${results.length} 条`);
        allResults.push(...results);
      } catch (err) {
        console.error(`  [${sheetName}] 解析失败: ${err.message}`);
      }
    }
  }

  console.log(`[航乐] 总计 ${allResults.length} 条`);
  return allResults;
}

module.exports = { parseHangle };
