/**
 * 星链专线 — 美线价格解析器
 *
 * 星链格式: 区域行 → 仓库列表 + 单价格 + 时效
 * 核心 Sheet "美国FBA-海运": 区域为行(美西/美中南/美中北/美东南/美东北)
 * 每个区域含仓库代码列表(用、分隔) + 12KG+ 统一价格 + 送仓时效
 */

const XLSX = require("xlsx");

const SUPPLIER = "星链专线";
const COUNTRY = "美国";

// ── 城市→区域映射 ──
const REGION_CITIES = {
  美西: ["深圳", "东莞", "广州", "中山", "惠州", "义乌", "上海", "宁波", "杭州", "厦门", "泉州", "福州", "汕头"],
  美中南: ["深圳", "东莞", "广州", "中山", "惠州", "义乌", "上海", "宁波", "杭州", "厦门", "泉州", "福州", "汕头"],
  美中北: ["深圳", "东莞", "广州", "中山", "惠州", "义乌", "上海", "宁波", "杭州", "厦门", "泉州", "福州", "汕头"],
  美东南: ["深圳", "东莞", "广州", "中山", "惠州", "义乌", "上海", "宁波", "杭州", "厦门", "泉州", "福州", "汕头"],
  美东北: ["深圳", "东莞", "广州", "中山", "惠州", "义乌", "上海", "宁波", "杭州", "厦门", "泉州", "福州", "汕头"],
  全美: ["深圳", "东莞", "广州", "中山", "惠州", "义乌", "上海", "宁波", "杭州", "厦门", "泉州", "福州", "汕头"],
};

// ── 解析仓库代码 ──
function parseWarehouses(cell) {
  const text = String(cell).replace(/\r?\n/g, "、").replace(/\s+/g, "");
  return text.split(/[、,，]/).map((s) => s.trim()).filter((s) => s && s.length >= 3);
}

// ── 解析时效 ──
function parseTransit(text) {
  const cleaned = String(text).replace(/\r?\n/g, " ").trim();
  // "20-25 天" or "41–44 天" or "22-25天"
  const match = cleaned.match(/(\d+)\s*[-–]\s*(\d+)\s*天/);
  if (match) return { min: parseInt(match[1]), max: parseInt(match[2]), desc: cleaned };
  const singleMatch = cleaned.match(/(\d+)\s*天/);
  if (singleMatch) return { min: parseInt(singleMatch[1]), max: parseInt(singleMatch[1]), desc: cleaned };
  return { min: null, max: null, desc: cleaned };
}

// ── 解析主价格表 (R5-R9 区域行) ──
function parseStandardFBA(data) {
  const results = [];
  // R4=header, R5-R9=data rows (index 4-8)
  const dataStartRow = 4; // 0-indexed, R5
  const dataEndRow = 8;   // 0-indexed, R9

  for (let ri = dataStartRow; ri <= dataEndRow && ri < data.length; ri++) {
    const row = data[ri];
    const region = String(row[0] || "").trim();
    const whCell = String(row[3] || "").trim();
    const priceCell = parseFloat(row[10]);
    const transitCell = String(row[12] || "");
    const notesCell = String(row[13] || "").replace(/\r?\n/g, " ").trim();

    if (!region || !whCell || isNaN(priceCell) || priceCell <= 0) continue;

    const warehouses = parseWarehouses(whCell);
    const transit = parseTransit(transitCell);
    const cities = REGION_CITIES[region] || REGION_CITIES["全美"];

    for (const wh of warehouses) {
      results.push(makeRecord({
        channelName: "星链直送-普船",
        vesselConfig: "普船直送",
        vesselTags: ["普船", "直送"],
        deliveryMethod: "直送",
        destCode: wh,
        destRegion: region,
        originCities: cities,
        billingType: "包税",
        minQty: "12KG+",
        minQtyValue: 12,
        price: priceCell,
        transitMin: transit.min,
        transitMax: transit.max,
        transitDesc: transit.desc,
        notes: notesCell,
      }));
    }
  }

  return results;
}

// ── 解析锁仓服务 (R13-R14) ──
function parseLockService(data) {
  const results = [];
  // R13=header for lock service (index 12), R14=data (index 13)
  if (data.length < 14) return results;

  const headerRow = data[12]; // R13
  const dataRow = data[13];   // R14

  const region = String(dataRow[0] || "").trim();
  const whCell = String(dataRow[3] || "").trim();
  const priceCell = parseFloat(dataRow[10]);
  const transitCell = String(dataRow[12] || "");
  const notesCell = String(dataRow[13] || "").replace(/\r?\n/g, " ").trim();

  if (!region || !whCell || isNaN(priceCell) || priceCell <= 0) return results;

  const warehouses = parseWarehouses(whCell);
  const transit = parseTransit(transitCell);
  const cities = REGION_CITIES["全美"];

  for (const wh of warehouses) {
    results.push(makeRecord({
      channelName: "星链锁仓-普船",
      vesselConfig: "普船锁仓",
      vesselTags: ["普船", "锁仓", "直送"],
      deliveryMethod: "直送",
      destCode: wh,
      destRegion: "全美锁仓",
      originCities: cities,
      billingType: "包税",
      minQty: "12KG+",
      minQtyValue: 12,
      price: priceCell,
      transitMin: transit.min,
      transitMax: transit.max,
      transitDesc: transit.desc,
      notes: notesCell,
    }));
  }

  return results;
}

// ── 辅助函数 ──
function makeRecord(opts) {
  return {
    supplier: SUPPLIER,
    country: COUNTRY,
    channel_name: opts.channelName,
    transport_mode: "海运",
    vessel_config: opts.vesselConfig,
    vessel_tags: opts.vesselTags || [],
    delivery_method: opts.deliveryMethod,
    destination_type: "warehouse",
    destination_code: opts.destCode,
    destination_region: opts.destRegion,
    origin_region: opts.destRegion,
    origin_cities: opts.originCities,
    billing_type: opts.billingType,
    tax_mode: opts.billingType,
    min_quantity: opts.minQty,
    min_quantity_value: opts.minQtyValue,
    unit_price: opts.price,
    price_unit: "元/KG",
    transit_time_min: opts.transitMin,
    transit_time_max: opts.transitMax,
    transit_time_desc: opts.transitDesc,
    claim_rule: opts.notes || "",
    effective_date: "",
    source_sheet: "美国FBA-海运",
  };
}

// ── 主入口 ──
function parseXinglian(filePath) {
  console.log("[星链] 开始解析:", filePath);
  const wb = XLSX.readFile(filePath);
  const allResults = [];

  // 美国FBA-海运
  if (wb.SheetNames.includes("美国FBA-海运")) {
    const ws = wb.Sheets["美国FBA-海运"];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });

    const standardResults = parseStandardFBA(data);
    console.log(`  [美国FBA-海运/标准] ${standardResults.length} 条`);
    allResults.push(...standardResults);

    const lockResults = parseLockService(data);
    console.log(`  [美国FBA-海运/锁仓] ${lockResults.length} 条`);
    allResults.push(...lockResults);
  }

  // 澳洲FBA-海运 (暂不解析，后续扩展)
  if (wb.SheetNames.includes("澳洲FBA-海运")) {
    console.log(`  [澳洲FBA-海运] 跳过 (暂不支持澳洲线)`);
  }

  console.log(`[星链] 总计 ${allResults.length} 条`);
  return allResults;
}

module.exports = { parseXinglian };
