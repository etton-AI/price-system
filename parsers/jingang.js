/**
 * 劲港物流 — 美线+加拿大线价格解析器
 */
const XLSX = require("xlsx");
const SUPPLIER = "劲港物流";

const CITY_GROUPS = [
  { label: "华南", cities: ["深圳", "东莞", "广州", "中山", "汕头"] },
  { label: "厦门", cities: ["厦门"] },
  { label: "福州/泉州", cities: ["福州", "泉州", "长沙"] },
  { label: "华东", cities: ["金华", "宁波", "苏州", "上海", "杭州", "温州", "义乌"] },
  { label: "青岛", cities: ["青岛", "郑州"] },
];

function mkr(o) {
  return {
    supplier: SUPPLIER, country: o.c || "美国", channel_name: o.cn || "", transport_mode: o.tm || "海运",
    vessel_config: o.vc || "", vessel_tags: o.vt || [], delivery_method: o.dm || "卡派",
    destination_type: o.dt || "warehouse", destination_code: o.dc || "", destination_region: o.dr || "",
    origin_region: o.or || "华南", origin_cities: o.oc || CITY_GROUPS[0].cities,
    billing_type: o.bt || "包税", tax_mode: o.tx || o.bt || "包税",
    min_quantity: o.mq || "", min_quantity_value: o.mv || 0, unit_price: o.p || 0, price_unit: o.pu || "元/KG",
    transit_time_min: o.tn || null, transit_time_max: o.tx2 || null, transit_time_desc: o.td || "",
    claim_rule: o.cr || "", effective_date: "", source_sheet: o.ss || "",
  };
}

/**
 * 通用解析：仓库行 × 城市列
 * config.tierCols: [{col, label, value, priceUnit, billingType, cityGroup, taxMode}]
 */
function parseGeneric(ws, config) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 6) return [];
  const results = [];
  let currentChannel = config.channelName;

  for (let ri = config.dataStartRow || 5; ri < data.length; ri++) {
    const row = data[ri];
    const col0 = String(row[0] || "").trim();
    const col1 = String(row[1] || "").trim();
    const col2 = String(row[2] || "").trim();

    if (!col2 || col2.length < 2) continue;
    if (col2.includes("仓库代码") || col2.includes("目的港") || col2.includes("区域")) continue;
    if (col2.includes("发货说明") || col2.includes("注意") || col2.includes("附加费")) continue;
    if (col0.includes("目的港") || col1.includes("下单渠道")) continue;

    // Detect channel from col1
    if (col1 && col1.length > 3 && !col1.includes("下单") && !col1.includes("渠道")) {
      currentChannel = col1;
    }

    // Detect destination
    let dest = col2;
    let destType = "warehouse";
    if (dest.includes("美西") || dest.includes("美中") || dest.includes("美东") || dest.includes("邮编")) {
      destType = "region";
      if (dest.includes("8") && dest.includes("9")) dest = "美西";
      else if (dest.includes("4") || dest.includes("5") || dest.includes("6") || dest.includes("7")) dest = "美中";
      else if (dest.includes("0") || dest.includes("1") || dest.includes("2") || dest.includes("3")) dest = "美东";
    }
    // CA sheets have warehouse codes
    if (dest.match(/^[A-Z]{2,}\d/) || dest.match(/^[A-Z]{2,}$/)) destType = "warehouse";

    for (const tc of config.tierCols) {
      if (tc.col >= row.length) continue;
      const price = parseFloat(row[tc.col]);
      if (isNaN(price) || price <= 0 || price > 99999) continue;

      const cg = tc.cityGroup || CITY_GROUPS[0];
      results.push(mkr({
        c: config.country, cn: currentChannel, tm: config.transportMode || "海运",
        vc: config.vesselConfig || config.transportMode || "海运",
        vt: [config.vesselConfig || config.transportMode || "海运"],
        dm: config.deliveryMethod || "卡派", dc: dest, dt: destType, dr: dest,
        or: cg.label, oc: cg.cities,
        bt: tc.billingType || "包税", tx: tc.taxMode || tc.billingType || "包税",
        mq: tc.label, mv: tc.value, p: price, pu: tc.priceUnit || "元/KG",
        ss: config.sheetName,
      }));
    }
  }
  return results;
}

/** Build tier configs from a row of weight headers */
function scanTierRow(row, cityGroup, startCol, count) {
  const tiers = [];
  for (let i = 0; i < count && (startCol + i) < row.length; i++) {
    const col = startCol + i;
    const cell = String(row[col] || "").trim();
    let label = "", value = 0, priceUnit = "元/KG", billingType = "包税";

    // Parse tier from cell text
    const kgm = cell.match(/(\d+\.?\d*)\s*KG\s*\+?\s*[\(（]?含税[\)）]?/i);
    const cbm = cell.match(/(\d+\.?\d*)\s*CBM\s*\+?\s*[\(（]?含税[\)）]?/i);
    const kgNoTax = cell.match(/(\d+\.?\d*)\s*KG\s*\+?\s*[\(（]?不含税[\)）]?/i);
    const cbmNoTax = cell.match(/(\d+\.?\d*)\s*CBM\s*\+?\s*[\(（]?不含税[\)）]?/i);

    if (cbm) { label = cbm[1] + "CBM+"; value = parseFloat(cbm[1]); priceUnit = "元/CBM"; }
    else if (kgm) { label = kgm[1] + "KG+"; value = parseFloat(kgm[1]); }
    else if (cbmNoTax) { label = cbmNoTax[1] + "CBM+"; value = parseFloat(cbmNoTax[1]); priceUnit = "元/CBM"; billingType = "不含税CBM"; }
    else if (kgNoTax) { label = kgNoTax[1] + "KG+"; value = parseFloat(kgNoTax[1]); billingType = "不包税"; }
    else if (cell.match(/(\d+\.?\d*)\s*KG/i)) {
      const m = cell.match(/(\d+\.?\d*)\s*KG/i); label = m[1] + "KG+"; value = parseFloat(m[1]);
    }
    else if (cell.match(/(\d+\.?\d*)\s*CBM/i)) {
      const m = cell.match(/(\d+\.?\d*)\s*CBM/i); label = m[1] + "CBM+"; value = parseFloat(m[1]); priceUnit = "元/CBM";
    }

    if (label) {
      tiers.push({ col, label, value, priceUnit, billingType, taxMode: billingType, cityGroup });
    }
  }
  return tiers;
}

// ── US 直送渠道 ──
function parseUSDirect(ws) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 6) return [];
  // Row 4 has weight tier labels
  const tierRow = data[4] || [];
  // Cols 3-6: 塘厦/广州/中山/汕头 (华南)
  // Cols 7-10: 厦门
  // Cols 11-14: 福州/泉州/长沙
  const tierCols = [
    ...scanTierRow(tierRow, CITY_GROUPS[0], 3, 4),  // 华南
    ...scanTierRow(tierRow, CITY_GROUPS[1], 7, 4),  // 厦门
    ...scanTierRow(tierRow, CITY_GROUPS[2], 11, 4), // 福州/泉州
  ];
  return parseGeneric(ws, {
    sheetName: "直送渠道", country: "美国", transportMode: "海运",
    deliveryMethod: "直送", channelName: "直送渠道", tierCols, dataStartRow: 5,
  });
}

// ── US 快递派渠道 ──
function parseUSExpress(ws) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 6) return [];
  const tierRow = data[5] || [];
  // Cols 3-4: 华东 12KG+/50KG+, 5-6: 华南, 7-8: 厦门, 9-10: 福州/泉州, 11-12: 青岛
  const tierCols = [
    ...scanTierRow(tierRow, CITY_GROUPS[3], 3, 2),  // 华东
    ...scanTierRow(tierRow, CITY_GROUPS[0], 5, 2),  // 华南
    ...scanTierRow(tierRow, CITY_GROUPS[1], 7, 2),  // 厦门
    ...scanTierRow(tierRow, CITY_GROUPS[2], 9, 2),  // 福州/泉州
    ...scanTierRow(tierRow, CITY_GROUPS[4], 11, 2), // 青岛
  ];
  return parseGeneric(ws, {
    sheetName: "快递派渠道", country: "美国", transportMode: "海运",
    deliveryMethod: "快递派", channelName: "快递派渠道", tierCols, dataStartRow: 6,
  });
}

// ── US 华东限时23 ──
function parseUSEastChina(ws) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 6) return [];
  const tierRow = data[3] || [];
  // Cols 3-6: 华东, Cols 7-10: 华南/其他
  const tierCols = [
    ...scanTierRow(tierRow, CITY_GROUPS[3], 3, 4),  // 华东
    ...scanTierRow(tierRow, CITY_GROUPS[0], 7, 4),  // 华南+其他
  ];
  return parseGeneric(ws, {
    sheetName: "华东限时23", country: "美国", transportMode: "海运",
    deliveryMethod: "卡派", channelName: "华东限时23", tierCols, dataStartRow: 4,
  });
}

// ── CA Sheets ──
const CA_CITY_GROUPS = [
  { label: "华南", cities: ["深圳", "东莞", "广州", "中山"] },
  { label: "华东", cities: ["金华", "宁波", "苏州", "上海", "杭州", "义乌"] },
];

function parseCAGeneric(ws, sheetName, channelName) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 6) return [];

  // Find the data header row with weight tiers
  let tierRow = null;
  let dataStartRow = 0;
  for (let ri = 2; ri < Math.min(10, data.length); ri++) {
    const r = data[ri];
    const cells = r.slice(2, 10).map(c => String(c || ""));
    if (cells.some(c => c.match(/\d+KG/i) || c.match(/\d+CBM/i))) {
      tierRow = r;
      dataStartRow = ri + 1;
      break;
    }
  }
  if (!tierRow) return [];

  // Dynamic tier columns based on what we find
  const allTiers = [];
  for (let ci = 2; ci < Math.min(tierRow.length, 14); ci++) {
    const cell = String(tierRow[ci] || "").trim();
    const kgm = cell.match(/(\d+\.?\d*)\s*KG/i);
    const cbm = cell.match(/(\d+\.?\d*)\s*CBM/i);
    if (kgm) {
      const cg = ci < 6 ? CA_CITY_GROUPS[0] : CA_CITY_GROUPS[1] || CA_CITY_GROUPS[0];
      allTiers.push({ col: ci, label: kgm[1] + "KG+", value: parseFloat(kgm[1]), priceUnit: "元/KG", billingType: "包税", taxMode: "包税", cityGroup: cg });
    } else if (cbm) {
      const cg = ci < 6 ? CA_CITY_GROUPS[0] : CA_CITY_GROUPS[1] || CA_CITY_GROUPS[0];
      allTiers.push({ col: ci, label: cbm[1] + "CBM+", value: parseFloat(cbm[1]), priceUnit: "元/CBM", billingType: "不含税CBM", taxMode: "不含税CBM", cityGroup: cg });
    }
  }

  return parseGeneric(ws, {
    sheetName, country: "加拿大", transportMode: "海运",
    deliveryMethod: sheetName.includes("快递") ? "快递派" : "卡派",
    channelName, tierCols: allTiers, dataStartRow,
  });
}

// ── 主入口 ──
function parseJingang(filePath) {
  console.log("[劲港] 开始解析:", filePath);
  const wb = XLSX.readFile(filePath);
  const all = [];

  // US sheets
  const usConfigs = [
    { name: "直送渠道", fn: parseUSDirect },
    { name: "快递派渠道", fn: parseUSExpress },
    { name: "华东限时23", fn: parseUSEastChina },
  ];
  for (const cfg of usConfigs) {
    if (wb.SheetNames.includes(cfg.name)) {
      try {
        const r = cfg.fn(wb.Sheets[cfg.name]);
        console.log(`  [${cfg.name}] ${r.length} 条`);
        all.push(...r);
      } catch (err) { console.error(`  [${cfg.name}] 失败: ${err.message}`); }
    }
  }

  // CA sheets
  const caConfigs = [
    ["美加20", "美加20-拆派"], ["美加25", "美加25-拆派"], ["美加30", "美加30-拆派"],
    ["加拿大ERS", "加拿大ERS-拆派"], ["加拿大ERS-商业卡派", "加拿大ERS-商业卡派"],
    ["加拿大经济", "加拿大经济-拆派"], ["加拿大经济-商业卡派", "加拿大经济-商业卡派"],
    ["加拿大ERS-多伦多快递派", "加拿大ERS-多伦多快递派"],
  ];
  for (const [sn, cn] of caConfigs) {
    if (wb.SheetNames.includes(sn)) {
      try {
        const r = parseCAGeneric(wb.Sheets[sn], sn, cn);
        console.log(`  [${sn}] ${r.length} 条 → 加拿大`);
        all.push(...r);
      } catch (err) { console.error(`  [${sn}] 失败: ${err.message}`); }
    }
  }

  console.log(`[劲港] 总计 ${all.length} 条`);
  return all;
}

module.exports = { parseJingang };
