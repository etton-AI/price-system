/**
 * 纽酷国际 — 美线FBA+加拿大线价格解析器
 * 格式: 多城市组 × 多重量段动态布局
 */
const XLSX = require("xlsx");
const SUPPLIER = "纽酷国际";

const CITY_MAP = {
  "华南": ["深圳", "东莞", "广州", "中山"],
  "华东/宁波/上海/苏州": ["义乌", "上海", "宁波", "苏州", "杭州"],
  "华东": ["义乌", "上海", "宁波", "杭州"],
  "宁波": ["宁波"],
  "青岛": ["青岛"],
  "福建": ["厦门", "泉州", "福州"],
  "福州": ["福州"],
  "天津": ["天津"],
  "苏州": ["苏州"],
  "中山/佛山": ["中山", "佛山"],
};

function resolveCity(label) {
  for (const [key, cities] of Object.entries(CITY_MAP)) {
    if (label.includes(key) || key.includes(label)) return { label: key, cities };
  }
  return { label, cities: ["深圳", "东莞"] };
}

function mkr(o) {
  return {
    supplier: SUPPLIER, country: o.c || "美国", channel_name: o.cn || "", transport_mode: o.tm || "海运",
    vessel_config: o.vc || "", vessel_tags: o.vt || [], delivery_method: o.dm || "卡派",
    destination_type: o.dt || "warehouse", destination_code: o.dc || "", destination_region: o.dr || "",
    origin_region: o.or || "华南", origin_cities: o.oc || ["深圳", "东莞", "广州", "中山"],
    billing_type: o.bt || "包税", tax_mode: o.tx || o.bt || "包税",
    min_quantity: o.mq || "", min_quantity_value: o.mv || 0, unit_price: o.p || 0, price_unit: o.pu || "元/KG",
    transit_time_min: o.tn || null, transit_time_max: o.tx2 || null, transit_time_desc: o.td || "",
    claim_rule: o.cr || "", effective_date: "", source_sheet: o.ss || "",
  };
}

/**
 * 动态扫描表头确定列布局
 * 返回: { channelCol, destCol, columns: [{col, cityLabel, cities, tierLabel, tierValue, priceUnit, billingType}] }
 */
function scanLayout(data, dataStartRow) {
  // Row 3 (index 3): channel series | channel name | tax | city group headers...
  // Row 4: region | | city name headers
  // Row 5: warehouse/region | | weight headers with values
  const headerRow3 = data[3] || [];
  const headerRow4 = data[4] || [];
  const headerRow5 = data[5] || [];

  const columns = [];
  let currentCity = "";

  for (let ci = 3; ci < Math.min(headerRow4.length, 16); ci++) {
    const cityLabel = String(headerRow4[ci] || "").trim();
    if (cityLabel && cityLabel.length > 1) {
      currentCity = cityLabel;
    }
    const tierCell = String(headerRow5[ci] || "").trim();
    // Parse quantity from tier cell (e.g., "50.00KG+", "0.50CBM+")
    const kgm = tierCell.match(/(\d+\.?\d*)\s*KG/i);
    const cbm = tierCell.match(/(\d+\.?\d*)\s*CBM/i);
    if (kgm) {
      const city = resolveCity(currentCity);
      columns.push({ col: ci, cityLabel: city.label, cities: city.cities, tierLabel: kgm[1] + "KG+", tierValue: parseFloat(kgm[1]), priceUnit: "元/KG", billingType: "包税" });
    } else if (cbm) {
      const city = resolveCity(currentCity);
      columns.push({ col: ci, cityLabel: city.label, cities: city.cities, tierLabel: cbm[1] + "CBM+", tierValue: parseFloat(cbm[1]), priceUnit: "元/CBM", billingType: "不含税CBM" });
    }
  }

  // Detect channel column: col0 or col1
  const channelCol = String(headerRow3[1] || "").trim().length > 2 ? 1 : 0;
  // Dest is always col2
  const destCol = 2;

  return { channelCol, destCol, columns, dataStartRow };
}

function parseGenericSheet(ws, sheetName, config) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 6) return [];

  // Detect FBN format: row3 col2 = "商业平台"
  const isFBN = String(data[3]?.[2] || "").includes("商业平台");

  const layout = scanLayout(data, config.dataStartRow || 6);
  const results = [];
  let currentChannel = config.channelName || sheetName;

  for (let ri = layout.dataStartRow; ri < data.length; ri++) {
    const row = data[ri];
    const chName = String(row[layout.channelCol] || "").trim();
    // FBN: dest in col3 (postal code); FBA: dest in col2
    const destCol = isFBN ? 3 : layout.destCol;
    const dest = String(row[destCol] || "").trim();

    if (!dest || dest.length < 2) continue;
    if (dest.includes("仓库代码") || dest.includes("区域") || dest.includes("邮编")) continue;
    if (dest.includes("备注") || dest.includes("说明") || dest.includes("头程/自提")) continue;

    // Update channel if col1 has a non-empty value
    if (chName && chName.length > 2 && !chName.includes("下单") && !chName.includes("渠道")) {
      currentChannel = chName;
    }

    // FBN: col3 is postal code; skip non-postal rows
    if (isFBN) {
      if (!/^\d{5}$/.test(dest)) continue; // must be 5-digit postal code
    }

    // Determine destination type
    let destType = isFBN ? "commercial" : "warehouse";
    let destRegion = dest;
    if (dest.includes("美西") || dest.includes("邮编8") || dest.includes("邮编 8")) { destType = "region"; destRegion = "美西"; }
    else if (dest.includes("美中") || dest.includes("邮编4") || dest.includes("邮编 4")) { destType = "region"; destRegion = "美中"; }
    else if (dest.includes("美东") || dest.includes("邮编0") || dest.includes("邮编 0")) { destType = "region"; destRegion = "美东"; }
    // 5-digit postal code = commercial address
    else if (/^\d{5}$/.test(dest)) { destType = "commercial"; destRegion = ""; }

    for (const col of layout.columns) {
      if (col.col >= row.length) continue;
      const price = parseFloat(row[col.col]);
      if (isNaN(price) || price <= 0 || price > 99999) continue;

      results.push(mkr({
        c: config.country || "美国", cn: currentChannel, tm: config.tm || "海运",
        vc: config.vc || currentChannel, vt: [config.vc || currentChannel],
        dm: config.dm || (sheetName.includes("海派") ? "快递派" : "卡派"),
        dc: dest, dt: destType, dr: destRegion,
        or: col.cityLabel, oc: col.cities,
        bt: col.billingType, tx: col.billingType,
        mq: col.tierLabel, mv: col.tierValue, p: price, pu: col.priceUnit,
        ss: sheetName,
      }));
    }
  }
  return results;
}

// ── 加拿大解析 ──
function parseCASheet(ws, sheetName, channelName) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 5) return [];
  const results = [];

  // Find data rows with warehouse codes
  for (let ri = 3; ri < data.length; ri++) {
    const row = data[ri];
    const col1 = String(row[1] || "").trim();
    const col2 = String(row[2] || "").trim();
    if (!col1 || col1.length < 2) continue;
    if (col1.includes("仓库") || col1.includes("地址") || col1.includes("下单渠道")) continue;

    let dest = col2 || col1;
    if (dest.match(/[A-Z]{2,}\d/)) {
      // It's a warehouse code
      for (let ci = 3; ci < Math.min(row.length, 10); ci++) {
        const p = parseFloat(row[ci]);
        if (!isNaN(p) && p > 0 && p < 99999) {
          results.push(mkr({ c: "加拿大", cn: channelName, tm: "海运", vc: "海运", vt: ["海运"],
            dm: "卡派", dc: dest.match(/[A-Z]{2,}\d/)[0], dt: "warehouse", dr: dest,
            mq: "50KG+", mv: 50, p, ss: sheetName }));
        }
      }
    }
  }
  return results;
}

// ── 主入口 ──
function parseNiuku(filePath) {
  console.log("[纽酷] 开始解析:", filePath);
  const wb = XLSX.readFile(filePath);
  const all = [];
  const sn = wb.SheetNames;

  // US sheets - use generic parser
  const usDataSheets = [
    "含税海派", "直送专线", "王牌渠道-全美25日达", "美森特快",
    "华东EXX快船", "ZIM快船", "合德快船", "单票单清", "先查后装",
    "萨瓦纳快线", "纽约25日达",
  ];

  for (const sh of usDataSheets) {
    if (sn.includes(sh)) {
      try {
        const r = parseGenericSheet(wb.Sheets[sh], sh, { country: "美国", tm: "海运" });
        console.log(`  [${sh}] ${r.length} 条`);
        all.push(...r);
      } catch (err) { console.error(`  [${sh}] 失败: ${err.message}`); }
    }
  }

  // Also handle FBN sheets
  const fbnSheets = sn.filter(s => s.includes("特快达") || s.includes("极速达") || s.includes("经济达") || s.includes("特惠达"));
  for (const sh of fbnSheets) {
    try {
      const r = parseGenericSheet(wb.Sheets[sh], sh, { country: "美国", tm: "海运" });
      console.log(`  [${sh}] ${r.length} 条`);
      all.push(...r);
    } catch (err) { console.error(`  [${sh}] 失败: ${err.message}`); }
  }

  // CA sheets
  const caSheets = sn.filter(s => s.includes("美转加") || s.includes("ERS") || s.includes("经济线") || s.includes("特惠"));
  for (const sh of caSheets) {
    try {
      const r = parseCASheet(wb.Sheets[sh], sh, sh);
      console.log(`  [${sh}] ${r.length} 条 → 加拿大`);
      all.push(...r);
    } catch (err) { console.error(`  [${sh}] 失败: ${err.message}`); }
  }

  console.log(`[纽酷] 总计 ${all.length} 条`);
  return all;
}

module.exports = { parseNiuku };
