/**
 * 皓鹏国际 — 美线价格解析器
 *
 * 核心数据来源（按Sheet分类）:
 * 1. "美森以星合德OA非OA海卡" — 仓×渠道，卡派，5个渠道组（美森正班/以星/合德/OA/非OA）
 * 2. "美西洛杉矶海卡特惠" — 仓×渠道，卡派，4个渠道组（美森正班/以星合德/OA/非OA）
 * 3. "美森以星合德OA海派" — 邮编区域×渠道，海派，6个渠道组
 * 4. "美中休斯顿海卡" — 仓×渠道，休斯顿卡派
 * 5. "美中芝加哥海卡" — 仓×渠道，芝加哥卡派
 * 6. "美东萨凡纳海卡" — 仓×渠道，萨凡纳卡派
 * 7. "美东纽约OA非OA海卡" — 仓×渠道，纽约卡派
 * 8. "美西美东商私卡" — 距离×渠道，商业/私人地址卡派
 * 9. "美国空运" — 邮编区域×渠道，空运
 *
 * 渠道→船司映射:
 *   美森正班 = Matson CLX 正班
 *   以星 = ZIM ZEX
 *   合德 = HEDE
 *   OA = COSCO/EMC/OOCL/CMA (OA联盟)
 *   非OA = ONE/WHL/YML/HMM
 */

const XLSX = require("xlsx");
const { detectCountry, detectAllCountries } = require("./country-detector");

const SUPPLIER = "皓鹏国际";

// ── 默认发货城市 ──
const DEFAULT_CITIES = ["深圳", "广州", "东莞", "中山", "惠州", "义乌", "上海", "宁波", "杭州", "厦门", "泉州", "福州"];

// ── 渠道→船司映射 ──
const CHANNEL_VESSEL = {
  "美森正班": { config: "美森正班CLX", tags: ["美森", "Matson", "CLX", "正班"] },
  "美森": { config: "美森正班CLX", tags: ["美森", "Matson", "CLX"] },
  "以星": { config: "以星ZIM-ZEX", tags: ["以星", "ZIM", "ZEX"] },
  "合德": { config: "合德HEDE", tags: ["合德", "HEDE"] },
  "OA": { config: "OA联盟普船", tags: ["OA", "COSCO", "EMC", "OOCL", "CMA", "普船"] },
  "非OA": { config: "非OA普船", tags: ["非OA", "ONE", "WHL", "YML", "HMM", "普船"] },
};

// ── 区域识别 ──
function identifyRegion(warehouseCode) {
  const wh = warehouseCode.toUpperCase().trim();
  // 美西仓库
  const west = ["ONT8", "LAX9", "LGB8", "SBD1", "SBD2", "SBD3", "SNA4", "LGB6", "LGB4", "LGB7",
    "LAS1", "LAS2", "LAS6", "LAS7", "GYR3", "PHX5", "PHX7", "PHX8", "OAK3", "OAK4", "OAK6",
    "SMF3", "SCK1", "SCK3", "SCK4", "XIX6", "RNO4", "SJC7", "FAT2", "ONT2", "ONT6", "ONT7",
    "ONT9", "KRB1", "KRB2", "KRB4", "KRB5", "PSP1", "SAN3", "SAN5", "BFL1", "ONT3", "ONT5",
    "POC1", "POC2", "POC3", "USWC1", "USWC2", "USWC5", "ABQ2", "PSC2", "LGB3", "LGB9",
    "VGT2", "LAS4", "LAS8", "PHX3", "PHX9", "TUS1", "TUS2", "LGB5", "GEG1", "PDX3", "PDX9",
    "BFI1", "BFI3", "BFI4", "BFI5", "BFI7", "RNO1", "BOI2", "SLC1", "SLC2", "SLC3"];
  if (west.some(w => wh.includes(w) || w.includes(wh.slice(0, 4)))) return "美西";

  // 美中仓库
  const central = ["MDW2", "MDW6", "MDW8", "MDW9", "ORD2", "ORD6", "IGQ2", "JVL1", "IND2",
    "IND3", "IND4", "IND5", "IND9", "MQJ1", "STL3", "STL4", "STL6", "MCI1", "MKC4", "MKC6",
    "FOE1", "ICT2", "OKC1", "OKC2", "DFW6", "DFW7", "FTW1", "FTW2", "FTW3", "FTW5", "FTW9",
    "DAL3", "IAH3", "HOU1", "HOU2", "HOU3", "HOU7", "HOU8", "SAT1", "SAT2", "SAT3", "SAT4",
    "SAT5", "SAT6", "AUS2", "AUS3", "RFD2", "RFD3", "RFD4", "FWA4", "DET1", "DET2",
    "CMH2", "CMH3", "CMH4", "CVG2", "CVG3", "CVG5", "CVG7", "SDF2", "SDF4", "SDF6", "SDF8",
    "SDF9", "LEX1", "LEX2", "LEX3", "MEM1", "MEM2", "MEM3", "MEM4", "MEM5", "MEM6",
    "BNA2", "BNA3", "BNA5", "CHA1", "CHA2", "HSV1", "HSV2", "ATL2", "ATL3", "ATL6", "ATL7",
    "ATL8", "MGE1", "MGE3", "MGE5", "SAV3", "GSP1", "CLT2", "CLT3", "CLT4", "CLT6",
    "CAE1", "CAE2", "CAE3", "PPO4", "MQJ1", "ORD6", "XIX1", "XIX2", "XIX3"];
  if (central.some(w => wh.includes(w) || w.includes(wh.slice(0, 4)))) return "美中";

  // 美东仓库
  const east = ["TEB3", "TEB4", "TEB6", "TEB9", "ABE2", "ABE3", "ABE4", "ABE5", "ABE8",
    "EWR4", "EWR5", "EWR6", "EWR7", "EWR9", "LGA9", "JFK2", "JFK8", "ACY1", "ACY2", "ACY3",
    "TTN2", "PHL1", "PHL3", "PHL4", "PHL5", "PHL6", "PHL7", "PHL8", "PHL9", "AVP1", "AVP3",
    "AVP9", "MDT1", "MDT4", "XEW5", "SWF1", "SWF2", "BDL6", "ALB1", "BOS7", "BOS1", "BOS2",
    "BWI2", "BWI5", "DCA1", "IAD1", "RIC1", "RIC2", "RIC3", "ORF2", "ORF3", "CHO1",
    "RMN3", "XRI3", "ILG1", "IUSL", "IUST", "PIT2", "HGR2", "WBW2", "HEA2", "RYY2",
    "TPA1", "TPA2", "TPA3", "TPA6", "MCO1", "MCO2", "MIA1", "MIA4", "MIA5", "MIA8",
    "FLL2", "PBI1", "PBI2", "PBI3", "JAX2", "JAX3", "JAX5", "JAX7", "TMB8", "TMB3",
    "RDG1", "HIA1", "RDU2", "RDU4", "GSO1", "XLX1", "XLX6", "TOL1", "MQY1", "MEM8",
    "BNA2", "BNA6", "IUSR", "XAV3", "WM-ATL1", "WM-ATL2", "WM-ATL3", "WM-MCO1", "XHH3"];
  if (east.some(w => wh.includes(w) || w.includes(wh.slice(0, 4)))) return "美东";

  return "美西"; // default
}

// ── 解析仓库列表 ──
function parseWarehouses(cell) {
  const text = String(cell).replace(/\r?\n/g, "/").replace(/\s+/g, "");
  return text.split(/[\/,，、]/).map(s => s.trim()).filter(s => s.length >= 3);
}

// ── 解析时效 ──
function parseTransit(text, extra) {
  const cleaned = String(text || "").replace(/\r?\n/g, " ").trim();
  const combined = [cleaned, String(extra || "").replace(/\r?\n/g, " ").trim()].filter(Boolean).join(" ");
  const match = combined.match(/(\d+)\s*[-–~约]*\s*(\d+)\s*(?:个)?(?:自然日|天|工作日)/);
  if (match) return { min: parseInt(match[1]), max: parseInt(match[2]), desc: combined };
  const single = combined.match(/(\d+)\s*(?:个)?(?:自然日|天|工作日)/);
  if (single) return { min: parseInt(single[1]), max: parseInt(single[1]), desc: combined };
  return { min: null, max: null, desc: combined };
}

// ── 通用记录生成 ──
function makeRecord(opts) {
  return {
    supplier: SUPPLIER,
    country: opts.country || detectCountry(opts.sourceSheet || "") || "美国",
    channel_name: opts.channelName || "",
    transport_mode: opts.transportMode || "海运",
    vessel_config: opts.vesselConfig || "",
    vessel_tags: opts.vesselTags || [],
    delivery_method: opts.deliveryMethod || "卡派",
    destination_type: opts.destType || "warehouse",
    destination_code: opts.destCode || "",
    destination_region: opts.destRegion || "",
    origin_region: opts.originRegion || "",
    origin_cities: opts.originCities || DEFAULT_CITIES,
    billing_type: opts.billingType || "包税",
    tax_mode: opts.taxMode || opts.billingType || "包税",
    min_quantity: opts.minQty || "",
    min_quantity_value: opts.minQtyValue || 0,
    unit_price: opts.price || 0,
    price_unit: "元/KG",
    cbm_price: opts.cbmPrice || null,
    transit_time_min: opts.transitMin || null,
    transit_time_max: opts.transitMax || null,
    transit_time_desc: opts.transitDesc || "",
    claim_rule: opts.claimRule || "",
    effective_date: "",
    source_sheet: opts.sourceSheet || "",
  };
}

// ═══════════════════════════════════════════════════════════════
// Sheet 解析器
// ═══════════════════════════════════════════════════════════════

// ── 美森以星合德OA海派 (邮编区域×渠道，海派) ──
function parseSeaExpress(ws) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 8) return [];

  const results = [];
  // R5: channel headers, R6: subtitles, R7: weight tiers, R8+: data rows
  const headerRow = data[4]; // R5 (0-indexed:4)
  const weightRow = data[6]; // R7

  // Identify channel groups from header
  const channels = [
    { name: "美森正班海派13日达", startCol: 2, vesselKey: "美森正班" },
    { name: "美森正班海派15日达", startCol: 5, vesselKey: "美森正班" },
    { name: "以星海派限时达", startCol: 8, vesselKey: "以星" },
    { name: "合德海派限时达", startCol: 11, vesselKey: "合德" },
    { name: "COSCO海派限时达", startCol: 14, vesselKey: "OA" },
    { name: "COSCO海派经济线", startCol: 17, vesselKey: "OA" },
  ];

  const regionMap = { "美国西部": "美西", "美国中部": "美中", "美国东部": "美东" };

  for (let ri = 7; ri < data.length; ri++) {
    const row = data[ri];
    const regionLabel = String(row[1] || "").replace(/\r?\n/g, " ").trim();
    if (!regionLabel) continue;
    // Extract region
    let destRegion = "";
    for (const [label, code] of Object.entries(regionMap)) {
      if (regionLabel.includes(label)) { destRegion = code; break; }
    }
    if (!destRegion) continue;

    for (const ch of channels) {
      for (let t = 0; t < 3; t++) {
        const price = parseFloat(row[ch.startCol + t]);
        const qtyLabels = ["12KG+", "45KG+", "100KG+"];
        const qtyValues = [12, 45, 100];
        if (!isNaN(price) && price > 0) {
          const vessel = CHANNEL_VESSEL[ch.vesselKey] || { config: "普船", tags: ["普船"] };
          const transitText = String(data[5][ch.startCol] || "").replace(/\r?\n/g, " ").trim();
          const transit = parseTransit(transitText);
          results.push(makeRecord({
            channelName: ch.name,
            vesselConfig: vessel.config,
            vesselTags: vessel.tags,
            deliveryMethod: "海派",
            destCode: destRegion,
            destType: "region",
            destRegion,
            billingType: "包税",
            minQty: qtyLabels[t],
            minQtyValue: qtyValues[t],
            price,
            transitMin: transit.min,
            transitMax: transit.max,
            transitDesc: transit.desc,
            sourceSheet: "美森以星合德OA海派",
          }));
        }
      }
    }
  }
  return results;
}

// ── 美森以星合德OA非OA海卡 (主表 — 仓×多渠道，卡派) ──
function parseSeaTruckMain(ws) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 8) return [];

  const results = [];
  // R5-R7: multi-tier headers, R8+: data
  const channels = [
    { name: "美森正班卡派限时达", startCol: 3, vesselKey: "美森正班", weightTiers: [
      { col: 3, qty: "12KG+", val: 12 }, { col: 4, qty: "100KG+", val: 100 },
    ], cbmCol: 5, noteCol: 6 },
    { name: "以星海卡", startCol: 7, vesselKey: "以星", weightTiers: [
      { col: 7, qty: "12KG+", val: 12 }, { col: 8, qty: "100KG+", val: 100 },
    ], cbmCol: 9, noteCol: 10 },
    { name: "合德海卡", startCol: 11, vesselKey: "合德", weightTiers: [
      { col: 11, qty: "12KG+", val: 12 }, { col: 12, qty: "100KG+", val: 100 },
    ], cbmCol: 13, noteCol: 14 },
    { name: "美西洛杉矶海卡-OA", startCol: 15, vesselKey: "OA", weightTiers: [
      { col: 15, qty: "12KG+", val: 12 }, { col: 16, qty: "100KG+", val: 100 },
    ], cbmCol: 17, noteCol: 18 },
    { name: "美西洛杉矶海卡-非OA", startCol: 19, vesselKey: "非OA", weightTiers: [
      { col: 19, qty: "12KG+", val: 12 },
    ], singlePrice: true },
  ];

  for (let ri = 7; ri < data.length; ri++) {
    const row = data[ri];
    const whCell = String(row[2] || "").trim();
    if (!whCell || whCell.length < 3) continue;
    if (whCell.includes("亚马逊仓库") || whCell.includes("渠道说明")) continue;

    const warehouses = parseWarehouses(whCell);
    const region = identifyRegion(warehouses[0] || "");

    for (const ch of channels) {
      for (const tier of ch.weightTiers) {
        const price = parseFloat(row[tier.col]);
        if (!isNaN(price) && price > 0) {
          const vessel = CHANNEL_VESSEL[ch.vesselKey] || { config: "普船", tags: ["普船"] };
          for (const wh of warehouses) {
            const cbmPrice = ch.cbmCol ? parseFloat(row[ch.cbmCol]) : null;
            results.push(makeRecord({
              channelName: ch.name,
              vesselConfig: vessel.config,
              vesselTags: vessel.tags,
              deliveryMethod: "卡派",
              destCode: wh,
              destRegion: region,
              billingType: "包税",
              minQty: tier.qty,
              minQtyValue: tier.val,
              price,
              cbmPrice: !isNaN(cbmPrice) && cbmPrice > 0 ? cbmPrice : null,
              claimRule: ch.noteCol ? String(row[ch.noteCol] || "").trim() : "",
              sourceSheet: "美森以星合德OA非OA海卡",
            }));
          }
        }
      }
      // CBM-only price for "不包1CBM+"
      if (ch.cbmCol) {
        const cbmP = parseFloat(row[ch.cbmCol]);
        if (!isNaN(cbmP) && cbmP > 0) {
          // CBM prices stored as cbm_price, mainly for reference
        }
      }
    }
  }
  return results;
}

// ── 美西洛杉矶海卡特惠 ──
function parseLASpecial(ws) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 8) return [];

  const results = [];
  const channels = [
    { name: "美森正班-美西洛杉矶海卡特惠", vesselKey: "美森正班", startCol: 3, tiers: [
      { col: 3, qty: "12KG+", val: 12 }, { col: 4, qty: "100KG+", val: 100 }, { col: 5, qty: "500KG+", val: 500 },
    ], cbmCol: 6, transitCol: 7 },
    { name: "以星合德统配-美西洛杉矶海卡特惠", vesselKey: "以星", startCol: 8, tiers: [
      { col: 8, qty: "12KG+", val: 12 }, { col: 9, qty: "100KG+", val: 100 }, { col: 10, qty: "500KG+", val: 500 },
    ], cbmCol: 11, transitCol: 12 },
    { name: "OA统配-美西洛杉矶海卡特惠", vesselKey: "OA", startCol: 13, tiers: [
      { col: 13, qty: "12KG+", val: 12 }, { col: 14, qty: "100KG+", val: 100 }, { col: 15, qty: "500KG+", val: 500 },
    ], cbmCol: 16, transitCol: 17 },
    { name: "非OA-美西洛杉矶海卡特惠", vesselKey: "非OA", startCol: 18, tiers: [
      { col: 18, qty: "12KG+", val: 12 }, { col: 19, qty: "100KG+", val: 100 },
    ]},
  ];

  for (let ri = 7; ri < data.length; ri++) {
    const row = data[ri];
    const whCell = String(row[2] || "").trim();
    if (!whCell || whCell.length < 3) continue;
    if (whCell.includes("亚马逊仓库")) continue;

    const warehouses = parseWarehouses(whCell);
    const region = "美西";

    for (const ch of channels) {
      for (const tier of ch.tiers) {
        const price = parseFloat(row[tier.col]);
        if (!isNaN(price) && price > 0) {
          const vessel = CHANNEL_VESSEL[ch.vesselKey] || { config: "普船", tags: ["普船"] };
          const transit = ch.transitCol ? parseTransit(row[ch.transitCol]) : { min: null, max: null, desc: "" };
          for (const wh of warehouses) {
            results.push(makeRecord({
              channelName: ch.name,
              vesselConfig: vessel.config,
              vesselTags: vessel.tags,
              deliveryMethod: "卡派",
              destCode: wh,
              destRegion: region,
              billingType: "包税",
              minQty: tier.qty,
              minQtyValue: tier.val,
              price,
              cbmPrice: ch.cbmCol ? parseFloat(row[ch.cbmCol]) : null,
              transitMin: transit.min,
              transitMax: transit.max,
              transitDesc: transit.desc,
              sourceSheet: "美西洛杉矶海卡特惠",
            }));
          }
        }
      }
    }
  }
  return results;
}

// ── 美中休斯顿海卡 ──
function parseHouston(ws) {
  return parseSimpleRegionSheet(ws, "美中休斯顿海卡", "美中", [
    { col: 3, qty: "12KG+", val: 12 },
    { col: 4, qty: "100KG+", val: 100 },
  ], 5, 6, 6);
}

// ── 美中芝加哥海卡 ──
function parseChicago(ws) {
  return parseSimpleRegionSheet(ws, "美中芝加哥海卡", "美中", [
    { col: 3, qty: "100KG+", val: 100 },
    { col: 4, qty: "300KG+", val: 300 },
    { col: 5, qty: "1000KG+", val: 1000 },
  ], 6, 7, 7);
}

// ── 美东萨凡纳海卡 ──
function parseSavannah(ws) {
  return parseSimpleRegionSheet(ws, "美东萨凡纳海卡", "美东", [
    { col: 3, qty: "100KG+", val: 100 },
    { col: 4, qty: "300KG+", val: 300 },
  ], 5, 6, 7);
}

// ── 美东纽约OA非OA海卡 ──
function parseNewYork(ws) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 7) return [];

  const results = [];
  const channels = [
    { name: "美东纽约OA海卡", vesselKey: "OA", tiers: [
      { col: 3, qty: "12KG+", val: 12 }, { col: 4, qty: "100KG+", val: 100 },
    ], cbmCol: 5, noteCol: 6 },
    { name: "美东纽约海卡经济线", vesselKey: "非OA", tiers: [
      { col: 8, qty: "12KG+", val: 12 }, { col: 9, qty: "100KG+", val: 100 },
    ], cbmCol: 10, noteCol: 11 },
  ];

  for (let ri = 6; ri < data.length; ri++) {
    const row = data[ri];
    const whCell = String(row[2] || "").trim();
    if (!whCell || whCell.length < 3) continue;
    if (whCell.includes("亚马逊仓库")) continue;

    const warehouses = parseWarehouses(whCell);
    const region = "美东";

    for (const ch of channels) {
      for (const tier of ch.tiers) {
        const price = parseFloat(row[tier.col]);
        if (!isNaN(price) && price > 0) {
          const vessel = CHANNEL_VESSEL[ch.vesselKey] || { config: "普船", tags: ["普船"] };
          const noteText = ch.noteCol ? String(row[ch.noteCol] || "").trim() : "";
          const transit = parseTransit(noteText);
          for (const wh of warehouses) {
            results.push(makeRecord({
              channelName: ch.name,
              vesselConfig: vessel.config,
              vesselTags: vessel.tags,
              deliveryMethod: "卡派",
              destCode: wh,
              destRegion: region,
              billingType: "包税",
              minQty: tier.qty,
              minQtyValue: tier.val,
              price,
              cbmPrice: ch.cbmCol ? parseFloat(row[ch.cbmCol]) : null,
              transitMin: transit.min,
              transitMax: transit.max,
              transitDesc: transit.desc || noteText,
              claimRule: noteText,
              sourceSheet: "美东纽约OA非OA海卡",
            }));
          }
        }
      }
    }
  }
  return results;
}

// ── 美西美东商私卡 ──
function parseCommercialCard(ws) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 9) return [];

  const results = [];
  // Distance label in col1, prices from col2
  // R5: col1=目的地距离, col2=美森正班, col6=以星海卡, col10=美西洛杉矶, col14=美东纽约
  const channels = [
    { name: "美森正班-商私卡", vesselKey: "美森正班", tiers: [
      { col: 2, qty: "300KG+", val: 300 }, { col: 3, qty: "500KG+", val: 500 },
      { col: 4, qty: "1000KG+", val: 1000 }, { col: 5, qty: "2000KG+", val: 2000 },
    ]},
    { name: "以星海卡-商私卡", vesselKey: "以星", tiers: [
      { col: 6, qty: "300KG+", val: 300 }, { col: 7, qty: "500KG+", val: 500 },
      { col: 8, qty: "1000KG+", val: 1000 }, { col: 9, qty: "2000KG+", val: 2000 },
    ]},
    { name: "美西洛杉矶海卡-商私卡", vesselKey: "OA", tiers: [
      { col: 10, qty: "300KG+", val: 300 }, { col: 11, qty: "500KG+", val: 500 },
      { col: 12, qty: "1000KG+", val: 1000 }, { col: 13, qty: "2000KG+", val: 2000 },
    ]},
    { name: "美东纽约海卡-商私卡", vesselKey: "非OA", tiers: [
      { col: 14, qty: "300KG+", val: 300 },
    ]},
  ];

  for (let ri = 8; ri < data.length; ri++) {
    const row = data[ri];
    const distCell = String(row[1] || "").trim(); // Distance in col 1
    if (!distCell || distCell.includes("单询") || distCell.includes("以上")) continue;
    if (!distCell.includes("英里")) continue;

    const distMatch = distCell.match(/(\d+)/);
    const distance = distMatch ? parseInt(distMatch[1]) : 0;
    const destCode = `商私_${distance}英里`;

    for (const ch of channels) {
      for (const tier of ch.tiers) {
        const price = parseFloat(row[tier.col]);
        if (!isNaN(price) && price > 0) {
          const vessel = CHANNEL_VESSEL[ch.vesselKey] || { config: "普船", tags: ["普船"] };
          results.push(makeRecord({
            channelName: ch.name,
            vesselConfig: vessel.config,
            vesselTags: vessel.tags,
            deliveryMethod: "卡派",
            destCode,
            destType: "commercial",
            destRegion: "全美",
            billingType: "包税",
            minQty: tier.qty,
            minQtyValue: tier.val,
            price,
            sourceSheet: "美西美东商私卡",
          }));
        }
      }
    }
  }
  return results;
}

// ── 美国空运 ──
function parseAirFreight(ws) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 9) return [];

  const results = [];
  const channels = [
    { name: "空运5日提-普货", transportMode: "空运", startCol: 3, numWeightTiers: 6 },
    { name: "空运5日提-带电/敏感", transportMode: "空运", startCol: 9, numWeightTiers: 6 },
    { name: "空运普货", transportMode: "空运", startCol: 15, numWeightTiers: 5 },
  ];

  const weightLabels = ["12KG+", "45KG+", "100KG+", "300KG+", "500KG+", "1000KG+"];
  const weightValues = [12, 45, 100, 300, 500, 1000];

  const regionMap = { "美国西部": "美西", "美国中部": "美中", "美国东部": "美东" };

  for (let ri = 8; ri < data.length; ri++) {
    const row = data[ri];
    // Region/warehouse label in col 2 (not col 1)
    const regionLabel = String(row[2] || "").replace(/\r?\n/g, " ").trim();
    if (!regionLabel) continue;

    let destRegion = "";
    for (const [label, code] of Object.entries(regionMap)) {
      if (regionLabel.includes(label)) { destRegion = code; break; }
    }

    // Check if warehouse-based row
    let warehouses = [];
    let isWarehouseBased = false;

    if (!destRegion && regionLabel.length >= 3) {
      // Could be warehouse or special location row
      const whText = String(row[2] || "").replace(/\r?\n/g, "/").replace(/\s+/g, "");
      if (whText.includes("(") || whText.includes("/")) {
        // Extract warehouse codes from parentheses or slash-separated list
        const parenMatch = whText.match(/\(([^)]+)\)/);
        if (parenMatch) {
          warehouses = parenMatch[1].split(/[\/,，、]/).map(s => s.trim()).filter(s => s.length >= 3 && s.match(/[A-Z0-9]/));
        }
        if (warehouses.length === 0) {
          warehouses = whText.replace(/\([^)]*\)/g, "").split(/[\/,，、]/).map(s => s.trim()).filter(s => s.length >= 3);
        }
      }
      // If it mentions 卡派直送 or specific warehouse pattern
      if (warehouses.length > 0 || regionLabel.includes("卡派") || regionLabel.match(/[A-Z]{2,}\d/)) {
        isWarehouseBased = true;
      } else if (regionLabel.includes("代发仓")) {
        isWarehouseBased = true;
        const parenMatch = whText.match(/\(([^)]+)\)/);
        if (parenMatch) {
          warehouses = parenMatch[1].split(/[\/,，、]/).map(s => s.trim()).filter(s => s.length >= 3);
        } else {
          warehouses = [regionLabel]; // use full label as warehouse
        }
      }
    }

    for (const ch of channels) {
      const numTiers = ch.numWeightTiers;
      for (let t = 0; t < numTiers; t++) {
        const price = parseFloat(row[ch.startCol + t]);
        if (!isNaN(price) && price > 0) {
          if (isWarehouseBased && warehouses.length > 0) {
            for (const wh of warehouses) {
              results.push(makeRecord({
                channelName: ch.name,
                transportMode: ch.transportMode,
                vesselConfig: "空运",
                vesselTags: ["空运"],
                deliveryMethod: wh.includes("卡派") ? "卡派" : "快递派",
                destCode: wh.replace(/卡派直送|直送/g, "").trim(),
                destRegion: identifyRegion(wh),
                billingType: "包税",
                minQty: weightLabels[t],
                minQtyValue: weightValues[t],
                price,
                sourceSheet: "美国空运",
              }));
            }
          } else if (destRegion) {
            results.push(makeRecord({
              channelName: ch.name,
              transportMode: ch.transportMode,
              vesselConfig: "空运",
              vesselTags: ["空运"],
              deliveryMethod: "快递派",
              destCode: destRegion,
              destType: "region",
              destRegion,
              billingType: "包税",
              minQty: weightLabels[t],
              minQtyValue: weightValues[t],
              price,
              sourceSheet: "美国空运",
            }));
          }
        }
      }
    }
  }
  return results;
}

// ── 通用单渠道×仓库 Sheet 解析 ──
function parseSimpleRegionSheet(ws, sheetName, defaultRegion, weightTiers, cbmCol, transitCol, dataStartRow) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < dataStartRow) return [];

  const results = [];
  for (let ri = dataStartRow - 1; ri < data.length; ri++) {
    const row = data[ri];
    const whCell = String(row[2] || "").trim();
    if (!whCell || whCell.length < 3) continue;
    if (whCell.includes("亚马逊仓库") || whCell.includes("自提")) continue;

    const warehouses = parseWarehouses(whCell);
    const transit = transitCol ? parseTransit(row[transitCol]) : { min: null, max: null, desc: "" };
    const noteText = transitCol ? String(row[transitCol] || "").trim() : "";

    for (const tier of weightTiers) {
      const price = parseFloat(row[tier.col]);
      if (!isNaN(price) && price > 0) {
        for (const wh of warehouses) {
          const region = identifyRegion(wh) || defaultRegion;
          results.push(makeRecord({
            channelName: sheetName,
            vesselConfig: "普船",
            vesselTags: ["普船"],
            deliveryMethod: "卡派",
            destCode: wh,
            destRegion: region,
            billingType: "包税",
            minQty: tier.qty,
            minQtyValue: tier.val,
            price,
            cbmPrice: cbmCol ? parseFloat(row[cbmCol]) : null,
            transitMin: transit.min,
            transitMax: transit.max,
            transitDesc: transit.desc || noteText,
            claimRule: noteText,
            sourceSheet: sheetName,
          }));
        }
      }
    }
  }
  return results;
}

// ═══════════════════════════════════════════════════════════════
// 多国通用解析器（英国 / 欧洲 / 加拿大）
// ═══════════════════════════════════════════════════════════════

/** 从超大件 Sheet 提取渠道组配置（邮编×渠道×重量阶梯 模式） */
function parseOversizePostalMatrix(ws, config) {
  // config: { dataStartRow, postalCol, channels: [{name, startCol, numTiers, weightLabels, weightValues, transportMode}] }
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < config.dataStartRow) return [];
  const results = [];

  for (let ri = config.dataStartRow - 1; ri < data.length; ri++) {
    const row = data[ri];
    const postalRaw = String(row[config.postalCol] || "").trim();
    if (!postalRaw || postalRaw.length < 2) continue;
    if (postalRaw.includes("渠道说明") || postalRaw.includes("邮编") || postalRaw.includes("超大件")) continue;
    // Skip header/metadata rows
    if (postalRaw.includes("注意事项") || postalRaw.includes("返回目录")) continue;

    for (const ch of config.channels) {
      for (let t = 0; t < ch.numTiers; t++) {
        const price = parseFloat(row[ch.startCol + t]);
        if (!isNaN(price) && price > 0) {
          results.push(makeRecord({
            channelName: ch.name,
            transportMode: ch.transportMode || "空运",
            vesselConfig: ch.vesselConfig || ch.name,
            vesselTags: ch.vesselTags || [],
            deliveryMethod: ch.deliveryMethod || "卡派",
            destCode: postalRaw.slice(0, 50),
            destType: "postal",
            destRegion: config.destRegion || "",
            billingType: "包税",
            minQty: (ch.weightLabels || [])[t] || `${(ch.weightValues || [])[t] || 0}KG+`,
            minQtyValue: (ch.weightValues || [])[t] || 0,
            price,
            price_unit: "元/KG",
            sourceSheet: config.sourceSheet || "",
          }));
        }
      }
    }
  }
  return results;
}

/** 英国超大件 (邮编 × 空运普货/空运带电/卡航) */
function parseUKOversize(ws) {
  return parseOversizePostalMatrix(ws, {
    dataStartRow: 10,
    postalCol: 2,
    destRegion: "英国",
    sourceSheet: "英国超大件",
    channels: [
      { name: "英国超大件-空运普货", startCol: 3, numTiers: 4, transportMode: "空运",
        weightLabels: ["45KG+","100KG+","300KG+","500KG+"], weightValues: [45,100,300,500],
        vesselConfig: "空运", vesselTags: ["空运"], deliveryMethod: "卡派" },
      { name: "英国超大件-空运带电", startCol: 7, numTiers: 4, transportMode: "空运",
        weightLabels: ["45KG+","100KG+","300KG+","500KG+"], weightValues: [45,100,300,500],
        vesselConfig: "空运", vesselTags: ["空运","带电"], deliveryMethod: "卡派" },
      { name: "英国超大件-卡航", startCol: 11, numTiers: 4, transportMode: "卡航",
        weightLabels: ["45KG+","100KG+","300KG+","500KG+"], weightValues: [45,100,300,500],
        vesselConfig: "卡航", vesselTags: ["卡航"], deliveryMethod: "卡派" },
    ],
  });
}

/** 欧洲超大件 (国家 × 邮编 × 空派普货/空派带电/卡航) */
function parseEUOversize(ws) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 10) return [];
  const results = [];

  // 渠道: col3-7=空派普货(68,100,300,500,1000), col8-12=空派带电, col13-14=卡航
  const channels = [
    { name: "欧洲超大件-空派普货", startCol: 3, numTiers: 5, transportMode: "空运",
      weightLabels: ["68KG","100KG","300KG+","500KG+","1000KG+"], weightValues: [68,100,300,500,1000],
      vesselConfig: "空运", vesselTags: ["空运"], deliveryMethod: "快递派" },
    { name: "欧洲超大件-空派带电", startCol: 8, numTiers: 5, transportMode: "空运",
      weightLabels: ["68KG","100KG","300KG+","500KG+","1000KG+"], weightValues: [68,100,300,500,1000],
      vesselConfig: "空运", vesselTags: ["空运","带电"], deliveryMethod: "快递派" },
    { name: "欧洲超大件-卡航", startCol: 13, numTiers: 2, transportMode: "卡航",
      weightLabels: ["68KG","100KG"], weightValues: [68,100],
      vesselConfig: "卡航", vesselTags: ["卡航"], deliveryMethod: "快递派" },
  ];

  let currentCountry = "";
  for (let ri = 9; ri < data.length; ri++) {
    const row = data[ri];
    const countryCell = String(row[1] || "").trim();
    const postalRaw = String(row[2] || "").trim();

    if (countryCell && countryCell.length >= 2 && !countryCell.match(/^\d/)) {
      currentCountry = countryCell;
    }
    if (!postalRaw || postalRaw.length < 2) continue;
    if (postalRaw.includes("渠道说明") || postalRaw.includes("国家")) continue;

    for (const ch of channels) {
      for (let t = 0; t < ch.numTiers; t++) {
        const price = parseFloat(row[ch.startCol + t]);
        if (!isNaN(price) && price > 0) {
          results.push(makeRecord({
            country: detectCountry(currentCountry) || "欧线",
            channelName: `${ch.name}(${currentCountry})`,
            transportMode: ch.transportMode,
            vesselConfig: ch.vesselConfig,
            vesselTags: ch.vesselTags,
            deliveryMethod: ch.deliveryMethod,
            destCode: `${currentCountry}_${postalRaw.slice(0, 30)}`,
            destType: "postal",
            destRegion: currentCountry,
            billingType: "包税",
            minQty: ch.weightLabels[t],
            minQtyValue: ch.weightValues[t],
            price,
            price_unit: "元/KG",
            sourceSheet: "欧洲超大件",
          }));
        }
      }
    }
  }
  return results;
}

/** 英国常规 (非超大件: 空运/卡航/铁路/海运) */
function parseUKAirSeaRail(ws) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 9) return [];
  const results = [];

  // 这个 Sheet 有多个渠道组，每组有独立重量阶梯
  // 基本模式: col0=送达方式, col1=税务类型, 后续按渠道组排列
  const channelGroups = [
    // DPD派 - 空运5日提普货: cols 2-5 (21/45/100/1000)
    { name: "英国空运5日提普货-DPD", startCol: 2, numTiers: 4, transportMode: "空运",
      weightLabels: ["21KG+","45KG+","100KG+","1000KG+"], weightValues: [21,45,100,1000],
      vesselConfig: "空运5日提", vesselTags: ["空运"], deliveryMethod: "快递派" },
    // 空运5日提带电: cols 6-9
    { name: "英国空运5日提带电-DPD", startCol: 6, numTiers: 4, transportMode: "空运",
      weightLabels: ["21KG+","45KG+","100KG+","1000KG+"], weightValues: [21,45,100,1000],
      vesselConfig: "空运5日提", vesselTags: ["空运","带电"], deliveryMethod: "快递派" },
    // 空运9日提普货: cols 10-11
    { name: "英国空运9日提普货-DPD", startCol: 10, numTiers: 2, transportMode: "空运",
      weightLabels: ["21KG+","45KG+"], weightValues: [21,45],
      vesselConfig: "空运9日提", vesselTags: ["空运"], deliveryMethod: "快递派" },
  ];

  // 遍历数据行
  for (let ri = 8; ri < data.length; ri++) {
    const row = data[ri];
    const col0 = String(row[0] || "").trim();
    const col1 = String(row[1] || "").trim();

    // 跳过空行和章节标题
    if (!col0 && !col1) continue;
    if (col0.includes("渠道说明") || col0.includes("空运") || col0.includes("临时")) continue;

    const destLabel = (col0 + "\n" + col1).trim();
    for (const ch of channelGroups) {
      for (let t = 0; t < ch.numTiers; t++) {
        const price = parseFloat(row[ch.startCol + t]);
        if (!isNaN(price) && price > 0 && price < 999) {
          results.push(makeRecord({
            country: "英国",
            channelName: ch.name,
            transportMode: ch.transportMode,
            vesselConfig: ch.vesselConfig,
            vesselTags: ch.vesselTags,
            deliveryMethod: ch.deliveryMethod,
            destCode: destLabel.slice(0, 60),
            destType: "region",
            destRegion: "英国",
            billingType: destLabel.includes("包税") ? "包税" : destLabel.includes("递延") ? "递延" : "自税",
            minQty: ch.weightLabels[t],
            minQtyValue: ch.weightValues[t],
            price,
            price_unit: "元/KG",
            sourceSheet: "英国空运海运铁路卡航",
          }));
        }
      }
    }
  }
  return results;
}

/** 欧洲常规 (非超大件: 仓库×渠道) */
function parseEUAirSeaRail(ws) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 9) return [];
  const results = [];

  // 渠道: col2-4=空运6日提普货(21/45/100), col5-7=空运6日提带电, col8-10=空运9日提普货, col11-13=空运9日提带电
  const channels = [
    { name: "欧洲空运6日提普货", startCol: 2, numTiers: 3, transportMode: "空运",
      weightLabels: ["21KG+","45KG+","100KG+"], weightValues: [21,45,100],
      vesselConfig: "空运6日提", vesselTags: ["空运"], deliveryMethod: "快递派" },
    { name: "欧洲空运6日提带电", startCol: 5, numTiers: 3, transportMode: "空运",
      weightLabels: ["21KG+","45KG+","100KG+"], weightValues: [21,45,100],
      vesselConfig: "空运6日提", vesselTags: ["空运","带电"], deliveryMethod: "快递派" },
    { name: "欧洲空运9日提普货", startCol: 8, numTiers: 3, transportMode: "空运",
      weightLabels: ["21KG+","45KG+","100KG+"], weightValues: [21,45,100],
      vesselConfig: "空运9日提", vesselTags: ["空运"], deliveryMethod: "快递派" },
    { name: "欧洲空运9日提带电", startCol: 11, numTiers: 3, transportMode: "空运",
      weightLabels: ["21KG+","45KG+","100KG+"], weightValues: [21,45,100],
      vesselConfig: "空运9日提", vesselTags: ["空运","带电"], deliveryMethod: "快递派" },
  ];

  for (let ri = 8; ri < data.length; ri++) {
    const row = data[ri];
    const col0 = String(row[0] || "").trim(); // country
    const col1 = String(row[1] || "").trim(); // warehouse info

    if (!col0 && !col1) continue;
    if (col0.includes("渠道说明") || col0.includes("国家")) continue;

    // Extract warehouse code: look for patterns like "38350-HAJ1" or "HAJ1"
    const whMatch = (col0 + " " + col1).match(/([A-Z]{3,4}\d|[A-Z]{2,3}\d{1,2}[A-Z]?)/g);
    const whCodes = whMatch ? [...new Set(whMatch)] : [];
    const destLabel = (col0 + " " + col1).replace(/\n/g, " ").trim().slice(0, 100);

    for (const ch of channels) {
      for (let t = 0; t < ch.numTiers; t++) {
        const price = parseFloat(row[ch.startCol + t]);
        if (!isNaN(price) && price > 0 && price < 999) {
          if (whCodes.length > 0) {
            for (const wh of whCodes) {
              results.push(makeRecord({
                country: "欧线",
                channelName: ch.name,
                transportMode: ch.transportMode,
                vesselConfig: ch.vesselConfig,
                vesselTags: ch.vesselTags,
                deliveryMethod: ch.deliveryMethod,
                destCode: wh,
                destType: "warehouse",
                destRegion: "欧线",
                billingType: "包税",
                minQty: ch.weightLabels[t],
                minQtyValue: ch.weightValues[t],
                price,
                price_unit: "元/KG",
                sourceSheet: "欧洲空运海运铁路卡航",
              }));
            }
          } else {
            results.push(makeRecord({
              country: "欧线",
              channelName: ch.name,
              transportMode: ch.transportMode,
              vesselConfig: ch.vesselConfig,
              vesselTags: ch.vesselTags,
              deliveryMethod: ch.deliveryMethod,
              destCode: destLabel.slice(0, 60),
              destType: "region",
              destRegion: "欧线",
              billingType: "包税",
              minQty: ch.weightLabels[t],
              minQtyValue: ch.weightValues[t],
              price,
              price_unit: "元/KG",
              sourceSheet: "欧洲空运海运铁路卡航",
            }));
          }
        }
      }
    }
  }
  return results;
}

// ═══════════════════════════════════════════════════════════════
// 主入口
// ═══════════════════════════════════════════════════════════════

function parseHaopeng(filePath) {
  console.log("[皓鹏] 开始解析:", filePath);
  const wb = XLSX.readFile(filePath);
  const allResults = [];

  // 已知 Sheet → 解析器（所有国家共用同一套解析器，国家由 Sheet 名检测）
  const sheetParsers = {
    "美森以星合德OA海派": parseSeaExpress,
    "美森以星合德OA非OA海卡": parseSeaTruckMain,
    "美西洛杉矶海卡特惠": parseLASpecial,
    "美中休斯顿海卡": parseHouston,
    "美中芝加哥海卡": parseChicago,
    "美东萨凡纳海卡": parseSavannah,
    "美西美东商私卡": parseCommercialCard,
    "美国空运": parseAirFreight,
    "美国空运小货特快": parseAirFreight,
  };

  // 英国 Sheet（名称含尾部空格，模糊匹配）
  const ukOversizeName = wb.SheetNames.find(n => n.includes("英国") && n.includes("超大件"));
  const ukNormalName = wb.SheetNames.find(n => n.trim() === "英国空运海运铁路卡航");
  // 欧洲 Sheet
  const euOversizeName = wb.SheetNames.find(n => n.includes("欧洲") && n.includes("超大件") && n.includes("铁路卡航"));
  const euNormalName = wb.SheetNames.find(n => n.trim() === "欧洲空运海运铁路卡航");

  if (ukOversizeName) sheetParsers[ukOversizeName] = parseUKOversize;
  if (ukNormalName) sheetParsers[ukNormalName] = parseUKAirSeaRail;
  if (euOversizeName) sheetParsers[euOversizeName] = parseEUOversize;
  if (euNormalName) sheetParsers[euNormalName] = parseEUAirSeaRail;

  // Handle NY sheet with possible trailing space
  const nySheetName = wb.SheetNames.find(n => n.includes("美东纽约OA非OA海卡"));
  if (nySheetName) {
    sheetParsers[nySheetName] = parseNewYork;
  }

  // 报告检测到的国家
  const allCountries = detectAllCountries(wb);
  console.log(`[皓鹏] 检测到国家: ${allCountries.join(", ") || "仅美国"}`);

  const skippedSheets = [];

  // 按 Sheet 名逐个匹配解析器
  for (const sheetName of wb.SheetNames) {
    if (sheetName.includes("渠道目录") || sheetName.includes("目录") || sheetName.startsWith("暂停")) continue;

    const parser = sheetParsers[sheetName];
    if (parser) {
      try {
        const results = parser(wb.Sheets[sheetName]);
        // ✅ 根据 Sheet 名自动覆盖国家（核心改进）
        const sheetCountry = detectCountry(sheetName);
        for (const r of results) {
          if (sheetCountry) r.country = sheetCountry;
        }
        console.log(`  [${sheetName.trim()}] ${results.length} 条 → ${sheetCountry || "美国"}`);
        allResults.push(...results);
      } catch (err) {
        console.error(`  [${sheetName.trim()}] 解析失败: ${err.message}`);
      }
    } else {
      const country = detectCountry(sheetName);
      if (country) {
        skippedSheets.push(`${sheetName.trim()} (${country})`);
      } else {
        skippedSheets.push(sheetName.trim());
      }
    }
  }

  if (skippedSheets.length > 0) {
    console.log(`[皓鹏] ⚠ 暂未解析的Sheet (需补充解析器): ${skippedSheets.join(" | ")}`);
  }

  console.log(`[皓鹏] 总计 ${allResults.length} 条`);
  return allResults;
}

module.exports = { parseHaopeng };
