/**
 * 皓辉国际 — 美线价格解析器
 *
 * 核心数据来源:
 * 1. "FBA渠道汇总" Sheet — 扁平结构(仓×渠道)，382条有效记录
 * 2. "海派服务" Sheet — 邮编区域×渠道，快递派送
 *
 * 渠道→船司映射:
 *   皓森达 = 美森正班 CLX
 *   皓速达 = 以星 ZIM-ZEX
 *   皓速达带托 = 以星 ZIM-ZEX 带车架
 *   皓快达 = COSCO/EMC/OOCL (OA普船)
 *   皓东达 = COSCO/OOCL/ZIM 美东直航
 *   纽约皓速达 = EXX快船 IPI海铁联运
 */

const XLSX = require("xlsx");

const SUPPLIER = "皓辉国际";
const COUNTRY = "美国";

// ── 渠道→船司映射 ──
const CHANNEL_VESSEL_MAP = {
  "皓森达": { config: "美森正班CLX", tags: ["美森", "Matson", "CLX"] },
  "皓速达": { config: "以星ZIM-ZEX", tags: ["以星", "ZIM", "ZEX"] },
  "皓速达带托": { config: "以星ZIM-ZEX带车架", tags: ["以星", "ZIM", "ZEX", "带车架"] },
  "皓快达": { config: "COSCO/EMC/OOCL普船", tags: ["COSCO", "EMC", "OOCL", "OA", "普船"] },
  "皓东达": { config: "COSCO/OOCL/ZIM美东直航", tags: ["COSCO", "OOCL", "ZIM", "美东直航"] },
  "纽约皓速达": { config: "EXX快船IPI海铁联运", tags: ["EXX", "IPI", "海铁联运"] },
};

// ── 默认发货城市 ──
const DEFAULT_CITIES = ["深圳", "广州", "东莞", "中山", "惠州", "义乌", "上海", "宁波", "杭州", "厦门", "泉州", "福州"];

// ── 推断船配置 ──
function inferVessel(channelName) {
  for (const [key, val] of Object.entries(CHANNEL_VESSEL_MAP)) {
    if (channelName.includes(key)) return val;
  }
  return { config: "普船", tags: ["普船"] };
}

// ── 解析时效 ──
function parseTransit(text) {
  const cleaned = String(text).replace(/\r?\n/g, " ").trim();
  // "17-21个自然日" or "开船后20-25个自然日提取"
  const match = cleaned.match(/(\d+)\s*[-–]\s*(\d+)\s*(?:个)?自然日/);
  if (match) return { min: parseInt(match[1]), max: parseInt(match[2]), desc: cleaned };
  const singleMatch = cleaned.match(/(\d+)\s*(?:个)?自然日/);
  if (singleMatch) return { min: parseInt(singleMatch[1]), max: parseInt(singleMatch[1]), desc: cleaned };
  return { min: null, max: null, desc: cleaned };
}

// ── 解析索赔规则 ──
function parseClaim(text) {
  const cleaned = String(text).replace(/\r?\n/g, " ").trim();
  if (!cleaned || cleaned === "*" || cleaned === "0") return "";
  return cleaned;
}

// ── 解析 FBA渠道汇总 (主数据源) ──
function parseChannelSummary(ws) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 2) return [];

  const results = [];
  // R1=header, R2+=data
  // Col 0: 匹配码, Col 1: 仓库代码, Col 2: 渠道名称, Col 3: 12KG+, Col 4: 51KG+,
  // Col 5: 签收/提取时效, Col 6: 理赔时效

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const whCode = String(row[1] || "").trim();
    const channelName = String(row[2] || "").trim();
    const price12 = parseFloat(row[3]);
    const price51 = parseFloat(row[4]);
    const transitText = String(row[5] || "");
    const claimText = String(row[6] || "");

    if (!whCode || !channelName) continue;
    if (whCode.length < 3) continue;

    const vessel = inferVessel(channelName);
    const transit = parseTransit(transitText);
    const claim = parseClaim(claimText);
    const isHaiPai = channelName.includes("海派");
    const deliveryMethod = isHaiPai ? "海派" : "卡派";

    // 12KG+ 价格
    if (!isNaN(price12) && price12 > 0) {
      results.push(makeRecord({
        channelName: channelName.replace("-FBA海卡", "").replace("-海派", ""),
        vesselConfig: vessel.config,
        vesselTags: vessel.tags,
        deliveryMethod,
        destCode: whCode,
        billingType: "包税",
        minQty: "12KG+",
        minQtyValue: 12,
        price: price12,
        transitMin: transit.min,
        transitMax: transit.max,
        transitDesc: transit.desc,
        claimRule: claim,
        sourceSheet: "FBA渠道汇总",
      }));
    }

    // 51KG+ 价格
    if (!isNaN(price51) && price51 > 0) {
      results.push(makeRecord({
        channelName: channelName.replace("-FBA海卡", "").replace("-海派", ""),
        vesselConfig: vessel.config,
        vesselTags: vessel.tags,
        deliveryMethod,
        destCode: whCode,
        billingType: "包税",
        minQty: "51KG+",
        minQtyValue: 51,
        price: price51,
        transitMin: transit.min,
        transitMax: transit.max,
        transitDesc: transit.desc,
        claimRule: claim,
        sourceSheet: "FBA渠道汇总",
      }));
    }
  }

  return results;
}

// ── 解析 海派服务 (邮编区域×渠道) ──
function parseHaiPai(ws) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 8) return [];

  const results = [];
  let currentChannel = "";

  for (let i = 7; i < data.length; i++) {
    const row = data[i];
    const col0 = String(row[0] || "").trim();
    const col1 = String(row[1] || "").trim();

    // 检测渠道名行
    if (col0 && (col0.includes("皓") || col0.includes("海派"))) {
      currentChannel = col0.replace(/\r?\n/g, " ").trim();
      continue;
    }

    // 空行或说明行
    if (!col1 || !currentChannel) continue;
    if (!col1.includes("邮编") && !col1.includes("美国")) continue;

    // 解析区域
    const regionText = col1.replace(/\r?\n/g, " ").trim();
    let destRegion = "";
    if (regionText.includes("西部") || regionText.includes("8-96") || regionText.includes("97-99")) {
      destRegion = "美西";
    } else if (regionText.includes("中部") || regionText.includes("5、6、7")) {
      destRegion = "美中";
    } else if (regionText.includes("东部") || regionText.includes("0、1、2、3、4")) {
      destRegion = "美东";
    }

    const vessel = inferVessel(currentChannel);
    const transitText = i === 7 ? String(row[8] || "") : ""; // 只在第一行有时效
    const transit = parseTransit(transitText || getDefaultTransit(currentChannel));

    // 广州仓价格: col 2,3,4 (12KG+, 51KG+, 101KG+)
    const gzPrice12 = parseFloat(row[2]);
    const gzPrice51 = parseFloat(row[3]);
    const gzPrice101 = parseFloat(row[4]);

    const tiers = [
      { qty: "12KG+", val: 12, price: gzPrice12 },
      { qty: "51KG+", val: 51, price: gzPrice51 },
      { qty: "101KG+", val: 101, price: gzPrice101 },
    ];

    for (const tier of tiers) {
      if (!isNaN(tier.price) && tier.price > 0) {
        results.push(makeRecord({
          channelName: currentChannel.replace(/\r?\n.*/s, "").trim() + "-海派",
          vesselConfig: vessel.config,
          vesselTags: vessel.tags,
          deliveryMethod: "海派",
          destCode: destRegion,
          destType: "region",
          billingType: "包税",
          minQty: tier.qty,
          minQtyValue: tier.val,
          price: tier.price,
          transitMin: transit.min,
          transitMax: transit.max,
          transitDesc: transit.desc,
          claimRule: "",
          sourceSheet: "海派服务",
        }));
      }
    }
  }

  return results;
}

function getDefaultTransit(channelName) {
  if (channelName.includes("皓森达")) return "12-15个自然日";
  if (channelName.includes("皓速达")) return "14-17个自然日";
  if (channelName.includes("皓快达")) return "20-25个自然日";
  return "";
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
    destination_type: opts.destType || "warehouse",
    destination_code: opts.destCode,
    destination_region: opts.destRegion || "",
    origin_region: "广州仓/深圳仓",
    origin_cities: DEFAULT_CITIES,
    billing_type: opts.billingType,
    tax_mode: opts.billingType,
    min_quantity: opts.minQty,
    min_quantity_value: opts.minQtyValue,
    unit_price: opts.price,
    price_unit: "元/KG",
    transit_time_min: opts.transitMin,
    transit_time_max: opts.transitMax,
    transit_time_desc: opts.transitDesc,
    claim_rule: opts.claimRule || "",
    effective_date: "",
    source_sheet: opts.sourceSheet,
  };
}

// ── 主入口 ──
function parseHaohui(filePath) {
  console.log("[皓辉] 开始解析:", filePath);
  const wb = XLSX.readFile(filePath);
  const allResults = [];

  // 1. FBA渠道汇总 (主要数据源)
  if (wb.SheetNames.includes("FBA渠道汇总")) {
    const results = parseChannelSummary(wb.Sheets["FBA渠道汇总"]);
    console.log(`  [FBA渠道汇总] ${results.length} 条`);
    allResults.push(...results);
  }

  // 2. 海派服务
  if (wb.SheetNames.includes("海派服务")) {
    const results = parseHaiPai(wb.Sheets["海派服务"]);
    console.log(`  [海派服务] ${results.length} 条`);
    allResults.push(...results);
  }

  // 3. FBA海卡服务美西 (补充数据 — 包含更多重量段及带托架渠道)
  if (wb.SheetNames.includes("FBA海卡服务美西")) {
    const results = parseFBAWestDetailed(wb.Sheets["FBA海卡服务美西"]);
    console.log(`  [FBA海卡美西/详细] ${results.length} 条`);
    allResults.push(...results);
  }

  // 4. FBA海卡服务美东 (补充数据 — 包含纽约皓速达EXX等)
  if (wb.SheetNames.includes("FBA海卡服务美东（纽约拆柜）")) {
    const results = parseFBAEastDetailed(wb.Sheets["FBA海卡服务美东（纽约拆柜）"]);
    console.log(`  [FBA海卡美东/详细] ${results.length} 条`);
    allResults.push(...results);
  }

  console.log(`[皓辉] 总计 ${allResults.length} 条`);
  return allResults;
}

// ── 解析 FBA海卡服务美西 (详细版，含更多重量段) ──
function parseFBAWestDetailed(ws) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 9) return [];

  const results = [];
  // R9+ = data rows (index 8+)
  for (let i = 8; i < data.length; i++) {
    const row = data[i];
    const regionCol = String(row[0] || "").trim();
    const whCell = String(row[1] || "").trim();

    if (!whCell || whCell.length < 3) continue;
    // 跳过非仓库行
    if (whCell.includes("分区") || whCell.includes("渠道") || whCell.includes("仓库")) continue;

    const warehouses = whCell.split(/[,，、]/).map(s => s.trim()).filter(s => s.length >= 3);
    const destRegion = regionCellToRegion(regionCol);

    // 渠道1: 皓快达 (cols 2-3: 12KG+, 51KG+, col 4: transit)
    parseChannelBlock(row, 2, "皓快达", warehouses, destRegion, results);

    // 渠道2: 皓速达-带托架 (cols 5-6: 12KG+, 51KG+, cols 7-8: transit/claim)
    parseChannelBlock(row, 5, "皓速达带托", warehouses, destRegion, results);

    // 渠道3: 皓速达 (cols 9-10: 12KG+, 51KG+, cols 11-12: transit/claim)
    parseChannelBlock(row, 9, "皓速达", warehouses, destRegion, results);

    // 渠道4: 皓森达 (col 13: 12KG+)
    parseChannelBlockSimple(row, 13, "皓森达", warehouses, destRegion, results);
  }

  return results;
}

function regionCellToRegion(cell) {
  if (!cell) return "";
  if (cell.includes("一区")) return "美西一区";
  if (cell.includes("二区")) return "美西二区";
  if (cell.includes("三区")) return "美西三区";
  return "美西";
}

function parseChannelBlock(row, startCol, channelName, warehouses, destRegion, results) {
  const price12 = parseFloat(row[startCol]);
  const price51 = parseFloat(row[startCol + 1]);

  if (!isNaN(price12) && price12 > 0) {
    for (const wh of warehouses) {
      results.push(makeRecord({
        channelName,
        vesselConfig: (CHANNEL_VESSEL_MAP[channelName] || inferVessel(channelName)).config,
        vesselTags: (CHANNEL_VESSEL_MAP[channelName] || inferVessel(channelName)).tags,
        deliveryMethod: "卡派",
        destCode: wh,
        destRegion,
        billingType: "包税",
        minQty: "12KG+",
        minQtyValue: 12,
        price: price12,
        sourceSheet: "FBA海卡服务美西",
      }));
    }
  }
  if (!isNaN(price51) && price51 > 0) {
    for (const wh of warehouses) {
      results.push(makeRecord({
        channelName,
        vesselConfig: (CHANNEL_VESSEL_MAP[channelName] || inferVessel(channelName)).config,
        vesselTags: (CHANNEL_VESSEL_MAP[channelName] || inferVessel(channelName)).tags,
        deliveryMethod: "卡派",
        destCode: wh,
        destRegion,
        billingType: "包税",
        minQty: "51KG+",
        minQtyValue: 51,
        price: price51,
        sourceSheet: "FBA海卡服务美西",
      }));
    }
  }
}

function parseChannelBlockSimple(row, col, channelName, warehouses, destRegion, results) {
  const price = parseFloat(row[col]);
  if (!isNaN(price) && price > 0) {
    for (const wh of warehouses) {
      results.push(makeRecord({
        channelName,
        vesselConfig: (CHANNEL_VESSEL_MAP[channelName] || inferVessel(channelName)).config,
        vesselTags: (CHANNEL_VESSEL_MAP[channelName] || inferVessel(channelName)).tags,
        deliveryMethod: "卡派",
        destCode: wh,
        destRegion,
        billingType: "包税",
        minQty: "12KG+",
        minQtyValue: 12,
        price,
        sourceSheet: "FBA海卡服务美西",
      }));
    }
  }
}

// ── 解析 FBA海卡服务美东 (详细版) ──
function parseFBAEastDetailed(ws) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 10) return [];

  const results = [];
  let currentChannel = "";

  for (let i = 7; i < data.length; i++) {
    const row = data[i];
    const col0 = String(row[0] || "").trim();
    const col1 = String(row[1] || "").trim();
    const whCell = String(row[2] || "").trim();

    // 检测渠道标题行
    if (col0 === "分区" || (col1 && col1.includes("FBA仓库代码"))) {
      const titleRow = row;
      for (let c = 0; c < titleRow.length; c++) {
        const cell = String(titleRow[c] || "");
        if (cell.includes("皓速达") || cell.includes("皓东达") || cell.includes("纽约")) {
          currentChannel = cell.replace(/\r?\n.*/s, "").trim();
          break;
        }
      }
      continue;
    }

    // 确定渠道: 前半部分=纽约皓速达, 后半部分=皓东达
    if (whCell === "分区" && col1 === "FBA仓库代码") {
      const titleRow = row;
      for (let c = 0; c < titleRow.length; c++) {
        const cell = String(titleRow[c] || "");
        if (cell.includes("纽约皓速达") || cell.includes("皓东达")) {
          if (cell.includes("纽约")) {
            currentChannel = "纽约皓速达";
          } else {
            currentChannel = "皓东达";
          }
        }
      }
    }

    if (!whCell || whCell.length < 3) continue;
    if (whCell.includes("分区") || whCell.includes("仓库代码") || whCell.includes("TikTok")) continue;

    const warehouses = whCell.split(/[,，]/).map(s => s.trim()).filter(s => s.length >= 3);
    const regionCol = String(row[1] || "").trim();

    // 推断渠道 (根据是否有价格数据)
    // 纽约皓速达: price columns at 4,5,6 (12KG+,51KG+,101KG+), col 7=transit
    const price301_1 = parseFloat(row[7]);
    const price12_1 = parseFloat(row[4]);
    if (!isNaN(price12_1) && price12_1 > 0) {
      for (const wh of warehouses) {
        parseWeightTiersEast(row, 4, "纽约皓速达", wh, results);
      }
    }

    // Check second channel block later in the row
    // For 皓东达, columns are shifted right (after col 8)
    const price12_2 = parseFloat(row[10]); // Assuming 皓东达 starts around col 9-10
    if (!isNaN(price12_2) && price12_2 > 0) {
      for (const wh of warehouses) {
        parseWeightTiersEast(row, 10, "皓东达", wh, results);
      }
    }
  }

  return results;
}

function parseWeightTiersEast(row, startCol, channelName, wh, results) {
  const tiers = [
    { qty: "12KG+", val: 12, col: startCol },
    { qty: "51KG+", val: 51, col: startCol + 1 },
    { qty: "101KG+", val: 101, col: startCol + 2 },
    { qty: "301KG+", val: 301, col: startCol + 3 },
  ];

  const vessel = CHANNEL_VESSEL_MAP[channelName] || inferVessel(channelName);

  for (const tier of tiers) {
    const price = parseFloat(row[tier.col]);
    if (!isNaN(price) && price > 0) {
      results.push(makeRecord({
        channelName,
        vesselConfig: vessel.config,
        vesselTags: vessel.tags,
        deliveryMethod: "卡派",
        destCode: wh,
        billingType: "包税",
        minQty: tier.qty,
        minQtyValue: tier.val,
        price,
        sourceSheet: "FBA海卡服务美东",
      }));
    }
  }
}

module.exports = { parseHaohui };
