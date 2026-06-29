/**
 * 心一供应链 — 欧洲线价格解析器
 *
 * 结构特点：
 * - 按国家行（非仓行），国家组以逗号分隔需拆分
 * - 每Sheet含多个子表（自税/含税变体）
 * - 运输方式: 空运(多个Sheet) + 海运(多个Sheet) + 中欧专车 + 铁路(铁卡)
 * - 空运重量段: 23/46/100 或 21/45/100/500
 * - 海运重量段: 25/71/100/1000
 * - 铁卡/专车重量段: 23/46/100 或 25/71/100
 */

const XLSX = require("xlsx");

const SUPPLIER = "心一供应链";
const COUNTRY = "欧洲";

// ── 国家组拆分 ──
const COUNTRY_NAMES = ["德国", "法国", "意大利", "西班牙", "波兰", "捷克", "荷兰", "奥地利",
  "比利时", "卢森堡", "丹麦", "瑞典", "芬兰", "匈牙利", "希腊", "葡萄牙", "爱尔兰",
  "罗马尼亚", "保加利亚", "克罗地亚", "斯洛文尼亚", "斯洛伐克", "立陶宛", "拉脱维亚", "爱沙尼亚"];

function parseCountries(cell) {
  const text = String(cell).replace(/\r?\n/g, " ").trim();
  for (const cn of COUNTRY_NAMES) {
    if (text === cn) return [cn];
  }
  // Split by comma/、
  const parts = text.split(/[,，、]/).map(s => s.trim()).filter(Boolean);
  const result = [];
  for (const p of parts) {
    for (const cn of COUNTRY_NAMES) {
      if (p.includes(cn)) { result.push(cn); break; }
    }
  }
  return result.length > 0 ? result : [text]; // fallback
}

// ── 仓库代码提取 ──
function parseWarehousesFromCell(cell) {
  const text = String(cell).replace(/\r?\n/g, " ").trim();
  const matches = text.match(/[A-Z]{2,}\d[A-Z\d]*/g);
  if (matches) return matches;
  // Postal code style: 44145, 38350 etc
  const nums = text.match(/\d{4,6}/g);
  if (nums) return nums;
  return [text];
}

// ── 时效解析 ──
function parseTransit(text) {
  const cleaned = String(text || "").replace(/\r?\n/g, " ").trim();
  const match = cleaned.match(/(\d+)\s*[-–~]\s*(\d+)\s*(?:个)?(?:自然日|天|工作日)/);
  if (match) return { min: parseInt(match[1]), max: parseInt(match[2]), desc: cleaned };
  const single = cleaned.match(/(\d+)\s*(?:个)?(?:自然日|天|工作日)/);
  if (single) return { min: parseInt(single[1]), max: parseInt(single[1]), desc: cleaned };
  return { min: null, max: null, desc: cleaned };
}

// ── 通用记录 ──
function makeRecord(opts) {
  return {
    supplier: SUPPLIER,
    country: COUNTRY,
    channel_name: opts.channelName || "",
    transport_mode: opts.transportMode || "海运",
    vessel_config: opts.vesselConfig || "",
    vessel_tags: opts.vesselTags || [],
    delivery_method: opts.deliveryMethod || "快递派",
    destination_type: opts.destType || "country",
    destination_code: opts.destCode || "",
    destination_region: opts.destRegion || "",
    origin_region: "华南",
    origin_cities: ["深圳", "东莞", "广州", "中山", "惠州"],
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

// ── 通用：解析国家行Sheet（含多个子表） ──
function parseCountryBasedSheet(ws, sheetName, transportMode, subTableConfigs) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 4) return [];

  const results = [];
  let currentConfig = null;

  for (let ri = 0; ri < data.length; ri++) {
    const row = data[ri];
    const col0 = String(row[0] || "").replace(/\r?\n/g, " ").trim();
    const col1 = String(row[1] || "").replace(/\r?\n/g, " ").trim();

    // 检测子表头（渠道名行）
    for (const cfg of subTableConfigs) {
      if (col0.includes(cfg.keyword) || col1.includes(cfg.keyword)) {
        currentConfig = cfg;
        break;
      }
    }

    // 跳过表头/空行/说明行
    if (!col1 || !currentConfig) continue;
    if (col1 === "国家" || col1 === "服务" || col1 === "渠道") continue;
    if (col0.includes("拒收") || col0.includes("备注") || col0.includes("特别提醒") || col0.includes("自用VAT")) continue;
    if (col1.includes("拒收") || col1.includes("备注")) continue;

    // 检测子表内部的新渠道（某些sheet的渠道在col0）
    for (const cfg of subTableConfigs) {
      if (col0.includes(cfg.keyword)) {
        currentConfig = cfg;
        continue; // skip this row, it's a header
      }
    }
    // If col0 is a channel keyword and col1 is empty or "国家", skip
    if (currentConfig && (col0.includes(currentConfig.keyword) && (!col1 || col1 === "国家"))) continue;

    // Parse data row
    const isWarehouseRow = col1.match(/[A-Z]{2,}\d/) || col1.match(/^\d{4,6}/);
    let destinations;

    if (isWarehouseRow) {
      destinations = parseWarehousesFromCell(col1);
    } else {
      destinations = parseCountries(col1);
    }

    if (destinations.length === 0) continue;
    if (destinations[0] === "国家" || destinations[0].length < 2) continue;

    const weightTiers = currentConfig.weightTiers;
    const transitText = String(row[currentConfig.transitCol] || "").trim();
    const transit = parseTransit(transitText);
    const claimText = String(row[currentConfig.claimCol || currentConfig.transitCol] || "").replace(/\r?\n/g, " ").trim();

    for (const dest of destinations) {
      for (const tier of weightTiers) {
        const price = parseFloat(row[tier.col]);
        if (!isNaN(price) && price > 0) {
          results.push(makeRecord({
            channelName: currentConfig.channelName,
            transportMode: currentConfig.transportMode || transportMode,
            vesselConfig: currentConfig.vesselConfig || "",
            vesselTags: currentConfig.vesselTags || [],
            deliveryMethod: currentConfig.deliveryMethod || (isWarehouseRow ? "卡派" : "快递派"),
            destCode: dest,
            destType: isWarehouseRow ? "warehouse" : "country",
            destRegion: dest,
            billingType: currentConfig.taxMode || "含税",
            taxMode: currentConfig.taxMode || "含税",
            minQty: tier.qty,
            minQtyValue: tier.val,
            price,
            transitMin: transit.min,
            transitMax: transit.max,
            transitDesc: transit.desc || transitText,
            claimRule: claimText,
            sourceSheet: sheetName,
          }));
        }
      }
    }
  }
  return results;
}

// ══════════════════════════════════════════════════
// Sheet 配置
// ══════════════════════════════════════════════════

// 欧洲空运普货自税、含税DPD
function parseAirDPD(ws) {
  return parseCountryBasedSheet(ws, "欧洲空运普货自税含税DPD", "空运", [
    { keyword: "普货含税", channelName: "欧洲空运普货-含税DPD", transportMode: "空运", taxMode: "含税",
      vesselConfig: "空运普货", vesselTags: ["空运", "DPD"],
      deliveryMethod: "快递派", transitCol: 5, weightTiers: [
        { col: 2, qty: "23KG+", val: 23 }, { col: 3, qty: "46KG+", val: 46 }, { col: 4, qty: "100KG+", val: 100 },
    ]},
    { keyword: "普货自税", channelName: "欧洲空运普货-自税DPD", transportMode: "空运", taxMode: "自税",
      vesselConfig: "空运普货", vesselTags: ["空运", "DPD"],
      deliveryMethod: "快递派", transitCol: 5, weightTiers: [
        { col: 2, qty: "23KG+", val: 23 }, { col: 3, qty: "46KG+", val: 46 }, { col: 4, qty: "100KG+", val: 100 },
    ]},
  ]);
}

// 欧洲空运-普货
function parseAirStandard(ws) {
  return parseCountryBasedSheet(ws, "欧洲空运-普货", "空运", [
    { keyword: "普货自税", channelName: "欧洲空运-普货自税", transportMode: "空运", taxMode: "自税",
      vesselConfig: "空运普货", vesselTags: ["空运"], deliveryMethod: "快递派", transitCol: 5, weightTiers: [
        { col: 2, qty: "21KG+", val: 21 }, { col: 3, qty: "45KG+", val: 45 },
        { col: 4, qty: "100KG+", val: 100 }, { col: 5, qty: "500KG+", val: 500 },
    ]},
    { keyword: "普货含税", channelName: "欧洲空运-普货含税", transportMode: "空运", taxMode: "含税",
      vesselConfig: "空运普货", vesselTags: ["空运"], deliveryMethod: "快递派", transitCol: 5, weightTiers: [
        { col: 2, qty: "21KG+", val: 21 }, { col: 3, qty: "45KG+", val: 45 },
        { col: 4, qty: "100KG+", val: 100 }, { col: 5, qty: "500KG+", val: 500 },
    ]},
  ]);
}

// 欧洲空运-快线八日提 (简化：支持含税/自税)
function parseAirExpress(ws, sheetName, channelLabel, transitDays) {
  return parseCountryBasedSheet(ws, sheetName, "空运", [
    { keyword: "含税", channelName: `欧洲空运${channelLabel}-含税`, transportMode: "空运", taxMode: "含税",
      vesselConfig: `空运${channelLabel}`, vesselTags: ["空运", "快线"], deliveryMethod: "快递派", transitCol: 5, weightTiers: [
        { col: 2, qty: "23KG+", val: 23 }, { col: 3, qty: "46KG+", val: 46 }, { col: 4, qty: "100KG+", val: 100 },
    ]},
    { keyword: "自税", channelName: `欧洲空运${channelLabel}-自税`, transportMode: "空运", taxMode: "自税",
      vesselConfig: `空运${channelLabel}`, vesselTags: ["空运", "快线"], deliveryMethod: "快递派", transitCol: 5, weightTiers: [
        { col: 2, qty: "23KG+", val: 23 }, { col: 3, qty: "46KG+", val: 46 }, { col: 4, qty: "100KG+", val: 100 },
    ]},
  ]);
}

// 欧洲海运含税（国家行DPD）
function parseSeaDPD(ws, sheetName, channelLabel, transitDesc) {
  return parseCountryBasedSheet(ws, sheetName, "海运", [
    { keyword: "德国", channelName: `欧洲海运${channelLabel}-含税DPD`, transportMode: "海运", taxMode: "含税",
      vesselConfig: "海运", vesselTags: ["海运", "DPD"], deliveryMethod: "快递派", transitCol: 5, weightTiers: [
        { col: 1, qty: "25KG+", val: 25 }, { col: 2, qty: "71KG+", val: 71 },
        { col: 3, qty: "100KG+", val: 100 }, { col: 4, qty: "1000KG+", val: 1000 },
    ]},
  ]);
}

// 欧洲海运45日达-含税（仓行+国家行混合）
function parseSea45Tax(ws) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 4) return [];

  const results = [];
  let currentChannel = "";

  for (let ri = 3; ri < data.length; ri++) {
    const row = data[ri];
    const col0 = String(row[0] || "").replace(/\r?\n/g, " ").trim();
    const col1 = String(row[1] || "").replace(/\r?\n/g, " ").trim();
    const col2 = String(row[2] || "").replace(/\r?\n/g, " ").trim();

    // Detect channel
    if (col0.includes("海卡直送")) { currentChannel = "直送"; continue; }
    if (col0.includes("快递派") || col0.includes("DPD")) { currentChannel = "快递派"; continue; }
    if (col0.includes("海派")) { currentChannel = "快递派"; continue; }

    if (!col1 && !col2) continue;
    if (col0.includes("拒收") || col0.includes("备注") || col0.includes("特别提醒")) continue;

    // Warehouse rows (have warehouse code in col2)
    if (col2 && col2.match(/[A-Z]{2,}\d/)) {
      const wh = col2.trim();
      const country = col1.replace("-DE", "").replace("德国", "德国").trim() || "德国";
      for (let t = 0; t < 3; t++) {
        const price = parseFloat(row[3 + t]);
        const labels = ["25KG+", "71KG+", "100KG+"];
        const vals = [25, 71, 100];
        if (!isNaN(price) && price > 0) {
          results.push(makeRecord({
            channelName: `欧洲海运45日达-${currentChannel}-含税`,
            transportMode: "海运", vesselConfig: "OA联盟快船", vesselTags: ["OA", "COSCO", "EMC", "OOCL", "CMA", "海运"],
            deliveryMethod: currentChannel === "直送" ? "直送" : "卡派",
            destCode: wh, destType: "warehouse", destRegion: country,
            billingType: "含税", taxMode: "含税",
            minQty: labels[t], minQtyValue: vals[t], price,
            transitMin: 30, transitMax: 45, transitDesc: "开船后30-45个自然日交仓",
            sourceSheet: "欧洲海运45日达-含税",
          }));
        }
      }
    } else if (col1 && parseCountries(col1).length > 0 && !col2) {
      // Country rows
      const countries = parseCountries(col1);
      for (const cn of countries) {
        for (let t = 0; t < 3; t++) {
          const price = parseFloat(row[3 + t]);
          const labels = ["25KG+", "71KG+", "100KG+"];
          const vals = [25, 71, 100];
          if (!isNaN(price) && price > 0) {
            results.push(makeRecord({
              channelName: `欧洲海运45日达-${currentChannel}-含税`,
              transportMode: "海运", vesselConfig: "OA联盟快船", vesselTags: ["OA", "COSCO", "EMC", "OOCL", "CMA", "海运"],
              deliveryMethod: "快递派", destCode: cn, destType: "country", destRegion: cn,
              billingType: "含税", taxMode: "含税",
              minQty: labels[t], minQtyValue: vals[t], price,
              transitMin: 30, transitMax: 45, transitDesc: "开船后30-45个自然日交仓",
              sourceSheet: "欧洲海运45日达-含税",
            }));
          }
        }
      }
    }
  }
  return results;
}

// 欧洲铁卡
function parseRailTruck(ws) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 4) return [];

  const results = [];
  let currentChannel = "";

  for (let ri = 3; ri < data.length; ri++) {
    const row = data[ri];
    const col0 = String(row[0] || "").replace(/\r?\n/g, " ").trim();
    const col1 = String(row[1] || "").replace(/\r?\n/g, " ").trim();

    if (col0.includes("铁卡直送")) { currentChannel = "铁卡直送-含税"; continue; }
    if (col0.includes("铁卡-含税")) { currentChannel = "铁卡-含税"; continue; }
    if (col0.includes("铁卡-卡派")) { currentChannel = "铁卡-卡派"; continue; }
    if (col0.includes("拒收") || col0.includes("备注") || col0.includes("渠道路线")) continue;
    if (col1 === "国家" || col1 === "服务") continue;

    if (!currentChannel) continue;

    const isWarehouseRow = col1.match(/[A-Z]{2,}\d/) || col1.match(/^\d{4,6}/);
    let destinations;
    if (isWarehouseRow) {
      destinations = parseWarehousesFromCell(col1);
    } else {
      destinations = parseCountries(col1);
    }
    if (destinations.length === 0) continue;

    const transit = parseTransit(String(row[5] || ""));

    for (const dest of destinations) {
      const tiers = [
        { col: 2, qty: "25KG+", val: 25 }, { col: 3, qty: "71KG+", val: 71 }, { col: 4, qty: "100KG+", val: 100 },
      ];
      for (const tier of tiers) {
        const price = parseFloat(row[tier.col]);
        if (!isNaN(price) && price > 0) {
          results.push(makeRecord({
            channelName: currentChannel, transportMode: "铁路",
            vesselConfig: "中欧铁卡", vesselTags: ["铁路", "中欧班列"],
            deliveryMethod: currentChannel.includes("直送") ? "直送" : (isWarehouseRow ? "卡派" : "快递派"),
            destCode: dest, destType: isWarehouseRow ? "warehouse" : "country", destRegion: dest,
            billingType: "含税", taxMode: "含税",
            minQty: tier.qty, minQtyValue: tier.val, price,
            transitMin: transit.min, transitMax: transit.max,
            transitDesc: transit.desc || "25-28个自然日",
            sourceSheet: "欧洲铁卡",
          }));
        }
      }
    }
  }
  return results;
}

// 中欧专车
function parseTruckExpress(ws) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 4) return [];

  const results = [];
  let currentChannel = "";

  for (let ri = 3; ri < data.length; ri++) {
    const row = data[ri];
    const col0 = String(row[0] || "").replace(/\r?\n/g, " ").trim();
    const col1 = String(row[1] || "").replace(/\r?\n/g, " ").trim();
    const col2 = String(row[2] || "").replace(/\r?\n/g, " ").trim();

    if (col0.includes("自税(卡派)") || col0.includes("自税（卡派）")) { currentChannel = "中欧专车-自税-卡派直送"; continue; }
    if (col0.includes("自税") && !col0.includes("(")) { currentChannel = "中欧专车-自税-快递派"; continue; }
    if (col0.includes("含税(卡派)") || col0.includes("含税（卡派）")) { currentChannel = "中欧专车-含税-卡派直送"; continue; }
    if (col0.includes("含税") && !col0.includes("(")) { currentChannel = "中欧专车-含税-快递派"; continue; }
    if (col0.includes("拒收") || col0.includes("备注")) continue;
    if (col1 === "国家" || col1 === "渠道") continue;

    if (!currentChannel) continue;
    if (!col1) continue;

    const deliveryMethod = currentChannel.includes("直送") ? "直送" : (col2.includes("快递") ? "快递派" : "卡派");
    const isWarehouseRow = col1.match(/[A-Z]{2,}\d/) || col1.match(/^\d{4,6}/);
    let destinations;
    if (isWarehouseRow) {
      destinations = parseWarehousesFromCell(col1);
    } else {
      destinations = parseCountries(col1);
    }
    if (destinations.length === 0) continue;

    const transitText = String(row[6] || "");
    const transit = parseTransit(transitText);

    for (const dest of destinations) {
      const tiers = [
        { col: 3, qty: "23KG+", val: 23 }, { col: 4, qty: "46KG+", val: 46 }, { col: 5, qty: "100KG+", val: 100 },
      ];
      for (const tier of tiers) {
        const price = parseFloat(row[tier.col]);
        if (!isNaN(price) && price > 0) {
          results.push(makeRecord({
            channelName: currentChannel, transportMode: "中欧专车",
            vesselConfig: "中欧专车", vesselTags: ["中欧专车"],
            deliveryMethod, destCode: dest,
            destType: isWarehouseRow ? "warehouse" : "country", destRegion: dest,
            billingType: currentChannel.includes("自税") ? "自税" : "含税",
            taxMode: currentChannel.includes("自税") ? "自税" : "含税",
            minQty: tier.qty, minQtyValue: tier.val, price,
            transitMin: transit.min, transitMax: transit.max,
            transitDesc: transit.desc || "22-28个自然日",
            sourceSheet: "中欧专车",
          }));
        }
      }
    }
  }
  return results;
}

// ══════════════════════════════════════════════════
// 主入口
// ══════════════════════════════════════════════════

function parseXinyi(filePath) {
  console.log("[心一] 开始解析:", filePath);
  const wb = XLSX.readFile(filePath);
  const allResults = [];

  const sheetParsers = {
    "欧洲空运普货自税、含税DPD": parseAirDPD,
    "欧洲空运快线八日提-普货": (ws) => parseAirExpress(ws, "欧洲空运快线八日提-普货", "快线八日提-普货"),
    "欧洲空运快线十日提-普货": (ws) => parseAirExpress(ws, "欧洲空运快线十日提-普货", "快线十日提-普货"),
    "欧洲空运-普货": parseAirStandard,
    "欧洲空运-带电": (ws) => parseCountryBasedSheet(ws, "欧洲空运-带电", "空运", [
      { keyword: "自税", channelName: "欧洲空运-带电自税", transportMode: "空运", taxMode: "自税",
        vesselConfig: "空运带电", vesselTags: ["空运", "带电"], deliveryMethod: "快递派", transitCol: 5, weightTiers: [
          { col: 2, qty: "21KG+", val: 21 }, { col: 3, qty: "45KG+", val: 45 }, { col: 4, qty: "100KG+", val: 100 }, { col: 5, qty: "500KG+", val: 500 },
      ]},
      { keyword: "含税", channelName: "欧洲空运-带电含税", transportMode: "空运", taxMode: "含税",
        vesselConfig: "空运带电", vesselTags: ["空运", "带电"], deliveryMethod: "快递派", transitCol: 5, weightTiers: [
          { col: 2, qty: "21KG+", val: 21 }, { col: 3, qty: "45KG+", val: 45 }, { col: 4, qty: "100KG+", val: 100 }, { col: 5, qty: "500KG+", val: 500 },
      ]},
    ]),
    "欧洲空运带电自税，含税DPD": (ws) => parseAirExpress(ws, "欧洲空运带电含税DPD", "带电DPD"),
    "欧洲空运快线十日提-带电": (ws) => parseAirExpress(ws, "欧洲空运快线十日提-带电", "快线十日提-带电"),
    "欧洲海运含税": (ws) => parseSeaDPD(ws, "欧洲海运含税", "含税DPD"),
    "欧洲海运快速达含税": (ws) => parseSeaDPD(ws, "欧洲海运快速达含税", "快速达含税DPD"),
    "欧洲海运限时达含税": (ws) => parseSeaDPD(ws, "欧洲海运限时达含税", "限时达含税DPD"),
    "欧洲海运含税比雷": (ws) => parseSeaDPD(ws, "欧洲海运含税比雷", "含税比雷DPD"),
    "欧洲海运直送": (ws) => parseCountryBasedSheet(ws, "欧洲海运直送", "海运", [
      { keyword: "含税", channelName: "欧洲海运直送-含税", transportMode: "海运", taxMode: "含税",
        vesselConfig: "海运直送", vesselTags: ["海运", "直送"], deliveryMethod: "直送", transitCol: 5, weightTiers: [
          { col: 2, qty: "25KG+", val: 25 }, { col: 3, qty: "71KG+", val: 71 }, { col: 4, qty: "100KG+", val: 100 },
      ]},
    ]),
    "欧洲海运45日达-含税": parseSea45Tax,
    "欧洲铁卡": parseRailTruck,
    "中欧专车": parseTruckExpress,
    "中欧专车25日达": parseTruckExpress,
  };

  for (const [sheetName, parser] of Object.entries(sheetParsers)) {
    const actualName = wb.SheetNames.find(n => n.trim() === sheetName.trim());
    if (actualName) {
      try {
        const results = parser(wb.Sheets[actualName]);
        console.log(`  [${sheetName}] ${results.length} 条`);
        allResults.push(...results);
      } catch (err) {
        console.error(`  [${sheetName}] 解析失败: ${err.message}`);
      }
    }
  }

  console.log(`[心一] 总计 ${allResults.length} 条`);
  return allResults;
}

module.exports = { parseXinyi };
