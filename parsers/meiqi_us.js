/**
 * 美琦国际 — 美/加/墨线价格解析器
 *
 * Sheet 结构：仓库×城市组×单价(KG/CBM)×渠道
 * 常见列布局: col0=区域标签, col1=仓库列表, col2+=KG价/CBM价交替
 */

const XLSX = require("xlsx");
const { detectCountry } = require("./country-detector");

const SUPPLIER = "美琦国际";
const COUNTRY = "美国";
const DEFAULT_CITIES = ["深圳", "中山", "广州", "东莞", "惠州", "汕头", "厦门", "义乌", "上海", "苏州", "宁波", "杭州", "青岛"];

// ── 区域识别 ──
function identifyRegion(wh) {
  const code = wh.toUpperCase().trim();
  const west = ["ONT8","LAX9","LGB8","SBD1","SBD2","SBD3","SNA4","LGB4","LGB6","LGB7","LGB9",
    "LAS1","VGT2","LAS6","GYR2","GYR3","GEU2","GEU3","MIT2","SMF3","SCK1","SCK3","SCK4","SCK8",
    "OAK3","SJC7","SMF6","QXY5","QZZ7","FAT2","MCC1","MCE1","HLI2","TCY1","TCY2","PSP3",
    "POC1","POC2","POC3","IUSJ","IUSP","IUSQ","IUSW","IUS1","IUT1","IUSI","IUTE",
    "ABQ2","PSC2","PHX5","PHX7","AZA4","RNO4","BFI3","PDX7","GEG2","SLC2","DEN2","DEN8"];
  if (west.some(w => code === w || code.includes(w))) return "美西";

  const central = ["MDW2","IND9","MQJ1","MEM1","FWA4","FTW1","IAH3","HOU8","SAT1","SAT4",
    "OKC2","MCI1","MKC4","MKC6","STL3","STL4","STL6","ICT2","DFW6","DFW7","FTW2","FTW3","FTW5",
    "DAL3","HOU1","HOU2","HOU7","SAT2","SAT3","AUS2","AUS3","ORD2","ORD6","IGQ2","JVL1","IND2",
    "IND3","IND4","IND5","RFD2","FWA4","DET1","DET2","CMH2","CMH3","CMH4","CVG2","CVG3","CVG5",
    "SDF2","SDF4","SDF6","SDF8","LEX1","LEX2","LEX3","MEM2","MEM3","MEM4","MEM5","MEM6",
    "BNA2","BNA3","BNA5","BNA6","CHA1","CHA2","HSV1","HSV2","ATL2","ATL3","ATL6","ATL7",
    "ATL8","MGE1","MGE3","MGE5","SAV3","GSP1","CLT2","CLT3","CLT4","CLT6","CAE1","CAE2"];
  if (central.some(w => code === w || code.includes(w))) return "美中";

  const east = ["TEB3","TEB4","TEB6","TEB9","ABE2","ABE3","ABE4","ABE5","ABE8","EWR4","EWR5",
    "EWR6","EWR7","EWR9","LGA9","JFK2","JFK8","ACY1","ACY2","ACY3","TTN2","PHL1","PHL3",
    "PHL4","PHL5","PHL6","PHL7","PHL8","PHL9","AVP1","AVP3","AVP9","MDT1","MDT2","MDT4",
    "XEW5","SWF1","SWF2","BDL2","BDL3","BDL6","ALB1","BOS7","BWI1","BWI2","BWI4","BWI5",
    "DCA1","DCA6","IAD1","RIC1","RIC2","RIC3","ORF2","ORF3","CHO1","RMN3","XRI3","ILG1",
    "PIT2","LBE1","HGR2","HGR6","WBW2","HEA2","RYY2","TPA1","TPA2","TPA3","TPA6",
    "MCO1","MCO2","MIA1","MIA4","MIA5","MIA8","FLL2","PBI1","PBI2","PBI3","JAX2","JAX3",
    "JAX5","JAX7","TMB8","TMB3","RDG1","HIA1","RDU1","RDU2","RDU4","GSO1","XLX1","XLX6",
    "XLX7","TOL1","TOL3","DTW3","MQY1","AKR1","IUSL","IUST","IUSR","XPH1","CMH2","CMH3"];
  if (east.some(w => code === w || code.includes(w))) return "美东";

  return "美西";
}

function parseWarehouses(cell) {
  const text = String(cell).replace(/\r?\n/g, " ").replace(/\s+/g, " ");
  return text.split(/[\/,，、;；]/).map(s => s.trim()).filter(s => s.length >= 3 && s.match(/[A-Z0-9]/));
}

// ── 通用记录生成 ──
function makeRecord(opts) {
  return {
    supplier: SUPPLIER,
    country: opts.country || COUNTRY,
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
    price_unit: opts.priceUnit || "元/KG",
    cbm_price: opts.cbmPrice || null,
    transit_time_min: opts.transitMin || null,
    transit_time_max: opts.transitMax || null,
    transit_time_desc: opts.transitDesc || "",
    claim_rule: opts.claimRule || "",
    effective_date: "",
    source_file: "",
    source_sheet: opts.sourceSheet || "",
  };
}

/**
 * 通用仓库×价格矩阵解析器
 * 适用于美琦大部分 Sheet（col1=仓库, col2+=KG价/CBM价 交替）
 * cityColumns: [{label, kgCol, cbmCol}] 每城市组占2列（KG + CBM）
 */
function parseWarehousePriceMatrix(ws, config) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < config.dataStartRow) return [];
  const results = [];

  for (let ri = config.dataStartRow - 1; ri < data.length; ri++) {
    const row = data[ri];
    const whCell = String(row[config.warehouseCol] || "").trim();
    if (!whCell || whCell.length < 3) continue;

    // Skip header / info rows
    const skipKeywords = ["仓库地址", "下单渠道", "热门仓", "外州", "正常派送", "海外仓", "加州", "备注", "产品说明",
      "时效赔付", "上架延误", "渠道说明", "返回目录", "亚马逊仓库", "以下", "不足"];
    if (skipKeywords.some(k => whCell.includes(k)) && !whCell.match(/[A-Z]{2,}\d/)) continue;

    const warehouses = parseWarehouses(whCell);
    if (warehouses.length === 0) continue;

    for (const cityGroup of config.cityGroups) {
      for (const tier of config.weightTiers) {
        const kgCol = cityGroup.kgStartCol + tier.colOffset;
        const kgPrice = parseFloat(row[kgCol]);
        if (!isNaN(kgPrice) && kgPrice > 0 && kgPrice < 999) {
          const cbmCol = cityGroup.cbmStartCol + tier.colOffset;
          const cbmPrice = parseFloat(row[cbmCol]);
          for (const wh of warehouses) {
            const region = identifyRegion(wh);
            results.push(makeRecord({
              channelName: config.channelName,
              transportMode: config.transportMode || "海运",
              vesselConfig: config.vesselConfig || "普船",
              vesselTags: config.vesselTags || ["普船"],
              deliveryMethod: config.deliveryMethod || "卡派",
              destCode: wh,
              destRegion: region,
              originRegion: cityGroup.label,
              originCities: cityGroup.cities || DEFAULT_CITIES,
              billingType: tier.taxMode || "包税",
              minQty: tier.label,
              minQtyValue: tier.value,
              price: kgPrice,
              priceUnit: "元/KG",
              cbmPrice: !isNaN(cbmPrice) && cbmPrice > 0 ? cbmPrice : null,
              transitMin: config.transitMin,
              transitMax: config.transitMax,
              transitDesc: config.transitDesc || "",
              sourceSheet: config.sourceSheet || "",
            }));
          }
        }
      }
    }
  }
  return results;
}

/** 第二段（50KG+/1CBM+）继续解析，通常 dataStartRow 之后有 "亚马逊仓库地址（以下地址 不足50KG 需+2）" 行 */
function parseSecondSection(ws, config) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  // Find the second section marker
  let startRow = -1;
  for (let ri = config.dataStartRow; ri < data.length; ri++) {
    const row = data[ri];
    const cell0 = String(row[0] || "").trim();
    const cell1 = String(row[1] || "").trim();
    if ((cell0 + cell1).includes("50KG") || (cell0 + cell1).includes("不足50") || cell0.includes("外州")) {
      startRow = ri + 1; // Next row after the marker
      break;
    }
  }
  if (startRow < 0) return [];

  // Update weight tiers
  const heavyTiers = [{ label: "50KG+", value: 50, colOffset: 0, taxMode: "包税" },
    { label: "1CBM+", value: 1, colOffset: 1, taxMode: "不包税" }];
  // Create a 2-tier config where the original columns are interpreted differently
  const heavyConfig = {
    ...config,
    dataStartRow: startRow,
    weightTiers: heavyTiers,
    sourceSheet: config.sourceSheet + "(重货)",
  };
  return parseWarehousePriceMatrix(ws, heavyConfig);
}

// ═══════════════════════════════════════════════════════════════
// 各 Sheet 解析器
// ═══════════════════════════════════════════════════════════════

/** Match系列 — 美西核心渠道 */
function parseMatchSeries(ws) {
  const results = [];
  // Section 1: "热门仓" rows (21KG+ 包税 / 1CBM+ 不包税), R7+
  const config1 = {
    channelName: "Match12-卡派",
    transportMode: "海运",
    vesselConfig: "美森MATSON(CLX正班/MAX加班)",
    vesselTags: ["美森", "MATSON", "CLX正班", "MAX加班"],
    deliveryMethod: "卡派",
    dataStartRow: 7,
    warehouseCol: 2,
    sourceSheet: "Match系列",
    transitMin: 13, transitMax: 17, transitDesc: "13-17天",
    cityGroups: [
      { label: "华南/厦门", kgStartCol: 3, cbmStartCol: 4, cities: ["深圳","中山","广州","东莞","惠州","汕头","厦门"] },
      { label: "义乌/苏州", kgStartCol: 5, cbmStartCol: 6, cities: ["义乌","上海","苏州","宁波","杭州"] },
      { label: "青岛", kgStartCol: 7, cbmStartCol: 8, cities: ["青岛"] },
    ],
    weightTiers: [
      { label: "21KG+", value: 21, colOffset: 0, taxMode: "包税" },
      { label: "1CBM+", value: 1, colOffset: 1, taxMode: "不包税" },
    ],
  };
  results.push(...parseWarehousePriceMatrix(ws, config1));
  // Section 2: 50KG+/1CBM+ rows
  results.push(...parseSecondSection(ws, config1));

  // Match15 渠道（col11-17: col11=KG_华南, col12=CBM_华南, col13=KG_义乌, col14=CBM_义乌, col15=KG_青岛, col16=CBM_青岛）
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  const match15Config = {
    channelName: "Match15-卡派",
    vesselConfig: "ZIM带车架/EXX统配",
    vesselTags: ["ZIM","EXX"],
    transportMode: "海运",
    deliveryMethod: "卡派",
    sourceSheet: "Match系列",
    cityGroups: [
      { label: "华南/厦门", kgStartCol: 11, cbmStartCol: 12, cities: ["深圳","中山","广州","东莞","惠州","汕头","厦门"] },
      { label: "义乌/苏州", kgStartCol: 13, cbmStartCol: 14, cities: ["义乌","上海","苏州","宁波","杭州"] },
      { label: "青岛", kgStartCol: 15, cbmStartCol: 16, cities: ["青岛"] },
    ],
    weightTiers: [
      { label: "21KG+", value: 21, colOffset: 0, taxMode: "包税" },
      { label: "1CBM+", value: 1, colOffset: 1, taxMode: "不包税" },
    ],
  };
  for (let ri = 6; ri < data.length; ri++) {
    const row = data[ri];
    const whCell = String(row[2] || "").trim();
    if (!whCell || whCell.length < 3) continue;
    const skipKws = ["仓库地址","下单渠道","时效赔付","上架延误","以下","不足"];
    if (skipKws.some(k => whCell.includes(k)) && !whCell.match(/[A-Z]{2,}\d/)) continue;
    const warehouses = parseWarehouses(whCell);
    if (warehouses.length === 0) continue;
    for (const cg of match15Config.cityGroups) {
      for (const tier of match15Config.weightTiers) {
        const price = parseFloat(row[cg.kgStartCol + tier.colOffset]);
        if (!isNaN(price) && price > 0 && price < 999) {
          for (const wh of warehouses) {
            results.push(makeRecord({
              channelName: match15Config.channelName,
              vesselConfig: match15Config.vesselConfig,
              vesselTags: match15Config.vesselTags,
              transportMode: match15Config.transportMode,
              deliveryMethod: match15Config.deliveryMethod,
              destCode: wh, destRegion: identifyRegion(wh),
              originRegion: cg.label, originCities: cg.cities || DEFAULT_CITIES,
              billingType: tier.taxMode, minQty: tier.label, minQtyValue: tier.value,
              price, priceUnit: "元/KG",
              cbmPrice: !isNaN(parseFloat(row[cg.cbmStartCol + tier.colOffset])) ? parseFloat(row[cg.cbmStartCol + tier.colOffset]) : null,
              transitMin: 15, transitMax: 19, transitDesc: "15-19天",
              sourceSheet: "Match系列",
            }));
          }
        }
      }
    }
  }

  return results;
}

/** OA-拆送 */
function parseOAChaiSong(ws) {
  const config = {
    channelName: "OA-拆送专线",
    transportMode: "海运",
    vesselConfig: "OA联盟(COSCO/OOCL/EMC/CMA)",
    vesselTags: ["OA","COSCO","OOCL","EMC","CMA"],
    deliveryMethod: "卡派",
    dataStartRow: 7,
    warehouseCol: 2,
    sourceSheet: "OA-拆送",
    transitMin: 24, transitMax: 32, transitDesc: "24-32天",
    cityGroups: [
      { label: "华南/华东", kgStartCol: 3, cbmStartCol: 4, cities: ["深圳","中山","广州","东莞","惠州","汕头","厦门","义乌","苏州"] },
      { label: "青岛", kgStartCol: 5, cbmStartCol: 6, cities: ["青岛"] },
    ],
    weightTiers: [
      { label: "21KG+", value: 21, colOffset: 0, taxMode: "包税" },
      { label: "1CBM+", value: 1, colOffset: 1, taxMode: "不包税" },
    ],
  };
  const results = parseWarehousePriceMatrix(ws, config);
  results.push(...parseSecondSection(ws, config));
  return results;
}

/** OA-直送 */
function parseOAZhiSong(ws) {
  const config = {
    channelName: "OA-直送专线",
    transportMode: "海运",
    vesselConfig: "OA联盟(COSCO/OOCL/EMC/CMA)",
    vesselTags: ["OA","COSCO","OOCL","EMC","CMA","直送"],
    deliveryMethod: "整柜直送",
    dataStartRow: 7,
    warehouseCol: 2,
    sourceSheet: "OA-直送",
    transitMin: 28, transitMax: 38, transitDesc: "28-38天",
    cityGroups: [
      { label: "华南/华东", kgStartCol: 3, cbmStartCol: 4, cities: ["深圳","中山","广州","东莞","惠州","汕头","厦门","义乌","苏州"] },
      { label: "青岛", kgStartCol: 5, cbmStartCol: 6, cities: ["青岛"] },
    ],
    weightTiers: [
      { label: "21KG+", value: 21, colOffset: 0, taxMode: "包税" },
      { label: "1CBM+", value: 1, colOffset: 1, taxMode: "不包税" },
    ],
  };
  return parseWarehousePriceMatrix(ws, config);
}

/** 美森美东秒送 */
function parseMeisenEast(ws) {
  const config = {
    channelName: "美森美东秒送",
    transportMode: "海运",
    vesselConfig: "美森MATSON GES",
    vesselTags: ["美森","MATSON","GES"],
    deliveryMethod: "卡派",
    dataStartRow: 7,
    warehouseCol: 2,
    sourceSheet: "美森美东秒送",
    transitMin: 21, transitMax: 25, transitDesc: "21-25天",
    cityGroups: [
      { label: "华南/华东", kgStartCol: 3, cbmStartCol: 4, cities: ["深圳","中山","广州","东莞","惠州","汕头","厦门","义乌","苏州"] },
      { label: "青岛", kgStartCol: 5, cbmStartCol: 6, cities: ["青岛"] },
    ],
    weightTiers: [
      { label: "50KG+", value: 50, colOffset: 0, taxMode: "包税" },
      { label: "1CBM+", value: 1, colOffset: 1, taxMode: "不包税" },
    ],
  };
  return parseWarehousePriceMatrix(ws, config);
}

/** MT-专线 */
function parseMTLine(ws) {
  const config = {
    channelName: "MT-专线",
    transportMode: "海运",
    vesselConfig: "2M/THE联盟统配",
    vesselTags: ["普船","2M","THE"],
    deliveryMethod: "卡派",
    dataStartRow: 7,
    warehouseCol: 2,
    sourceSheet: "MT-专线",
    transitMin: 32, transitMax: 41, transitDesc: "32-41天",
    cityGroups: [
      { label: "华南/华东", kgStartCol: 3, cbmStartCol: 4, cities: ["深圳","中山","广州","东莞","惠州","汕头","厦门","义乌","苏州"] },
      { label: "青岛", kgStartCol: 5, cbmStartCol: 6, cities: ["青岛"] },
    ],
    weightTiers: [
      { label: "21KG+", value: 21, colOffset: 0, taxMode: "包税" },
      { label: "1CBM+", value: 1, colOffset: 1, taxMode: "不包税" },
    ],
  };
  return parseWarehousePriceMatrix(ws, config);
}

/** OA-奥克兰 */
function parseOAAuckland(ws) {
  const config = {
    channelName: "OA-奥克兰",
    transportMode: "海运",
    vesselConfig: "OA联盟直航奥克兰",
    vesselTags: ["OA","COSCO","OOCL","奥克兰"],
    deliveryMethod: "卡派",
    dataStartRow: 7,
    warehouseCol: 2,
    sourceSheet: "OA-奥克兰",
    transitMin: 24, transitMax: 30, transitDesc: "24-30天",
    cityGroups: [
      { label: "华南/华东", kgStartCol: 3, cbmStartCol: 4, cities: ["深圳","中山","广州","东莞","惠州","汕头","厦门","义乌","苏州"] },
      { label: "青岛", kgStartCol: 5, cbmStartCol: 6, cities: ["青岛"] },
    ],
    weightTiers: [
      { label: "21KG+", value: 21, colOffset: 0, taxMode: "包税" },
      { label: "1CBM+", value: 1, colOffset: 1, taxMode: "不包税" },
    ],
  };
  return parseWarehousePriceMatrix(ws, config);
}

/** 海派渠道 */
function parseSeaExpress(ws) {
  // 海派渠道结构有所不同，更简单
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  const results = [];
  for (let ri = 5; ri < data.length; ri++) {
    const row = data[ri];
    const label = String(row[1] || row[0] || "").trim();
    if (!label) continue;

    const whMatch = label.match(/([A-Z]{2,}\d+)/g);
    const warehouses = whMatch || [label.slice(0, 20)];

    // 海派通常是 col2=KG价
    const kgPrice = parseFloat(row[2]);
    if (!isNaN(kgPrice) && kgPrice > 0 && kgPrice < 999) {
      for (const wh of warehouses) {
        results.push(makeRecord({
          channelName: "海派渠道",
          transportMode: "海运",
          vesselConfig: "海派UPS",
          vesselTags: ["海派","UPS"],
          deliveryMethod: "海派",
          destCode: wh,
          destRegion: identifyRegion(wh),
          billingType: "包税",
          minQty: "12KG+", minQtyValue: 12,
          price: kgPrice,
          priceUnit: "元/KG",
          sourceSheet: "海派渠道",
        }));
      }
    }
  }
  return results;
}

/** 商业/私人地址 */
function parseCommercialAddress(ws) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  const results = [];
  for (let ri = 7; ri < data.length; ri++) {
    const row = data[ri];
    const addrLabel = String(row[1] || "").trim();
    if (!addrLabel || addrLabel.includes("地址说明")) continue;

    const kgPrice = parseFloat(row[2]);
    if (!isNaN(kgPrice) && kgPrice > 0 && kgPrice < 999) {
      results.push(makeRecord({
        channelName: "商业/私人地址-卡派",
        transportMode: "海运",
        vesselConfig: "普船",
        vesselTags: ["普船"],
        deliveryMethod: "卡派",
        destCode: addrLabel.slice(0, 50),
        destType: "address",
        destRegion: "全美",
        billingType: "包税",
        minQty: "21KG+", minQtyValue: 21,
        price: kgPrice,
        priceUnit: "元/KG",
        sourceSheet: "商业|私人地址",
      }));
    }
  }
  return results;
}

/** 加拿大直航 */
function parseCanadaDirect(ws) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  const results = [];
  for (let ri = 7; ri < data.length; ri++) {
    const row = data[ri];
    const label = String(row[1] || row[0] || "").trim();
    if (!label || label.includes("仓库地址") || label.includes("加拿大")) continue;

    for (let t = 0; t < 4; t++) {
      const kgPrice = parseFloat(row[2 + t * 2]);
      const cbmPrice = parseFloat(row[3 + t * 2]);
      const qtyLabels = ["21KG+","50KG+","100KG+","500KG+"];
      const qtyVals = [21, 50, 100, 500];
      if (!isNaN(kgPrice) && kgPrice > 0 && kgPrice < 999) {
        results.push(makeRecord({
          country: "加拿大",
          channelName: "加拿大直航",
          transportMode: "海运",
          vesselConfig: "COSCO直航",
          vesselTags: ["COSCO","加拿大","直航"],
          deliveryMethod: "卡派",
          destCode: label.slice(0, 40),
          destType: "warehouse",
          destRegion: "加拿大",
          billingType: "包税",
          minQty: qtyLabels[t], minQtyValue: qtyVals[t],
          price: kgPrice,
          priceUnit: "元/KG",
          sourceSheet: "加拿大",
        }));
      }
    }
  }
  return results;
}

/** 美转加 */
function parseUSToCanada(ws) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  const results = [];
  for (let ri = 7; ri < data.length; ri++) {
    const row = data[ri];
    const label = String(row[1] || row[0] || "").trim();
    if (!label || label.includes("仓库地址") || label.includes("加拿大")) continue;

    const kgPrice = parseFloat(row[2]);
    if (!isNaN(kgPrice) && kgPrice > 0 && kgPrice < 999) {
      results.push(makeRecord({
        country: "加拿大",
        channelName: "美转加",
        transportMode: "海运",
        vesselConfig: "美转加(洛杉矶中转)",
        vesselTags: ["美转加","洛杉矶"],
        deliveryMethod: "卡派",
        destCode: label.slice(0, 40),
        destType: "warehouse",
        destRegion: "加拿大",
        billingType: "包税",
        minQty: "21KG+", minQtyValue: 21,
        price: kgPrice,
        priceUnit: "元/KG",
        sourceSheet: "美转加",
      }));
    }
  }
  return results;
}

/** 美转墨 (按货物品类 × CBM阶梯) */
function parseUSToMexico(ws, channelName) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  const results = [];
  // 美转墨结构: R3+ 品类行, col0=品类描述, col2=1CBM+, col3=5CBM+, col4=10CBM+
  for (let ri = 2; ri < data.length; ri++) {
    const row = data[ri];
    const category = String(row[0] || "").trim();
    if (!category || category.includes("产品类目") || category.includes("海派渠道") || category.includes("备注")) continue;
    if (category.includes("体积") || category.includes("计费方式") || category.includes("品牌查询") || category.includes("统配")) continue;

    const cbmTiers = [
      { label: "1CBM+", value: 1, col: 2 },
      { label: "5CBM+", value: 5, col: 3 },
      { label: "10CBM+", value: 10, col: 4 },
    ];
    for (const tier of cbmTiers) {
      const price = parseFloat(row[tier.col]);
      if (!isNaN(price) && price > 0 && price < 99999) {
        results.push(makeRecord({
          country: "墨西哥",
          channelName,
          transportMode: "海运",
          vesselConfig: "美转墨(WHL/COSCO/YM)",
          vesselTags: ["美转墨","WHL","COSCO","YM"],
          deliveryMethod: "卡派",
          destCode: "墨西哥城",
          destType: "warehouse",
          destRegion: "墨西哥",
          billingType: "包税",
          minQty: tier.label,
          minQtyValue: tier.value,
          price,
          priceUnit: "元/CBM",
          transitMin: 35, transitMax: 40, transitDesc: "35-40天",
          sourceSheet: channelName,
        }));
      }
    }
  }
  return results;
}

/** TH特惠 / 纽约 / 芝加哥 / 休斯敦 / 萨瓦纳 等专用线 */
function parseRegionalSheet(ws, config) {
  return parseWarehousePriceMatrix(ws, config);
}

// ═══════════════════════════════════════════════════════════════
// 主入口
// ═══════════════════════════════════════════════════════════════

function parseMeiQi(filePath) {
  console.log("[美琦] 开始解析:", filePath);
  const wb = XLSX.readFile(filePath);
  const allResults = [];

  // 映射 Sheet 名（含空格前缀）到解析器
  const sheetParserMap = new Map();

  // 美线核心
  const matchSheet = wb.SheetNames.find(n => n.includes("Match系列"));
  if (matchSheet) sheetParserMap.set(matchSheet, parseMatchSeries);

  const oaChai = wb.SheetNames.find(n => n.includes("OA-拆送"));
  if (oaChai) sheetParserMap.set(oaChai, parseOAChaiSong);

  const oaZhi = wb.SheetNames.find(n => n.includes("OA-直送"));
  if (oaZhi) sheetParserMap.set(oaZhi, parseOAZhiSong);

  const meisenEast = wb.SheetNames.find(n => n.includes("美森美东秒送"));
  if (meisenEast) sheetParserMap.set(meisenEast, parseMeisenEast);

  const mtLine = wb.SheetNames.find(n => n.trim() === "MT-专线");
  if (mtLine) sheetParserMap.set(mtLine, parseMTLine);

  const oaAuckland = wb.SheetNames.find(n => n.includes("OA-奥克兰"));
  if (oaAuckland) sheetParserMap.set(oaAuckland, parseOAAuckland);

  const seaExpress = wb.SheetNames.find(n => n.includes("海派渠道"));
  if (seaExpress) sheetParserMap.set(seaExpress, parseSeaExpress);

  const commAddr = wb.SheetNames.find(n => n.includes("商业") && n.includes("私人"));
  if (commAddr) sheetParserMap.set(commAddr, parseCommercialAddress);

  // 区域专线
  // 区域专线（与 Match 系列结构一致：col2=仓库, col3-4=华南, col5-6=青岛）
  const regionalConfigs = [
    { match: "纽约", channelName: "美东极速快车", vesselConfig: "ZIM带车架海铁联运", vesselTags: ["ZIM","海铁"], transitDesc: "26-38天", transitMin: 26, transitMax: 38 },
    { match: "休斯敦", channelName: "休斯顿专线", vesselConfig: "COSCO/EMC/OOCL", vesselTags: ["COSCO","EMC","OOCL","休斯顿"], transitDesc: "38-45天", transitMin: 38, transitMax: 45 },
    { match: "萨瓦纳", channelName: "萨瓦纳专线", vesselConfig: "OA联盟海铁联运", vesselTags: ["OA","海铁","萨瓦纳"], transitDesc: "40-45天", transitMin: 40, transitMax: 45 },
  ];
  for (const rc of regionalConfigs) {
    const sn = wb.SheetNames.find(n => n.includes(rc.match) && !n.includes("OA") && !n.includes("美森"));
    if (sn) {
      sheetParserMap.set(sn, (wss) => parseRegionalSheet(wss, {
        channelName: rc.channelName, vesselConfig: rc.vesselConfig, vesselTags: rc.vesselTags,
        transportMode: "海运", deliveryMethod: "卡派",
        dataStartRow: 7, warehouseCol: 2, sourceSheet: rc.match,
        transitMin: rc.transitMin, transitMax: rc.transitMax, transitDesc: rc.transitDesc,
        cityGroups: [
          { label:"华南/华东", kgStartCol:3, cbmStartCol:4, cities: DEFAULT_CITIES },
          { label:"青岛", kgStartCol:5, cbmStartCol:6, cities:["青岛"] },
        ],
        weightTiers: [
          { label:"50KG+", value:50, colOffset:0, taxMode:"包税" },
          { label:"1CBM+", value:1, colOffset:1, taxMode:"不包税" },
        ],
      }));
    }
  }

  // 芝加哥（双渠道：海铁快车 + 海铁联运）
  const chicagoSheet = wb.SheetNames.find(n => n.includes("芝加哥"));
  if (chicagoSheet) {
    sheetParserMap.set(chicagoSheet, (wss) => {
      const r1 = parseRegionalSheet(wss, {
        channelName: "芝加哥海铁快车", vesselConfig: "COSCO稳速达海铁", vesselTags: ["COSCO","海铁","芝加哥"],
        transportMode: "海运", deliveryMethod: "卡派",
        dataStartRow: 7, warehouseCol: 2, sourceSheet: "芝加哥",
        transitMin: 28, transitMax: 34, transitDesc: "28-34天",
        cityGroups: [
          { label:"华南/华东", kgStartCol:3, cbmStartCol:4, cities: DEFAULT_CITIES },
          { label:"青岛", kgStartCol:5, cbmStartCol:6, cities:["青岛"] },
        ],
        weightTiers: [
          { label:"50KG+", value:50, colOffset:0, taxMode:"包税" },
          { label:"1CBM+", value:1, colOffset:1, taxMode:"不包税" },
        ],
      });
      // 第二组: col9=KG_华南, col10=CBM_华南
      const data = XLSX.utils.sheet_to_json(wss, { header: 1, defval: "" });
      for (let ri = 6; ri < data.length; ri++) {
        const row = data[ri];
        const whCell = String(row[2] || "").trim();
        if (!whCell || whCell.length < 3) continue;
        const warehouses = parseWarehouses(whCell);
        if (warehouses.length === 0) continue;
        const kgPrice = parseFloat(row[9]);
        const cbmPrice = parseFloat(row[10]);
        if (!isNaN(kgPrice) && kgPrice > 0 && kgPrice < 999) {
          for (const wh of warehouses) {
            r1.push(makeRecord({
              channelName: "芝加哥海铁联运", vesselConfig: "普船海铁联运", vesselTags: ["普船","海铁","芝加哥"],
              transportMode: "海运", deliveryMethod: "卡派",
              destCode: wh, destRegion: identifyRegion(wh),
              originRegion: "华南/华东", originCities: DEFAULT_CITIES,
              billingType: "包税", minQty: "50KG+", minQtyValue: 50,
              price: kgPrice, priceUnit: "元/KG",
              cbmPrice: !isNaN(cbmPrice) && cbmPrice > 0 ? cbmPrice : null,
              transitMin: 30, transitMax: 38, transitDesc: "30-38天",
              sourceSheet: "芝加哥",
            }));
          }
        }
      }
      return r1;
    });
  }

  for (const rc of regionalConfigs) {
    const sn = wb.SheetNames.find(n => n.includes(rc.match) && !n.includes("OA") && !n.includes("美森"));
    if (sn) {
      sheetParserMap.set(sn, (wss) => parseRegionalSheet(wss, {
        channelName: rc.channelName, vesselConfig: rc.vesselConfig, vesselTags: rc.vesselTags,
        transportMode: "海运", deliveryMethod: "卡派",
        dataStartRow: 7, warehouseCol: 2, sourceSheet: rc.match,
        transitMin: rc.transitMin, transitMax: rc.transitMax, transitDesc: rc.transitDesc,
        cityGroups: [
          { label:"华南/华东", kgStartCol:3, cbmStartCol:4, cities: DEFAULT_CITIES },
          { label:"青岛", kgStartCol:5, cbmStartCol:6, cities:["青岛"] },
        ],
        weightTiers: [
          { label:"50KG+", value:50, colOffset:0, taxMode:"包税" },
          { label:"1CBM+", value:1, colOffset:1, taxMode:"不包税" },
        ],
      }));
    }
  }

  // 加拿大
  const canadaSheet = wb.SheetNames.find(n => n.trim() === "加拿大");
  if (canadaSheet) sheetParserMap.set(canadaSheet, parseCanadaDirect);

  const usToCa = wb.SheetNames.find(n => n.includes("美转加"));
  if (usToCa) sheetParserMap.set(usToCa, parseUSToCanada);

  // TH特惠
  const thSheet = wb.SheetNames.find(n => n.includes("TH特惠"));
  if (thSheet) {
    sheetParserMap.set(thSheet, (wss) => parseRegionalSheet(wss, {
      channelName: "TH特惠", vesselConfig: "普船特惠", vesselTags: ["普船","特惠"],
      transportMode: "海运", deliveryMethod: "卡派",
      dataStartRow: 7, warehouseCol: 2, sourceSheet: "TH特惠",
      transitMin: 26, transitMax: 32, transitDesc: "26-32天",
      cityGroups: [
        { label:"华南/华东", kgStartCol:3, cbmStartCol:4, cities: DEFAULT_CITIES },
        { label:"青岛", kgStartCol:5, cbmStartCol:6, cities:["青岛"] },
      ],
      weightTiers: [
        { label:"21KG+", value:21, colOffset:0, taxMode:"包税" },
        { label:"1CBM+", value:1, colOffset:1, taxMode:"不包税" },
      ],
    }));
  }

  // 墨西哥
  const mexicoSheets = [
    { match: "美转墨限时达", name: "美转墨限时达" },
    { match: "美转墨快线", name: "美转墨快线" },
    { match: "美转墨普线", name: "美转墨普线" },
    { match: "美转墨铭感货", name: "美转墨敏感货" },
  ];
  for (const ms of mexicoSheets) {
    const sn = wb.SheetNames.find(n => n.includes(ms.match));
    if (sn) sheetParserMap.set(sn, (wss) => parseUSToMexico(wss, ms.name));
  }

  const skippedSheets = [];

  for (const sheetName of wb.SheetNames) {
    // 跳过元数据 Sheet
    if (["目录","快捷查价","发货须知","合作流程","分区","反倾销列表",
      "美国亚马逊仓库","海外仓地址","墨西哥仓库","墨西哥海外仓",
      "暂存渠道","MATCH闪送","超大件海派"].some(k => sheetName.includes(k))) continue;

    const parser = sheetParserMap.get(sheetName);
    if (parser) {
      try {
        const results = parser(wb.Sheets[sheetName]);
        // 根据 Sheet 名自动设置国家
        const sheetCountry = detectCountry(sheetName);
        let countryOverride = sheetCountry;
        if (sheetName.includes("加拿大") || sheetName.includes("美转加")) countryOverride = "加拿大";
        if (sheetName.includes("墨西哥") || sheetName.includes("美转墨")) countryOverride = "墨西哥";
        for (const r of results) {
          if (countryOverride) r.country = countryOverride;
        }
        console.log(`  [${sheetName.trim()}] ${results.length} 条 → ${countryOverride || "美国"}`);
        allResults.push(...results);
      } catch (err) {
        console.error(`  [${sheetName.trim()}] 解析失败: ${err.message}`);
        console.error(err.stack);
      }
    } else {
      if (!["目录","快捷查价","发货须知","合作流程","分区","反倾销列表",
        "美国亚马逊仓库","海外仓地址","墨西哥仓库","墨西哥海外仓",
        "暂存渠道","MATCH闪送","超大件海派"].some(k => sheetName.includes(k))) {
        skippedSheets.push(sheetName.trim());
      }
    }
  }

  if (skippedSheets.length > 0) {
    console.log(`[美琦] ⚠ 暂未解析的Sheet: ${skippedSheets.join(" | ")}`);
  }

  console.log(`[美琦] 总计 ${allResults.length} 条`);
  return allResults;
}

module.exports = { parseMeiQi };
