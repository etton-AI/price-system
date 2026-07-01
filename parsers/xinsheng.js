/**
 * 新胜供应链 — 英国线+欧洲线价格解析器
 * 覆盖: 英国海运/空派/卡航/铁路/超大件/小包 + 欧洲海运/空派/卡航/铁路
 * 支持: 义乌新胜 + 深圳新胜 两种文件格式
 */
const XLSX = require("xlsx");
const SUPPLIER = "新胜供应链";

const EU_COUNTRIES = ["德国","法国","意大利","西班牙","波兰","捷克","荷兰","奥地利","比利时","卢森堡","丹麦","瑞典","芬兰","匈牙利","葡萄牙","希腊","爱尔兰","罗马尼亚","保加利亚","克罗地亚","斯洛文尼亚","斯洛伐克"];

function parseCities(text) {
  const t = String(text).replace(/\r?\n/g, " ").trim();
  if (t.includes("华南")) return ["深圳","东莞","广州","中山","惠州"];
  if (t.includes("华东")) return ["上海","义乌","宁波","杭州","苏州"];
  return ["深圳","东莞","广州","中山"];
}

function parseCountry(text) {
  const t = String(text).replace(/\r?\n/g, " ").trim();
  for (const cn of EU_COUNTRIES) {
    if (t === cn || t.startsWith(cn)) return cn;
  }
  // Check if any EU country is in the text
  for (const cn of EU_COUNTRIES) {
    if (t.includes(cn)) return cn;
  }
  return t;
}

function parseWeightTier(text) {
  const t = String(text).replace(/\r?\n/g, " ").trim();
  const m = t.match(/(\d+)\s*[-–~]?\s*(\d+)?\s*KG\+?/i);
  if (m) {
    const v = parseInt(m[1]);
    return { label: t, value: v, qty: (m[2] ? m[1]+"-"+m[2] : m[1]) + "KG+" };
  }
  const m2 = t.match(/(\d+)\s*CBM\+?/i);
  if (m2) return { label: t, value: parseInt(m2[1]), qty: m2[0] };
  return null;
}

function mkr(o) {
  return {
    supplier: SUPPLIER, country: o.c || "英国", channel_name: o.cn || "", transport_mode: o.tm || "海运",
    vessel_config: o.vc || "", vessel_tags: o.vt || [], delivery_method: o.dm || "快递派",
    destination_type: o.dt || "warehouse", destination_code: o.dc || "", destination_region: o.dr || "",
    origin_region: "华南", origin_cities: ["深圳","东莞","广州","中山"],
    billing_type: o.bt || "包税", tax_mode: o.tx || o.bt || "包税",
    min_quantity: o.mq || "", min_quantity_value: o.mv || 0, unit_price: o.p || 0, price_unit: "元/KG",
    transit_time_min: o.tn || null, transit_time_max: o.tx2 || null, transit_time_desc: o.td || "",
    claim_rule: o.cr || "", effective_date: "", source_sheet: o.ss || "",
  };
}

/**
 * 通用表格解析器
 * 扫描 Sheet，识别 "渠道名称" 行 → 读取下一行的重量段 → 读取后续数据行
 */
function parseGenericSheet(ws, sheetName, defaults) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 4) return [];
  const results = [];

  // Infer country from sheet name
  let country = defaults.country || "英国";
  const sn = sheetName.toLowerCase();
  if (sn.includes("欧洲") || sn.includes("欧")) country = "欧洲";
  if (sn.includes("英国") || sn.includes("uk")) country = "英国";

  // Infer transport mode from sheet name
  let transportMode = defaults.tm || "海运";
  if (sn.includes("空派") || sn.includes("空运")) transportMode = "空运";
  if (sn.includes("卡航")) transportMode = "卡航";
  if (sn.includes("铁路")) transportMode = "铁路";
  if (sn.includes("超大件")) transportMode = transportMode; // keep parent mode

  let currentSection = null; // { channelName, taxMode, weightTiers: [{col, label, value}], deliveryMethod }

  for (let ri = 0; ri < data.length; ri++) {
    const row = data[ri];
    const c0 = String(row[0] || "").trim();
    const c1 = String(row[1] || "").trim();
    const c2 = String(row[2] || "").trim();

    // Skip empty rows and metadata
    if (!c0 && !c1 && !c2 && !String(row[3] || "").trim()) continue;

    // Detect section header: row with "渠道名称" in col0 or col1
    if (c0.includes("渠道名称") || c1.includes("渠道名称")) {
      // Next rows will be channel entries
      continue;
    }

    // Detect channel code row (K01, E03A, A01, B06, etc.)
    const chMatch = c1.match(/^([A-Z]\d+[A-Z]?)/);
    const isChannelRow = chMatch || (c0.match(/^([A-Z]\d+[A-Z]?)/) && c0.length <= 6);

    if (isChannelRow || (c1.length > 5 && (c1.includes("英国") || c1.includes("欧洲") || c1.includes("中欧") || c1.includes("空派") || c1.includes("海派") || c1.includes("海卡") || c1.includes("卡派")))) {
      // This is a channel row — extract channel name from c1 (or c0)
      let chName = c1 || c0;
      // Clean up channel name
      chName = chName.replace(/\r?\n/g, " ").trim();

      // Determine tax mode from channel name
      let taxMode = "包税";
      if (chName.includes("自税") || chName.includes("递延") || chName.includes("VAT") || chName.includes("vat")) {
        taxMode = "自税/递延";
      } else if (chName.includes("包税")) {
        taxMode = "包税";
      } else if (chName.includes("不包税") || chName.includes("不含税")) {
        taxMode = "不包税";
      }

      // Determine delivery method
      let dm = "快递派";
      if (chName.includes("海卡") || chName.includes("卡派")) dm = "卡派";
      if (chName.includes("海派") || chName.includes("快递派") || chName.includes("DPD") || chName.includes("UPS")) dm = "快递派";
      if (chName.includes("超大件")) dm = "卡派";

      // Find weight tiers: look at current row cols 3+, and also next row
      const weightTiers = [];
      // First check current row for weight tiers
      for (let col = 3; col < Math.min(row.length, 12); col++) {
        const cell = String(row[col] || "").trim();
        if (!cell) continue;
        const wt = parseWeightTier(cell);
        if (wt && wt.value > 0 && wt.value < 100000) {
          weightTiers.push({ col, ...wt });
        }
      }

      // If no weight tiers found in current row, check row index+1
      if (weightTiers.length === 0 && ri + 1 < data.length) {
        const nextRow = data[ri + 1];
        // Check if next row looks like weight tiers (all numeric-like headers)
        const nextRowTiers = [];
        for (let col = 3; col < Math.min(nextRow.length, 12); col++) {
          const cell = String(nextRow[col] || "").trim();
          if (!cell) continue;
          const wt = parseWeightTier(cell);
          if (wt && wt.value > 0 && wt.value < 100000) {
            nextRowTiers.push({ col, ...wt });
          }
        }
        if (nextRowTiers.length >= 2) {
          // Next row IS the weight tier row
          weightTiers.push(...nextRowTiers);
        }
      }

      if (weightTiers.length === 0) {
        // Default tiers based on sheet name
        const defaults_ = sn.includes("超大件") ? [{col:3,label:"100KG+",value:100,qty:"100KG+"},{col:4,label:"300KG+",value:300,qty:"300KG+"},{col:5,label:"500KG+",value:500,qty:"500KG+"}]
          : [{col:3,label:"21KG+",value:21,qty:"21KG+"},{col:4,label:"100KG+",value:100,qty:"100KG+"},{col:5,label:"500KG+",value:500,qty:"500KG+"}];
        weightTiers.push(...defaults_);
      }

      currentSection = { channelName: chName, taxMode, weightTiers, deliveryMethod: dm };
      continue;
    }

    // Check if this is a data row under a current section
    if (!currentSection) continue;

    // Extract destination (col2) and determine if warehouse or country
    const destText = c2 || c0;
    if (!destText || destText.includes("渠道名称") || destText.includes("国家") || destText.includes("重量") || destText.includes("计费")) continue;

    // Skip section headers and notes
    if (c0.includes("欧洲") && c0.includes("专线") && !c1) continue;
    if (c0.includes("申报货值") || c0.includes("说明") || c0.includes("备注")) continue;

    let destinations = [];
    let destType = "warehouse";

    // Determine destination
    if (country === "欧洲") {
      const cn = parseCountry(destText);
      if (cn && EU_COUNTRIES.some(c => cn.includes(c))) {
        destinations = [cn];
        destType = "country";
      } else if (destText.match(/[A-Z]{2,}\d/)) {
        destinations = [destText];
        destType = "warehouse";
      } else if (destText.includes("FBA") || destText.includes("亚马逊")) {
        destinations = ["欧洲FBA"];
        destType = "warehouse";
      } else {
        // Might be multiple countries
        const parts = destText.split(/[\/,，、]/);
        for (const p of parts) {
          const cn2 = parseCountry(p.trim());
          if (cn2 && EU_COUNTRIES.some(c => cn2.includes(c))) destinations.push(cn2);
        }
        if (destinations.length > 0) destType = "country";
        else continue; // can't parse destination
      }
    } else {
      // UK — extract warehouse codes
      const whMatch = destText.match(/[A-Z]{2,}\d/g);
      if (whMatch) {
        destinations = whMatch;
        destType = "warehouse";
      } else if (destText.includes("FBA") || destText.includes("海外仓") || destText.includes("其他")) {
        destinations = [destText.replace(/\r?\n/g, " ").trim().slice(0, 40)];
        destType = "warehouse";
      } else {
        destinations = [destText.slice(0, 40)];
        destType = "country";
      }
    }

    if (destinations.length === 0) continue;

    // Parse prices for each weight tier
    for (const wt of currentSection.weightTiers) {
      const price = parseFloat(row[wt.col]);
      if (!isNaN(price) && price > 0 && price < 99999) {
        for (const dest of destinations) {
          // Infer transit time from c8/c9 if present
          let transitDesc = "";
          let transitMin = null, transitMax = null;
          const transitCell = String(row[8] || row[9] || "").trim();
          if (transitCell) {
            transitDesc = transitCell;
            const tm_ = transitCell.match(/(\d+)\s*[-–~]\s*(\d+)/);
            if (tm_) { transitMin = parseInt(tm_[1]); transitMax = parseInt(tm_[2]); }
          }

          results.push(mkr({
            c: country, cn: currentSection.channelName, tm: transportMode,
            vc: transportMode, vt: [transportMode], dm: currentSection.deliveryMethod,
            dc: dest.slice(0, 50), dt: destType, dr: dest.slice(0, 50),
            bt: currentSection.taxMode, tx: currentSection.taxMode,
            mq: wt.qty, mv: wt.value, p: price,
            td: transitDesc, tn: transitMin, tx2: transitMax,
            ss: sheetName,
          }));
        }
      }
    }
  }

  return results;
}

/**
 * 简单价格表解析器 — 用于超大件/小包等简单格式
 * R0-R2: title, R3+: col0=描述, col2+=价格×重量段
 */
function parseSimpleSheet(ws, sheetName, country, channelName, transportMode, tiers) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 4) return [];
  const results = [];

  for (let ri = 3; ri < data.length; ri++) {
    const row = data[ri];
    const label = String(row[0] || row[1] || "").replace(/\r?\n/g, " ").trim();
    if (!label || label.includes("渠道") || label.includes("重量")) continue;

    const whMatch = label.match(/[A-Z]{2,}\d/g);
    const dest = whMatch ? whMatch[0] : label.slice(0, 40);
    const destType = whMatch ? "warehouse" : "country";

    for (let ti = 0; ti < tiers.length; ti++) {
      const price = parseFloat(row[2 + ti]);
      if (!isNaN(price) && price > 0 && price < 99999) {
        results.push(mkr({
          c: country, cn: channelName, tm: transportMode,
          vc: transportMode, vt: [transportMode], dm: "卡派",
          dc: dest, dt: destType, dr: dest,
          bt: "包税", tx: "包税",
          mq: tiers[ti].q, mv: tiers[ti].v, p: price, ss: sheetName,
        }));
      }
    }
  }
  return results;
}

function parseXinsheng(filePath) {
  console.log("[新胜] 开始解析:", filePath);
  const wb = XLSX.readFile(filePath);
  const all = [];

  // Auto-process all sheets that contain price data
  const skipSheets = ["目录", "出货须知", "出货必读", "查验仓储费", "附加费", "仓库分区", "偏远", "产品附加费",
    "超重超尺寸", "卡航附加费", "包税渠道产品附加费", "海运包税产品附加费", "海运快线产品附加费", "空派附加费"];

  for (const sheetName of wb.SheetNames) {
    if (skipSheets.some(k => sheetName.includes(k))) continue;

    try {
      const results = parseGenericSheet(wb.Sheets[sheetName], sheetName, {});
      if (results.length > 0) {
        console.log(`  [${sheetName}] ${results.length} 条`);
        all.push(...results);
      } else {
        // Try simple format
        let tiers = [{q:"21KG+",v:21},{q:"100KG+",v:100},{q:"500KG+",v:500}];
        let country = sheetName.includes("欧洲") ? "欧洲" : "英国";
        let tm = "海运";
        if (sheetName.includes("空")) tm = "空运";
        if (sheetName.includes("卡航")) tm = "卡航";
        if (sheetName.includes("铁路")) tm = "铁路";

        const r2 = parseSimpleSheet(wb.Sheets[sheetName], sheetName, country,
          sheetName.replace(/\r?\n/g, " ").trim(), tm, tiers);
        if (r2.length > 0) {
          console.log(`  [${sheetName}] ${r2.length} 条 (simple)`);
          all.push(...r2);
        }
      }
    } catch (err) {
      console.error(`  [${sheetName}] 失败: ${err.message}`);
    }
  }

  console.log(`[新胜] 总计 ${all.length} 条`);
  return all;
}
module.exports = { parseXinsheng };
