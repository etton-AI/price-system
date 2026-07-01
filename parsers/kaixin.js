/**
 * 凯鑫科技 — 美/欧/英/加线价格解析器
 * 覆盖: 美国空派 / 欧洲卡派+快递派 / 欧洲超大件 / 英国 / 罗马尼亚专线 / A50加拿大
 */
const XLSX = require("xlsx");
const SUPPLIER = "凯鑫科技";

const EU_COUNTRIES = ["德国","法国","意大利","西班牙","波兰","捷克","荷兰","奥地利","比利时","卢森堡","丹麦","瑞典","芬兰","匈牙利","希腊","葡萄牙","爱尔兰","罗马尼亚","保加利亚","克罗地亚","斯洛文尼亚","斯洛伐克","爱沙尼亚","立陶宛","拉脱维亚"];

function mkr(o) {
  return {
    supplier: SUPPLIER, country: o.c || "欧线", channel_name: o.cn || "", transport_mode: o.tm || "铁路",
    vessel_config: o.vc || "", vessel_tags: o.vt || [], delivery_method: o.dm || "卡派",
    destination_type: o.dt || "warehouse", destination_code: o.dc || "", destination_region: o.dr || "",
    origin_region: "华南", origin_cities: ["深圳","东莞","广州","中山"],
    billing_type: o.bt || "包税", tax_mode: o.tx || o.bt || "包税",
    min_quantity: o.mq || "", min_quantity_value: o.mv || 0, unit_price: o.p || 0, price_unit: "元/KG",
    transit_time_min: o.tn || null, transit_time_max: o.tx2 || null, transit_time_desc: o.td || "",
    claim_rule: o.cr || "", effective_date: "", source_sheet: o.ss || "",
  };
}

/**
 * 通用分组解析: 检测表头行 → 识别列组(渠道×税模式) → 解析数据行
 * 适用于凯鑫大部分 Sheet
 */
function parseGroupedSheet(ws, sheetName, defaults) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 4) return [];
  const results = [];

  // Determine country from sheet name
  let country = defaults.country || "欧线";
  if (sheetName.includes("英国")) country = "英国";
  if (sheetName.includes("加拿大")) country = "加拿大";
  if (sheetName.includes("美国")) country = "美国";
  if (sheetName.includes("罗马尼亚")) country = "欧线";

  // Determine transport mode
  let tm = defaults.tm || "铁路";
  if (sheetName.includes("海运")) tm = "海运";
  if (sheetName.includes("空派") || sheetName.includes("空运")) tm = "空运";

  // Determine delivery method
  let dm = "卡派";
  if (sheetName.includes("快递派")) dm = "快递派";

  // Find header rows: look for rows with weight tier patterns (21KG+, 51KG+, etc.)
  let headerRowIdx = -1;
  let subHeaderRowIdx = -1;
  for (let ri = 1; ri < Math.min(data.length, 6); ri++) {
    const row = data[ri];
    let kgCount = 0;
    for (let c = 2; c < Math.min(row.length, 14); c++) {
      const cell = String(row[c] || "").trim();
      if (cell.match(/\d+\s*KG\+/i) || cell.match(/\d+\s*CBM\+/i)) kgCount++;
    }
    if (kgCount >= 2) {
      if (headerRowIdx < 0) headerRowIdx = ri;
      else if (subHeaderRowIdx < 0) subHeaderRowIdx = ri;
      else break;
    }
  }

  if (headerRowIdx < 0) return [];

  // Find channel/tax grouping rows (usually rows 1-3)
  // Look for rows with "包税", "递延", "自税", transport mode names
  const groupRows = [];
  for (let ri = 1; ri < headerRowIdx; ri++) {
    const row = data[ri];
    const text = row.slice(1).map(c => String(c || "").trim()).join(" ");
    if (text.includes("包税") || text.includes("递延") || text.includes("自税") || text.includes("铁路") || text.includes("海运") || text.includes("快铁") || text.includes("卡航")) {
      groupRows.push(ri);
    }
  }

  // Find channel column groups from header row
  const headerRow = data[headerRowIdx];
  const subHeaderRow = subHeaderRowIdx >= 0 ? data[subHeaderRowIdx] : null;

  // Identify channel groups: scan columns 2+ for contiguous blocks
  const channelGroups = [];
  let currentGroup = null;

  for (let col = 2; col < Math.min(headerRow.length, 18); col++) {
    const hdrCell = String(headerRow[col] || "").trim();
    const subCell = subHeaderRow ? String(subHeaderRow[col] || "").trim() : "";
    const wtMatch = hdrCell.match(/(\d+)\s*KG\+/i) || hdrCell.match(/(\d+)\s*CBM\+/i) || subCell.match(/(\d+)\s*KG\+/i) || subCell.match(/(\d+)\s*CBM\+/i);

    if (wtMatch || (currentGroup && (!hdrCell || hdrCell === "" || wtMatch))) {
      // Part of current group or new weight tier
      if (!currentGroup) {
        // Look back to find channel name from group rows
        currentGroup = { startCol: col, endCol: col, tiers: [], channelName: "", taxMode: "包税" };
      }
      let tierLabel = hdrCell || subCell;
      let mv = parseInt(wtMatch?.[1] || "21");
      currentGroup.tiers.push({ col, label: tierLabel, value: mv, qty: tierLabel });
      currentGroup.endCol = col;
    } else if (currentGroup && hdrCell && !wtMatch) {
      // End of group
      channelGroups.push(currentGroup);
      currentGroup = null;
    }
  }
  if (currentGroup && currentGroup.tiers.length > 0) {
    channelGroups.push(currentGroup);
  }

  // Try to infer channel names and tax modes from group rows
  if (channelGroups.length > 0 && groupRows.length > 0) {
    // For each channel group, look at the group rows to infer name
    for (let gi = 0; gi < channelGroups.length; gi++) {
      const cg = channelGroups[gi];
      const midCol = Math.floor((cg.startCol + cg.endCol) / 2);

      // Look for text in group rows around this column
      const labels = [];
      for (const gr of groupRows) {
        const cell = String(data[gr][midCol] || data[gr][cg.startCol] || "").trim();
        if (cell && cell.length > 1 && !cell.match(/^\d/)) labels.push(cell);
      }

      // Infer channel name
      if (labels.length > 0) {
        // Determine transport mode from labels
        let modeLabel = tm;
        if (labels.some(l => l.includes("铁路"))) modeLabel = "铁路";
        if (labels.some(l => l.includes("快铁"))) modeLabel = "快铁";
        if (labels.some(l => l.includes("海运"))) modeLabel = "海运";
        if (labels.some(l => l.includes("卡航"))) modeLabel = "卡航";

        let taxLabel = "包税";
        if (labels.some(l => l.includes("递延"))) taxLabel = "递延";
        if (labels.some(l => l.includes("自税"))) taxLabel = "自税";

        cg.channelName = `欧洲${modeLabel}-${dm}${taxLabel}`;
        if (country === "英国") cg.channelName = `英国${modeLabel}-${dm}${taxLabel}`;
        if (country === "加拿大") cg.channelName = `加拿大${modeLabel}-${dm}${taxLabel}`;
        cg.taxMode = taxLabel;
      } else {
        cg.channelName = `${country}${tm}-${dm}`;
      }
    }
  }

  // If no channel groups detected, use simple approach
  if (channelGroups.length === 0) {
    // Simple: all cols 2+ are one channel group
    const cg = { startCol: 2, endCol: Math.min(headerRow.length - 1, 12), tiers: [], channelName: `${country}${tm}-${dm}`, taxMode: "包税" };
    for (let col = 2; col <= cg.endCol; col++) {
      const cell = String(headerRow[col] || "").trim();
      const wtMatch = cell.match(/(\d+)\s*KG\+/i);
      if (wtMatch) {
        cg.tiers.push({ col, label: cell, value: parseInt(wtMatch[1]), qty: cell });
      }
    }
    if (cg.tiers.length > 0) channelGroups.push(cg);
  }

  // Parse data rows
  for (let ri = headerRowIdx + 1; ri < data.length; ri++) {
    const row = data[ri];
    const c0 = String(row[0] || "").trim();
    const c1 = String(row[1] || "").trim();

    if (!c1 && !c0) continue;
    if (c0.includes("下单渠道") || c0.includes("清关费") || c0.includes("报关费") || c0.includes("费用说明") || c0.includes("注意事项") || c0.includes("包装要求") || c0.includes("小货附加费") || c0.includes("运行路线")) continue;
    if (c1.includes("下单渠道") || c1.includes("清关费") || c1.includes("渠道名称")) continue;

    // Extract destination
    let warehouses = [];
    let destType = "warehouse";
    const destText = c1 || c0;

    const whMatch = destText.match(/[A-Z]{2,}\d[\d-]*/g);
    if (whMatch) {
      warehouses = whMatch;
      destType = "warehouse";
    } else {
      // Check for country names
      for (const cn of EU_COUNTRIES) {
        if (destText === cn || destText.startsWith(cn)) {
          warehouses = [cn];
          destType = "country";
          break;
        }
      }
      if (warehouses.length === 0) {
        // Multi-country
        const found = [];
        for (const cn of EU_COUNTRIES) {
          if (destText.includes(cn)) found.push(cn);
        }
        if (found.length > 0) {
          warehouses = found;
          destType = "country";
        } else if (destText.includes("四大仓")) {
          warehouses = ["亚马逊四大仓(BHX4/LBA4/BHX8/LBA8)"];
          destType = "warehouse";
        } else if (destText.includes("除四大仓")) {
          warehouses = ["英国其他亚马逊仓"];
          destType = "warehouse";
        } else if (destText.includes("英国全部")) {
          warehouses = ["英国全部亚马逊仓"];
          destType = "warehouse";
        } else if (destText.match(/[A-Z][a-z]/) && destText.length > 2) {
          warehouses = [destText.slice(0, 50)];
          destType = "country";
        } else {
          continue;
        }
      }
    }

    // Extract transit time from nearby columns
    let transitDesc = "";
    let transitMin = null, transitMax = null;
    for (let c = channelGroups[channelGroups.length-1]?.endCol + 1 || 7; c < Math.min(row.length, 14); c++) {
      const cell = String(row[c] || "").trim();
      if (cell && (cell.includes("天") || cell.includes("日") || cell.includes("工作日"))) {
        transitDesc = cell;
        const tm_ = cell.match(/(\d+)\s*[-–~]*\s*(\d+)/);
        if (tm_) { transitMin = parseInt(tm_[1]); transitMax = parseInt(tm_[2]); }
        break;
      }
    }

    for (const cg of channelGroups) {
      for (const tier of cg.tiers) {
        const price = parseFloat(row[tier.col]);
        if (isNaN(price) || price <= 0 || price > 99999) continue;

        for (const wh of warehouses) {
          results.push(mkr({
            c: country, cn: cg.channelName, tm, vc: tm, vt: [tm], dm,
            dc: wh, dt: destType, dr: wh,
            bt: cg.taxMode, tx: cg.taxMode,
            mq: tier.qty, mv: tier.value, p: price,
            td: transitDesc, tn: transitMin, tx2: transitMax, ss: sheetName,
          }));
        }
      }
    }
  }

  return results;
}

/**
 * 英国 Sheet 专用解析 (多种渠道+税模式混合)
 */
function parseUKSheet(ws) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 5) return [];
  const results = [];

  // R1: ["派送方式","仓点","铁路","",...]
  // R2: ["","","英国铁路","","","","","","中英快铁","",...]
  // R3: ["","","包税","","","自税/递延","","","包税","","","自税/递延"]
  // R4: ["","","21KG+","51KG+","100KG+","21KG+","51KG+","100KG+","21KG+","51KG+","100KG+","21KG+"]
  // R5+: data

  const r2 = data[2] || [];
  const r3 = data[3] || [];
  const r4 = data[4] || [];

  // Identify channel groups
  const groups = [];
  let cg = null;
  for (let col = 2; col < Math.min(r4.length, 14); col++) {
    const r2cell = String(r2[col] || "").trim();
    const r3cell = String(r3[col] || "").trim();
    const r4cell = String(r4[col] || "").trim();

    if (r2cell && (r2cell.includes("铁路") || r2cell.includes("快铁") || r2cell.includes("海运") || r2cell.includes("卡航"))) {
      if (cg) groups.push(cg);
      cg = { startCol: col, endCol: col, tiers: [], name: r2cell, taxMode: r3cell || "包税" };
    }
    if (cg && r4cell.match(/\d+\s*KG\+/i)) {
      const mv = parseInt(r4cell.match(/(\d+)/)[1]);
      cg.tiers.push({ col, label: r4cell, value: mv, qty: r4cell });
      cg.endCol = col;
    }
  }
  if (cg) groups.push(cg);

  for (let i = 0; i < groups.length; i++) {
    const g = groups[i];
    g.channelName = `英国${g.name}-卡派${g.taxMode}`;
  }

  for (let ri = 5; ri < data.length; ri++) {
    const row = data[ri];
    const c0 = String(row[0] || "").trim();
    const c1 = String(row[1] || "").trim();
    if (!c1 || c0.includes("下单渠道") || c0.includes("清关费")) continue;

    const dm = c0.includes("快递") ? "快递派" : "卡派";
    let warehouses = [];
    let destType = "warehouse";

    if (c1.includes("四大仓")) {
      warehouses = ["BHX4","LBA4","BHX8","LBA8"];
    } else if (c1.includes("除四大仓")) {
      warehouses = ["英国其他亚马逊仓"];
    } else if (c1.includes("全部亚马逊")) {
      warehouses = ["英国全部亚马逊仓"];
    } else if (c1.match(/[A-Z]{2,}\d/)) {
      warehouses = [c1.match(/[A-Z]{2,}\d/)[0]];
    } else {
      warehouses = ["英国"];
      destType = "country";
    }

    let transitDesc = "";
    let transitMin = null, transitMax = null;

    for (const g of groups) {
      for (const tier of g.tiers) {
        const price = parseFloat(row[tier.col]);
        if (isNaN(price) || price <= 0 || price > 99999) continue;

        for (const wh of warehouses) {
          results.push(mkr({
            c: "英国", cn: g.channelName.replace("卡派", dm === "快递派" ? "快递派" : "卡派"),
            tm: g.name.includes("铁路") || g.name.includes("快铁") ? "铁路" : g.name.includes("海运") ? "海运" : "卡航",
            vc: g.name, vt: [g.name], dm,
            dc: wh, dt: destType, dr: "英国",
            bt: g.taxMode, tx: g.taxMode,
            mq: tier.qty, mv: tier.value, p: price,
            td: transitDesc, tn: transitMin, tx2: transitMax, ss: "英国",
          }));
        }
      }
    }
  }
  return results;
}

function parseKaixin(filePath) {
  console.log("[凯鑫] 开始解析:", filePath);
  const wb = XLSX.readFile(filePath);
  const all = [];

  const skipSheets = ["目录", "装柜计划", "时效参考", "大货中转", "海外仓服务", "凯鑫发货必读", "产品附加费"];

  for (const sheetName of wb.SheetNames) {
    if (skipSheets.some(k => sheetName.includes(k))) continue;

    try {
      let results = [];
      if (sheetName === "英国") {
        results = parseUKSheet(wb.Sheets[sheetName]);
      } else {
        results = parseGroupedSheet(wb.Sheets[sheetName], sheetName, {});
      }
      console.log(`  [${sheetName}] ${results.length} 条`);
      all.push(...results);
    } catch (err) {
      console.error(`  [${sheetName}] 失败: ${err.message}`);
    }
  }

  console.log(`[凯鑫] 总计 ${all.length} 条`);
  return all;
}
module.exports = { parseKaixin };
