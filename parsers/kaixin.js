/**
 * 凯鑫科技 — 美线+欧洲线+英国线价格解析器
 * 覆盖: 美国空派 / 欧洲卡派+快递派(铁路/海运/卡航) / 英国(铁路/海运/卡航)
 */
const XLSX = require("xlsx");
const SUPPLIER = "凯鑫科技";

const EU_COUNTRIES = ["德国", "法国", "意大利", "西班牙", "波兰", "捷克", "荷兰", "奥地利", "比利时", "卢森堡", "丹麦", "瑞典", "芬兰", "匈牙利", "希腊", "葡萄牙", "爱尔兰", "罗马尼亚", "保加利亚", "克罗地亚", "斯洛文尼亚", "斯洛伐克"];

function pCountries(cell) {
  const t = String(cell).replace(/\r?\n/g, " ").trim();
  for (const cn of EU_COUNTRIES) { if (t === cn) return [cn]; }
  const parts = t.split(/[ ,，、]+/).filter(Boolean);
  const r = [];
  for (const p of parts) { for (const cn of EU_COUNTRIES) { if (p === cn || p.includes(cn)) { r.push(cn); break; } } }
  return r.length > 0 ? r : [t];
}

function mkr(o) {
  return {
    supplier: SUPPLIER, country: o.c || "美国", channel_name: o.cn || "", transport_mode: o.tm || "海运",
    vessel_config: o.vc || "", vessel_tags: o.vt || [], delivery_method: o.dm || "卡派",
    destination_type: o.dt || "country", destination_code: o.dc || "", destination_region: o.dr || "",
    origin_region: "华南", origin_cities: ["深圳", "东莞", "广州", "中山"],
    billing_type: o.bt || "包税", tax_mode: o.tx || o.bt || "包税",
    min_quantity: o.mq || "", min_quantity_value: o.mv || 0, unit_price: o.p || 0, price_unit: "元/KG",
    transit_time_min: o.tn || null, transit_time_max: o.tx2 || null, transit_time_desc: o.td || "",
    claim_rule: o.cr || "", effective_date: "", source_sheet: o.ss || "",
  };
}

// ── 欧洲快递派 (铁路/海运/卡航 sub-tables within one sheet) ──
function parseEUDelivery(ws, sheetName) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 5) return [];
  const results = [];
  let currentMode = "", currentCh = "";

  for (let ri = 1; ri < data.length; ri++) {
    const row = data[ri];
    const c1 = String(row[1] || "").trim();
    const c2 = String(row[2] || "").trim();

    if (c1.includes("铁路")) { currentMode = "铁路"; continue; }
    if (c1.includes("海运")) { currentMode = "海运"; continue; }
    if (c1.includes("卡航")) { currentMode = "卡航"; continue; }
    if (c2.includes("包税")) { currentCh = "包税"; continue; }
    if (c2.includes("递延") || c2.includes("自税")) { currentCh = "递延/自税"; continue; }

    if (!currentMode || !currentCh) continue;
    if (!c1 || c1 === "国家/渠道" || c1 === "渠道" || c1 === "派送方式") continue;

    const countries = pCountries(c1);
    if (countries.length === 0) continue;

    const taxCols = currentCh.includes("递延") ? { start: 4, labels: ["21KG+", "51KG+", "101KG+"], vals: [21, 51, 101] }
      : { start: 1, labels: ["21KG+", "51KG+", "101KG+"], vals: [21, 51, 101] };

    if (currentCh === "包税") {
      for (let ti = 0; ti < 3; ti++) {
        const p = parseFloat(row[2 + ti]); // cols 2,3,4
        if (!isNaN(p) && p > 0) {
          for (const cn of countries) {
            results.push(mkr({ c: "欧洲", cn: `欧洲${currentMode}-快递派包税`, tm: currentMode, vc: currentMode, vt: [currentMode],
              dm: "快递派", dc: cn, dt: "country", dr: cn, bt: "包税",
              mq: taxCols.labels[ti], mv: taxCols.vals[ti], p, ss: sheetName }));
          }
        }
      }
      // 递延 cols 5,6,7
      for (let ti = 0; ti < 3; ti++) {
        const p = parseFloat(row[5 + ti]);
        if (!isNaN(p) && p > 0) {
          for (const cn of countries) {
            results.push(mkr({ c: "欧洲", cn: `欧洲${currentMode}-快递派递延`, tm: currentMode, vc: currentMode, vt: [currentMode],
              dm: "快递派", dc: cn, dt: "country", dr: cn, bt: "递延", tx: "递延",
              mq: taxCols.labels[ti], mv: taxCols.vals[ti], p, ss: sheetName }));
          }
        }
      }
    }
  }
  return results;
}

// ── 英国 (铁路/海运/卡航) ──
function parseUK(ws) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 5) return [];
  const results = [];
  let currentMode = "", currentCh = "";

  for (let ri = 1; ri < data.length; ri++) {
    const row = data[ri];
    const c0 = String(row[0] || "").trim();
    const c1 = String(row[1] || "").trim();
    const c2 = String(row[2] || "").trim();

    if (c2.includes("铁路")) { currentMode = "铁路"; continue; }
    if (c2.includes("海运")) { currentMode = "海运"; continue; }
    if (c2.includes("卡航")) { currentMode = "卡航"; continue; }
    if (c2.includes("包税")) { currentCh = "包税"; continue; }
    if (c2.includes("自税") || c2.includes("递延")) { currentCh = "自税/递延"; continue; }

    if (!currentMode || !currentCh) continue;
    const warehouse = c1.replace(/\r?\n/g, " ").trim();
    if (!warehouse || warehouse === "仓点") continue;

    let dest = warehouse;
    if (warehouse.includes("四大仓")) dest = "亚马逊四大仓(BHX4/LBA4/BHX8/LBA8)";
    else if (warehouse.includes("除四大仓")) dest = "英国其他亚马逊仓";

    if (currentCh === "包税") {
      for (let ti = 0; ti < 3; ti++) {
        const p = parseFloat(row[3 + ti]);
        if (!isNaN(p) && p > 0) {
          results.push(mkr({ c: "英国", cn: `英国${currentMode}-卡派包税`, tm: currentMode, vc: currentMode, vt: [currentMode],
            dm: "卡派", dc: dest, dt: "warehouse", dr: "英国", bt: "包税",
            mq: ["21KG+", "51KG+", "100KG+"][ti], mv: [21, 51, 100][ti], p, ss: "英国" }));
        }
      }
    } else {
      for (let ti = 0; ti < 3; ti++) {
        const p = parseFloat(row[6 + ti]);
        if (!isNaN(p) && p > 0) {
          results.push(mkr({ c: "英国", cn: `英国${currentMode}-卡派自税`, tm: currentMode, vc: currentMode, vt: [currentMode],
            dm: "卡派", dc: dest, dt: "warehouse", dr: "英国", bt: "自税", tx: "自税",
            mq: ["21KG+", "51KG+", "100KG+"][ti], mv: [21, 51, 100][ti], p, ss: "英国" }));
        }
      }
    }
  }
  return results;
}

// ── 美国空派 ──
function parseUSAir(ws) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 5) return [];
  const results = [];
  let currentCh = "";

  for (let ri = 3; ri < data.length; ri++) {
    const row = data[ri];
    const c1 = String(row[1] || "").trim();
    const c2 = String(row[2] || "").trim();

    if (c1.match(/A\d+/) || c1.includes("空派")) { currentCh = c1; continue; }
    if (!currentCh || !c2) continue;

    let dr = "";
    if (c2.includes("8.9") || c2.includes("西岸")) dr = "美西";
    else if (c2.includes("4.5.6.7") || c2.includes("中部")) dr = "美中";
    else if (c2.includes("0.1.2.3") || c2.includes("东部")) dr = "美东";
    else continue;

    const tiers = [{ c: 3, q: "21KG+", v: 21 }, { c: 4, q: "45KG+", v: 45 }, { c: 5, q: "71KG+", v: 71 }, { c: 6, q: "101KG+", v: 101 }];
    for (const t of tiers) {
      const p = parseFloat(row[t.c]);
      if (!isNaN(p) && p > 0) {
        results.push(mkr({ c: "美国", cn: `美国空派-${currentCh}`, tm: "空运", vc: "空运", vt: ["空运"],
          dm: "快递派", dc: dr, dt: "region", dr, bt: "包税", mq: t.q, mv: t.v, p, ss: "美国空派" }));
      }
    }
  }
  return results;
}

function parseKaixin(filePath) {
  console.log("[凯鑫] 开始解析:", filePath);
  const wb = XLSX.readFile(filePath);
  const all = [];

  const configs = [
    { name: "美国空派", fn: parseUSAir },
    { name: "欧洲-快递派", fn: (ws) => parseEUDelivery(ws, "欧洲-快递派") },
    { name: "英国", fn: parseUK },
  ];

  for (const cfg of configs) {
    if (wb.SheetNames.includes(cfg.name)) {
      try {
        const r = cfg.fn(wb.Sheets[cfg.name]);
        console.log(`  [${cfg.name}] ${r.length} 条`);
        all.push(...r);
      } catch (err) { console.error(`  [${cfg.name}] 失败: ${err.message}`); }
    }
  }
  console.log(`[凯鑫] 总计 ${all.length} 条`);
  return all;
}
module.exports = { parseKaixin };
