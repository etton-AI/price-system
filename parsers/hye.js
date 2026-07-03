/**
 * HYE环洋国际 — 美线价格解析器
 * 格式: 海卡仓库组行 + 海派区域行
 */
const XLSX = require("xlsx");
const SUPPLIER = "HYE环洋";

const DEFAULT_CITIES = ["深圳", "东莞", "广州", "中山"];

function mkr(o) {
  return {
    supplier: SUPPLIER, country: "美国", channel_name: o.cn || "", transport_mode: o.tm || "海运",
    vessel_config: o.vc || "", vessel_tags: o.vt || [], delivery_method: o.dm || "卡派",
    destination_type: o.dt || "warehouse", destination_code: o.dc || "", destination_region: o.dr || "",
    origin_region: o.or || "华南", origin_cities: o.oc || DEFAULT_CITIES,
    billing_type: o.bt || "包税", tax_mode: o.tx || o.bt || "包税",
    min_quantity: o.mq || "", min_quantity_value: o.mv || 0, unit_price: o.p || 0, price_unit: o.pu || "元/KG",
    transit_time_min: o.tn || null, transit_time_max: o.tx2 || null, transit_time_desc: o.td || "",
    claim_rule: o.cr || "", effective_date: "", source_sheet: o.ss || "",
  };
}

/** Split slash-separated warehouse groups, filter noise */
function splitWarehouses(cell) {
  return String(cell || "").split("/").map(s => s.trim()).filter(s =>
    s && s.length >= 3 && !s.includes("海外仓") && !s.includes("自提") && !s.includes("商业地址") && !s.includes("头程")
  );
}

/** 海卡类: col1=仓库组, col3=KG价格, col4=CBM价格 */
function parseSeaCard(ws, sheetName, channelName, vesselConfig) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 6) return [];
  const results = [];

  for (let ri = 5; ri < data.length; ri++) {
    const row = data[ri];
    const col1 = String(row[1] || "").trim();
    if (!col1 || col1.length < 2) continue;
    if (col1.includes("仓库代码") || col1.includes("备注") || col1.includes("海关")) continue;

    const warehouses = splitWarehouses(col1);
    if (warehouses.length === 0) continue;

    // KG price (col3)
    const pKg = parseFloat(row[3]);
    // CBM price (col4)
    const pCbm = parseFloat(row[4]);

    for (const wh of warehouses) {
      if (wh.length < 3) continue;
      if (!isNaN(pKg) && pKg > 0) {
        results.push(mkr({ cn: channelName, tm: "海运", vc: vesselConfig, vt: [vesselConfig],
          dm: "卡派", dc: wh, dt: "warehouse", dr: wh,
          mq: "12KG+", mv: 12, p: pKg, ss: sheetName }));
      }
      if (!isNaN(pCbm) && pCbm > 0) {
        results.push(mkr({ cn: channelName, tm: "海运", vc: vesselConfig, vt: [vesselConfig],
          dm: "卡派", dc: wh, dt: "warehouse", dr: wh,
          bt: "不含税CBM", tx: "不含税CBM", mq: "2CBM+", mv: 2, p: pCbm, pu: "元/CBM", ss: sheetName }));
      }
    }
  }
  return results;
}

/** 海派: region rows with price columns */
function parseSeaParcel(ws, sheetName) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 6) return [];
  const results = [];
  let currentChannel = "";

  for (let ri = 3; ri < data.length; ri++) {
    const row = data[ri];
    const col1 = String(row[1] || "").trim();
    const col2 = String(row[2] || "").trim();

    // Channel name in col1
    if (col1 && col1.includes("HYE") && col1.includes("海派")) {
      currentChannel = col1; continue;
    }
    if (!currentChannel) continue;

    // Region row
    let dr = "";
    if (col1.includes("西岸") || col2.includes("80000")) dr = "美西";
    else if (col1.includes("中部") || col2.includes("40000")) dr = "美中";
    else if (col1.includes("东岸") || col2.includes("00000")) dr = "美东";
    else continue;

    // Prices in col4 (12-100KG) and col5 (101+KG)
    const tiers = [{ col: 4, label: "12-100KG", val: 12 }, { col: 5, label: "101KG+", val: 101 }];
    for (const t of tiers) {
      const p = parseFloat(row[t.col]);
      if (!isNaN(p) && p > 0) {
        results.push(mkr({ cn: currentChannel, tm: "海运", vc: "海运", vt: ["海运", "海派"],
          dm: "快递派", dc: dr, dt: "region", dr,
          mq: t.label, mv: t.val, p, ss: sheetName }));
      }
    }
  }
  return results;
}

/** 头程报价 */
function parseHeadhaul(ws) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 6) return [];
  const results = [];

  for (let ri = 5; ri < data.length; ri++) {
    const row = data[ri];
    const col2 = String(row[2] || "").trim();
    if (!col2 || !col2.includes("HYE")) continue;

    const channelName = col2;
    const p = parseFloat(row[5]); // col5 = 1KG+ price
    if (!isNaN(p) && p > 0) {
      results.push(mkr({ cn: channelName, tm: "海运", vc: "海运", vt: ["海运", "头程"],
        dm: "头程", dc: "*", dt: "none", dr: "",
        mq: "1KG+", mv: 1, p, ss: "海派头程报价汇总" }));
    }
  }
  return results;
}

// ── 主入口 ──
function parseHYE(filePath) {
  console.log("[HYE环洋] 开始解析:", filePath);
  const wb = XLSX.readFile(filePath);
  const all = [];
  const sn = wb.SheetNames;

  // 海派汇总 (多个HYE海派渠道)
  if (sn.includes("HYE海派渠道汇总")) {
    const r = parseSeaParcel(wb.Sheets["HYE海派渠道汇总"], "HYE海派渠道汇总");
    console.log(`  [HYE海派渠道汇总] ${r.length} 条`);
    all.push(...r);
  }

  // 海卡类 (仓库行)
  const seaCardConfigs = [
    ["HYE快船海卡", "HYE快提统配美西海卡", "快提统配"],
    ["HYE普船海卡", "HYE普船统配美西LA海卡", "普船统配LA"],
    ["HYE美西特惠普船LA海卡", "HYE美西特惠普船LA海卡", "特惠普船LA"],
  ];
  for (const [snName, chName, vc] of seaCardConfigs) {
    if (sn.includes(snName)) {
      const r = parseSeaCard(wb.Sheets[snName], snName, chName, vc);
      console.log(`  [${snName}] ${r.length} 条`);
      all.push(...r);
    }
  }

  // 商私卡
  ["HYE美西商私卡", "HYE美东商私卡"].forEach(sh => {
    if (sn.includes(sh)) {
      const dr = sh.includes("美东") ? "美东" : "美西";
      const data = XLSX.utils.sheet_to_json(wb.Sheets[sh], { header: 1, defval: "" });
      const results = [];
      for (let ri = 3; ri < data.length; ri++) {
        const row = data[ri];
        const col1 = String(row[1] || "").trim();
        if (!col1 || col1.length < 3) continue;
        const p = parseFloat(row[3]); // KG price
        if (!isNaN(p) && p > 0) {
          results.push(mkr({ cn: sh, tm: "海运", vc: "海运", vt: ["海运"], dm: "卡派",
            dc: col1, dt: "warehouse", dr, mq: "12KG+", mv: 12, p, ss: sh }));
        }
      }
      console.log(`  [${sh}] ${results.length} 条`);
      all.push(...results);
    }
  });

  // 头程
  if (sn.includes("海派头程报价汇总")) {
    const r = parseHeadhaul(wb.Sheets["海派头程报价汇总"]);
    console.log(`  [海派头程] ${r.length} 条`);
    all.push(...r);
  }

  console.log(`[HYE环洋] 总计 ${all.length} 条`);
  return all;
}

module.exports = { parseHYE };
