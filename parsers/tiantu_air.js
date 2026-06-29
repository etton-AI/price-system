/**
 * 天图通逊 — 英美空运价格解析器
 * 处理 "天图通逊英美空运&同行VIP价" 文件 (美国+英国空运)
 */
const XLSX = require("xlsx");
const SUPPLIER = "天图通逊";

const CITY_MAP = {
  华南: ["深圳", "广州", "中山", "东莞南城", "惠州"],
  重庆: ["重庆"],
  汕头: ["汕头"],
  "厦门/泉州/福州": ["厦门", "泉州", "福州"],
  华东: ["义乌", "上海", "宁波", "苏州", "杭州", "绍兴"],
  "青岛/郑州等": ["青岛", "郑州", "温州", "台州", "连云港", "南京", "合肥"],
  "天津/南昌/石家庄": ["天津", "南昌", "石家庄"],
  "济南/潍坊": ["济南", "潍坊"],
  "西安/沧州/保定": ["西安", "沧州", "保定"],
  "武汉/长沙/成都": ["武汉", "长沙", "成都"],
};

const CITY_COLS = [
  { key: "华南", cities: ["深圳", "广州", "中山", "东莞南城", "惠州"] },
  { key: "重庆", cities: ["重庆"] },
  { key: "厦门/泉州/福州", cities: ["厦门", "泉州", "福州"] },
  { key: "华东", cities: ["义乌", "上海", "宁波", "苏州", "杭州", "绍兴"] },
  { key: "青岛/郑州等", cities: ["青岛", "郑州", "温州", "台州", "连云港", "南京", "合肥"] },
  { key: "天津/南昌/石家庄", cities: ["天津", "南昌", "石家庄"] },
  { key: "济南/潍坊", cities: ["济南", "潍坊"] },
];

function mkr(o) {
  return {
    supplier: SUPPLIER, country: o.c || "美国", channel_name: o.cn || "", transport_mode: o.tm || "空运",
    vessel_config: o.vc || "", vessel_tags: o.vt || [], delivery_method: o.dm || "快递派",
    destination_type: o.dt || "warehouse", destination_code: o.dc || "", destination_region: o.dr || "",
    origin_region: o.or || "华南", origin_cities: o.oc || ["深圳"],
    billing_type: o.bt || "包税", tax_mode: o.tx || o.bt || "包税",
    min_quantity: o.mq || "", min_quantity_value: o.mv || 0, unit_price: o.p || 0, price_unit: "元/KG",
    transit_time_min: o.tn || null, transit_time_max: o.tx2 || null, transit_time_desc: o.td || "",
    claim_rule: o.cr || "", effective_date: "", source_sheet: o.ss || "",
  };
}

// 美国空运 (仓行×城市列)
function parseUSAirSheet(ws, sheetName, channelName) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 6) return [];
  const results = [];

  // R5 = city row, R6+ = data
  for (let ri = 5; ri < data.length; ri++) {
    const row = data[ri];
    const dest = String(row[0] || "").trim();
    if (!dest || dest.length < 2) continue;
    // Skip if not warehouse/postal code
    if (dest.includes("美国") || dest.includes("空运") || dest.includes("U.S.")) continue;
    if (dest.includes("备注") || dest.includes("说明")) continue;

    const destType = dest.match(/[A-Z]{2,}\d/) ? "warehouse" : "region";

    for (let ci = 1; ci < row.length && ci < CITY_COLS.length + 1; ci++) {
      const cityInfo = CITY_COLS[ci - 1];
      if (!cityInfo) continue;
      const p = parseFloat(row[ci]);
      if (!isNaN(p) && p > 0) {
        results.push(mkr({
          c: "美国", cn: channelName, tm: "空运", vc: "空运", vt: ["空运"],
          dm: channelName.includes("卡") ? "空卡" : "快递派",
          dc: dest, dt: destType, dr: dest,
          or: cityInfo.key, oc: cityInfo.cities,
          bt: "包税", mq: "500KG+", mv: 500, p,
          ss: sheetName,
        }));
      }
    }
  }
  return results;
}

// 美国空运 区域行 (邮编区域)
function parseUSAirRegion(ws, sheetName, channelName) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 7) return [];
  const results = [];

  // R5=city headers, R6=weight tiers, R7+=data
  for (let ri = 7; ri < data.length; ri++) {
    const row = data[ri];
    const region = String(row[0] || "").trim();
    if (!region || region.includes("备注")) continue;

    let dr = "";
    if (region.includes("8.9") || region.includes("西岸")) dr = "美西";
    else if (region.includes("97.98.99") || region.includes("西北")) dr = "美西北";
    else if (region.includes("4.5.6.7") || region.includes("中部")) dr = "美中";
    else if (region.includes("0.1.2.3") || region.includes("东岸")) dr = "美东";
    else continue;

    // Parse weight tiers per city
    const tierGroups = [
      { offset: 1, labels: ["12KG+", "71KG+", "100KG+"], vals: [12, 71, 100] },
      { offset: 4, labels: ["12KG+", "71KG+", "100KG+"], vals: [12, 71, 100] },
      { offset: 7, labels: ["12KG+", "71KG+", "100KG+"], vals: [12, 71, 100] },
    ];

    for (let gi = 0; gi < tierGroups.length; gi++) {
      const g = tierGroups[gi];
      const cityInfo = CITY_COLS[gi] || CITY_COLS[0];
      for (let ti = 0; ti < g.labels.length; ti++) {
        const p = parseFloat(row[g.offset + ti]);
        if (!isNaN(p) && p > 0) {
          results.push(mkr({
            c: "美国", cn: channelName, tm: "空运", vc: "空运", vt: ["空运"],
            dm: "快递派", dc: dr, dt: "region", dr,
            or: cityInfo.key, oc: cityInfo.cities,
            bt: "包税", mq: g.labels[ti], mv: g.vals[ti], p,
            ss: sheetName,
          }));
        }
      }
    }
  }
  return results;
}

// 英国空运 (城市行×包税/不包税)
function parseUKAir(ws) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 6) return [];
  const results = [];

  // Channel columns: 包税(cols 2-5), 不包税(cols 6-9)
  for (let ri = 5; ri < data.length; ri++) {
    const row = data[ri];
    const cityLabel = String(row[0] || "").trim();
    if (!cityLabel) continue;

    let cityName = "华南";
    if (cityLabel.includes("华南") || cityLabel.includes("深圳") || cityLabel.includes("广州")) cityName = "华南";
    else if (cityLabel.includes("华东") || cityLabel.includes("义乌")) cityName = "华东";
    else if (cityLabel.includes("重庆")) cityName = "重庆";

    const weightRow = data[ri];
    const weightLabels = ["12KG+", "71KG+", "100KG+", "300KG+"];
    const weightVals = [12, 71, 100, 300];

    // 包税 (cols 2-5)
    for (let ti = 0; ti < 4; ti++) {
      const p = parseFloat(weightRow[2 + ti]);
      if (!isNaN(p) && p > 0) {
        results.push(mkr({
          c: "英国", cn: "英国普货五日提-包税", tm: "空运", vc: "空运", vt: ["空运"],
          dm: "快递派", dc: "英国", dt: "country", dr: "英国",
          or: cityName, oc: [cityName],
          bt: "包税", mq: weightLabels[ti], mv: weightVals[ti], p,
          td: "3-5日提取", ss: "英国空运",
        }));
      }
    }
    // 不包税 (cols 6-9)
    for (let ti = 0; ti < 4; ti++) {
      const p = parseFloat(weightRow[6 + ti]);
      if (!isNaN(p) && p > 0) {
        results.push(mkr({
          c: "英国", cn: "英国普货五日提-不包税", tm: "空运", vc: "空运", vt: ["空运"],
          dm: "快递派", dc: "英国", dt: "country", dr: "英国",
          or: cityName, oc: [cityName],
          bt: "不包税", tx: "不包税", mq: weightLabels[ti], mv: weightVals[ti], p,
          td: "3-5日提取", ss: "英国空运",
        }));
      }
    }
    // Only parse first city group (华南)
    if (cityLabel.includes("华南") || cityLabel.includes("深圳")) break;
  }
  return results;
}

function parseTiantuAir(filePath) {
  console.log("[天图空运] 开始解析:", filePath);
  const wb = XLSX.readFile(filePath);
  const all = [];

  const configs = [
    { name: "美国空卡（8）", fn: (ws) => parseUSAirSheet(ws, "美国空卡8日达", "美国空卡-8日达") },
    { name: "美国空运&美西普货(5-8) ", fn: (ws) => parseUSAirRegion(ws, "美西普货5-8日达", "美国空派-美西普货5日提") },
  ];

  for (const cfg of configs) {
    const actual = wb.SheetNames.find(n => n.trim() === cfg.name.trim());
    if (actual) {
      try {
        const r = cfg.fn(wb.Sheets[actual]);
        console.log(`  [${cfg.name.trim()}] ${r.length} 条`);
        all.push(...r);
      } catch (err) { console.error(`  [${cfg.name.trim()}] 失败: ${err.message}`); }
    }
  }

  // UK air
  const ukSheet = wb.SheetNames.find(n => n.includes("英国空运"));
  if (ukSheet) {
    try {
      const r = parseUKAir(wb.Sheets[ukSheet]);
      console.log(`  [英国空运] ${r.length} 条`);
      all.push(...r);
    } catch (err) { console.error(`  [英国空运] 失败: ${err.message}`); }
  }

  console.log(`[天图空运] 总计 ${all.length} 条`);
  return all;
}
module.exports = { parseTiantuAir };
