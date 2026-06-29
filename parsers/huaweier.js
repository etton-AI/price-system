/**
 * 华威尔 — 美线+欧洲线价格解析器
 *
 * 覆盖: 美国空派(大陆飞/小货/超大件) + 美国海运(海派/海卡) + 欧洲空派(普货/带电)
 */
const XLSX = require("xlsx");
const SUPPLIER = "华威尔";

const SZ_CITIES = ["深圳", "东莞", "广州", "中山"];
const YW_CITIES = ["义乌", "上海", "宁波", "杭州"];

function mkr(o) {
  return {
    supplier: SUPPLIER, country: o.c || "美国", channel_name: o.cn || "", transport_mode: o.tm || "海运",
    vessel_config: o.vc || "", vessel_tags: o.vt || [], delivery_method: o.dm || "卡派",
    destination_type: o.dt || "warehouse", destination_code: o.dc || "", destination_region: o.dr || "",
    origin_region: o.or || "华南", origin_cities: o.oc || SZ_CITIES,
    billing_type: o.bt || "包税", tax_mode: o.tx || o.bt || "包税",
    min_quantity: o.mq || "", min_quantity_value: o.mv || 0, unit_price: o.p || 0, price_unit: "元/KG",
    transit_time_min: o.tn || null, transit_time_max: o.tx2 || null, transit_time_desc: o.td || "",
    claim_rule: o.cr || "", effective_date: "", source_sheet: o.ss || "",
  };
}

function parseTransit(text) {
  const c = String(text || "").replace(/\r?\n/g, " ").trim();
  const m = c.match(/(\d+)\s*[-–~]*\s*(\d+)\s*(?:个)?(?:自然日|天|工作日)/);
  if (m) return { min: parseInt(m[1]), max: parseInt(m[2]), desc: c };
  return { min: null, max: null, desc: c };
}

// ── 美国空派大陆飞 (邮编区域行) ──
function parseUSAir(ws) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 5) return [];
  const results = [];
  let currentCh = "";

  for (let ri = 4; ri < data.length; ri++) {
    const row = data[ri];
    const c1 = String(row[1] || "").trim();
    const c2 = String(row[2] || "").trim();
    if (!c1 || c1.includes("渠道代码")) continue;
    if (c1.match(/^A\d?-?[AB]?/) || c1.match(/特快|经济/)) { currentCh = c1; continue; }
    if (!currentCh) continue;

    const regionText = c2;
    let dr = "";
    if (regionText.includes("7-9") || regionText.includes("8-9")) dr = "美西";
    else if (regionText.includes("0-6") || regionText.includes("0-4")) dr = "美东";
    else continue;

    const tiers = [{ col: 3, q: "10KG+", v: 10 }, { col: 4, q: "21KG+", v: 21 },
      { col: 5, q: "71KG+", v: 71 }, { col: 6, q: "101KG+", v: 101 }];
    for (const t of tiers) {
      const p = parseFloat(row[t.col]);
      if (!isNaN(p) && p > 0) {
        results.push(mkr({ c: "美国", cn: `美国空派-${currentCh}`, tm: "空运", vc: "空运", vt: ["空运"],
          dm: "快递派", dc: dr, dt: "region", dr, bt: "包税", mq: t.q, mv: t.v, p, ss: "美国空派大陆飞" }));
      }
    }
  }
  return results;
}

// ── 美国海运快递派 (邮编区域行, 华南/华东两列) ──
function parseUSSeaExpress(ws) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 6) return [];
  const results = [];
  let currentVessel = "";

  for (let ri = 4; ri < data.length; ri++) {
    const row = data[ri];
    const c0 = String(row[0] || "").trim();
    const c1 = String(row[1] || "").trim();
    if (!c1 && c0) continue;

    if (c0.includes("美森") || c0.includes("以星") || c0.includes("OA") || c0.includes("COSCO") || c0.includes("快船")) {
      currentVessel = c0; continue;
    }
    if (!currentVessel || !c1) continue;

    // Region row
    let dr = "";
    if (c1.includes("80000-96999") || c1.includes("美西")) dr = "美西";
    else if (c1.includes("97000-99999") || c1.includes("美中")) dr = "美中";
    else if (c1.match(/[0-7]/)) dr = "美东";
    else continue;

    // 华南: cols 2-3, 华东: cols 5-6
    const tiers = [
      { c: 2, q: "21KG+", v: 21, or: "华南", oc: SZ_CITIES },
      { c: 3, q: "101KG+", v: 101, or: "华南", oc: SZ_CITIES },
      { c: 5, q: "21KG+", v: 21, or: "华东", oc: YW_CITIES },
      { c: 6, q: "101KG+", v: 101, or: "华东", oc: YW_CITIES },
    ];
    for (const t of tiers) {
      const p = parseFloat(row[t.c]);
      if (!isNaN(p) && p > 0) {
        results.push(mkr({ c: "美国", cn: `美国海派-${currentVessel}`, tm: "海运", vc: currentVessel, vt: ["海运", "海派"],
          dm: "海派", dc: dr, dt: "region", dr, or: t.or, oc: t.oc,
          bt: "包税", mq: t.q, mv: t.v, p, ss: "美国海运快递派" }));
      }
    }
  }
  return results;
}

// ── 美西FBA海卡 (仓行, 以星/88快船) ──
function parseUSSeaTruck(ws) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 7) return [];
  const results = [];

  for (let ri = 6; ri < data.length; ri++) {
    const row = data[ri];
    const wh = String(row[1] || "").trim();
    if (!wh || wh.length < 3) continue;

    const configs = [
      { cn: "以星", c: 2, or: "华南" }, { cn: "88快船", c: 3, or: "华南" },
      { cn: "以星", c: 5, or: "华东" }, { cn: "88快船", c: 6, or: "华东" },
    ];
    for (const cfg of configs) {
      const p = parseFloat(row[cfg.c]);
      if (!isNaN(p) && p > 0) {
        results.push(mkr({ c: "美国", cn: `美西FBA海卡-${cfg.cn}`, tm: "海运", vc: cfg.cn, vt: ["海运", cfg.cn],
          dm: "卡派", dc: wh, dt: "warehouse", dr: "美西", or: cfg.or, oc: cfg.or === "华南" ? SZ_CITIES : YW_CITIES,
          bt: "包税", mq: "100KG+", mv: 100, p, ss: "美西FBA海卡" }));
      }
    }
  }
  return results;
}

// ── 欧洲空派 (国家行) ──
function parseEUAir(ws) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 6) return [];
  const results = [];
  const EU = ["德国", "法国", "意大利", "西班牙", "波兰", "捷克", "荷兰", "奥地利", "比利时", "卢森堡", "丹麦", "瑞典", "芬兰", "匈牙利", "葡萄牙", "斯洛伐克", "斯洛文尼亚", "罗马尼亚", "保加利亚", "克罗地亚", "拉脱维亚", "立陶宛", "爱沙尼亚"];

  for (let ri = 5; ri < data.length; ri++) {
    const row = data[ri];
    const c0 = String(row[0] || "").trim();
    if (!c0 || c0.includes("国家名")) continue;

    const transit = String(row[6] || "").replace(/\r?\n/g, " ").trim();
    const td = parseTransit(transit);

    const countries = [];
    for (const cn of EU) { if (c0.includes(cn)) countries.push(cn); }
    if (countries.length === 0) continue;

    const tiers = [{ c: 1, q: "21KG+", v: 21 }, { c: 2, q: "45KG+", v: 45 },
      { c: 3, q: "71KG+", v: 71 }, { c: 4, q: "100KG+", v: 100 }, { c: 5, q: "500KG+", v: 500 }];

    for (const cn of countries) {
      for (const t of tiers) {
        const p = parseFloat(row[t.c]);
        if (!isNaN(p) && p > 0) {
          results.push(mkr({ c: "欧洲", cn: "欧洲空派-普货包税P1快线", tm: "空运", vc: "空运", vt: ["空运"],
            dm: "快递派", dc: cn, dt: "country", dr: cn, bt: "包税",
            mq: t.q, mv: t.v, p, tn: td.min, tx2: td.max, td: td.desc, ss: "欧洲空派普货" }));
        }
      }
    }
  }
  return results;
}

// ── 主入口 ──
function parseHuaweier(filePath) {
  console.log("[华威尔] 开始解析:", filePath);
  const wb = XLSX.readFile(filePath);
  const all = [];

  const configs = [
    { name: "美国空派大陆飞综合报价（A A1 A2重货价)", fn: parseUSAir },
    { name: "美国海运快递派（综合报价）", fn: parseUSSeaExpress },
    { name: "美西FBA海卡（综合报价）", fn: parseUSSeaTruck },
    { name: "欧洲空派综合报价（普货）", fn: parseEUAir },
    { name: "欧洲空派综合报价（带电）", fn: (ws) => parseEUAir(ws, "带电") },
    { name: "欧洲空派超大件综合报价", fn: (ws) => parseEUAir(ws, "超大件") },
  ];

  for (const cfg of configs) {
    if (wb.SheetNames.includes(cfg.name)) {
      try {
        const r = cfg.fn(wb.Sheets[cfg.name]);
        console.log(`  [${cfg.name.slice(0, 25)}] ${r.length} 条`);
        all.push(...r);
      } catch (err) { console.error(`  [${cfg.name.slice(0, 20)}] 失败: ${err.message}`); }
    }
  }
  console.log(`[华威尔] 总计 ${all.length} 条`);
  return all;
}
module.exports = { parseHuaweier };
