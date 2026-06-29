/**
 * 深圳新胜供应链 — 英国线+欧洲线价格解析器
 * 覆盖: 英国海运/空运/卡航 + 欧洲海运/卡航/空派 + 超大件
 */
const XLSX = require("xlsx");
const SUPPLIER = "新胜供应链";

const EU_C = ["德国", "法国", "意大利", "西班牙", "波兰", "捷克", "荷兰", "奥地利", "比利时", "卢森堡", "丹麦", "瑞典", "芬兰", "匈牙利", "葡萄牙", "希腊"];

function pC(cell) {
  const t = String(cell).replace(/\r?\n/g, " ").trim();
  for (const cn of EU_C) { if (t === cn || t.includes(cn)) return [cn]; }
  const parts = t.split(/[\/,，、]+/).filter(Boolean);
  for (const p of parts) { for (const cn of EU_C) { if (p.includes(cn)) return [cn]; } }
  return [t];
}
function mkr(o) {
  return {
    supplier: SUPPLIER, country: o.c || "欧洲", channel_name: o.cn || "", transport_mode: o.tm || "海运",
    vessel_config: o.vc || "", vessel_tags: o.vt || [], delivery_method: o.dm || "快递派",
    destination_type: o.dt || "country", destination_code: o.dc || "", destination_region: o.dr || "",
    origin_region: "华南", origin_cities: ["深圳", "东莞", "广州", "中山"],
    billing_type: o.bt || "包税", tax_mode: o.tx || o.bt || "包税",
    min_quantity: o.mq || "", min_quantity_value: o.mv || 0, unit_price: o.p || 0, price_unit: "元/KG",
    transit_time_min: o.tn || null, transit_time_max: o.tx2 || null, transit_time_desc: o.td || "",
    claim_rule: o.cr || "", effective_date: "", source_sheet: o.ss || "",
  };
}
function parseTransit(t) {
  const c = String(t || "").replace(/\r?\n/g, " ").trim();
  const m = c.match(/(\d+)\s*[-–~]*\s*(\d+)\s*[个]?(?:自然日|天)/);
  if (m) return { min: parseInt(m[1]), max: parseInt(m[2]), desc: c };
  return { min: null, max: null, desc: c };
}

// 通用解析：渠道行 + 国家/分区行 + 重量列
function parseChannelSheet(ws, sheetName, transportMode, countryCode, colConfig) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 5) return [];
  const results = [];
  let currentCh = "";

  for (let ri = 4; ri < data.length; ri++) {
    const row = data[ri];
    const c0 = String(row[0] || "").trim();
    const c1 = String(row[1] || "").trim();

    if (c1.match(/K\d+|C\d+|E\d+/) || c0.match(/K\d+|C\d+|E\d+/)) {
      currentCh = (c1.match(/[KC]E?\d+/) || c0.match(/[KC]E?\d+/))?.[0] || c1;
      continue;
    }
    if (!currentCh || !c1) continue;
    if (c1.includes("渠道名称") || c1.includes("分区") || c1.includes("重量")) continue;

    const channelLabel = c1;
    const destText = String(row[2] || "").trim();
    let destinations = pC(destText);
    if (destinations[0] === destText && !EU_C.some(cn => destText.includes(cn))) {
      // Could be warehouse or region
      destinations = [destText];
    }

    const dt = destinations.some(d => d.match(/FBA|亚马逊|仓库|仓/)) ? "warehouse" : "country";

    for (let ti = 0; ti < Math.min(colConfig.tiers.length, row.length - 3); ti++) {
      const p = parseFloat(row[3 + ti]);
      if (!isNaN(p) && p > 0 && ti < colConfig.tiers.length) {
        for (const dest of destinations) {
          results.push(mkr({
            c: countryCode, cn: `${currentCh}-${transportMode}`, tm: transportMode,
            vc: transportMode, vt: [transportMode], dm: colConfig.dm || "快递派",
            dc: dest, dt, dr: dest,
            bt: currentCh.includes("包税") || !currentCh.includes("不包税") ? "包税" : "不包税",
            tx: currentCh.includes("递延") ? "递延" : currentCh.includes("不包税") ? "不包税" : "包税",
            mq: colConfig.tiers[ti].q, mv: colConfig.tiers[ti].v, p, ss: sheetName,
          }));
        }
      }
    }
  }
  return results;
}

function parseXinsheng(filePath) {
  console.log("[新胜] 开始解析:", filePath);
  const wb = XLSX.readFile(filePath);
  const all = [];

  const configs = [
    { name: "英国海运自税、包税", tm: "海运", c: "英国", tiers: [{q:"21KG+",v:21},{q:"100KG+",v:100},{q:"500KG+",v:500},{q:"1000KG+",v:1000}] },
    { name: "英国卡航", tm: "卡航", c: "英国", tiers: [{q:"21KG+",v:21},{q:"51KG+",v:51},{q:"100KG+",v:100},{q:"1000KG+",v:1000}] },
    { name: "欧洲海运", tm: "海运", c: "欧洲", tiers: [{q:"50KG+",v:50},{q:"71KG+",v:71},{q:"100KG+",v:100},{q:"1000KG+",v:1000}] },
    { name: "欧洲卡航", tm: "卡航", c: "欧洲", tiers: [{q:"21KG+",v:21},{q:"100KG+",v:100},{q:"500KG+",v:500}] },
    { name: "欧洲空派-FBA（普货、带电）", tm: "空运", c: "欧洲", tiers: [{q:"21KG+",v:21},{q:"45KG+",v:45},{q:"100KG+",v:100},{q:"500KG+",v:500}] },
    { name: "英国空派限时达，经济线（普货、带电）", tm: "空运", c: "英国", tiers: [{q:"21KG+",v:21},{q:"51KG+",v:51},{q:"100KG+",v:100}] },
  ];

  for (const cfg of configs) {
    if (wb.SheetNames.includes(cfg.name)) {
      try {
        const r = parseChannelSheet(wb.Sheets[cfg.name], cfg.name, cfg.tm, cfg.c, { tiers: cfg.tiers, dm: "快递派" });
        console.log(`  [${cfg.name}] ${r.length} 条`);
        all.push(...r);
      } catch (err) { console.error(`  [${cfg.name}] 失败: ${err.message}`); }
    }
  }
  console.log(`[新胜] 总计 ${all.length} 条`);
  return all;
}
module.exports = { parseXinsheng };
