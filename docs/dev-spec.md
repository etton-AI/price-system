# Price System — 开发者规格文档

> 最后更新: 2026-07-01 | 维护者: berry-bi

---

## 1. 项目概述与技术栈

**FBA 比价查询系统** — 从 15 家供应商的 Excel 报价表中自动解析物流价格，统一为结构化 JSON 数据，提供 Web 端多维度比价查询界面。

### 技术栈

| 类别 | 技术 | 版本 |
|------|------|------|
| 框架 | Next.js 15 (App Router) | `^15.3.3` |
| UI | React | `^19.1.0` |
| 语言 | TypeScript (Web) + JavaScript (Parsers) | `^5.8.3` |
| 样式 | Tailwind CSS 4 (PostCSS 插件) | `^4.1.8` |
| Excel 解析 | xlsx (SheetJS) | `^0.18.5` |
| Lint | ESLint 9 flat config | `^9.27.0` |
| 运行时 | Node.js 22 (Alpine) | — |
| 数据格式 | 单文件 JSON (~34MB, 43000+ 条记录) | — |
| 部署 | Docker → GHCR → Sealos K8s | — |

### 仓库信息

- **GitHub**: `etton-ai/price-system`
- **容器镜像**: `ghcr.io/etton-ai/price-system:latest`
- **Sealos Ingress**: `wlylcsujbziw.cloud.sealos.io`

---

## 2. 目录结构与各模块职责

```
price-system/
├── .github/workflows/
│   └── docker-build.yml              # CI: push main → 构建 Docker 镜像 → 推送 GHCR
├── k8s/
│   └── deploy.yaml                   # K8s Deployment + Service + Ingress
├── excels/                           # 原始供应商 Excel 报价表 (18 个文件)
│   ├── ETTON 易通科技物流价格表*.xlsx (2 个)
│   ├── 天图通逊英美加同行VIP价*.xlsx (1 个)
│   ├── 英美跨境-美线/空派/加拿大公布价*.xlsx (3 个)
│   ├── 皓辉国际供应链VIP价格*.xlsx (1 个)
│   ├── 皓鹏国际同行VIP*.xlsx (1 个)
│   ├── 星链专线报价表*.xlsx (1 个)
│   ├── 心一欧洲同行VIP报价表*.xlsx (1 个)
│   ├── 航乐报价表*.xlsx (1 个)
│   ├── 丰运同行VIP报价*.xlsx (1 个)
│   ├── 华威尔同行协议价*.xlsx (1 个)
│   ├── 凯鑫科技VIP报价表*.xls (1 个)
│   ├── 新胜供应链报价表*.xls (2 个)
│   └── 美琦同行VIP报价表*.xlsx (1 个)
├── parsers/                          # 解析器 + 构建工具 (全部 CommonJS)
│   ├── build_db.js                   # ★ 构建入口：扫描 excels/ → 调用解析器 → 输出 prices.json
│   ├── query.js                      # CLI 查价工具
│   ├── country-detector.js           # 共享国家检测器（Sheet名/文件名 → 国家）
│   ├── etton_us.js                   # 易通ETTON (美/加/澳)
│   ├── fengyun.js                    # 丰运跨境 (欧/英)
│   ├── hangle.js                     # 航乐国际 (英/欧)
│   ├── haohui_us.js                  # 皓辉国际 (美)
│   ├── haopeng_us.js                 # 皓鹏国际 (美/英/欧/加/墨/巴西/澳/DG/TEMU)
│   ├── huaweier.js                   # 华威尔 (美/欧)
│   ├── kaixin.js                     # 凯鑫科技 (欧/英/加)
│   ├── meiqi_us.js                   # 美琦国际 (美/加/墨)
│   ├── tiantu_air.js                 # 天图通逊空运 (美/英)
│   ├── tiantu_uk.js                  # 天图通逊英国
│   ├── tiantu_us.js                  # 天图通逊 (美/加，自动委托英国 Sheet 给 tiantu_uk)
│   ├── xinglian_us.js                # 星链专线 (美)
│   ├── xinsheng.js                   # 新胜供应链 (英/欧)
│   ├── xinyi_eu.js                   # 心一供应链 (欧)
│   └── yingmei_us.js                 # 英美跨境 (美/加)
├── data/
│   └── prices.json                   # 构建输出 (~34MB)
├── public/
│   └── data/                         # Web 运行时数据目录
│       └── prices.json               # 同 data/prices.json (build_db 自动复制)
├── src/
│   ├── app/
│   │   ├── layout.tsx                # 根布局: html lang=zh-CN
│   │   ├── page.tsx                  # 首页 → redirect("/price-query")
│   │   ├── globals.css               # Tailwind CSS 4 入口
│   │   ├── price-query/
│   │   │   └── page.tsx              # ★ 比价查询主页面 (817 行客户端组件)
│   │   └── api/
│   │       └── price-query/
│   │           └── route.ts          # ★ GET 查询 API (305 行)
│   └── lib/
│       └── price-store.ts            # 数据加载模块（文件缓存）
├── Dockerfile                        # 双阶段构建（含 build_db 步骤）
├── next.config.ts                    # output: "standalone"
├── postcss.config.mjs
├── eslint.config.mjs
├── tsconfig.json
├── package.json
└── CLAUDE.md                         # Claude Code 项目指引
```

---

## 3. 数据流架构

```
┌─────────────────────┐
│  excels/*.xlsx(xls) │  原始供应商 Excel 报价表 (18 个文件)
└────────┬────────────┘
         │ 文件名关键词匹配识别供应商
         ▼
┌─────────────────────┐
│  parsers/build_db.js │  扫描 excels/ → 调用对应解析器
│  + 15 个供应商解析器  │  每个解析器: XLSX.readFile → 逐 Sheet 解析 → 返回记录数组
└────────┬────────────┘
         │ 去重 (supplier|channel|dest|region|billing|minQty)
         │ 提取生效日期 (文件名正则)
         │ 设置 source_file 字段
         ▼
┌─────────────────────┐
│  data/prices.json    │  统一 JSON (~34MB, ~43000 条)
│  public/data/        │  (自动复制)
└────────┬────────────┘
         │
    ┌────┴────┐
    ▼         ▼
┌────────┐ ┌──────────┐
│ Web UI │ │ CLI Tool │
│ /price │ │ query.js │
│ -query │ │          │
└───┬────┘ └────┬─────┘
    │           │
    ▼           ▼
 price-store.ts  直接 fs.readFile
 (内存缓存)      (本地路径)
```

---

## 4. 已实现的模块

### 4.1 数据模型 (PriceEntry)

**定义位置**: `src/lib/price-store.ts`

```typescript
interface PriceEntry {
  supplier: string;            // 供应商名称（如 "易通ETTON"）
  country?: string;            // 目的国家（美国/英国/欧线/加拿大/墨西哥/巴西/澳大利亚）
  channel_name: string;        // 渠道名称（如 "美森MATSON CLX(正班)-卡派包税"）
  vessel_config: string;       // 船司配置（如 "美森正班CLX"）
  vessel_tags: string[];       // 搜索标签（如 ["matson","clx"]）
  transport_mode: string;      // 运输方式（海运/空运/卡航/铁路/快递）
  delivery_method: string;     // 送仓方式（卡派/海派/快递派/整柜直送/自提）
  destination_code: string;    // 目的仓代码（如 "ONT8"）或国家名
  destination_type: string;    // "warehouse" | "country" | "none"
  destination_region: string;  // 目的地区域（如 "美西"）
  origin_region: string;       // 起运地区域（如 "华南"）
  origin_cities: string[];     // 起运地城市列表
  billing_type: string;        // 税模式（包税/递延/自税/不包税）
  tax_mode: string;            // 税务模式（同 billing_type）
  min_quantity: string;        // 最小起运量标签（如 "21KG+"）
  min_quantity_value: number;  // 最小起运量数值（如 21）
  unit_price: number;          // 单价
  price_unit: string;          // 价格单位（"元/KG" | "元/CBM"）
  cbm_price?: number;          // 按方单价（部分渠道）
  transit_time_min: number | null;  // 最短时效(天)
  transit_time_max: number | null;  // 最长时效(天)
  transit_time_desc: string;        // 时效文字描述
  claim_rule: string;          // 赔偿规则
  effective_date: string;      // 生效日期 (YYYY-MM-DD)
  source_file: string;         // 来源文件名
  source_sheet?: string;       // 来源 Sheet 名
}
```

---

### 4.2 build_db.js — 构建入口

**文件**: `parsers/build_db.js` (229 行)

#### 流程
1. 扫描 `../excels/` 下所有 `.xlsx` / `.xls` 文件（过滤 `~$` 临时文件）
2. `identifySupplier(fileName)` — 基于文件名关键词匹配供应商
3. 调用对应解析器的导出函数
4. 从文件名提取生效日期（正则匹配 `YYYY年M月D日` / `M月D日` / `M.D`）
5. 设置每条记录的 `effective_date` 和 `source_file`
6. 去重: 唯一键 `supplier|channel_name|destination_code|origin_region|billing_type|min_quantity`
7. 输出压缩 JSON 到 `data/prices.json` + `public/data/prices.json`

#### 供应商识别规则

| 文件名关键词 | 供应商 ID | 解析器模块 |
|-------------|-----------|-----------|
| `etton` / `易通` | `etton` | `./etton_us` |
| `天图` + `英国` (不含美) | `tiantu_uk` | `./tiantu_uk` |
| `天图` + `空运`/`air` | `tiantu_air` | `./tiantu_air` |
| `天图` / `tiantu` | `tiantu` | `./tiantu_us` |
| `英美` / `yingmei` | `yingmei` | `./yingmei_us` |
| `皓辉` / `haohui` | `haohui` | `./haohui_us` |
| `皓鹏` / `haopeng` | `haopeng` | `./haopeng_us` |
| `星链` / `xinglian` | `xinglian` | `./xinglian_us` |
| `心一` / `xinyi` | `xinyi` | `./xinyi_eu` |
| `航乐` / `hangle` | `hangle` | `./hangle` |
| `丰运` / `fengyun` | `fengyun` | `./fengyun` |
| `华威尔` / `huaweier` | `huaweier` | `./huaweier` |
| `凯鑫` / `kaixin` | `kaixin` | `./kaixin` |
| `新胜` / `xinsheng` | `xinsheng` | `./xinsheng` |
| `美琦` / `meiqi` | `meiqi` | `./meiqi_us` |

---

### 4.3 country-detector.js — 共享国家检测器

**文件**: `parsers/country-detector.js` (71 行)

#### 导出函数

| 函数 | 说明 |
|------|------|
| `detectCountry(sheetName)` | 从 Sheet 名检测国家，返回国家名或 `null` |
| `detectCountryFromFileName(fileName)` | 从文件名检测国家（兜底） |
| `detectAllCountries(workbook)` | 扫描所有 Sheet，返回去重国家列表 |

#### 国家关键词模式

| 国家 | 关键词 |
|------|--------|
| 美国 | 美国, 美线, 美西, 美中, 美东, USA, 美森, 洛杉矶, 休斯顿, 芝加哥, 萨凡纳, 纽约 |
| 英国 | 英国, 英线, UK, 伦敦, 曼城, Felixstowe, Southampton |
| 欧线 | 欧洲, 欧线, EU, 德国, 法国, 意大利, 西班牙, 荷兰, 比利时, 波兰, 捷克, 匈牙利, 罗马尼亚 |
| 加拿大 | 加拿大, 加线, Canada, 温哥华, 多伦多, 蒙特利尔 |
| 墨西哥 | 墨西哥, 墨线, Mexico |
| 巴西 | 巴西, Brazil |
| 澳大利亚 | 澳大利亚, 澳洲, 澳线, Australia, 悉尼, 墨尔本, 布里斯班 |
| 日本 | 日本, 日线, Japan, 东京, 大阪 |

---

### 4.4 供应商解析器详情

#### 通用解析模式

所有解析器遵循统一模式：
1. `XLSX.readFile(filePath)` 读取 Excel
2. 遍历 `wb.SheetNames`，跳过目录/说明等非价格 Sheet
3. 基于 Sheet 名或文件名 `detectCountry()` 确定国家
4. `XLSX.utils.sheet_to_json(ws, {header:1, defval:""})` 转为二维数组
5. 定位表头行 → 识别渠道列组 → 解析数据行 → 调用 `makeRecord()` / `mkr()` 生成记录
6. 返回 `PriceEntry[]` 数组

#### makeRecord / mkr 模式

每个解析器内部定义 `makeRecord(opts)` 或 `mkr(opts)` 函数，用于构造标准 PriceEntry：

```javascript
function mkr(o) {
  return {
    supplier: SUPPLIER,
    country: o.c || DEFAULT_COUNTRY,
    channel_name: o.cn || "",
    transport_mode: o.tm || "海运",
    vessel_config: o.vc || "",
    vessel_tags: o.vt || [],
    delivery_method: o.dm || "卡派",
    destination_type: o.dt || "warehouse",
    destination_code: o.dc || "",
    destination_region: o.dr || "",
    origin_region: o.origin || "华南",
    origin_cities: o.cities || [...],
    billing_type: o.bt || "包税",
    tax_mode: o.tx || o.bt || "包税",
    min_quantity: o.mq || "",           // ★ 注意: 参数是 opts.minQty
    min_quantity_value: o.mv || 0,      // ★ 注意: 参数是 opts.minQtyValue
    unit_price: o.p || 0,
    price_unit: "元/KG",
    transit_time_min: o.tn || null,
    transit_time_max: o.tx2 || null,
    transit_time_desc: o.td || "",
    claim_rule: o.cr || "",
    effective_date: "",
    source_sheet: o.ss || "",
  };
}
```

#### ⚠️ 关键 Gotcha: `min_quantity` 参数名

`makeRecord` 内部 key 是 `min_quantity`，但参数名是 `opts.minQty`（不是 `opts.min_quantity`）。如果写错为 `min_quantity: value`，该值会丢失，导致去重 key 碰撞 → **大量记录被错误去重**。

---

#### 4.4.1 易通ETTON — `etton_us.js` (912 行)

**覆盖国家**: 美国、加拿大、澳大利亚

**美国 Sheet**: 12 个区域专线（美西12-15日达、17-20日达、22-25日达、FBA渠道、整柜直送、美西北/芝加哥/休斯顿/纽约/萨凡纳专线、DG海卡）

**渠道-船司映射规则**:
| 渠道 | 船司 |
|------|------|
| 12日达/15日达 | 美森正班CLX |
| 17日达 | EXX/以星带托架/合德快提统配 |
| 20日达 | 以星不带托架/合德普提 |
| 22日达 | COSCO/OA |
| 25日达 | OA/MSC/ONE/WHL/HMM/TSL普船统配 |

**数据结构**: 4 固定区域（东莞/中山/广州 | 嘉兴/义乌 | 汕头/厦门/泉州 | 武汉/长沙），仓库代码在 A 列（"/" 分隔需拆分）

**加/澳解析**: 自动检测 Sheet 名，启发式列检测（R1-R2 渠道名，R3 城市组，R4-R5 重量梯度），通用解析

---

#### 4.4.2 天图通逊 — `tiantu_us.js` (598 行) + `tiantu_uk.js` (260 行) + `tiantu_air.js` (213 行)

**覆盖国家**: 美国、加拿大、英国

**美国 Sheet**: 美西直送/美森/以星&普船、美东直航/海铁、美东南直航/海陆、美中休斯顿/芝加哥、美西北

**加拿大 Sheet**: 直航-普船、普船-特惠、商业/私人地址×2、小包系列

**英国 Sheet** (tiantu_uk.js): 铁运专线、中英专车、卡航20日达、海运海派/卡派、商私地址

**空运 Sheet** (tiantu_air.js): 美国空卡、美国空运&美西普货、英国空运

**数据结构**: 仓库行 × 11 城市列（华南/重庆/汕头/厦泉福/华东/青岛/天津/济南/西安/武汉），左侧包税(50KG+) / 右侧不包税(0.5CBM+)

**委托机制**: `tiantu_us.js` 自动导入 `./tiantu_uk`，当检测到英国 Sheet 名时委托解析

---

#### 4.4.3 英美跨境 — `yingmei_us.js` (601 行)

**覆盖国家**: 美国、加拿大

**渠道编码体系**:
- **A 系列 (海派)**: A1(美森CLX), A2(美森MAX), A4_18(合德快提/ZIM带车架), A4_21(合德/ZIM), A5(普船)
- **B 系列 (卡派)**: B1(美森CLX), B2(美森MAX), B4(合德/ZIM), B5(普船LA), B7(普船NY), B8(海外仓自提, 含子渠道), B10(普船OAK), B12(COSCO/OOCL), B14/B15(WHL/HMM/海领)

**加拿大 Sheet**: 直航快船、直航统配特惠、合德美转加、COSCO美转加、美森美转加、商私地址卡派

**区域划分**: 3 区域（义乌/上海/宁波/杭州/温州 | 东莞/宝安/中山/广州/南城/汕头/深圳 | 福州/厦门/泉州/合肥/青岛/温州/汕头）

---

#### 4.4.4 皓辉国际 — `haohui_us.js` (487 行)

**覆盖国家**: 美国

**渠道-船司映射**:
| 渠道 | 船司 |
|------|------|
| 皓森达 | 美森正班CLX |
| 皓速达 | 以星ZIM-ZEX |
| 皓速达带托 | 以星ZIM-ZEX带车架 |
| 皓快达 | COSCO/EMC/OOCL普船 |
| 皓东达 | COSCO/OOCL/ZIM美东直航 |
| 纽约皓速达 | EXX快船IPI海铁联运 |

---

#### 4.4.5 皓鹏国际 — `haopeng_us.js` (1196 行) ★ 最大解析器

**覆盖国家**: 美国、英国、欧洲、加拿大、巴西、澳大利亚、墨西哥

**美国 Sheet**: 美森以星合德OA非OA海卡、洛杉矶海卡特惠、海派、休斯顿/芝加哥/萨凡纳/纽约海卡、商私卡、空运

**其他 Sheet**: 英国超大件/空运海运铁路卡航、欧洲超大件/空运海运铁路卡航、巴西海运、澳大利亚空运海运、墨西哥空派美转墨直航、加拿大空运海运/超大件、欧英美加海运空运DG、TEMU-Y2专线

**已知状态**:
- 墨西哥空派美转墨直航: 所有价格显示"渠道暂停" → 返回 0 条
- 欧洲空运海运-超大件: 所有价格显示"单询" → 不解析（正确行为）

---

#### 4.4.6 美琦国际 — `meiqi_us.js` (814 行)

**覆盖国家**: 美国、加拿大、墨西哥

**美国 Sheet**: Match系列(Match12/Match15)、OA拆送/直送、美森美东秒送、MT专线、OA奥克兰、海派、商业私人地址、纽约/休斯敦/萨瓦纳/芝加哥/TH特惠

**加拿大 Sheet**: 加拿大、美转加

**美转墨**: 限时达/快线/普线/敏货，需自动检测列偏移（部分 Sheet col0 为空，数据从 col1 开始）

---

#### 4.4.7 星链专线 — `xinglian_us.js` (189 行)

**覆盖国家**: 美国

**渠道**: 星链直送-普船、星链锁仓-普船

**数据结构**: 区域行（美西/美中南/美中北/美东南/美东北），每行含仓库列表（中文逗号分隔），单价格 + 时效

---

#### 4.4.8 心一供应链 — `xinyi_eu.js` (486 行)

**覆盖国家**: 欧洲

**16 个 Sheet**: 空运普货/带电（DPD自税/含税、快线八日提/十日提）、海运（含税/快速达/限时达/比雷/直送/45日达）、铁卡、中欧专车

**数据结构**: 国家行（非仓库），子表检测（含税/自税），可配置重量梯度

---

#### 4.4.9 航乐国际 — `hangle.js` (272 行)

**覆盖国家**: 英国、欧洲

**Sheet**: 英国卡航海运、欧洲卡航/海运/铁路（快递派+专仓卡派）

**数据结构**: 仓库代码（"+" 分隔），CBM 行（10CBM+），21/101/501KG 梯度

---

#### 4.4.10 丰运跨境 — `fengyun.js` (187 行)

**覆盖国家**: 欧洲、英国

**Sheet**: 欧洲海运/空运/卡航/铁路一口价、英国海运/铁路/卡航/空运一口价

**数据结构**: 城市列（深圳/广州/义乌），国家行，包税/递延双模式

---

#### 4.4.11 华威尔 — `huaweier.js` (201 行)

**覆盖国家**: 美国、欧洲

**Sheet**: 美国空派大陆飞、美国海运快递派、美西FBA海卡、欧洲空派（普货/带电/超大件）

**数据结构**: 城市列（华南/华东），国家匹配

---

#### 4.4.12 凯鑫科技 — `kaixin.js` (370 行)

**覆盖国家**: 欧洲、英国、加拿大、美国

**解析策略**: 自动检测所有非跳过 Sheet
- `parseUKSheet()` — 英国专用（多种渠道+税模式混合）
- `parseGroupedSheet()` — 通用分组解析（自动检测表头行、渠道组、重量梯度）

**通用解析能力**: 扫描行 1-6 寻找重量梯度模式（`/\d+\s*KG\+/i`），自动识别渠道组边界，推断渠道名和税模式（包税/递延/自税）

---

#### 4.4.13 新胜供应链 — `xinsheng.js` (323 行)

**覆盖国家**: 英国、欧洲

**解析策略**:
- `parseGenericSheet()` — 扫描 "渠道名称" 关键字，识别渠道编码（K01, E03A 等），自动检测重量梯度
- `parseSimpleSheet()` — 兜底（超大件、小包等简单格式）

---

### 4.5 查询 API

**文件**: `src/app/api/price-query/route.ts` (305 行)

#### GET `/api/price-query`

**查询参数**:

| 参数 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `country` | string | `"美国"` | 国家过滤 |
| `transport_mode` | string | — | 运输方式（海运/空运/卡航/铁路） |
| `dest` | string | — | 目的仓代码（如 ONT8） |
| `origin` | string | — | 发货城市（如 深圳） |
| `weight` | number | — | 重量(KG)，匹配最近重量梯度 |
| `vessel` | string | — | 船司/渠道关键词 |
| `method` | string | — | 送仓方式（卡派/海派/整柜直送/自提） |
| `supplier` | string | — | 供应商（逗号分隔多选） |
| `top` | number | — | 返回前 N 条 |
| `best` | string | — | `"1"` 或 `"true"` 仅返回最优价 |
| `meta` | string | — | `"1"` 返回供应商元数据 |

**查询流程**:
1. 加载数据（`price-store.ts` 文件缓存）
2. 国家过滤
3. 运输方式过滤
4. 目的仓精确匹配
5. 发货城市 → 供应商区域展开（ETTON/天图/英美有城市→区域映射表）
6. 重量梯度匹配（找 `min_quantity_value >= weight` 的最大梯度，按 supplier+channel+dest+region 去重）
7. 船司关键词匹配（vessel_config 或 channel_name）
8. 送仓方式过滤
9. 供应商过滤（逗号分隔，模糊匹配）
10. 按 unit_price 升序 → transit_time_min 升序排序
11. Top N / Best

**返回格式**:
```json
{
  "success": true,
  "query": { "country": "美国", ... },
  "results": [ /* PriceEntry[] */ ],
  "total": 42,
  "best": { /* 最低价 PriceEntry */ },
  "stats": { "total": 43033, "generated_at": "2026-07-01T..." }
}
```

#### ⚠️ 城市→区域映射表

仅 **ETTON、天图、英美** 三家有完整的城市→起运区域映射表（`CITY_TO_ORIGIN`），其他供应商直接使用城市名进行包含匹配。新增供应商时需评估是否需要添加映射。

---

### 4.6 查询页面 (Web UI)

**文件**: `src/app/price-query/page.tsx` (817 行)

#### 功能模块
1. **线路选择器**: 8 条线路 Tab（美国/英国/欧线/加拿大/墨西哥/巴西/澳大利亚/日本）
2. **运输方式过滤**: 全部/海运/空运/卡航+专车/铁路
3. **多供应商过滤器**: 下拉多选框，每个供应商有独立颜色徽章
4. **查询表单**: 目的仓、发货城市、重量(KG)、渠道关键词、送仓方式
5. **结果数量控制**: 10/30/50/100/全部
6. **快捷预设**: 每条线路有预设查询组合（如 美国→ONT8→深圳→100KG→卡派）
7. **结果表格**: 供应商徽章、渠道名、运输方式、船司配置、单价（最低价绿色高亮）、时效、赔偿规则、起运区域、生效日期
8. **最优推荐横幅**: 黄色渐变卡片，展示最低价方案
9. **导出**: CSV 下载 + JSON 展示/复制
10. **Excel 上传区**: 拖拽上传供应商 Excel 文件（⚠️ 上传 API 未实现，见已知坑）
11. **供应商元数据**: 页面加载时自动获取

---

### 4.7 数据加载模块

**文件**: `src/lib/price-store.ts` (66 行)

- `getData()` — 读取 `public/data/prices.json`，内存缓存，后续调用直接返回缓存
- `refreshCache()` — 清除缓存（上传新数据后调用）
- `getDataPath()` — 返回数据文件路径

---

### 4.8 CLI 查价工具

**文件**: `parsers/query.js` (435 行)

```bash
# 基本查价
node parsers/query.js -d ONT8 -o 深圳 -w 100

# 指定船司 + 送仓方式
node parsers/query.js -d ONT8 -o 深圳 -w 100 -v EXX -m 卡派

# 显示最低价
node parsers/query.js -d ONT8 --best

# 限制结果数 + 导出
node parsers/query.js -d ONT8 -t 10 --export csv
```

**参数**:
| 参数 | 说明 |
|------|------|
| `-d, --dest` | 目的仓代码 |
| `-o, --origin` | 发货城市 |
| `-w, --weight` | 重量(KG) |
| `-v, --vessel` | 船司关键词 |
| `-m, --method` | 送仓方式 |
| `-s, --supplier` | 供应商 |
| `-t, --top N` | 限制输出条数 |
| `--best` | 仅最优价 |
| `--export csv\|json` | 导出格式 |

---

## 5. 非目标（明确没做的）

- ❌ **上传 API (`POST /api/price-query/upload`)**: 页面引用了上传端点，但 `route.ts` 中未实现
- ❌ **用户认证/权限**: 公开查询接口
- ❌ **增量更新**: 每次 `build_db` 全量重建 34MB JSON
- ❌ **数据库**: 无 SQL/NoSQL，纯文件 JSON
- ❌ **版本化价格历史**: 不保留历史报价，每次都覆盖
- ❌ **价格单位转换**: 仅支持"元/KG"和"元/CBM"，不支持其他货币/单位
- ❌ **移动端优化**: Web 页面设计针对桌面端
- ❌ **自动化测试**: 无单元测试/集成测试
- ❌ **日志/监控**: 仅 console.log
- ❌ **解析器配置化**: 每个供应商硬编码解析函数，无配置文件驱动的通用解析
- ❌ **星链澳洲线**: `xinglian_us.js` 中 "澳洲FBA-海运" Sheet 标记为跳过

---

## 6. 已知坑 / 绕过的 Hack / 待重构项

### 已知坑

1. **`makeRecord` 的 `minQty` 参数名陷阱** ★★★
   - 解析器中 `makeRecord({ minQty: t.label })` → 内部映射为 `min_quantity`
   - 如果误写为 `min_quantity: t.label`，参数直接进入 undefined path，默认为 `""` → 所有记录 min_quantity 相同 → 去重碰撞 → 大量记录丢失
   - **历史事故**: 皓鹏新增巴西/澳洲/加拿大/DG/TEMU 解析器时全部写错，导致 287 条记录被错误去重，后通过 `sed` 修复 13 处

2. **CommonJS vs ES Module 混杂**
   - 解析器全部使用 `require/module.exports`（CommonJS）
   - Next.js App Router 使用 ES Module (`import/export`)
   - **结果**: 解析器只能通过 `node` 直接运行，不能从 Next.js 代码中 import
   - `build_db.js` 在 `npm run build` 时独立执行（`node parsers/build_db.js`），不在 Next.js 构建流程中

3. **DG Sheet 国家覆盖 Bug**
   - Sheet "欧英美加海运空运DG" 的 `detectCountry` 返回 "美国"（因为含"美"字）
   - 导致 DG 中加拿大/英国/欧洲的记录 `country` 被错误覆盖为 "美国"
   - **影响范围**: ~8 条 DG 加拿大记录 → 当前未修复

4. **`detectCountry` 优先级问题**
   - 关键词列表中"美森"触发"美国"匹配，因为含"美"字
   - 如果 Sheet 名同时含美国和加拿大关键词，`detectCountry` 返回第一个匹配（美国优先）
   - 部分解析器（如 haopeng、meiqi）在循环中手动覆盖 `country` 字段来绕过

5. **港口/仓库名称歧义**
   - "奥克兰" 同时是美国和新西兰城市 → `country-detector.js` 未包含此关键词
   - "温哥华" 在美国和加拿大都有 → 已正确处理为加拿大

6. **Excel 临时文件干扰**
   - `~$` 开头的 Office 临时文件会被 `readdirSync` 捕获
   - `build_db.js` 已过滤 `!f.startsWith("~$")`，但如果新解析器自己实现文件扫描需注意

7. **34MB JSON 全量加载**
   - Web 服务每次启动/缓存失效时加载 34MB JSON → ~500ms 解析时间
   - Node.js 内存占用增加 ~100MB
   - **缓解**: `price-store.ts` 内存缓存 + 启动时一次性加载

8. **xlsx 库的日期单元格**
   - SheetJS 将 Excel 日期转为数字（自 1900-01-01 的天数），需要手动转换
   - 部分解析器需处理日期格式的生效日期列

9. **`$` 临时 Excel 文件**
   - excels/ 目录中可能有 `~$` 开头的 Office 锁定文件
   - `build_db.js` 第 60 行正确过滤了这些文件

### 待重构项

- [ ] **实现 `POST /api/price-query/upload`**: 页面已有 UI，API 路由未实现
- [ ] **将 34MB JSON 拆分为按国家分片**: 减少单次加载量，支持按需加载
- [ ] **修复 DG Sheet 国家覆盖**: 在 haopeng 解析器的 DG 循环中使用 per-sheet country override
- [ ] **统一解析器模块格式**: 全部迁移到 ES Module 或全部保持 CommonJS（当前混用）
- [ ] **提取通用解析器**: `parseGroupedSheet`/`parseGenericSheet` 等模式出现在多个解析器中，应提取为共享模块
- [ ] **添加解析器测试**: 至少对 `makeRecord`/`mkr` 输出格式做快照测试
- [ ] **供应商元数据动态化**: 当前 supplier color/name 映射硬编码在页面中
- [ ] **支持 `detectCountry` 多国家返回**: 当前一个 Sheet 只能返回一个国家，DG 类 Sheet 需要支持多国家
- [ ] **日本线路数据**: `country-detector.js` 已定义日本关键词，但无日本价格数据

---

## 7. 运行 & 测试命令

### 开发环境

```bash
# 安装依赖
npm install

# 仅重新构建数据库（解析 Excel → prices.json）
npm run build-db

# 启动开发服务器 (默认 http://localhost:3000)
npm run dev

# 启动并暴露给局域网
npx next dev -H 0.0.0.0 -p 3001
```

### CLI 查价

```bash
# 在项目根目录
node parsers/query.js -d ONT8 -o 深圳 -w 100
node parsers/query.js -d LGB8 -o 义乌 -v 美森 -m 卡派 --best
node parsers/query.js -d ONT8 -o 深圳 -w 500 -t 20 --export csv > result.csv
```

### 生产构建

```bash
# 完整构建（含 build-db + next build）
npm run build

# 运行生产版本
npm run start
```

### 代码检查

```bash
npm run lint
```

### Docker 构建

```bash
docker build -t price-system .
docker run -p 3000:3000 price-system
```

### 手动验证新解析器

```bash
# 在 Node.js REPL 中
node -e "
const { parseETTON } = require('./parsers/etton_us');
const results = parseETTON('./excels/ETTON报价表.xlsx');
console.log('Total:', results.length);
console.log('Sample:', JSON.stringify(results[0], null, 2));
"
```

---

## 8. 对外 API / DB Schema

### 8.1 查询 API

**`GET /api/price-query`**

详见 4.5 节。支持 `country`, `transport_mode`, `dest`, `origin`, `weight`, `vessel`, `method`, `supplier`, `top`, `best`, `meta` 参数。

### 8.2 数据文件

**`public/data/prices.json`** — 唯一数据源

```json
{
  "generated_at": "2026-07-01T06:00:00.000Z",
  "total_records": 43033,
  "stats": {
    "etton": 3284,
    "tiantu": 10234,
    "tiantu_uk": 1234,
    "yingmei": 3210,
    "haohui": 2100,
    "haopeng": 6373,
    "xinglian": 320,
    "xinyi": 2800,
    "hangle": 1100,
    "fengyun": 987,
    "huaweier": 890,
    "tiantu_air": 876,
    "kaixin": 286,
    "xinsheng": 266,
    "meiqi": 5432
  },
  "data": [
    {
      "supplier": "易通ETTON",
      "country": "美国",
      "channel_name": "易·12日达卡派",
      "transport_mode": "海运",
      "vessel_config": "美森正班CLX",
      "vessel_tags": ["matson", "clx"],
      "delivery_method": "卡派",
      "destination_code": "ONT8",
      "destination_type": "warehouse",
      "destination_region": "美西",
      "origin_region": "东莞/中山/广州",
      "origin_cities": ["深圳", "东莞", "广州", "中山"],
      "billing_type": "包税",
      "tax_mode": "包税",
      "min_quantity": "21KG+",
      "min_quantity_value": 21,
      "unit_price": 8.5,
      "price_unit": "元/KG",
      "transit_time_min": 12,
      "transit_time_max": 15,
      "transit_time_desc": "12-15工作日",
      "claim_rule": "",
      "effective_date": "2026-06-30",
      "source_file": "2026年6月30日 ETTON 易通科技物流价格表.xlsx",
      "source_sheet": "美西12-15日达"
    }
  ]
}
```

### 8.3 无外部 API 依赖

本项目不调用任何外部 API。所有数据来自本地 Excel 解析。

### 8.4 CI/CD 流程

```
Git push main
  → GitHub Actions: docker-build.yml
    → docker/build-push-action@v5
      → ghcr.io/etton-ai/price-system:latest + :sha
        → Sealos 控制台手动重新部署
```

---

## 附录 A: 供应商覆盖矩阵

| 供应商 | 美国 | 英国 | 欧洲 | 加拿大 | 墨西哥 | 巴西 | 澳大利亚 |
|--------|:----:|:----:|:----:|:------:|:------:|:----:|:--------:|
| 易通ETTON | ✅ | — | — | ✅ | — | — | ✅ |
| 天图通逊 | ✅ | ✅ | — | ✅ | — | — | — |
| 英美跨境 | ✅ | — | — | ✅ | — | — | — |
| 皓辉国际 | ✅ | — | — | — | — | — | — |
| 皓鹏国际 | ✅ | ✅ | ✅ | ✅ | ⚠️暂停 | ✅ | ✅ |
| 星链专线 | ✅ | — | — | — | — | — | — |
| 心一供应链 | — | — | ✅ | — | — | — | — |
| 航乐国际 | — | ✅ | ✅ | — | — | — | — |
| 丰运跨境 | — | ✅ | ✅ | — | — | — | — |
| 华威尔 | ✅ | — | ✅ | — | — | — | — |
| 凯鑫科技 | ✅ | ✅ | ✅ | ✅ | — | — | — |
| 新胜供应链 | — | ✅ | ✅ | — | — | — | — |
| 美琦国际 | ✅ | — | — | ✅ | ✅ | — | — |

## 附录 B: 数据记录数 (2026-07-01)

| 国家 | 记录数 |
|------|--------|
| 🇺🇸 美国 | 30,039 |
| 🇨🇦 加拿大 | 7,877 |
| 🇪🇺 欧洲/欧线 | 3,305 |
| 🇬🇧 英国 | 1,364 |
| 🇦🇺 澳大利亚 | 420 |
| 🇲🇽 墨西哥 | 24 |
| 🇧🇷 巴西 | 4 |
| **总计** | **43,033** |
