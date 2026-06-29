# CLAUDE.md

## 项目简介
FBA 比价查询系统 — 基于 Next.js 的多供应商物流价格查询与对比平台。

## 技术栈
- 框架: Next.js 15.3 (App Router) + React 19 + TypeScript 5.8
- 样式: Tailwind CSS 4
- Excel 解析: xlsx (SheetJS)
- 部署: Docker (standalone 模式)

## 快速开始
```powershell
npm install
npm run build-db   # 解析 Excel 生成 prices.json
npm run dev        # 启动开发服务器
```
浏览器访问 http://localhost:3000/price-query

## 命令
| npm run dev   | 启动开发服务器 |
| npm run build | 构建（含 build-db） |
| npm run build-db | 仅重新生成价格数据 |

## 目录结构
```
src/app/price-query/       # 比价查询页面
src/app/api/price-query/   # 查询 + 上传 API
src/lib/price-store.ts     # 数据加载模块
parsers/                   # 供应商 Excel 解析器 (build_db.js)
excels/                    # 供应商 Excel 报价表
data/ + public/data/       # 价格 JSON 数据
```
