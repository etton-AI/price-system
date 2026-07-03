# CLAUDE.md

## 项目简介
FBA 比价查询系统 — 基于 Next.js 的多供应商物流价格查询与对比平台。

## 技术栈
- 框架: Next.js 15.3 (App Router) + React 19 + TypeScript 5.8
- 样式: Tailwind CSS 4
- Excel 解析: xlsx (SheetJS)
- JWT: jose（兼容 Edge + Node.js 运行时）
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
src/app/api/auth/          # 登录 + 用户信息 API
src/lib/price-store.ts     # 数据加载模块
src/lib/auth.ts            # 权限管理模块（JWT/密码/日志）
src/middleware.ts           # Basic Auth 页面级防护
parsers/                   # 供应商 Excel 解析器 (build_db.js)
excels/                    # 供应商 Excel 报价表
data/ + public/data/       # 价格 JSON 数据 + 用户数据
```

## 权限管理规则

### 两层防护
1. **页面级 Basic Auth**: `src/middleware.ts` 拦截 `/price-query` 路由，校验 `BASIC_USER`/`BASIC_PASS` 环境变量。未配置时自动跳过（兼容本地开发）。
2. **角色级 JWT 登录**: 页面内登录获取 JWT Token，角色分为 `admin`（可上传报价表）和 `viewer`（仅查询）。

### 环境变量
| 变量 | 必填 | 说明 |
|------|------|------|
| `JWT_SECRET` | 生产必填 | JWT 签名密钥（32字节随机 hex） |
| `ADMIN_PASSWORD` | 否 | 默认 admin 账号初始密码（默认 `admin123`） |
| `BASIC_USER` | 否 | 页面级 Basic Auth 用户名 |
| `BASIC_PASS` | 否 | 页面级 Basic Auth 密码 |

### 用户管理
- 用户数据存储在 `data/users.json`（JSON 文件）
- 密码使用 Node.js `crypto.scryptSync` 哈希存储
- 首次启动自动创建默认 `admin` 账号
- 用户操作: 通过 `src/lib/auth.ts` 的 `createUser()`/`changePassword()` 管理

### 上传管控
- **鉴权**: 上传接口 `/api/price-query/upload` 仅接受 `admin` 角色的 JWT Token
- **格式**: 仅 `.xlsx`（前端 + 后端双重校验）
- **大小**: ≤15MB（`src/app/api/price-query/upload/route.ts` 中 `MAX_FILE_SIZE`）
- **日志**: 每次上传记录到 `data/upload-log.json`（保留最近 500 条）
- **瘦身模式**: 前端已预留复选框勾子（`slimMode` state），后续可实现 SheetJS 裁剪

### API 权限矩阵
| 接口 | 权限 |
|------|------|
| `GET /api/price-query` | 公开（无需登录） |
| `GET /api/price-query?meta=1` | 公开 |
| `POST /api/price-query/upload` | admin JWT |
| `POST /api/auth/login` | 公开 |
| `GET /api/auth/me` | JWT（任意角色） |

### 添加新用户
```javascript
// 在 Node.js 环境中执行
const { createUser } = require("./src/lib/auth"); // 注意: 需要构建后才能 import
createUser("username", "password", "viewer"); // 或 "admin"
```
