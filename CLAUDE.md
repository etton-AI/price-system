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
src/app/price-query/       # 比价查询页面（含登录面板）
src/app/api/price-query/   # 查询 + 上传 API
src/app/api/auth/          # 登录 + 用户信息 API
src/lib/price-store.ts     # 数据加载模块
src/lib/auth.ts            # 权限管理模块（JWT/密码/日志）
parsers/                   # 供应商 Excel 解析器 (build_db.js)
excels/                    # 供应商 Excel 报价表
data/ + public/data/       # 价格 JSON 数据 + 用户数据
```

## 权限管理规则

### 登录鉴权
- **页面级保护**: 未登录用户只能看到全屏登录页面，不可访问任何查询功能
- **JWT 登录**: 页面内登录获取 JWT Token，角色分为 `admin`（可上传报价表）和 `viewer`（仅查询）
- **无 Basic Auth**: 已改为页面内 JWT 登录面板，不再使用 middleware Basic Auth

### 默认账号
| 用户名 | 默认密码 | 角色 | 说明 |
|--------|----------|------|------|
| `admin` | `etton2026` | 管理员 | 可查询 + 上传更新价格 |
| `viewer` | `visit20260703` | 访客 | 仅可查询比价 |

密码通过环境变量配置：`ADMIN_PASSWORD` / `VIEWER_PASSWORD`（不设置则使用上方默认值）。

### 密码过期机制
- 密码自创建/修改之日起 **30 天**有效（可配 `PASSWORD_EXPIRY_DAYS`）
- 登录时检查 `passwordChangedAt`，超过 30 天返回 `expired: true`
- 前端弹出阻断弹窗："密码已过期，请联系管理员重置密码"
- 过期用户无法进入系统，必须由管理员通过 `changePassword()` 重置

### 环境变量
| 变量 | 必填 | 说明 |
|------|------|------|
| `JWT_SECRET` | 生产必填 | JWT 签名密钥（32字节随机 hex） |
| `ADMIN_PASSWORD` | 否 | admin 账号密码（默认 `etton2026`） |
| `VIEWER_PASSWORD` | 否 | viewer 账号密码（默认 `visit20260703`） |
| `PASSWORD_EXPIRY_DAYS` | 否 | 密码过期天数（默认 30） |

### 上传管控
- **鉴权**: 上传接口仅接受 `admin` 角色的 JWT Token
- **格式**: 仅 `.xlsx`（前端 + 后端双重校验）
- **大小**: ≤15MB
- **日志**: 记录到 `data/upload-log.json`（保留最近 500 条）
- **瘦身模式**: 前端已预留复选框勾子（`slimMode` state）

### API 权限矩阵
| 接口 | 权限 |
|------|------|
| `GET /api/price-query` | JWT（任意角色） |
| `GET /api/price-query?meta=1` | JWT（任意角色） |
| `POST /api/price-query/upload` | admin JWT |
| `POST /api/auth/login` | 公开 |
| `GET /api/auth/me` | JWT（任意角色） |

### 添加新用户
```javascript
const { createUser } = require("./src/lib/auth");
createUser("username", "password", "viewer"); // 或 "admin"
```

### 重置密码
```javascript
const { changePassword } = require("./src/lib/auth");
changePassword("admin", "new-password");
```
