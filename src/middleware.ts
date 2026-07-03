/**
 * Next.js 中间件 — 对 /price-query 页面 + API 实施 Basic Auth 页面级防护
 *
 * 环境变量：
 *   BASIC_USER / BASIC_PASS — 配置 Basic Auth 凭据（缺一则跳过保护）
 *
 * 受保护路由：
 *   /price-query, /price-query/:path*   — 前端页面
 *   /api/price-query, /api/price-query/:path*  — 查询 + 上传 API
 *
 * 不受保护：
 *   /api/auth/:path*  — 登录/用户信息接口需公开访问
 */

import { NextResponse } from "next/server";
import type { NextRequest } from "next/server";

export function middleware(request: NextRequest) {
  const basicUser = process.env.BASIC_USER;
  const basicPass = process.env.BASIC_PASS;

  // 未完整配置则跳过 Basic Auth（兼容本地开发）
  if (!basicUser || !basicPass) {
    return NextResponse.next();
  }

  const authHeader = request.headers.get("authorization");

  if (authHeader) {
    const basicPrefix = "Basic ";
    if (authHeader.startsWith(basicPrefix)) {
      const base64 = authHeader.slice(basicPrefix.length);
      const decoded = Buffer.from(base64, "base64").toString("utf-8");
      const [user, pass] = decoded.split(":");

      if (user === basicUser && pass === basicPass) {
        return NextResponse.next();
      }
    }
  }

  // 未认证 → 返回 401 + WWW-Authenticate
  return new NextResponse("Authentication Required", {
    status: 401,
    headers: {
      "WWW-Authenticate": 'Basic realm="FBA Price Query", charset="UTF-8"',
    },
  });
}

/** 仅拦截需要保护的路由 */
export const config = {
  matcher: [
    "/price-query",
    "/price-query/:path*",
    "/api/price-query",
    "/api/price-query/:path*",
  ],
};
