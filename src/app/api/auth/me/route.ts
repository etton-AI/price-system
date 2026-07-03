/**
 * 获取当前用户信息
 * GET /api/auth/me
 * Header: Authorization: Bearer <token>
 * 返回: { user: { username, role } }
 */

import { NextRequest, NextResponse } from "next/server";
import { extractBearerToken, verifyToken } from "@/lib/auth";

export async function GET(request: NextRequest) {
  try {
    const token = extractBearerToken(request.headers.get("authorization"));
    if (!token) {
      return NextResponse.json(
        { success: false, error: "未提供认证令牌" },
        { status: 401 }
      );
    }

    const payload = await verifyToken(token);
    return NextResponse.json({
      success: true,
      user: {
        username: payload.sub,
        role: payload.role,
      },
    });
  } catch {
    return NextResponse.json(
      { success: false, error: "令牌无效或已过期" },
      { status: 401 }
    );
  }
}
