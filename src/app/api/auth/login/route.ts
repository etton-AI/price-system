/**
 * 登录接口（身份选择模式）
 * POST /api/auth/login
 * Body: { identity: "admin" | "guest", password: string }
 * 返回: { success, token, user } 或 { success:false, expired:true, error }
 */

import { NextRequest, NextResponse } from "next/server";
import { findUser, verifyPassword, signToken, isPasswordExpired, passwordRemainingDays } from "@/lib/auth";

export async function POST(request: NextRequest) {
  try {
    const body = await request.json();
    const { identity, password } = body || {};

    if (!identity || !password) {
      return NextResponse.json(
        { success: false, error: "请选择身份并输入密码" },
        { status: 400 }
      );
    }

    if (identity !== "admin" && identity !== "guest") {
      return NextResponse.json(
        { success: false, error: "无效的身份类型" },
        { status: 400 }
      );
    }

    const user = findUser(identity);
    if (!user) {
      return NextResponse.json(
        { success: false, error: `账号 "${identity}" 不存在` },
        { status: 401 }
      );
    }

    const valid = verifyPassword(password, user.salt, user.passwordHash);
    if (!valid) {
      return NextResponse.json(
        { success: false, error: "密码错误" },
        { status: 401 }
      );
    }

    // 密码过期检查
    if (isPasswordExpired(user)) {
      const days = Math.abs(passwordRemainingDays(user));
      console.log(`[auth] ${identity} 密码已过期 ${days} 天`);
      return NextResponse.json({
        success: false,
        expired: true,
        error: `密码已过期 ${days} 天，请联系管理员重置密码`,
      });
    }

    const token = await signToken(user.username, user.role);
    console.log(`[auth] ${identity} (${user.role}) 登录成功`);

    return NextResponse.json({
      success: true,
      token,
      user: { username: user.username, role: user.role },
    });
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : "登录失败";
    console.error("[auth] 登录错误:", message);
    return NextResponse.json({ success: false, error: message }, { status: 500 });
  }
}
