/**
 * 登录接口
 * POST /api/auth/login
 * Body: { username: string; password: string }
 * 返回: { token: string; user: { username, role } }
 */

import { NextRequest, NextResponse } from "next/server";
import { findUser, verifyPassword, signToken } from "@/lib/auth";

export async function POST(request: NextRequest) {
  try {
    const body = await request.json();
    const { username, password } = body || {};

    if (!username || !password) {
      return NextResponse.json(
        { success: false, error: "用户名和密码不能为空" },
        { status: 400 }
      );
    }

    const user = findUser(username);
    if (!user) {
      return NextResponse.json(
        { success: false, error: "用户名或密码错误" },
        { status: 401 }
      );
    }

    const valid = verifyPassword(password, user.salt, user.passwordHash);
    if (!valid) {
      return NextResponse.json(
        { success: false, error: "用户名或密码错误" },
        { status: 401 }
      );
    }

    const token = await signToken(user.username, user.role);

    console.log(`[auth] 用户 ${username} (${user.role}) 登录成功`);

    return NextResponse.json({
      success: true,
      token,
      user: {
        username: user.username,
        role: user.role,
      },
    });
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : "登录失败";
    console.error("[auth] 登录错误:", message);
    return NextResponse.json({ success: false, error: message }, { status: 500 });
  }
}
