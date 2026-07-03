/**
 * 权限管理模块
 * - JWT 签发 / 验证
 * - 密码哈希（scrypt）
 * - 用户 CRUD（JSON 文件存储）
 * - Basic Auth 校验
 * - 操作日志记录
 */

import { SignJWT, jwtVerify } from "jose";
import crypto from "crypto";
import fs from "fs";
import path from "path";

// ── 类型 ──

export interface User {
  username: string;
  /** scrypt 哈希后的密码 (hex) */
  passwordHash: string;
  /** scrypt salt (hex) */
  salt: string;
  role: "admin" | "viewer";
  createdAt: string;
  updatedAt?: string;
}

export interface JwtPayload {
  sub: string;       // username
  role: "admin" | "viewer";
  iat: number;
  exp: number;
}

export interface UploadLogEntry {
  timestamp: string;
  username: string;
  fileName: string;
  fileSize: number;
  recordCount: number;
  supplier: string;
  result: "success" | "error";
  error?: string;
}

// ── 环境变量 ──

const JWT_SECRET_RAW = process.env.JWT_SECRET || "";
const BASIC_USER = process.env.BASIC_USER || "";
const BASIC_PASS = process.env.BASIC_PASS || "";

/** JWT Secret 编码为 Uint8Array（jose 要求） */
function getJwtSecret(): Uint8Array {
  if (!JWT_SECRET_RAW) {
    throw new Error("[auth] JWT_SECRET 环境变量未设置，无法签发/验证令牌");
  }
  return new TextEncoder().encode(JWT_SECRET_RAW);
}

// ── 用户存储路径 ──

const DATA_DIR = path.join(process.cwd(), "data");
const USERS_PATH = path.join(DATA_DIR, "users.json");
const UPLOAD_LOG_PATH = path.join(DATA_DIR, "upload-log.json");

/** 确保 data 目录存在 */
function ensureDataDir(): void {
  if (!fs.existsSync(DATA_DIR)) {
    fs.mkdirSync(DATA_DIR, { recursive: true });
  }
}

// ── 密码工具 ──

const SCRYPT_KEYLEN = 64;
const SCRYPT_OPTIONS = { N: 16384, r: 8, p: 1 };

/** 哈希密码，返回 { hash (hex), salt (hex) } */
export function hashPassword(password: string): { hash: string; salt: string } {
  const salt = crypto.randomBytes(32).toString("hex");
  const hash = crypto.scryptSync(password, salt, SCRYPT_KEYLEN, SCRYPT_OPTIONS).toString("hex");
  return { hash, salt };
}

/** 验证密码 */
export function verifyPassword(password: string, salt: string, storedHash: string): boolean {
  const hash = crypto.scryptSync(password, salt, SCRYPT_KEYLEN, SCRYPT_OPTIONS).toString("hex");
  return crypto.timingSafeEqual(Buffer.from(hash), Buffer.from(storedHash));
}

// ── 用户管理 ──

/** 读取所有用户 */
export function getUsers(): User[] {
  ensureDataDir();
  if (!fs.existsSync(USERS_PATH)) {
    // 初始化默认管理员
    const adminPassword = process.env.ADMIN_PASSWORD || "admin123";
    if (!process.env.ADMIN_PASSWORD) {
      console.warn("[auth] ⚠ ADMIN_PASSWORD 未设置，使用默认密码 admin123（仅开发环境安全）");
    }
    const { hash, salt } = hashPassword(adminPassword);
    const defaultUser: User = {
      username: "admin",
      passwordHash: hash,
      salt,
      role: "admin",
      createdAt: new Date().toISOString(),
    };
    fs.writeFileSync(USERS_PATH, JSON.stringify({ users: [defaultUser] }, null, 2), "utf-8");
    console.log("[auth] 已创建默认管理员账号: admin");
    return [defaultUser];
  }
  const raw = JSON.parse(fs.readFileSync(USERS_PATH, "utf-8"));
  return raw.users || [];
}

/** 根据用户名查找用户 */
export function findUser(username: string): User | undefined {
  return getUsers().find((u) => u.username === username);
}

/** 保存用户列表 */
function saveUsers(users: User[]): void {
  ensureDataDir();
  fs.writeFileSync(USERS_PATH, JSON.stringify({ users }, null, 2), "utf-8");
}

/** 创建新用户 */
export function createUser(username: string, password: string, role: "admin" | "viewer"): User {
  const users = getUsers();
  if (users.find((u) => u.username === username)) {
    throw new Error(`用户 "${username}" 已存在`);
  }
  const { hash, salt } = hashPassword(password);
  const user: User = {
    username,
    passwordHash: hash,
    salt,
    role,
    createdAt: new Date().toISOString(),
  };
  users.push(user);
  saveUsers(users);
  console.log(`[auth] 创建用户: ${username} (${role})`);
  return user;
}

/** 修改用户密码 */
export function changePassword(username: string, newPassword: string): void {
  const users = getUsers();
  const idx = users.findIndex((u) => u.username === username);
  if (idx === -1) throw new Error(`用户 "${username}" 不存在`);
  const { hash, salt } = hashPassword(newPassword);
  users[idx].passwordHash = hash;
  users[idx].salt = salt;
  users[idx].updatedAt = new Date().toISOString();
  saveUsers(users);
  console.log(`[auth] 用户 ${username} 密码已更新`);
}

// ── JWT ──

const JWT_EXPIRY = "24h";

/** 签发 JWT Token */
export async function signToken(username: string, role: "admin" | "viewer"): Promise<string> {
  const secret = getJwtSecret();
  const token = await new SignJWT({ sub: username, role })
    .setProtectedHeader({ alg: "HS256" })
    .setIssuedAt()
    .setExpirationTime(JWT_EXPIRY)
    .sign(secret);
  return token;
}

/** 验证 JWT Token，返回 payload */
export async function verifyToken(token: string): Promise<JwtPayload> {
  const secret = getJwtSecret();
  const { payload } = await jwtVerify(token, secret);
  return payload as unknown as JwtPayload;
}

// ── Basic Auth ──

/** 检查 Basic Auth 是否启用 */
export function isBasicAuthEnabled(): boolean {
  return Boolean(BASIC_USER && BASIC_PASS);
}

/** 校验 Basic Auth 凭据，返回 true 表示通过 */
export function checkBasicAuth(authHeader: string | null): boolean {
  if (!isBasicAuthEnabled()) {
    // 未配置 Basic Auth 环境变量 → 跳过校验
    return true;
  }
  if (!authHeader || !authHeader.startsWith("Basic ")) {
    return false;
  }
  const base64 = authHeader.slice(6);
  const decoded = Buffer.from(base64, "base64").toString("utf-8");
  const [user, pass] = decoded.split(":");
  return user === BASIC_USER && pass === BASIC_PASS;
}

/** 从 Authorization header 提取 Bearer token */
export function extractBearerToken(authHeader: string | null): string | null {
  if (!authHeader || !authHeader.startsWith("Bearer ")) return null;
  return authHeader.slice(7);
}

// ── 操作日志 ──

/** 记录上传操作 */
export function logUpload(entry: UploadLogEntry): void {
  ensureDataDir();
  let logs: UploadLogEntry[] = [];
  if (fs.existsSync(UPLOAD_LOG_PATH)) {
    try { logs = JSON.parse(fs.readFileSync(UPLOAD_LOG_PATH, "utf-8")); } catch { /* ignore */ }
  }
  // 保留最近 500 条
  logs.unshift(entry);
  if (logs.length > 500) logs = logs.slice(0, 500);
  fs.writeFileSync(UPLOAD_LOG_PATH, JSON.stringify(logs, null, 2), "utf-8");
  console.log(`[upload-log] ${entry.username} 上传 ${entry.fileName} → ${entry.result} (${entry.recordCount}条)`);
}

/** 读取操作日志 */
export function getUploadLogs(limit = 50): UploadLogEntry[] {
  if (!fs.existsSync(UPLOAD_LOG_PATH)) return [];
  try {
    const logs: UploadLogEntry[] = JSON.parse(fs.readFileSync(UPLOAD_LOG_PATH, "utf-8"));
    return logs.slice(0, limit);
  } catch { return []; }
}
