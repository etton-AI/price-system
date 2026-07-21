/**
 * 权限管理模块
 * - JWT 签发 / 验证
 * - 密码哈希（scrypt）
 * - 用户 CRUD（JSON 文件存储）
 * - 密码过期检查（30天）
 * - 操作日志记录
 */

import { SignJWT, jwtVerify } from "jose";
import crypto from "crypto";
import fs from "fs";
import path from "path";
import { execSync } from "child_process";

// ── 类型 ──

export interface User {
  username: string;
  passwordHash: string;
  salt: string;
  role: "admin" | "guest";
  createdAt: string;
  passwordChangedAt: string;
  updatedAt?: string;
}

export interface JwtPayload {
  sub: string;
  role: "admin" | "guest";
  iat: number;
  exp: number;
}

export interface LoginResult {
  success: boolean;
  token?: string;
  user?: { username: string; role: string };
  expired?: boolean;
  error?: string;
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

// ── 运行时环境变量读取 ──
// Next.js standalone webpack 会替换 process.env 为构建时的 shim，
// 导致运行时 K8s 环境变量无法被读取。
// 用 execSync 直接调用 OS printenv 命令，从系统层面获取环境变量，
// 彻底绕过所有 Node.js / webpack 层面的拦截。

function readEnv(key: string, fallback = ""): string {
  try {
    const val = execSync(`printenv ${key}`, {
      encoding: "utf8",
      timeout: 1000,
    }).trim();
    if (val) return val;
  } catch { /* fallback */ }
  return fallback;
}

function getAdminPassword(): string {
  return readEnv("ADMIN_PASSWORD", "etton2026");
}

function getGuestPassword(): string {
  return readEnv("GUEST_PASSWORD", "visit20260703");
}

// ── 常量 ──

export const PASSWORD_EXPIRY_DAYS = 30;

// ── JWT Secret ──

function getJwtSecret(): Uint8Array {
  const secret = readEnv("JWT_SECRET");
  if (!secret) {
    throw new Error("[auth] JWT_SECRET 环境变量未设置");
  }
  return new TextEncoder().encode(secret);
}

// ── 用户存储路径 ──

const DATA_DIR = path.join(process.cwd(), "data");
const USERS_PATH = path.join(DATA_DIR, "users.json");
const UPLOAD_LOG_PATH = path.join(DATA_DIR, "upload-log.json");

function ensureDataDir(): void {
  if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
}

// ── 密码工具 ──

const SCRYPT_KEYLEN = 64;
const SCRYPT_OPTIONS = { N: 16384, r: 8, p: 1 };

export function hashPassword(password: string): { hash: string; salt: string } {
  const salt = crypto.randomBytes(32).toString("hex");
  const hash = crypto.scryptSync(password, salt, SCRYPT_KEYLEN, SCRYPT_OPTIONS).toString("hex");
  return { hash, salt };
}

export function verifyPassword(password: string, salt: string, storedHash: string): boolean {
  const hash = crypto.scryptSync(password, salt, SCRYPT_KEYLEN, SCRYPT_OPTIONS).toString("hex");
  return crypto.timingSafeEqual(Buffer.from(hash), Buffer.from(storedHash));
}

// ── 密码过期检查 ──

export function isPasswordExpired(user: User): boolean {
  const changedAt = user.passwordChangedAt || user.createdAt;
  const diffDays = (Date.now() - new Date(changedAt).getTime()) / (1000 * 60 * 60 * 24);
  return diffDays > PASSWORD_EXPIRY_DAYS;
}

export function passwordRemainingDays(user: User): number {
  const changedAt = user.passwordChangedAt || user.createdAt;
  const diffDays = (Date.now() - new Date(changedAt).getTime()) / (1000 * 60 * 60 * 24);
  return Math.ceil(PASSWORD_EXPIRY_DAYS - diffDays);
}

// ── 用户管理 ──

function makeUser(username: string, password: string, role: "admin" | "guest"): User {
  const { hash, salt } = hashPassword(password);
  const now = new Date().toISOString();
  return { username, passwordHash: hash, salt, role, createdAt: now, passwordChangedAt: now };
}

export function getUsers(): User[] {
  ensureDataDir();
  if (!fs.existsSync(USERS_PATH)) {
    if (!readEnv("ADMIN_PASSWORD")) console.warn("[auth] ADMIN_PASSWORD 未设置，使用默认密码 etton2026");
    if (!readEnv("GUEST_PASSWORD")) console.warn("[auth] GUEST_PASSWORD 未设置，使用默认密码 visit20260703");
    const defaultUsers: User[] = [
      makeUser("admin", getAdminPassword(), "admin"),
      makeUser("guest", getGuestPassword(), "guest"),
    ];
    fs.writeFileSync(USERS_PATH, JSON.stringify({ users: defaultUsers }, null, 2), "utf-8");
    console.log("[auth] 已创建默认账号: admin (管理员) + guest (访客)");
    return defaultUsers;
  }
  const raw = JSON.parse(fs.readFileSync(USERS_PATH, "utf-8"));
  const users = raw.users || [];
  // 兼容旧数据：补充 passwordChangedAt 字段
  let patched = false;
  for (const u of users) {
    if (!u.passwordChangedAt) {
      u.passwordChangedAt = u.updatedAt || u.createdAt;
      patched = true;
    }
  }
  if (patched) saveUsers(users);
  return users;
}

export function findUser(username: string): User | undefined {
  return getUsers().find((u) => u.username === username);
}

function saveUsers(users: User[]): void {
  ensureDataDir();
  fs.writeFileSync(USERS_PATH, JSON.stringify({ users }, null, 2), "utf-8");
}

export function createUser(username: string, password: string, role: "admin" | "guest"): User {
  const users = getUsers();
  if (users.find((u) => u.username === username)) throw new Error(`用户 "${username}" 已存在`);
  const user = makeUser(username, password, role);
  users.push(user);
  saveUsers(users);
  console.log(`[auth] 创建用户: ${username} (${role})`);
  return user;
}

export function changePassword(username: string, newPassword: string): void {
  const users = getUsers();
  const idx = users.findIndex((u) => u.username === username);
  if (idx === -1) throw new Error(`用户 "${username}" 不存在`);
  const { hash, salt } = hashPassword(newPassword);
  users[idx].passwordHash = hash;
  users[idx].salt = salt;
  users[idx].passwordChangedAt = new Date().toISOString();
  users[idx].updatedAt = new Date().toISOString();
  saveUsers(users);
  console.log(`[auth] 用户 ${username} 密码已更新`);
}

// ── JWT ──

const JWT_EXPIRY = "24h";

export async function signToken(username: string, role: "admin" | "guest"): Promise<string> {
  const secret = getJwtSecret();
  return await new SignJWT({ sub: username, role })
    .setProtectedHeader({ alg: "HS256" })
    .setIssuedAt()
    .setExpirationTime(JWT_EXPIRY)
    .sign(secret);
}

export async function verifyToken(token: string): Promise<JwtPayload> {
  const secret = getJwtSecret();
  const { payload } = await jwtVerify(token, secret);
  return payload as unknown as JwtPayload;
}

// ── Bearer Token 提取 ──

export function extractBearerToken(authHeader: string | null): string | null {
  if (!authHeader || !authHeader.startsWith("Bearer ")) return null;
  return authHeader.slice(7);
}

// ── 操作日志 ──

export function logUpload(entry: UploadLogEntry): void {
  ensureDataDir();
  let logs: UploadLogEntry[] = [];
  if (fs.existsSync(UPLOAD_LOG_PATH)) {
    try { logs = JSON.parse(fs.readFileSync(UPLOAD_LOG_PATH, "utf-8")); } catch { /* */ }
  }
  logs.unshift(entry);
  if (logs.length > 500) logs = logs.slice(0, 500);
  fs.writeFileSync(UPLOAD_LOG_PATH, JSON.stringify(logs, null, 2), "utf-8");
  console.log(`[upload-log] ${entry.username} 上传 ${entry.fileName} → ${entry.result} (${entry.recordCount}条)`);
}

export function getUploadLogs(limit = 50): UploadLogEntry[] {
  if (!fs.existsSync(UPLOAD_LOG_PATH)) return [];
  try {
    const logs: UploadLogEntry[] = JSON.parse(fs.readFileSync(UPLOAD_LOG_PATH, "utf-8"));
    return logs.slice(0, limit);
  } catch { return []; }
}
