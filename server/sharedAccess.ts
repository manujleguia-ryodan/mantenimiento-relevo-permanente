import { createHash, createHmac, randomBytes, timingSafeEqual } from "node:crypto";
import { parse as parseCookie } from "cookie";
import type { Request } from "express";
import type { User } from "../drizzle/schema";

export const SHARED_ACCESS_COOKIE = "mantenimiento_shared_access";
export const SHARED_ACCESS_TTL_MS = 12 * 60 * 60 * 1000;
export const SHARED_PIN_MIN_LENGTH = 4;
export const MASTER_PIN = "9999";
export const MAX_FAILED_PIN_ATTEMPTS = 8;
export const PIN_ATTEMPT_WINDOW_MS = 15 * 60 * 1000;

const failedPinAttempts = new Map<string, { count: number; resetAt: number }>();

export function getConfiguredSharedPin(): string | null {
  const pin = process.env.SHARED_ACCESS_PIN || MASTER_PIN;
  return /^\d{4}$/.test(pin) ? pin : MASTER_PIN;
}

export function matchesSharedPin(
  candidate: string,
  configuredPin: string | null
) {
  if (!configuredPin || candidate.length !== configuredPin.length) return false;
  const candidateBytes = Buffer.from(candidate, "utf8");
  const configuredBytes = Buffer.from(configuredPin, "utf8");
  return (
    candidateBytes.length === configuredBytes.length &&
    timingSafeEqual(candidateBytes, configuredBytes)
  );
}

function signature(payload: string, secret: string) {
  return createHmac("sha256", secret).update(payload).digest("base64url");
}

export function createSharedAccessToken(
  secret: string,
  expiresAt = Date.now() + SHARED_ACCESS_TTL_MS,
  role: SharedAccessRole = "admin",
  guard: SharedAccessGuard = "A",
  name = "Acceso compartido"
) {
  const encodedName = Buffer.from(name, "utf8").toString("base64url");
  const payload = `${expiresAt}|${role}|${guard}|${encodedName}`;
  return `${payload}.${signature(payload, secret)}`;
}

export function verifySharedAccessToken(
  token: string | undefined,
  secret: string | undefined,
  now = Date.now()
) {
  if (!token || !secret) return false;
  const separator = token.lastIndexOf(".");
  if (separator <= 0) return false;
  const payload = token.slice(0, separator);
  const suppliedSignature = token.slice(separator + 1);
  const [expiresAtText] = payload.split("|");
  if (!/^\d+$/.test(expiresAtText) || !suppliedSignature) return false;
  const expiresAt = Number(expiresAtText);
  if (!Number.isSafeInteger(expiresAt) || expiresAt <= now) return false;

  const expectedSignature = signature(payload, secret);
  const suppliedBytes = Buffer.from(suppliedSignature, "base64url");
  const expectedBytes = Buffer.from(expectedSignature, "base64url");
  return (
    suppliedBytes.length === expectedBytes.length &&
    timingSafeEqual(suppliedBytes, expectedBytes)
  );
}

export function hasSharedAccess(req: Pick<Request, "headers">) {
  const rawCookieHeader = req.headers.cookie;
  const cookieHeader = Array.isArray(rawCookieHeader)
    ? rawCookieHeader.join(";")
    : (rawCookieHeader ?? "");
  const token = parseCookie(cookieHeader)[SHARED_ACCESS_COOKIE];
  return verifySharedAccessToken(token, process.env.JWT_SECRET);
}

export function getSharedAccessRole(
  req: Pick<Request, "headers">
): SharedAccessRole {
  const rawCookieHeader = req.headers.cookie;
  const cookieHeader = Array.isArray(rawCookieHeader)
    ? rawCookieHeader.join(";")
    : (rawCookieHeader ?? "");
  const token = parseCookie(cookieHeader)[SHARED_ACCESS_COOKIE];
  if (!token || !verifySharedAccessToken(token, process.env.JWT_SECRET)) return "admin";
  const payload = token.slice(0, token.lastIndexOf("."));
  const role = payload.split("|")[1];
  return role === "facilitator" || role === "supervisor" ? role : "admin";
}

export function getSharedAccessGuard(
  req: Pick<Request, "headers">
): SharedAccessGuard {
  const rawCookieHeader = req.headers.cookie;
  const cookieHeader = Array.isArray(rawCookieHeader)
    ? rawCookieHeader.join(";")
    : (rawCookieHeader ?? "");
  const token = parseCookie(cookieHeader)[SHARED_ACCESS_COOKIE];
  if (!token || !verifySharedAccessToken(token, process.env.JWT_SECRET)) return "A";
  const payload = token.slice(0, token.lastIndexOf("."));
  const guard = payload.split("|")[2];
  return guard === "B" || guard === "C" || guard === "D" ? guard : "A";
}

export function getSharedAccessName(req: Pick<Request, "headers">) {
  const rawCookieHeader = req.headers.cookie;
  const cookieHeader = Array.isArray(rawCookieHeader)
    ? rawCookieHeader.join(";")
    : (rawCookieHeader ?? "");
  const token = parseCookie(cookieHeader)[SHARED_ACCESS_COOKIE];
  if (!token || !verifySharedAccessToken(token, process.env.JWT_SECRET)) return "Acceso compartido";
  const payload = token.slice(0, token.lastIndexOf("."));
  try {
    const encodedName = payload.split("|")[3];
    if (!encodedName) return "Acceso compartido";
    return Buffer.from(encodedName, "base64url").toString("utf8") || "Acceso compartido";
  } catch {
    return "Acceso compartido";
  }
}

export type SharedAccessRole = "admin" | "facilitator" | "supervisor";
export type SharedAccessGuard = "A" | "B" | "C" | "D";

export function createSharedAccessUser(
  role: SharedAccessRole = "admin",
  now = new Date(),
  guard: SharedAccessGuard = "A",
  name = "Acceso compartido"
): User & { guard: SharedAccessGuard } {
  return {
    id: 0,
    openId: "shared-pin-access",
    name,
    email: null,
    loginMethod: "shared-pin",
    role,
    createdAt: now,
    updatedAt: now,
    lastSignedIn: now,
    guard,
  };
}

export function hashAccessCode(
  code: string,
  salt = randomBytes(16).toString("hex")
) {
  return {
    salt,
    hash: createHash("sha256").update(`${salt}:${code}`).digest("hex"),
  };
}

export function matchesAccessCode(
  code: string,
  salt: string,
  expectedHash: string
) {
  const actualBytes = Buffer.from(hashAccessCode(code, salt).hash, "hex");
  const expectedBytes = Buffer.from(expectedHash, "hex");
  return (
    actualBytes.length === expectedBytes.length &&
    timingSafeEqual(actualBytes, expectedBytes)
  );
}

export function canAttemptSharedPin(ip: string, now = Date.now()) {
  const bucket = failedPinAttempts.get(ip);
  if (!bucket || bucket.resetAt <= now) {
    failedPinAttempts.delete(ip);
    return true;
  }
  return bucket.count < MAX_FAILED_PIN_ATTEMPTS;
}

export function recordFailedSharedPinAttempt(ip: string, now = Date.now()) {
  const bucket = failedPinAttempts.get(ip);
  if (!bucket || bucket.resetAt <= now) {
    failedPinAttempts.set(ip, {
      count: 1,
      resetAt: now + PIN_ATTEMPT_WINDOW_MS,
    });
  } else {
    bucket.count += 1;
  }

  if (failedPinAttempts.size > 1000) {
    failedPinAttempts.forEach((value, key) => {
      if (value.resetAt <= now) failedPinAttempts.delete(key);
    });
  }
}

export function clearFailedSharedPinAttempts(ip: string) {
  failedPinAttempts.delete(ip);
}
