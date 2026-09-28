import { describe, expect, it } from "vitest";
import type { TrpcContext } from "./_core/context";
import { appRouter } from "./routers";
import {
  canAttemptSharedPin,
  createSharedAccessToken,
  getConfiguredSharedPin,
  MAX_FAILED_PIN_ATTEMPTS,
  PIN_ATTEMPT_WINDOW_MS,
  recordFailedSharedPinAttempt,
  SHARED_ACCESS_COOKIE,
  verifySharedAccessToken,
} from "./sharedAccess";

type CookieWrite = {
  name: string;
  value: string;
  options: Record<string, unknown>;
};

function createContext(cookieHeader = "", ip = `pin-test-${Math.random()}`) {
  const writtenCookies: CookieWrite[] = [];
  const req = {
    protocol: "https",
    headers: { cookie: cookieHeader },
    ip,
    socket: { remoteAddress: ip },
  } as unknown as TrpcContext["req"];
  const res = {
    cookie: (name: string, value: string, options: Record<string, unknown>) => {
      writtenCookies.push({ name, value, options });
    },
    clearCookie: () => undefined,
  } as unknown as TrpcContext["res"];
  return { ctx: { req, res, user: null } as TrpcContext, writtenCookies };
}

function createRoleContext(role: "supervisor" | "viewer") {
  const { ctx } = createContext();
  ctx.user = {
    id: 91,
    openId: `role-${role}`,
    name: `Usuario ${role}`,
    email: null,
    loginMethod: "test",
    role,
    createdAt: new Date(),
    updatedAt: new Date(),
    lastSignedIn: new Date(),
  };
  return ctx;
}

describe("shared PIN access", () => {
  it("uses a configured server-side PIN with at least eight characters", () => {
    const configuredPin = getConfiguredSharedPin();
    expect(configuredPin).not.toBeNull();
    expect(configuredPin?.length).toBeGreaterThanOrEqual(8);
  });

  it("signs session tokens and rejects expired or altered values", () => {
    const token = createSharedAccessToken("test-signing-secret", 2_000);
    expect(verifySharedAccessToken(token, "test-signing-secret", 1_999)).toBe(
      true
    );
    expect(verifySharedAccessToken(token, "test-signing-secret", 2_000)).toBe(
      false
    );
    expect(
      verifySharedAccessToken(
        `${token.slice(0, -1)}x`,
        "test-signing-secret",
        1_000
      )
    ).toBe(false);
    expect(verifySharedAccessToken(token, "different-secret", 1_000)).toBe(
      false
    );
  });

  it("limits repeated failed PIN attempts for one client", () => {
    const ip = `rate-limit-test-${Math.random()}`;
    const now = 10_000;
    for (let i = 0; i < MAX_FAILED_PIN_ATTEMPTS; i += 1) {
      expect(canAttemptSharedPin(ip, now)).toBe(true);
      recordFailedSharedPinAttempt(ip, now);
    }
    expect(canAttemptSharedPin(ip, now)).toBe(false);
    expect(canAttemptSharedPin(ip, now + PIN_ATTEMPT_WINDOW_MS)).toBe(true);
  });

  it("validates the configured secret through auth.pinLogin and protects maintenance routes", async () => {
    const configuredPin = getConfiguredSharedPin();
    expect(configuredPin).not.toBeNull();
    const login = createContext();
    const user = await appRouter
      .createCaller(login.ctx)
      .auth.pinLogin({ pin: configuredPin! });

    expect(user).toMatchObject({
      id: 0,
      name: "Acceso compartido",
      email: null,
      role: "facilitator",
    });
    expect(login.writtenCookies).toHaveLength(1);
    expect(login.writtenCookies[0]?.name).toBe(SHARED_ACCESS_COOKIE);
    expect(login.writtenCookies[0]?.options).toMatchObject({
      httpOnly: true,
      secure: true,
      sameSite: "lax",
      path: "/",
    });

    const token = login.writtenCookies[0]!.value;
    const authenticated = createContext(`${SHARED_ACCESS_COOKIE}=${token}`);
    await expect(
      appRouter.createCaller(authenticated.ctx).auth.me()
    ).resolves.toMatchObject({
      role: "facilitator",
      name: "Acceso compartido",
    });

    const unauthenticated = createContext();
    await expect(
      appRouter.createCaller(unauthenticated.ctx).maintenance.bootstrap()
    ).rejects.toMatchObject({
      code: "UNAUTHORIZED",
    });
  });

  it("does not issue a session for a wrong PIN", async () => {
    const configuredPin = getConfiguredSharedPin();
    expect(configuredPin).not.toBeNull();
    const wrongPin = `${configuredPin!.startsWith("x") ? "y" : "x"}${configuredPin!.slice(1)}`;
    const attempt = createContext();

    await expect(
      appRouter.createCaller(attempt.ctx).auth.pinLogin({ pin: wrongPin })
    ).rejects.toMatchObject({
      code: "UNAUTHORIZED",
    });
    expect(attempt.writtenCookies).toHaveLength(0);
  });

  it.each(["supervisor", "viewer"] as const)(
    "rejects restricted logistics mutations for %s with 403",
    async role => {
      const caller = appRouter.createCaller(createRoleContext(role));
      await expect(
        caller.maintenance.updateReception({
          workOrderId: 1,
          weekId: 1,
          guard: "A",
          status: "RECIBIDA",
          warehouseCodes: [],
        })
      ).rejects.toMatchObject({ code: "FORBIDDEN" });
      await expect(
        caller.maintenance.generateHandover({
          weekId: 1,
          fromGuard: "A",
          toGuard: "B",
        })
      ).rejects.toMatchObject({ code: "FORBIDDEN" });
      await expect(
        caller.maintenance.importWeek({
          label: "SEM 1",
          weekNumber: 1,
          year: 2026,
          guard: "A",
          rows: [{ otNumber: "1", equipmentCode: "CA01" }],
        })
      ).rejects.toMatchObject({ code: "FORBIDDEN" });
      if (role === "viewer") {
        await expect(
          caller.maintenance.markOperationalExecution({
            workOrderId: 1,
            weekId: 1,
            guard: "A",
            closed: true,
          })
        ).rejects.toMatchObject({ code: "FORBIDDEN" });
      }
    }
  );
});
