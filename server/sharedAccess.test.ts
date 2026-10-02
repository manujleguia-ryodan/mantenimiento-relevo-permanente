import { describe, expect, it } from "vitest";
import { appRouter } from "./routers";
import {
  canAttemptSharedPin,
  createSharedAccessToken,
  hashAccessCode,
  MAX_FAILED_PIN_ATTEMPTS,
  PIN_ATTEMPT_WINDOW_MS,
  recordFailedSharedPinAttempt,
  matchesAccessCode,
  verifySharedAccessToken,
} from "./sharedAccess";

describe("acceso operativo directo", () => {
  it("permite consultar el bootstrap sin sesión ni PIN", async () => {
    await expect(
      appRouter.createCaller({ user: null } as never).maintenance.bootstrap()
    ).resolves.toMatchObject({
      weeks: expect.any(Array),
      equipment: expect.any(Array),
    });
  });

  it("mantiene las utilidades criptográficas legadas sin usarlas en el flujo UI", () => {
    const token = createSharedAccessToken("test-signing-secret", 2_000);
    expect(verifySharedAccessToken(token, "test-signing-secret", 1_999)).toBe(true);
    expect(verifySharedAccessToken(token, "test-signing-secret", 2_000)).toBe(false);
    expect(verifySharedAccessToken(`${token.slice(0, -1)}x`, "test-signing-secret", 1_000)).toBe(false);

    const access = hashAccessCode("legacy-demo", "fixed-salt");
    expect(access.hash).not.toContain("legacy-demo");
    expect(matchesAccessCode("legacy-demo", access.salt, access.hash)).toBe(true);
    expect(matchesAccessCode("wrong-code", access.salt, access.hash)).toBe(false);
  });

  it("mantiene aislado el limitador legado de intentos, aunque la UI ya no lo usa", () => {
    const ip = `legacy-rate-limit-${Math.random()}`;
    const now = 10_000;
    for (let i = 0; i < MAX_FAILED_PIN_ATTEMPTS; i += 1) {
      expect(canAttemptSharedPin(ip, now)).toBe(true);
      recordFailedSharedPinAttempt(ip, now);
    }
    expect(canAttemptSharedPin(ip, now)).toBe(false);
    expect(canAttemptSharedPin(ip, now + PIN_ATTEMPT_WINDOW_MS)).toBe(true);
  });
});
