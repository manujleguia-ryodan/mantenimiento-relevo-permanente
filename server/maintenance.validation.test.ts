import { describe, expect, it } from "vitest";
import {
  calculateReceptionPercent,
  canReviewReception,
  getPendingWarehouseLabel,
  getReservationValue,
  hasReservation,
  validateReceptionInput,
} from "@shared/maintenance";

describe("maintenance business rules", () => {
  it("derives the combined warehouse label from two boolean flags", () => {
    expect(getPendingWarehouseLabel(false, false)).toBeNull();
    expect(getPendingWarehouseLabel(true, false)).toBe("ALMACEN_1");
    expect(getPendingWarehouseLabel(false, true)).toBe("ALMACEN_2");
    expect(getPendingWarehouseLabel(true, true)).toBe("AMBOS");
  });

  it("requires at least one warehouse for partial or missing reception", () => {
    expect(validateReceptionInput("RECIBIDA", false, false)).toBeNull();
    expect(validateReceptionInput("PARCIAL", true, false)).toBeNull();
    expect(validateReceptionInput("FALTA_RECIBIR", false, true)).toBeNull();
    expect(validateReceptionInput("PARCIAL", false, false)).toContain("Selecciona");
    expect(validateReceptionInput("FALTA_RECIBIR", false, false)).toContain("Selecciona");
  });

  it("rejects contradictory normalized warehouse payloads", () => {
    expect(validateReceptionInput("RECIBIDA", ["CENTRAL"])).toContain("Completa");
    expect(validateReceptionInput("RECIBIDA", [])).toBeNull();
    expect(validateReceptionInput("PARCIAL", [])).toContain("almacén");
    expect(validateReceptionInput("FALTA_RECIBIR", [])).toContain("almacén");
    expect(validateReceptionInput("PARCIAL", ["CENTRAL", "DONALDSON"])).toBeNull();
    expect(validateReceptionInput("FALTA_RECIBIR", ["DONALDSON"])).toBeNull();
  });

  it("calculates the complete-reception percentage safely", () => {
    expect(calculateReceptionPercent(0, 0)).toBe(0);
    expect(calculateReceptionPercent(35, 22)).toBe(63);
    expect(calculateReceptionPercent(10, 15)).toBe(100);
    expect(calculateReceptionPercent(10, -2)).toBe(0);
  });

  it("detects real reservation numbers and excludes empty logistics records", () => {
    expect(getReservationValue('{"RESERVA":"3157775"}')).toBe("3157775");
    expect(hasReservation('{"Reserva":"3157775"}')).toBe(true);
    expect(hasReservation('{"RESERVA":""}')).toBe(false);
    expect(hasReservation('{"RESERVA":"0"}')).toBe(false);
    expect(hasReservation(null)).toBe(false);
  });

  it("allows only the first reception review", () => {
    expect(canReviewReception()).toBe(true);
    expect(canReviewReception("SIN_REVISAR")).toBe(true);
    expect(canReviewReception("RECIBIDA")).toBe(false);
    expect(canReviewReception("PARCIAL")).toBe(false);
    expect(canReviewReception("FALTA_RECIBIR")).toBe(false);
  });
});
