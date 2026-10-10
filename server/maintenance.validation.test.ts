import { describe, expect, it } from "vitest";
import {
  calculateReceptionPercent,
  areAllWorkOrdersClosed,
  canPurgeSapRescheduledWorkOrder,
  canReviewReception,
  filterActiveWorkshopEquipment,
  getPendingWarehouseLabel,
  getReservationValue,
  hasReservation,
  isEquipmentCycleFinalized,
  isWorkOrderClosed,
  matchesOtCycleFilter,
  isCorrectiveWorkOrder,
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

  it("allows repeated review of pending reception states", () => {
    expect(canReviewReception()).toBe(true);
    expect(canReviewReception("SIN_REVISAR")).toBe(true);
    expect(canReviewReception("RECIBIDA")).toBe(true);
    expect(canReviewReception("PARCIAL")).toBe(true);
    expect(canReviewReception("FALTA_RECIBIR")).toBe(true);
    expect(canReviewReception("REVISADO")).toBe(true);
  });

  it("does not treat pending reception states as closed", () => {
    expect(isWorkOrderClosed("PARCIAL")).toBe(false);
    expect(isWorkOrderClosed("FALTA_RECIBIR")).toBe(false);
    expect(isWorkOrderClosed("CERRADA_NOTIFICADA")).toBe(true);
  });

  it("separates corrective work orders from scheduled work orders", () => {
    expect(isCorrectiveWorkOrder('{"Tipo":"Correctiva"}')).toBe(true);
    expect(isCorrectiveWorkOrder('{"Tipo":"Programada"}')).toBe(false);
    expect(isCorrectiveWorkOrder(null)).toBe(false);
  });

  it("shows only equipment explicitly marked active in the workshop", () => {
    expect(filterActiveWorkshopEquipment([
      { code: "CA01", activo_en_taller: false },
      { code: "CA02", activo_en_taller: true },
      { code: "CA03" },
    ])).toEqual([{ code: "CA02", activo_en_taller: true }]);
  });

  it("purges a workshop equipment only when every related OT is closed", () => {
    expect(areAllWorkOrdersClosed(["CERRADA_NOTIFICADA", "CERRADA_NOTIFICADA"])).toBe(true);
    expect(areAllWorkOrdersClosed(["CERRADA_NOTIFICADA", "ABIERTA"])).toBe(false);
    expect(areAllWorkOrdersClosed([])).toBe(false);
  });

  it("classifies a cycle from every OT, not from a single OT", () => {
    expect(isEquipmentCycleFinalized([
      "CERRADA_NOTIFICADA",
      "CERRADA_NOTIFICADA",
      "CERRADA_NOTIFICADA",
      "CERRADA_NOTIFICADA",
    ])).toBe(true);
    expect(isEquipmentCycleFinalized([
      "CERRADA_NOTIFICADA",
      "CERRADA_NOTIFICADA",
      "CERRADA_NOTIFICADA",
      "PARCIAL",
    ])).toBe(false);
    expect(isEquipmentCycleFinalized(["CERRADA_NOTIFICADA", "REVISADO"])).toBe(false);
  });

  it("classifies individual OTs for the five operational filters", () => {
    expect(matchesOtCycleFilter({ hasReceptionState: false }, "SIN_REVISAR")).toBe(true);
    expect(matchesOtCycleFilter({ receptionStatus: "PARCIAL", hasReceptionState: true }, "PENDIENTES")).toBe(true);
    expect(matchesOtCycleFilter({ receptionStatus: "FALTA_RECIBIR", hasReceptionState: true }, "PENDIENTES")).toBe(true);
    expect(matchesOtCycleFilter({ receptionStatus: "REVISADO", hasReceptionState: true }, "REVISADAS")).toBe(true);
    expect(matchesOtCycleFilter({ receptionStatus: "REVISADO", executionStatus: "CERRADA_NOTIFICADA", hasReceptionState: true }, "REVISADAS")).toBe(false);
    expect(matchesOtCycleFilter({ executionStatus: "CERRADA_NOTIFICADA" }, "FINALIZADAS")).toBe(true);
    expect(matchesOtCycleFilter({ receptionStatus: "PARCIAL", executionStatus: "CERRADA_NOTIFICADA" }, "PENDIENTES")).toBe(false);
  });

  it("excludes closed OTs from a subsequent equipment re-entry", () => {
    const orders = [
      { ot: "OT1", status: "CERRADA_NOTIFICADA" },
      { ot: "OT2", status: "PARCIAL" },
      { ot: "OT3", status: undefined },
      { ot: "OT4", status: "CERRADA_NOTIFICADA" },
    ];
    expect(orders.filter(order => !isWorkOrderClosed(order.status)).map(order => order.ot)).toEqual(["OT2", "OT3"]);
    expect(isWorkOrderClosed("CERRADA_NOTIFICADA")).toBe(true);
    expect(isWorkOrderClosed("PARCIAL")).toBe(false);
  });

  it("purges only SAP OTs still sin revisar and keeps protected records", () => {
    expect(canPurgeSapRescheduledWorkOrder({ isInNewExcel: false, isCorrective: false, equipmentActive: false, statuses: ["SIN_REVISAR"] })).toBe(true);
    expect(canPurgeSapRescheduledWorkOrder({ isInNewExcel: true, isCorrective: false, equipmentActive: false, statuses: ["SIN_REVISAR"] })).toBe(false);
    expect(canPurgeSapRescheduledWorkOrder({ isInNewExcel: false, isCorrective: true, equipmentActive: false, statuses: ["SIN_REVISAR"] })).toBe(false);
    expect(canPurgeSapRescheduledWorkOrder({ isInNewExcel: false, isCorrective: false, equipmentActive: true, statuses: ["SIN_REVISAR"] })).toBe(false);
    expect(canPurgeSapRescheduledWorkOrder({ isInNewExcel: false, isCorrective: false, equipmentActive: false, statuses: ["RECIBIDA"] })).toBe(false);
    expect(canPurgeSapRescheduledWorkOrder({ isInNewExcel: false, isCorrective: false, equipmentActive: false, statuses: ["PARCIAL"] })).toBe(false);
  });
});
