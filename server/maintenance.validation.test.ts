import { describe, expect, it } from "vitest";
import {
  calculateReceptionPercent,
  areAllWorkOrdersClosed,
  canCloseWorkOrder,
  canPurgeSapRescheduledWorkOrder,
  canReviewReception,
  filterActiveWorkshopEquipment,
  getReceptionAction,
  getPendingWarehouseLabel,
  getReservationValue,
  hasReservation,
  isEquipmentCycleFinalized,
  isWorkOrderClosed,
  matchesOtCycleFilter,
  isCorrectiveWorkOrder,
  nextOptimisticVersion,
  selectLatestSharedReceptionState,
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
    expect(hasReservation('{"Reserva":"REPUESTO 123"}')).toBe(true);
    expect(hasReservation('{"Reserva":"-"}')).toBe(false);
    expect(hasReservation('{"Reserva":"—"}')).toBe(false);
    expect(hasReservation('{"Reserva":"   "}')).toBe(false);
    expect(hasReservation('{"Reserva":"  REPUESTO 123  "}')).toBe(true);
    expect(hasReservation('{"Reserva":" 0 "}')).toBe(false);
    expect(hasReservation(null)).toBe(false);
    expect(hasReservation(undefined)).toBe(false);
  });

  it("checks consolidated JSON reservation arrays element by element", () => {
    const withReservation = (value: unknown) => JSON.stringify({ Reserva: value });

    expect(hasReservation(withReservation('["0","—"]'))).toBe(false);
    expect(hasReservation(withReservation('["-","0"]'))).toBe(false);
    expect(hasReservation(withReservation('["0",""]'))).toBe(false);
    expect(hasReservation(withReservation("[]"))).toBe(false);
    expect(hasReservation(withReservation('["0","REPUESTO 123"]'))).toBe(true);
    expect(hasReservation(withReservation('["—","REPUESTO 123"]'))).toBe(true);
    expect(hasReservation(withReservation('["REPUESTO 123","REPUESTO 456"]'))).toBe(true);
  });

  it("preserves individual values and rejects nested non-scalar values", () => {
    const withReservation = (value: unknown) => JSON.stringify({ Reserva: value });

    expect(hasReservation(withReservation("TEXTO SIN FORMATO"))).toBe(true);
    expect(hasReservation(withReservation('["0",]'))).toBe(true);
    expect(hasReservation(withReservation(123))).toBe(true);
    expect(hasReservation(withReservation(true))).toBe(true);
    expect(hasReservation(withReservation(null))).toBe(false);
    expect(hasReservation(withReservation({ codigo: "REPUESTO 123" }))).toBe(false);
    expect(hasReservation(withReservation([["REPUESTO 123"]]))).toBe(false);
    expect(hasReservation(withReservation('[{"codigo":"REPUESTO 123"}]'))).toBe(false);
    expect(hasReservation(withReservation('[["REPUESTO 123"]]'))).toBe(false);

    expect(getReservationValue(withReservation('["0","—"]'))).toBe('["0","—"]');
  });

  it("allows repeated review of pending reception states", () => {
    expect(canReviewReception()).toBe(true);
    expect(canReviewReception("SIN_REVISAR")).toBe(true);
    expect(canReviewReception("RECIBIDA")).toBe(true);
    expect(canReviewReception("PARCIAL")).toBe(true);
    expect(canReviewReception("FALTA_RECIBIR")).toBe(true);
    expect(canReviewReception("REVISADO")).toBe(true);
  });

  it("shows only the correct reception action for each OT state", () => {
    expect(getReceptionAction("SIN_REVISAR")).toBe("REVIEW");
    expect(getReceptionAction("RECIBIDA")).toBe("REVIEW");
    expect(getReceptionAction("REVISADO")).toBe("COMPLETED");
    expect(getReceptionAction("PARCIAL")).toBe("RESOLVE_DIFFERENCES");
    expect(getReceptionAction("FALTA_RECIBIR")).toBe("RESOLVE_DIFFERENCES");
  });

  it("does not treat pending reception states as closed", () => {
    expect(isWorkOrderClosed("PARCIAL")).toBe(false);
    expect(isWorkOrderClosed("FALTA_RECIBIR")).toBe(false);
    expect(isWorkOrderClosed("CERRADA_NOTIFICADA")).toBe(true);
  });

  it("selects one shared reception state by recency regardless of guard", () => {
    const older = { id: 10, guard: "A", status: "PARCIAL", updatedAt: "2026-09-28T15:17:49Z" };
    const newer = { id: 11, guard: "D", status: "RECIBIDA", updatedAt: "2026-09-28T15:18:53Z" };
    expect(selectLatestSharedReceptionState([older, newer])).toEqual(newer);
    expect(selectLatestSharedReceptionState([
      { id: 20, guard: "A", updatedAt: "2026-09-28T15:18:53Z" },
      { id: 21, guard: "B", updatedAt: "2026-09-28T15:18:53Z" },
    ])?.id).toBe(21);
  });

  it("advances only a matching optimistic version", () => {
    expect(nextOptimisticVersion(1, 1)).toBe(2);
    expect(nextOptimisticVersion(8, 7)).toBeNull();
    expect(nextOptimisticVersion(0, 0)).toBeNull();
  });

  it("allows closing scheduled work orders without logistics and requires review when logistics apply", () => {
    expect(canCloseWorkOrder({ executionStatus: "ABIERTA", receptionStatus: "SIN_REVISAR", requiresLogistics: false })).toBe(true);
    expect(canCloseWorkOrder({ executionStatus: "ABIERTA", receptionStatus: "SIN_REVISAR", requiresLogistics: true })).toBe(false);
    expect(canCloseWorkOrder({ executionStatus: "ABIERTA", receptionStatus: "REVISADO", requiresLogistics: true })).toBe(true);
    expect(canCloseWorkOrder({ executionStatus: "CERRADA_NOTIFICADA", requiresLogistics: false })).toBe(false);
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
    expect(matchesOtCycleFilter({ receptionStatus: "RECIBIDA", hasReceptionState: true }, "REVISADAS")).toBe(true);
    expect(matchesOtCycleFilter({ receptionStatus: "REVISADO", hasReceptionState: true }, "REVISADAS")).toBe(true);
    expect(matchesOtCycleFilter({ receptionStatus: "REVISADO", executionStatus: "CERRADA_NOTIFICADA", hasReceptionState: true }, "REVISADAS")).toBe(false);
    expect(matchesOtCycleFilter({ executionStatus: "CERRADA_NOTIFICADA" }, "FINALIZADAS")).toBe(true);
    expect(matchesOtCycleFilter({ receptionStatus: "PARCIAL", executionStatus: "CERRADA_NOTIFICADA" }, "PENDIENTES")).toBe(false);
    expect(matchesOtCycleFilter({ isCorrective: true }, "SIN_REVISAR")).toBe(false);
    expect(matchesOtCycleFilter({ isCorrective: true }, "PENDIENTES")).toBe(false);
    expect(matchesOtCycleFilter({ isCorrective: true }, "TODAS")).toBe(true);
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
